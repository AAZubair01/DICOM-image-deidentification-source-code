#!/usr/bin/env python3
"""
dicom-cxr-deidentify - de-identification & cleaning pipeline for a DICOM
chest-X-ray dataset matched against a spreadsheet of radiology reports.

Run this ON THE MACHINE THAT HOLDS THE DICOMs (it never touches the network).

WORKFLOW
--------
STEP 1 (always run this first):
    python src/deidentify_dataset.py --config config.yaml --audit
    Only inspects folders, tag usage and matching quality; writes
    audit_report.txt into the output folder. Nothing is modified.
    Review it before the full run.

STEP 2:
    python src/deidentify_dataset.py --config config.yaml

STEP 3 (after EVERY full run):
    python src/verify_deidentification.py --config config.yaml
    See the companion script for its options (--check-dates, --ignore).

WHAT THE FULL RUN DOES
----------------------
1. Reads EVERY sheet of the metadata workbook (one row per patient),
   auto-detecting name/code/date/conclusion columns and skipping duplicated
   header rows.
2. Matches each DICOM to a metadata row:
     primary  - numeric CODE extracted from the filename / PatientID /
                PatientName / AccessionNumber, exact-matched on the CODE
                column of the spreadsheet
     fallback - fuzzy name matching (token-sorted, with exam-date bonus)
   Weak or ambiguous matches go to a manual review queue - never guessed.
3. De-identifies headers following the DICOM PS3.15 basic profile, at EVERY
   nesting level (tags hidden inside sequences are scrubbed too):
     - identifying tags + ALL private tags deleted
     - dates optionally shifted by a consistent per-patient random offset
       (--shift-dates or dates.shift in config; default keeps real dates,
       which is basic-profile compliant)
     - every UID deterministically remapped (sha256 -> '2.25.' OID)
     - PatientName / PatientID replaced by a pseudo-ID (<PREFIX>_0001 ...)
   PIXEL DATA IS NEVER MODIFIED.
4. Routes each image:
     <output.base_dir>/<clean_subdir>/     verified, shareable chest images
     <output.base_dir>/<qa_subdir>/        YOU review these before sharing
   Images are quarantined (not deleted) when they are: non-chest anatomy,
   non-image DICOM objects (e.g. modality PR 'Presentation State'), flagged
   BurnedInAnnotation=YES, or CD-export artefacts.
5. Writes:
     <output.base_dir>/metadata_for_colleague.csv   one row per clean image
     <output.base_dir>/qa_report.csv                why each file was held
     <output.mapping_dir>/real_to_pseudo_mapping.csv   NEVER share this

Requirements:  pip install -r requirements.txt
"""

import os
import re
import csv
import sys
import argparse
import hashlib
import random
from collections import Counter
from datetime import datetime, timedelta
from difflib import SequenceMatcher

try:
    import yaml
    import pydicom
    from pydicom.datadict import keyword_for_tag
    from openpyxl import load_workbook
except ImportError as exc:
    sys.exit("Missing dependency: %s\n"
             "Install with:  pip install -r requirements.txt" % exc)


# ============================== config ======================================
def load_config(path):
    if not os.path.isfile(path):
        sys.exit("Config file not found: %s\n"
                 "Copy config.example.yaml to config.yaml and fill in your "
                 "paths." % path)
    with open(path, encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh)
    if not isinstance(cfg, dict):
        sys.exit("Config file is empty or invalid: %s" % path)
    return cfg


# Filled from config in main() / run_audit(): keep module-level names so the
# rest of the code reads exactly like the documented pipeline.
METADATA_XLSX = None
SRC_FOLDERS   = []          # [(folder, label, prefix), ...]
OUT_DIR       = None
CLEAN_DIR     = None
QA_DIR        = None
MAP_DIR       = None
MIN_NAME_SCORE = 0.85
AMBIGUOUS_GAP  = 0.05
RANDOM_SEED    = 42
SHIFT_DATES    = False
SHIFT_MIN_DAYS = 60
SHIFT_MAX_DAYS = 700
SKIP_FILENAMES = {"DICOMDIR", "LABEL.TXT"}
# ============================================================================


# ----------------------------- helpers --------------------------------------
def norm_name(s):
    s = str(s).upper()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return " ".join(s.split())


def extract_codes(s):
    """Digit runs of length >= 4 - hospital codes, embedded dates, etc."""
    return set(re.findall(r"\d{4,}", str(s)))


def name_score(a, b):
    a, b = norm_name(a), norm_name(b)
    if not a or not b:
        return 0.0
    straight = SequenceMatcher(None, a, b).ratio()
    sorted_tok = SequenceMatcher(None, " ".join(sorted(a.split())),
                                 " ".join(sorted(b.split()))).ratio()
    return max(straight, sorted_tok)


def to_yyyymmdd(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 8:
        if 1900 <= int(digits[:4]) <= 2100:
            return digits
        for fmt in ("%d%m%Y", "%m%d%Y"):
            try:
                return datetime.strptime(digits, fmt).strftime("%Y%m%d")
            except ValueError:
                pass
    return None


def uid_remap(uid, salt):
    """Deterministic DICOM-valid UID remap (max 64 chars, numeric)."""
    if not uid:
        return uid
    h = hashlib.sha256((str(uid) + "|" + salt).encode()).hexdigest()
    return "2.25." + str(int(h[:32], 16))


# ------------------------------ Excel ---------------------------------------
def load_excel_rows(path):
    """Read EVERY sheet; auto-detect columns; skip duplicated header rows."""
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [norm_name(h) if h is not None else "" for h in rows[0]]
        i_name = next((i for i, h in enumerate(header)
                       if "NAME" in h and "DATE" not in h), None)
        i_code = next((i for i, h in enumerate(header) if "CODE" in h), None)
        i_date = next((i for i, h in enumerate(header) if "DATE" in h), None)
        i_conc = next((i for i, h in enumerate(header)
                       if any(k in h for k in
                              ("CONCLUSION", "REPORT", "DIAGNOSIS",
                               "FINDINGS", "IMPRESSION"))), None)
        n_before = len(out)
        n_skip = 0
        for r in rows[1:]:
            if not r or all(v is None for v in r):
                continue
            name = str(r[i_name]).strip() if (i_name is not None and i_name < len(r)
                                              and r[i_name] is not None) else ""
            code = str(r[i_code]).strip() if (i_code is not None and i_code < len(r)
                                              and r[i_code] is not None) else ""
            if norm_name(name) in ("NAME", "S N", "SN") and norm_name(code) == "CODE":
                n_skip += 1
                continue
            if not name and not code:
                continue
            out.append({
                "sheet": ws.title,
                "name": name,
                "code": code,
                "exam_date": to_yyyymmdd(r[i_date]) if i_date is not None else None,
                "conclusion": (str(r[i_conc]).strip()
                               if (i_conc is not None and i_conc < len(r)
                                   and r[i_conc] is not None) else ""),
            })
        print("      sheet '%s': %d patient rows (%d header rows skipped) | "
              "name=%s code=%s date=%s conclusion=%s"
              % (ws.title, len(out) - n_before, n_skip,
                 header[i_name] if i_name is not None else "-",
                 header[i_code] if i_code is not None else "-",
                 header[i_date] if i_date is not None else "-",
                 header[i_conc] if i_conc is not None else "-"))
    wb.close()
    return out


# ------------------------------ matching ------------------------------------
def build_code_index(excel_rows):
    idx = {}
    for i, row in enumerate(excel_rows):
        for c in extract_codes(row["code"]):
            idx.setdefault(c, []).append(i)
    return idx


def match_patient(candidates, excel_rows, code_idx):
    """Return (row, method, score, second_score); method 'code'|'name'|None."""
    hits = []
    for cand in candidates:
        for c in extract_codes(cand):
            if c in code_idx:
                hits.extend(code_idx[c])
    if hits:
        uniq = sorted(set(hits))
        if len(uniq) == 1:
            return excel_rows[uniq[0]], "code", 1.0, 1.0
        return None, "code_ambiguous", 0.0, 0.0

    best_score, best_row, second = 0.0, None, 0.0
    for cand in candidates:
        cand_norm = norm_name(cand)
        m = re.search(r"(19|20)\d{6}", cand_norm)
        cand_date = m.group(0) if m else None
        for row in excel_rows:
            s = name_score(cand, row["name"])
            if cand_date and row["exam_date"] and cand_date == row["exam_date"]:
                s = min(1.0, s + 0.05)
            if s > best_score:
                second = best_score
                best_score, best_row = s, row
            elif s > second:
                second = s
    if best_row is not None and best_score >= MIN_NAME_SCORE and \
       not (second >= MIN_NAME_SCORE - 0.10 and (best_score - second) < AMBIGUOUS_GAP):
        return best_row, "name", best_score, second
    return best_row, None, best_score, second


# ----------------------- header scrubbing (PS3.15) --------------------------
TAGS_TO_DELETE = [
    # patient identity
    "PatientName", "PatientID", "PatientBirthDate", "PatientBirthTime",
    "OtherPatientIDs", "OtherPatientNames", "OtherPatientIDsSequence",
    "IssuerOfPatientID", "IssuerOfPatientIDQualifiersSequence",
    "PatientAddress", "PatientTelephoneNumbers", "PatientTelecomInformation",
    "PatientComments", "PatientMotherBirthName", "MilitaryRank",
    "BranchOfService", "MedicalRecordLocator", "EthnicGroup", "Occupation",
    "PatientInstitutionResidence", "PatientState",
    "PatientPrimaryLanguageCodeSequence", "PatientInsurancePlanCodeSequence",
    "ResponsiblePerson", "ResponsibleOrganization",
    "ConfidentialityConstraintOnPatientDataDescription", "SpecialNeeds",
    "CurrentPatientLocation", "AdditionalPatientHistory",
    # study / ordering
    "AccessionNumber", "StudyID", "ReferringPhysicianName",
    "PerformingPhysicianName", "NameOfPhysiciansReadingStudy", "OperatorsName",
    "PhysiciansOfRecord", "RequestingPhysician",
    "ScheduledPerformingPhysicianName",
    # NOTE: InstitutionName is intentionally NOT listed here. It is removed
    # at runtime only when identity.retain_institution_name is false.
    # InstitutionAddress, InstitutionalDepartmentName and StationName are
    # always removed - they locate the machine/room far more precisely
    # than the hospital name.
    "InstitutionAddress", "InstitutionalDepartmentName",
    "StationName", "AdmittingDiagnosesDescription", "DischargeDiagnosisDescription",
    "FillerOrderNumberImagingServiceRequest",
    "PlacerOrderNumberImagingServiceRequest",
    "ReasonForTheImagingServiceRequest", "AdmissionID", "ServiceEpisodeID",
    "ServiceEpisodeDescription", "ProtocolName",
    # procedure / order IDs that often carry the hospital code
    "RequestedProcedureID", "ScheduledProcedureStepID",
    "PerformedProcedureStepID",
    # misc free-text comment fields
    "ImageComments", "IdentifyingComments",
    # graphic annotation authors (staff names stored in headers, not pixels)
    "ContentCreatorName", "ContentCreatorIdentificationCodeSequence",
    # device / system
    "DeviceSerialNumber", "GantryID", "DetectorID", "PlateID", "CassetteID",
    "GeneratorID", "PerformedLocation",
]
DATE_TAGS = ["StudyDate", "SeriesDate", "AcquisitionDate", "ContentDate",
             "LastMenstrualDate", "InstanceCreationDate", "AcquisitionDateTime"]
TIME_TAGS = ["StudyTime", "SeriesTime", "AcquisitionTime", "ContentTime",
             "InstanceCreationTime"]
UID_TAGS = ["SOPInstanceUID", "StudyInstanceUID", "SeriesInstanceUID",
            "FrameOfReferenceUID", "InstanceCreatorUID", "DeviceUID"]

DELETE_KW = set(TAGS_TO_DELETE) | set(TIME_TAGS)
DATE_KW   = set(DATE_TAGS)
UID_KW    = set(UID_TAGS)

_uid_cache = {}


def _iter_elements(ds):
    """Yield (owner_dataset, element) for EVERY element, recursing into
    sequences - nested copies of identifying tags must be scrubbed too."""
    snapshot = [(ds, e) for e in ds.elements()]
    for owner, elem in snapshot:
        yield owner, elem
        if elem.VR == "SQ":
            for sub in (elem.value or []):
                yield from _iter_elements(sub)


def shift_date_str(value, delta_days):
    for fmt in ("%Y%m%d%H%M%S.%f", "%Y%m%d%H%M%S", "%Y%m%d"):
        try:
            return (datetime.strptime(str(value).strip(), fmt)
                    + timedelta(days=delta_days)).strftime(fmt)
        except ValueError:
            continue
    return None


def scrub(ds, delta_days, pseudo_id):
    """De-identify one dataset in place, at EVERY nesting level.
    delta_days=None -> keep original dates (PS3.15-compliant default).
    Pixel data is never modified."""
    try:
        ds.remove_private_tags()
    except Exception:
        pass
    for owner, elem in _iter_elements(ds):
        kw = keyword_for_tag(elem.tag)   # tag -> keyword (NOT tag_for_keyword!)
        if not kw:
            continue
        try:
            if kw in DELETE_KW:
                del owner[elem.tag]
            elif delta_days is not None and kw in DATE_KW:
                new = shift_date_str(elem.value, delta_days)
                if new is None:
                    del owner[elem.tag]
                else:
                    owner[elem.tag].value = new
            elif kw in UID_KW:
                old = str(elem.value)
                key = (kw, old)
                if key not in _uid_cache:
                    _uid_cache[key] = uid_remap(old, pseudo_id)
                owner[elem.tag].value = _uid_cache[key]
        except Exception:
            pass
    try:
        if ds.file_meta and "MediaStorageSOPInstanceUID" in ds.file_meta:
            old = str(ds.file_meta.MediaStorageSOPInstanceUID)
            key = ("MediaStorageSOPInstanceUID", old)
            if key not in _uid_cache:
                _uid_cache[key] = uid_remap(old, pseudo_id)
            ds.file_meta.MediaStorageSOPInstanceUID = _uid_cache[key]
    except Exception:
        pass
    ds.PatientName = pseudo_id
    ds.PatientID = pseudo_id
    return ds


# --------------------------- chest-image filter ------------------------------
def chest_check(ds):
    """Return (list_of_reasons, view_position). Empty reasons = clean chest."""
    reasons = []
    mod = str(getattr(ds, "Modality", "") or "").upper()
    if mod not in ("CR", "DX"):
        reasons.append("modality=%s" % (mod or "missing"))
    bp = str(getattr(ds, "BodyPartExamined", "") or "").upper()
    series_desc = str(getattr(ds, "SeriesDescription", "") or "").upper()
    if bp and "CHEST" not in bp:
        reasons.append("body_part=%s" % bp)
    elif not bp and "CHEST" not in series_desc:
        reasons.append("body_part_missing_and_series_not_chest")
    vp = str(getattr(ds, "ViewPosition", "") or "").upper()
    if vp and vp not in ("PA", "AP"):
        reasons.append("view_position=%s" % vp)
    return reasons, vp


def burned_in(ds):
    try:
        return str(getattr(ds, "BurnedInAnnotation", "") or "").upper() == "YES"
    except Exception:
        return False


def identity_candidates(ds, fpath):
    cands = [os.path.splitext(os.path.basename(fpath))[0]]
    for kw in ("PatientName", "PatientID", "AccessionNumber"):
        try:
            v = ds.get(kw)
            if v is not None and str(v).strip():
                cands.append(str(v).replace("^", " "))
        except Exception:
            pass
    return cands


# ------------------------------- audit ---------------------------------------
def run_audit(excel_rows):
    code_idx = build_code_index(excel_rows)
    lines = ["AUDIT REPORT - generated %s" % datetime.now(),
             "Metadata rows read: %d (all sheets)" % len(excel_rows), ""]
    for src_dir, label, prefix in SRC_FOLDERS:
        lines.append("=" * 60)
        lines.append("FOLDER: %s   (label=%s)" % (src_dir, label))
        if not os.path.isdir(src_dir):
            lines.append("  !! FOLDER NOT FOUND")
            continue
        files = []
        for root, _, fns in os.walk(src_dir):
            for fn in fns:
                files.append(os.path.join(root, fn))
        lines.append("  files found: %d" % len(files))
        tally, match_methods = Counter(), Counter()
        name_samples, no_match_samples = [], []
        for fpath in files:
            base = os.path.basename(fpath).upper()
            if base in SKIP_FILENAMES:
                tally["skipped_artifact(%s)" % base] += 1
                continue
            tally["total"] += 1
            try:
                ds = pydicom.dcmread(fpath, stop_before_pixels=True)
                tally["readable_dicom"] += 1
            except Exception:
                tally["NOT_readable"] += 1
                if len(name_samples) < 20:
                    name_samples.append("  UNREADABLE: " + os.path.basename(fpath))
                continue
            for kw in ("Modality", "ViewPosition", "BodyPartExamined",
                       "BurnedInAnnotation"):
                v = str(getattr(ds, kw, "") or "(missing)")
                tally["%s=%s" % (kw, v)] += 1
            row, method, score, second = match_patient(
                identity_candidates(ds, fpath), excel_rows, code_idx)
            if method == "code":
                match_methods["matched_by_code"] += 1
            elif method == "name":
                match_methods["matched_by_name"] += 1
            elif method == "code_ambiguous":
                match_methods["CODE_AMBIGUOUS"] += 1
                if len(no_match_samples) < 20:
                    no_match_samples.append("  AMBIG-CODE: "
                                            + os.path.basename(fpath))
            else:
                match_methods["NO_MATCH"] += 1
                if len(no_match_samples) < 20:
                    no_match_samples.append("  no-match (best=%.2f): %s"
                                            % (score, os.path.basename(fpath)))
            if len(name_samples) < 20:
                name_samples.append("  %s | PN=%r PID=%r"
                                    % (os.path.basename(fpath),
                                       str(ds.get("PatientName", "") or ""),
                                       str(ds.get("PatientID", "") or "")))
        lines.append("  ---- tag tally ----")
        for k in sorted(tally):
            lines.append("    %-40s %d" % (k, tally[k]))
        lines.append("  ---- matching ----")
        for k in sorted(match_methods):
            lines.append("    %-20s %d" % (k, match_methods[k]))
        lines.append("  ---- sample filenames / header identity (up to 20) ----")
        lines.extend(name_samples)
        if no_match_samples:
            lines.append("  ---- unmatched / ambiguous samples (up to 20) ----")
            lines.extend(no_match_samples)
        lines.append("")
    report = os.path.join(OUT_DIR, "audit_report.txt")
    with open(report, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    print("Audit written to:\n  %s\n\nReview it before the full run." % report)


# --------------------------------- main --------------------------------------
def apply_config(cfg, shift_dates_override):
    global METADATA_XLSX, SRC_FOLDERS, OUT_DIR, CLEAN_DIR, QA_DIR, MAP_DIR
    global MIN_NAME_SCORE, AMBIGUOUS_GAP, RANDOM_SEED
    global SHIFT_DATES, SHIFT_MIN_DAYS, SHIFT_MAX_DAYS, SKIP_FILENAMES

    METADATA_XLSX = cfg["project"]["metadata_file"]
    SRC_FOLDERS = [(s["folder"], s["label"], s["prefix"])
                   for s in cfg["sources"]]
    OUT_DIR   = cfg["output"]["base_dir"]
    CLEAN_DIR = os.path.join(OUT_DIR, cfg["output"].get("clean_subdir", "clean_chest"))
    QA_DIR    = os.path.join(OUT_DIR, cfg["output"].get("qa_subdir", "qa_manual_review"))
    MAP_DIR   = cfg["output"]["mapping_dir"]

    MIN_NAME_SCORE = float(cfg.get("matching", {}).get("min_name_score", 0.85))
    AMBIGUOUS_GAP  = float(cfg.get("matching", {}).get("ambiguous_gap", 0.05))
    RANDOM_SEED    = int(cfg.get("matching", {}).get("random_seed", 42))

    dates = cfg.get("dates", {})
    SHIFT_DATES    = bool(dates.get("shift", False))
    SHIFT_MIN_DAYS = int(dates.get("shift_min_days", 60))
    SHIFT_MAX_DAYS = int(dates.get("shift_max_days", 700))
    if shift_dates_override:
        SHIFT_DATES = True

    SKIP_FILENAMES = set(cfg.get("files", {}).get(
        "skip_filenames", ["DICOMDIR", "LABEL.TXT"]))

    # institution-name policy
    retain = bool(cfg.get("identity", {}).get("retain_institution_name", True))
    if not retain:
        DELETE_KW.add("InstitutionName")


def main():
    rng = random.Random(RANDOM_SEED)
    for d in (OUT_DIR, CLEAN_DIR, QA_DIR, MAP_DIR):
        os.makedirs(d, exist_ok=True)

    print("[1/5] Reading metadata (all sheets) ...")
    excel_rows = load_excel_rows(METADATA_XLSX)
    print("      TOTAL: %d patient rows across all sheets." % len(excel_rows))
    code_idx = build_code_index(excel_rows)

    patients = {}
    qa_rows, metadata_rows = [], []
    counters = {}

    for src_dir, label, prefix in SRC_FOLDERS:
        print("[2/5] Processing folder: %s" % src_dir)
        files = []
        for root, _, fns in os.walk(src_dir):
            for fn in fns:
                files.append(os.path.join(root, fn))
        print("      %d files found." % len(files))

        for fpath in files:
            base = os.path.basename(fpath).upper()
            if base in SKIP_FILENAMES:
                qa_rows.append({"file": fpath, "issue": "export_artifact",
                                "detail": "non-image export file, skipped by design"})
                continue
            try:
                ds = pydicom.dcmread(fpath, stop_before_pixels=False)
            except Exception as exc:
                qa_rows.append({"file": fpath, "issue": "not_readable_dicom",
                                "detail": str(exc)[:120]})
                continue

            row, method, score, second = match_patient(
                identity_candidates(ds, fpath), excel_rows, code_idx)

            if method is None:
                reason = "no_match"
                if row is not None:
                    reason = ("low_or_ambiguous_name_match (best=%.2f "
                              "second=%.2f)" % (score, second))
                qa_rows.append({"file": fpath, "issue": reason,
                                "detail": "needs manual review"})
                continue

            key = (norm_name(row["name"]), row["code"], row["exam_date"])
            if key not in patients:
                n = counters.get(prefix, 0) + 1
                counters[prefix] = n
                patients[key] = {
                    "pseudo_id": "%s_%04d" % (prefix, n),
                    "label": label,
                    "row": row,
                    "delta_days": rng.randint(SHIFT_MIN_DAYS, SHIFT_MAX_DAYS),
                    "best_score": score,
                    "n_clean": 0, "n_qa": 0,
                }
            rec = patients[key]

            try:
                scrub(ds, rec["delta_days"] if SHIFT_DATES else None,
                      rec["pseudo_id"])
            except Exception as exc:
                qa_rows.append({"file": fpath, "issue": "scrub_failed",
                                "detail": str(exc)})
                continue

            reasons, vp = chest_check(ds)
            if burned_in(ds):
                reasons.append("burned_in_annotation_flagged_YES")

            if reasons:
                rec["n_qa"] += 1
                dest_dir = os.path.join(QA_DIR, rec["pseudo_id"])
                qa_rows.append({"file": fpath, "issue": "; ".join(reasons),
                                "detail": "pseudo=%s" % rec["pseudo_id"]})
            else:
                rec["n_clean"] += 1
                dest_dir = os.path.join(CLEAN_DIR, rec["pseudo_id"])
                metadata_rows.append({
                    "pseudo_patient_id": rec["pseudo_id"],
                    "image_file": "",
                    "label": rec["label"],
                    "view_position": vp,
                    "modality": str(getattr(ds, "Modality", "") or ""),
                    "exam_date": str(getattr(ds, "StudyDate", "") or ""),
                    "report_conclusion": rec["row"]["conclusion"],
                })
            os.makedirs(dest_dir, exist_ok=True)
            n = rec["n_clean"] + rec["n_qa"]
            out_name = "%s_%03d.dcm" % (rec["pseudo_id"], n)
            try:
                ds.save_as(os.path.join(dest_dir, out_name))
            except Exception as exc:
                qa_rows.append({"file": fpath, "issue": "save_failed",
                                "detail": str(exc)})
                continue
            if not reasons:
                metadata_rows[-1]["image_file"] = os.path.join(
                    rec["pseudo_id"], out_name)

    print("[3/5] Writing metadata_for_colleague.csv ...")
    with open(os.path.join(OUT_DIR, "metadata_for_colleague.csv"), "w",
              newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["pseudo_patient_id", "image_file",
                                           "label", "view_position", "modality",
                                           "exam_date", "report_conclusion"])
        w.writeheader()
        w.writerows(metadata_rows)

    print("[4/5] Writing qa_report.csv ...")
    with open(os.path.join(OUT_DIR, "qa_report.csv"), "w",
              newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["file", "issue", "detail"])
        w.writeheader()
        w.writerows(qa_rows)
    qc = Counter()
    for q in qa_rows:
        qc[q["issue"].split(";")[0].split(" (")[0]] += 1
    print("      QA breakdown:")
    for k, v in qc.most_common():
        print("        %-45s %d" % (k, v))

    print("[5/5] Writing mapping (KEEP LOCAL, DO NOT SHARE) ...")
    with open(os.path.join(MAP_DIR, "real_to_pseudo_mapping.csv"), "w",
              newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["pseudo_id", "label",
                                           "real_name", "real_code",
                                           "real_exam_date",
                                           "date_shift_days",
                                           "match_method", "match_score",
                                           "n_clean_images", "n_qa_images"])
        w.writeheader()
        for key, rec in patients.items():
            w.writerow({
                "pseudo_id": rec["pseudo_id"],
                "label": rec["label"],
                "real_name": rec["row"]["name"],
                "real_code": rec["row"]["code"],
                "real_exam_date": rec["row"]["exam_date"] or "",
                "date_shift_days": rec["delta_days"] if SHIFT_DATES else 0,
                "match_method": "code" if rec["best_score"] == 1.0 else "name",
                "match_score": "%.3f" % rec["best_score"],
                "n_clean_images": rec["n_clean"],
                "n_qa_images": rec["n_qa"],
            })

    print("\nDone.")
    print("  Patients mapped     : %d" % len(patients))
    print("  Clean images        : %d (in %s)" % (len(metadata_rows), CLEAN_DIR))
    print("  QA-quarantined files: %d (see qa_report.csv)" % len(qa_rows))
    print("\nNEXT: run verify_deidentification.py, review qa_report.csv, and")
    print("visually inspect a sample of images before sharing anything.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="De-identify & clean a DICOM chest-X-ray dataset.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__)
    ap.add_argument("--config", default="config.yaml",
                    help="path to your config file (default: config.yaml)")
    ap.add_argument("--audit", action="store_true",
                    help="only inspect folders + matching, write audit_report.txt")
    ap.add_argument("--shift-dates", action="store_true",
                    help="shift all dates by a fixed random per-patient offset "
                         "(overrides dates.shift in config)")
    args = ap.parse_args()
    apply_config(load_config(args.config), args.shift_dates)
    if args.audit:
        print("AUDIT MODE - nothing is modified except audit_report.txt\n")
        os.makedirs(OUT_DIR, exist_ok=True)
        run_audit(load_excel_rows(METADATA_XLSX))
    else:
        main()
