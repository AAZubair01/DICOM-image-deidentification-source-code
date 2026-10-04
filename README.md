# dicom-cxr-deidentify

De-identify and clean a DICOM chest-X-ray dataset (e.g. paediatric pneumonia
vs control) so it can be shared with a collaborator for machine-learning
pilot work — with matching to a spreadsheet of radiology reports, PS3.15-style
header scrubbing, chest-only filtering, and an automated re-identification
self-audit.

Built for a real two-folder, two-sheet workflow (59 pneumonia + 25 control
patients, CD-export style `DICOMDIR` / `LABEL.TXT` / `DICOMOBJ` structures),
but fully config-driven so it can be reused for the next phase of a study or
adapted to a different dataset.

---

## ⚠️ Read this first (ethics & safety)

- **You** are responsible for ethics approval / institutional sign-off before
  sharing patient data, even anonymized. This tool does not replace that.
- This is a research utility, **not a certified medical-device anonymizer**.
  Always run the verification step and **visually inspect a sample of images**
  before sharing anything.
- **Never commit outputs to git.** The `.gitignore` blocks the usual output
  locations; extend it if you change the folder layout. The
  `mapping_keep_local_DO_NOT_SHARE` folder contains the real identities and
  must stay on your machine, ideally encrypted.
- The pipeline **never modifies pixel data**. Burned-in text in the image
  itself is only *flagged* (`BurnedInAnnotation = YES` goes to the QA queue) —
  detecting text burned into pixels is a human/audit task.

## What it does

| Step | Detail |
|---|---|
| 1. Metadata | Reads **every sheet** of your spreadsheet (cases, controls), auto-detecting name/code/date/conclusion columns and skipping duplicated header rows |
| 2. Matching | Links each DICOM to a patient primarily by **numeric hospital code** (filename / PatientID / PatientName / AccessionNumber → CODE column), with fuzzy name matching as fallback. Weak/ambiguous matches go to a review queue — **never guessed** |
| 3. Scrubbing | Deletes identifying tags per DICOM PS3.15 basic profile **at every nesting level** (sequences included) + all private tags; deterministically remaps UIDs; replaces PatientName/ID with pseudo-IDs (`PNEU_0001`, `CTRL_0001`); optionally shifts dates per-patient (`--shift-dates`) |
| 4. Filtering | Keeps chest PA/AP CR/DX only; quarantines non-chest anatomy, presentation states (modality `PR`), and flagged burned-in annotations into `qa_manual_review/` |
| 5. Verification | A separate script hunts every header value for the real names/codes from your spreadsheet and reports the exact tag path of any hit |

## Quickstart

```bash
pip install -r requirements.txt

cp config.example.yaml config.yaml
# edit config.yaml: point to your spreadsheet, your DICOM folders, your outputs

# 1. audit first - inspects everything, modifies nothing
python src/deidentify_dataset.py --config config.yaml --audit

# 2. review audit_report.txt, then run the pipeline
python src/deidentify_dataset.py --config config.yaml

# 3. verify - attack your own data before anyone else can
python src/verify_deidentification.py --config config.yaml
python src/verify_deidentification.py --config config.yaml --ignore "AMIN,KANO,TEACHING,HOSPITAL"
```

## Configuration

Everything lives in `config.yaml` (see `config.example.yaml` for annotated
placeholders):

- `project.metadata_file` — your spreadsheet. Needs a **NAME** and a
  **CODE** column (auto-detected); DATE and CONCLUSION columns are optional.
- `sources` — one entry per DICOM folder, each with a `label` (becomes the
  class label) and `prefix` (becomes the pseudo-ID prefix). Add as many as
  you like.
- `matching` — fuzzy-match thresholds and the random seed (fixed seed =
  reproducible pseudo-IDs and date shifts across re-runs).
- `dates.shift` — `false` keeps real exam dates (PS3.15-compliant);
  `true` shifts each patient's dates by a consistent random offset.
- `identity.retain_institution_name` — keep the hospital name (`true`) or
  scrub it too (`false`). Address/department/station are always removed.
- `output` — where clean data, QA queue and the (never-shared) mapping go.

## Outputs

```
<output.base_dir>/
├── clean_chest/<PSEUDO_ID>/<PSEUDO_ID>_NNN.dcm   verified chest images
├── qa_manual_review/<PSEUDO_ID>/...              YOU review before sharing
├── metadata_for_colleague.csv                    pseudo_id, image_file, label,
│                                                 view_position, modality,
│                                                 exam_date, report_conclusion
├── qa_report.csv                                 why each file was quarantined
├── audit_report.txt                              from --audit runs
└── verification_report.csv                       from the verifier

<output.mapping_dir>/
└── real_to_pseudo_mapping.csv                    real name/code ↔ pseudo-ID
                                                  >>> NEVER share this <<<
```

## Understanding the QA report

Typical categories (all legitimate, nothing silently lost):

- `export_artifact` — CD/PACS export files (`DICOMDIR`, `LABEL.TXT`); not images.
- `modality=PR` — Presentation State objects (display transforms, not images).
- `body_part=NECK/ABDOMEN/...` — non-chest studies sitting in the folders.
- `burned_in_annotation_flagged_YES` — inspect these visually for embedded text.
- `no_match` / `low_or_ambiguous_name_match` — investigate manually against
  your spreadsheet.

## Understanding verification hits

After a correct run, remaining hits are almost always one of:

1. **Deliberately retained institution name** — whitelist it with `--ignore`.
2. **Coincidental digit collisions** — a 40-digit hashed UID can contain any
   5-digit code by chance (~5 expected per dataset); the `2.25.` prefix and
   the value snippet prove it's a hash, not an identifier.
3. **Binary sequence blobs** — raw bytes that happen to match a token.

A hit is only dangerous when the *snippet* shows a real name/code in a
meaningful field — the report tells you exactly which tag, so the fix is
obvious (add the tag to `TAGS_TO_DELETE` in `src/deidentify_dataset.py`).

## Packaging & sharing

```powershell
# checksums
Get-ChildItem -Recurse clean_chest | Get-FileHash -Algorithm SHA256 > checksums_sha256.txt
# encrypted archive (AES-256, header-encrypted)
7z a -pYOURSTRONGPASSWORD -mhe=on dataset_v1.7z clean_chest metadata_for_colleague.csv checksums_sha256.txt
```

Send the archive via an encrypted channel (institutional cloud / SFTP), the
password through a **different** channel, and ask the receiver to verify the
checksums. **Never** put patient-derived data on a public code host — this
repository is for the code only.

## Limitations

- Matching assumes codes/names appear in filenames or DICOM headers; unusual
  exports may need the audit-driven tweaks described above.
- Dates are kept by default (policy choice); pixel-level text detection is
  manual.
- This software is provided as-is under the MIT License (see `LICENSE`),
  without warranty; regulatory compliance remains the user's responsibility.

## Reuse for the next phase

Point `sources` at the new folders, update `project.metadata_file`, keep the
same `matching.random_seed` if you want globally consistent pseudo-IDs across
phases (or change it if phases must be unlinkable).
