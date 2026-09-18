---
name: security-review
description: Multi-phase, adversarially-verified security audit of this repo's WI/PM document pipeline. Use whenever code touching wi_pipeline.py, WI_Equipment.ipynb, or any .docx/.xlsx input-and-output handling changes, or when the user asks for a security review/audit of a diff or PR.
---

# Security audit for the WI/PM pipeline

Modeled on Cloudflare's `security-audit-skill` methodology (coverage-led
hunting + adversarial candidate validation + machine-readable, three-verdict
findings) but scoped to this repo's actual size and attack surface: a single
pipeline (`wi_pipeline.py`) that parses attacker-controllable `.docx` "Work
Instruction" files (`python-docx`/`lxml`) and emits `.xlsx` workbooks with
embedded images (`openpyxl`/Pillow), plus a manual notebook
(`WI_Equipment.ipynb`). Do not import Cloudflare's memory-safety, cloud, or
AI/LLM attack-class material — none of it applies here; the boundaries below
are the only ones that exist in this codebase.

The point of every phase is the same as theirs: **the agent that finds a
candidate is never the one that decides it's real.** A flat checklist review
produces plausible-sounding findings nobody checked; this process forces a
second, adversarial pass before anything is reported.

## Trust boundaries (recon — do this once per audit, update if the pipeline changes)

1. **Input boundary**: any `.docx` path passed to `process_wi_document` /
   `docx.Document(path)` — content and structure are attacker-controlled.
2. **Output boundary**: the `.xlsx` produced by `save_pm_workbook` /
   `build_pm_workbook` — consumed later by a human in Excel; anything the
   pipeline writes into a cell or filename that originated from boundary 1
   crosses back out here.
3. **Filesystem boundary**: `output_path`, sheet names derived from input
   filenames (`_sheet_name`), and any temp file the pipeline creates.

Every finding must name which boundary it crosses and the concrete call
chain from boundary-1 input to the sink. A finding that can't be traced to
one of these three boundaries is out of scope — reject it in phase 3, don't
report it.

## Phase 1 — Coverage-led hunting

Enumerate every function in `wi_pipeline.py` (and any new pipeline code)
that touches boundary 1 or 2, and check each one against this list. Track
which functions you've covered — an audit that "felt thorough" but skipped
a function is not complete.

- `docx.Document(path)` / any raw `lxml`/`zipfile` use on document content —
  `python-docx` does not enable `defusedxml` protection by default. Flag
  any parse path vulnerable to XXE/billion-laughs, and treat `.docx` as a
  zip archive that needs a member-count/uncompressed-size budget before
  extraction (zip bomb).
- `_image_parts` / `rel.target_part.blob` → `_embed_images` — blobs are
  attacker data fed to Pillow with no `Image.MAX_IMAGE_PIXELS`-style cap
  before decode (`_fit()` only bounds *display* size, not decoded buffer
  size). Decompression-bomb candidate.
- `save_pm_workbook(paths, output_path=...)`, `_sheet_name(pm_name)` — any
  path/filename built from document content or unsanitized caller input
  without confirming it can't escape the output directory (`..`, absolute
  paths, symlinks), and confirm `_sheet_name`'s `[]:*?/\\` stripping isn't
  bypassed by reusing the raw name elsewhere.
- `FIGURE_RE`, `CONCEPT_RULES`, `_trigrams`, and any new regex run over
  full (attacker-supplied, unbounded-length) document text — nested
  quantifiers are a ReDoS candidate.
- Cell values written from document text (`step.text`, `step.equipment`,
  group titles) into the `.xlsx` — formula/CSV injection if a value starts
  with `=`, `+`, `-`, `@` and is later exported/opened without Excel's
  default warning suppressed.
- Dependency pins (`python-docx`, `openpyxl`, `Pillow`) — flag a bump *or*
  a pin that reintroduces a version with a known XML/image-handling CVE.

Record every candidate here, even ones you suspect are non-issues — phase 2
disposes of those, don't pre-filter.

## Phase 2 — Adversarial candidate validation

For each candidate from phase 1, re-derive it from scratch as if you had
not just written it: find the actual call chain from a boundary-1 input to
the sink, with file:line at each hop. If you cannot construct a concrete,
bounded trace — a real input value and a real code path it flows through —
the candidate does not survive this phase.

Actively try to disprove each one: is the input actually attacker-reachable
in this repo's usage (a human runs the notebook on files they chose), or
does something upstream already bound it? A theoretical Pillow bomb where
the `.docx` files are supplied by the same person running the pipeline is
weaker than a supported CVE with no mitigation — say so in the verdict, don't
silently drop the severity distinction.

## Phase 3 — Structured verdicts

Classify every phase-1 candidate into exactly one of:

- `confirmed` — complete source trace from a named boundary through to the
  sink, with a concrete triggering input.
- `needs_validation` — a real gap exists (e.g., "no size cap before
  `ZipFile.read`") but you could not fully bound the impact or reproduce a
  trigger — report the fact, without inventing a severity you didn't
  verify.
- `rejected` — investigated and disproved (input isn't attacker-reachable
  here, or an existing check already bounds it). State why, so a re-audit
  doesn't re-open it without new information.

## Phase 4 — Independent re-verification

Before reporting, re-check each `confirmed`/`needs_validation` finding's
source trace one more time, specifically looking for a reason it's wrong
(wrong line, the mitigating check you missed, a call path that doesn't
actually exist). Only findings that survive this get reported. This mirrors
the source repo's separation between "found it" and "verified it" — don't
let phase 1's author be the only check.

## Phase 5 — Reporting

Report as `findings.json`-shaped entries (adapt field names to whatever the
calling tool expects, e.g. `ReportFindings`):

```json
{
  "verdict": "confirmed | needs_validation | rejected",
  "boundary": "input | output | filesystem",
  "file": "wi_pipeline.py",
  "line": 0,
  "summary": "one-sentence defect statement",
  "trace": ["boundary input -> function -> ... -> sink"],
  "fix": "minimal, concrete change (e.g. 'cap zip member size before ZipFile.read')"
}
```

Never emit a finding that skipped phase 2 or phase 4. `rejected` items are
worth listing briefly (so the same non-issue isn't re-litigated) but must
not appear as security risks.

## Explicitly out of scope

- Standard library `re`/`dataclasses`/`openpyxl` styling APIs used as
  documented.
- Web-app-style auth/session controls for `WI_Equipment.ipynb` — it's a
  local, manually-run notebook, not a service.
- Any Cloudflare attack-class category (memory safety, cloud/deployment,
  AI/LLM, network protocols) that has no corresponding code in this repo.
