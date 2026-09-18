---
name: security-review
description: Review changes to this repo's WI/PM document pipeline for security issues before merging. Use whenever code touching wi_pipeline.py, WI_Equipment.ipynb, or any .docx/.xlsx input-and-output handling changes, or when the user asks for a security review of a diff or PR.
---

# Security review for the WI/PM pipeline

This repo parses untrusted `.docx` "Work Instruction" files (via `python-docx`
/ `lxml`) and generates `.xlsx` workbooks with embedded images (via
`openpyxl` / Pillow). Both directions handle attacker-controllable file
content, so review changes with that threat model in mind rather than
generic advice.

## What to check on every change to the pipeline

1. **XML parsing of `.docx` input (`docx.Document(path)`, any raw `lxml`/
   `zipfile` use)**
   - `python-docx` parses the embedded XML with `lxml` and does **not**
     enable `defusedxml`/entity-expansion protection by default. Confirm no
     new code path introduces raw `lxml.etree.parse`/`fromstring` on
     document-supplied XML without disabling external entity resolution
     (`resolve_entities=False`, `no_network=True`) — this is the classic
     XXE / billion-laughs vector for `.docx` (which is just a zip of XML).
   - Treat any `.docx` as a zip archive: reject or cap deeply nested /
     absurdly large member counts and uncompressed sizes before extracting
     (zip bomb). Flag any new code that trusts `ZipFile.infolist()` sizes or
     iterates without a total-size budget.

2. **Path handling**
   - `save_pm_workbook(paths, output_path=...)` and `_sheet_name(pm_name)`
     derive filenames/sheet names from input file paths. Flag any change
     that writes to a path built from document content or an unsanitized
     caller-supplied string without confirming it can't escape the intended
     output directory (`..`, absolute paths, symlinks).
   - `_sheet_name` already strips `[]:*?/\\` for Excel sheet-name validity —
     don't regress that if refactored, and don't reuse the raw (unstripped)
     name anywhere a path is built from it.

3. **Embedded images (`_image_parts`, `_embed_images`)**
   - Images come straight from the `.docx`'s relationship parts
     (`rel.target_part.blob`) into `openpyxl`/Pillow. Flag any new code that
     loads these blobs without a size cap — Pillow decompression bombs
     (`Image.MAX_IMAGE_PIXELS`) are relevant since blobs are attacker
     data. `_fit()` bounds *display* size but not the decoded pixel buffer
     size before that.
   - Confirm `io.BytesIO(blob)` usage doesn't get replaced with something
     that writes the blob to a predictable temp path (symlink/race risk).

4. **Regex-driven parsing (`FIGURE_RE`, `CONCEPT_RULES`, `_trigrams`)**
   - These run over full document text. Flag any new regex with nested
     quantifiers (`(a+)+`, etc.) added to `CONCEPT_RULES` or similar —
     ReDoS is plausible given documents are attacker-supplied and text
     length is not bounded before matching.

5. **Dependency/output trust boundary**
   - `openpyxl`, `python-docx`, `Pillow` versions: flag if a change pins to
     or reintroduces a version with a known CVE in XML or image handling.
   - The generated `.xlsx` is later opened by end users — don't introduce
     macros, external references, or formulas built from unsanitized
     document text (formula/CSV injection if a cell value starts with
     `=`, `+`, `-`, `@`).

## What NOT to flag

- Standard library `re`, `dataclasses`, `openpyxl` styling APIs used as
  documented — not a supply-chain or injection concern by themselves.
- The notebook (`WI_Equipment.ipynb`) is a local/manual workflow, not a
  service; don't demand web-app-style auth/session controls for it.

## Output format

List findings as: file:line, the concrete attacker-controlled input that
reaches the sink, and the minimal fix (e.g. "cap zip member size before
`ZipFile.read`", not "add general input validation"). Skip anything you
can't tie to an actual data flow in this repo.
