---
name: doc-sorting
description: This skill should be used when a Chinese report or course lab document (local .docx) needs its blanks filled in, figure/table captions added, and layout normalized - including 图注/表注, section labels such as 程序截图/运行结果/文字说明, body first-line indentation, font sizes, bold emphasis, and splitting run-on paragraphs. Trigger phrases include 补全实验报告、加图注、加表注、字体调小、正文缩进、首行缩进、加粗、该有的换行、排版、美化报告, or any request to batch-format a .docx while keeping its existing images, tables and code screenshots intact. Encodes the hybrid editor-SDK + direct-XML workflow, the coordinate-drift / namespace-corruption / editor-save-round-trip traps, and Word-to-PDF-to-PNG visual verification.
agent_created: true
---

# Doc Sorting

## Overview

Fill in, caption, and normalize the layout of a Chinese .docx report **without
damaging what is already there** (existing screenshots, tables, question text).

The core architectural rule, learned from real damage:

> **Insert content through the editor SDK. Apply formatting through direct XML.**
> Never bulk-format through editor character offsets, and never treat an editor
> `save_file` as the authoritative write-back.

| Job | Tool | Why |
|---|---|---|
| Read / understand structure | `scripts/docx_scan.py` | read-only, no editor needed |
| Locate an anchor | editor `doc_find` | returns exact `begin`/`end`/`paragraph_id` |
| Insert paragraphs, captions, labels | `scripts/docx_insert.py` | atomic insert + format in one call |
| Bulk formatting (size/bold/align/indent) | `scripts/docx_format.py` | content-matched, coordinate-free, idempotent |
| Prove it is correct | `scripts/docx_verify.py` | XML-level assertions |
| See real pagination | `scripts/docx2pdf.ps1` → `scripts/pdf2png.py` | XML cannot show pagination |

## Before You Start

1. **Back up.** `docx_format.py --inplace` creates `<name>.docx.bak` automatically;
   for other paths copy manually.
2. **Read `references/pitfalls.md`.** All five traps there have corrupted a real
   document. Trap 0 (editor `save_file` re-serializes and degrades the whole
   document) is the most invisible.
3. Check the file is not open in WPS/Word - a held lock surfaces as
   `PermissionError` on write-back. Close the editor session
   (`edsdk_client.py close_file --fid ... --json '{"force":true}'`) and retry.

## Workflow

### Phase 1 - Recon (read-only)

```bash
PY=python                                  # use an isolated interpreter
S=~/.workbuddy/skills/doc-sorting/scripts

$PY $S/docx_scan.py report.docx                     > scan.tsv   # full model
$PY $S/docx_scan.py report.docx --images                          # image paragraphs
$PY $S/docx_scan.py report.docx --only "文字说明"                  # find the blanks
```

Note the real `style_id` and font size of body paragraphs - the rules in
Phase 3 must match them.

Also: extract `word/media/*` into a labelled contact sheet (Pillow grid) and
look at every screenshot, and read the companion source directory if there is
one. Captions must describe what the screenshot actually shows; filled-in text
must agree with the real code - never invent class names, method names or
output values.

### Phase 2 - Fill blanks (editor SDK)

```bash
$PY $S/docx_insert.py anchors report.docx "文字说明：" "运行结果："   # offsets
$PY $S/docx_insert.py images  report.docx                            # image offsets
$PY $S/docx_insert.py plan    plan.json                              # dry run
$PY $S/docx_insert.py run     plan.json                              # execute + save
```

Plan ops - the script applies them in **descending `idx` order** so earlier
insertions never invalidate later ones:

```json
{ "file_id": "D:/report.docx", "save": true,
  "ops": [
    {"after_image": 5429, "text": "图 2-5 完善后的 CSSE 类代码",
     "para": {"jc": "center"}, "font": {"size_pt": 9}},
    {"idx": 14702, "text": "运行结果输出了 5 个行政办公室。"}
  ] }
```

* `after_image` inserts at `image.index + 1` → caption lands directly below.
* **Missing line breaks** (a `运行结果…` sentence glued onto the end of a
  paragraph) are fixed by inserting an empty paragraph at the glue point:
  `{"idx": <glue position>, "text": ""}` - text before inherits the original
  paragraph properties.
* Tables: fill with editor `doc_set_table_cells`
  (`cells: [{row, col, text}]`) in one call; put the table caption above it
  with a normal insert.
* Filled prose must read like the student wrote it - see
  `references/layout-spec.md` → 语气.

### Phase 3 - Normalize layout (direct XML, coordinate-free)

```bash
cp $S/../assets/course-report.rules.json ./rules.json
# edit rules.json: body rule's style_id + size_pt must match Phase 1 findings

$PY $S/docx_format.py report.docx --spec rules.json --dry-run
$PY $S/docx_format.py report.docx --spec rules.json --out clean.docx
```

Rules select paragraphs by content signature (regex / exact text / style id /
current size / has-image / blank), never by offset - so the run is idempotent
and cannot hit neighbours. First matching rule wins; put specific rules first.

If a previous pass already damaged the file,
`--clear-stray-indent` removes the SDK-fingerprint indent
(`w:ind` with `firstLineChars` but no `firstLine`) from unmatched paragraphs.

### Phase 4 - Verify (two layers)

```bash
$PY $S/docx_verify.py clean.docx --spec rules.json --strict-stray   # must exit 0

powershell -ExecutionPolicy Bypass -File $S/docx2pdf.ps1 -Src clean.docx -Pdf clean.pdf
$PY $S/pdf2png.py clean.pdf pages/
```

Assertions catch: per-rule counts and formats, `mc:Ignorable` prefix loss
(Word's "file may be corrupted"), leftover `&#13;`, empty `w:rPr`, stray indents.
Rendering catches what XML cannot: pagination, caption placement, table column
width, question text accidentally enlarged. **Always look at the rendered pages.**

Then write back (`--inplace`, or copy over the target), and re-open the document
in the editor to refresh the preview.

## Hard Rules

* Never bulk-format through editor offsets. Every write shifts them.
* Never use `range.end + 1` - crossing the paragraph mark injects a literal
  `\r` (`&#13;`) into `w:t` and shifts every later offset.
* Never re-serialize `document.xml` with `xml.etree.ElementTree` - it drops the
  namespace prefixes that `mc:Ignorable` references and Word rejects the file.
  Use **lxml**.
* Never treat editor `save_file` as the final write-back for formatting you
  control in XML; it prunes empty paragraphs, drops paragraph-mark `rPr`, and
  re-applies indents inside table cells. Re-verify after any editor save.
* Do not silently rewrite a screenshot or a code result to match the question.
  If the code contradicts the assignment (e.g. the question says `3+9i` but
  `main` uses `3+2i`), report the discrepancy and let the user decide - this is
  the one place to stop and ask.
* Leave template question text untouched.

## Resources

* `scripts/edsdk_client.py` - auto-locates and drives the local editor SDK
  (`--where`, `--pool`, `--list`, `--schema`, or call any tool directly)
* `scripts/docx_scan.py` - read-only paragraph model (text, size, bold, align,
  indent, has-image, approximate offsets)
* `scripts/docx_insert.py` - anchor/image resolution + safe ordered insertion
* `scripts/docx_format.py` - rule-driven lxml formatting, `--dry-run`,
  `--inplace` (auto `.bak`), `--clear-stray-indent`
* `scripts/docx_verify.py` - XML assertions, exit code gates the write-back
* `scripts/docx2pdf.ps1`, `scripts/pdf2png.py` - Word COM → PDF → PNG
* `references/workflow.md` - the full walk-through with commands and a
  pre-delivery checklist
* `references/edsdk.md` - editor tool inventory, `idx` semantics, offset model
* `references/pitfalls.md` - the five corruption traps and how to recover
* `references/layout-spec.md` - Chinese report typography: font sizes (pt ↔
  half-points), 首行缩进 twips math, caption naming, prose tone
* `assets/course-report.rules.json` - starter rules spec; adjust the `body`
  rule's `style_id`/`size_pt` to the document at hand
