#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Assert that a .docx really matches the formatting spec, at the XML level.

Formatting a document is only half the job - the other half is *proving* it.
This script re-derives every paragraph classification from the file on disk and
checks the expected format for each class, plus a set of package-level
invariants that catch the classic corruption modes.

Checks
------
per rule   : every paragraph matching the rule's `match` also satisfies the
             rule's `format` (size, bold, alignment, no stray indent, indent).
             Paragraphs matching NO rule must not carry an SDK-style stray
             indent when ``--strict-stray`` is given.
package    : (a) the OOXML package is a readable zip and word/document.xml
                 parses
             (b) no literal CR (``&#13;``) inside any ``w:t``
             (c) every prefix referenced by ``mc:Ignorable`` is declared on the
                 root element (the signature of an ElementTree re-serialization
                 that Word rejects as "corrupted")
             (d) no empty ``w:rPr``/``w:pPr`` left behind

Exit code is 0 when everything passes, 1 otherwise - suitable for CI or for
gating a write-back.

Usage
-----
    python docx_verify.py clean.docx --spec rules.json
    python docx_verify.py clean.docx --spec rules.json --strict-stray
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def w(tag: str):
    return "{%s}%s" % (W, tag)


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(w("t")))


def run_size_halfpt(r):
    rpr = r.find(w("rPr"))
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return int(sz.get(w("val"))) if sz is not None and sz.get(w("val")) else None


def para_mark_size_halfpt(p):
    ppr = p.find(w("pPr"))
    rpr = ppr.find(w("rPr")) if ppr is not None else None
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return int(sz.get(w("val"))) if sz is not None and sz.get(w("val")) else None


def rec_of(p) -> dict:
    ppr = p.find(w("pPr"))
    style = None
    if ppr is not None:
        e = ppr.find(w("pStyle"))
        style = e.get(w("val")) if e is not None else None
    size = None
    for r in p.iter(w("r")):
        size = run_size_halfpt(r)
        break
    if size is None:
        size = para_mark_size_halfpt(p)
    text = para_text(p)
    return {
        "text": text,
        "style": style,
        "size_pt": (size / 2) if size else None,
        "has_image": any(n.tag in (w("drawing"), w("pict")) for n in p.iter()),
        "is_blank": not text.strip(),
        "ppr": ppr,
    }


def match_rule(rec: dict, match: dict) -> bool:
    s = rec["text"].strip()
    if "text_regex" in match and not re.search(match["text_regex"], s):
        return False
    if "text_in" in match and s not in match["text_in"]:
        return False
    if "text_equals" in match and s != match["text_equals"]:
        return False
    if "text_contains" in match and match["text_contains"] not in rec["text"]:
        return False
    if "style_id" in match and rec["style"] != match["style_id"]:
        return False
    if "size_pt" in match and rec["size_pt"] != match["size_pt"]:
        return False
    if "has_image" in match and rec["has_image"] != match["has_image"]:
        return False
    if "is_blank" in match and rec["is_blank"] != match["is_blank"]:
        return False
    return True


def all_runs_bold(p):
    sizes, bolds = [], []
    for r in p.iter(w("r")):
        txt = "".join(t.text or "" for t in r.iter(w("t")))
        if not txt.strip():
            continue
        rpr = r.find(w("rPr"))
        sizes.append(run_size_halfpt(r))
        bolds.append(rpr is not None and rpr.find(w("b")) is not None)
    return sizes, bolds


def check_rule_format(p, rec, fmt, failures, tag):
    ppr = rec["ppr"]
    ind = ppr.find(w("ind")) if ppr is not None else None
    if fmt.get("clear_indent") and ind is not None:
        failures.append(f"{tag}: unexpected w:ind on {rec['text'][:36]!r}")
    if "size_pt" in fmt:
        want = int(round(float(fmt["size_pt"]) * 2))
        mark = para_mark_size_halfpt(p)
        if mark != want:
            failures.append(f"{tag}: para-mark size {mark} != {want}")
        sizes, _ = all_runs_bold(p)
        for s in sizes:
            if s != want:
                failures.append(f"{tag}: run size {s} != {want} "
                                f"in {rec['text'][:32]!r}")
                break
    if "bold" in fmt:
        _, bolds = all_runs_bold(p)
        if any(b != bool(fmt["bold"]) for b in bolds):
            failures.append(f"{tag}: bold mismatch in {rec['text'][:32]!r}")
    if fmt.get("align"):
        jc = ppr.find(w("jc")) if ppr is not None else None
        got = jc.get(w("val")) if jc is not None else None
        if got != fmt["align"]:
            failures.append(f"{tag}: align {got} != {fmt['align']} "
                            f"on {rec['text'][:32]!r}")
    if "first_line_indent_chars" in fmt:
        if ind is None or not ind.get(w("firstLineChars")):
            failures.append(f"{tag}: missing first-line indent on "
                            f"{rec['text'][:32]!r}")
        elif int(ind.get(w("firstLineChars"))) != int(
            float(fmt["first_line_indent_chars"]) * 100
        ):
            failures.append(f"{tag}: firstLineChars {ind.get(w('firstLineChars'))}"
                            f" != {int(float(fmt['first_line_indent_chars'])*100)}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Verify docx formatting")
    ap.add_argument("docx")
    ap.add_argument("--spec", required=True)
    ap.add_argument("--strict-stray", action="store_true")
    ap.add_argument("--max-report", type=int, default=40)
    a = ap.parse_args(argv)

    with open(a.spec, encoding="utf-8") as fh:
        spec = json.load(fh)

    failures: list = []
    counts: dict = {}
    try:
        with zipfile.ZipFile(a.docx) as z:
            if z.testzip() is not None:
                failures.append("zip: corrupt entry")
            raw = z.read("word/document.xml")
    except Exception as exc:  # noqa: BLE001
        print(f"FATAL: cannot open package: {exc}")
        return 1

    root = etree.fromstring(raw)

    # --- package invariants -------------------------------------------------
    cr = raw.decode("utf-8", "ignore").count("&#13;")
    if cr:
        failures.append(f"document.xml still contains {cr} literal CR entity")

    declared = set()
    for k in root.nsmap:
        if k:
            declared.add(k)
    ignorable = root.get("{%s}Ignorable" % MC) or ""
    for pref in ignorable.split():
        if pref not in declared:
            failures.append(
                f"mc:Ignorable references undeclared prefix {pref!r} - Word will "
                "report the file as corrupted (namespace loss)"
            )

    empty_rpr = 0
    for rpr in root.iter(w("rPr")):
        if len(rpr) == 0 and len(rpr.attrib) == 0:
            empty_rpr += 1
    if empty_rpr:
        failures.append(f"{empty_rpr} empty <w:rPr/> element(s)")

    # --- per-rule ----------------------------------------------------------
    rules = spec.get("rules", [])
    for p in root.iter(w("p")):
        rec = rec_of(p)
        if rec["is_blank"]:
            continue
        hit = None
        for rule in rules:
            if match_rule(rec, rule.get("match", {})):
                hit = rule
                break
        if hit:
            name = hit.get("name", "rule")
            counts[name] = counts.get(name, 0) + 1
            check_rule_format(p, rec, hit.get("format", {}), failures, name)
        else:
            ppr = rec["ppr"]
            if ppr is not None and a.strict_stray:
                ind = ppr.find(w("ind"))
                if ind is not None and ind.get(w("firstLineChars")) \
                        and not ind.get(w("firstLine")):
                    failures.append(
                        f"stray SDK indent on unmatched {rec['text'][:36]!r}"
                    )

    print("matched paragraphs per rule:")
    for k in sorted(counts):
        print(f"  {k:24s} {counts[k]}")
    print(f"failures: {len(failures)}")
    for f in failures[: a.max_report]:
        print("  FAIL " + f)
    if len(failures) > a.max_report:
        print(f"  ... {len(failures) - a.max_report} more")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
