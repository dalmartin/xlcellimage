I had claude conduct research on the bug, and it found a mismatch between current rich value mapping and actual rich value mapping, as expected.

## Update 2026-08-22: fix implemented and independently verified, not yet merged

Root cause was more specific than a simple off-by-one: `workbook_parser.py` never parsed `xl/richData/rdrichvalue.xml` or `xl/richData/rdrichvaluestructure.xml` at all, so it skipped real hops of indirection (futureMetadata block -> rv index -> rv structure -> LocalImageIdentifier) and only produced correct results on simple single-image files by coincidence. It also never verified a `<bk>`'s `<rc>` was actually the `XLRICHVALUE`-typed one, so a cell with both `XLDAPR` (dynamic array) and `XLRICHVALUE` metadata would resolve wrong.

- Two-agent process: one implemented the fix against the full corrected resolution algorithm, a second independently re-derived the algorithm by hand-tracing the raw XML in the test fixture (not by reading the first agent's explanation) and checked its work.
- New regression test (`tests/data/test_workbook_multi_rc.xlsx` + `tests/data/test_workbook_parser_multi_rc.py`) built specifically to catch this bug: confirmed it **fails against the old code** (5/20 failing, wrong images resolved, matching the diagnosed bug exactly) and **passes against the fix** (20/20).
- Only file changed: `src/xlcellimage/workbook_parser.py`. Nothing committed, on purpose — everything's sitting in the working tree (`git diff` / `git status`) for review before merging, per your note above.
- One open judgment call worth a look: the fix extended graceful "missing file -> no images" handling to `richValueRel.xml.rels` too, not just the two new rich-data files, on the theory that a workbook without any Place-in-Cell images shouldn't crash on construction.

**Still needed before this ships**: run it against the actual private file that originally surfaced this bug (not available to Claude) to confirm real-world behavior, then decide on merging `bug/incorrect-image-mapping` into `main`.
