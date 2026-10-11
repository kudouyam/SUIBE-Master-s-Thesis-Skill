# Agent instructions — SUIBE Master's Thesis Skill

This is an **unofficial** formatting Skill for Shanghai University of International Business and Economics (SUIBE) 2020 master's thesis requirements.

1. First read [SKILL.md](SKILL.md), then read `references/01_学校规范逐项对照.md`. Treat official school rules as authoritative over external GitHub template conventions.
2. If a user provides a .docx, **preserve the original**. Do not alter content or fabricate missing data.
3. Ensure Python 3.10+ and dependencies from requirements.txt; run `python scripts/format_thesis.py "<input.docx>"`.
4. Deliver the output DOCX plus the Markdown/JSON revision reports.
5. Report low-confidence titles, complex tables, captions, page breaks, citations, footnotes, field refresh, and unverified current-year rules. Visual-review final Word layout. Never claim 100% conformance without verification.
6. The repository already contains the uncompressed implementation at `scripts/format_thesis.py`; do not fetch a private ZIP or invent an alternative formatter. If the file is missing, report the incomplete checkout.

For users passing only the repository link, fetch/clone it first. An Agent lacking filesystem, code execution, or GitHub access cannot apply this Skill.