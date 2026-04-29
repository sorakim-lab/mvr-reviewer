# mvr-reviewer

A rule-based document reviewer for **Method Validation Reports (MVR)** in GMP-regulated pharmaceutical environments.

> **Why rule-based, not ML?** In GMP environments, explainability is not a feature — it is a regulatory requirement. Every flagged item must trace back to a documented rule, not a model's confidence score. This tool is designed around that constraint.

---

## Background

In GMP pharmaceutical sites, Method Validation Reports follow strict internal conventions: terminology spacing, table header standards, mandatory sections (CE/CPA/LMW definitions, Revision History, signatures), bilingual captions, and equipment calibration formatting. Reviewers spend significant time catching the same recurring formatting issues across documents — issues that are deterministic, well-defined, and ideal for automation.

Most "AI document review" tools rely on machine learning, but ML carries two problems in regulated environments:

1. **Validation burden**: A black-box model cannot be qualified under GMP without extensive justification of its decision boundaries.
2. **Explainability gap**: Auditors and reviewers need to know *why* something was flagged. "The model gave it 0.73 confidence" is not an acceptable answer in front of a regulator.

This project takes the opposite approach: a **fully transparent rule-based reviewer** where every flag traces to a YAML-defined rule with an explicit ID, severity, rationale, and pattern.

## Design principles

- **Deterministic over probabilistic** — same document, same rules, always same result.
- **Rules as data, not code** — adding/modifying rules requires editing YAML, not Python. The YAML file itself can be attached to an SOP as the documented review specification.
- **Severity-tiered, not binary** — `required` (must fix), `recommended` (style), `info` (manual confirmation needed). Reviewers retain final judgment.
- **Layered output** — JSON for machine consumption, Markdown for quick review, DOCX-with-comments for working directly in Word.
- **Operates as an external layer** — does not modify or interact with validated systems. The original document is never altered; review output is always a separate file.

## Research context

This tool is part of a broader **Regulatory HCI** research program examining how to design tools that simultaneously satisfy regulatory constraints and improve worker conditions in GMP environments. The contribution is not the rule engine itself (well-understood since the 1970s) but the explicit framing of *why we did not use machine learning* as a design choice rather than a limitation. Trade-offs between performance and explainability are not absolute — they are determined by institutional context.

The rule catalog is grounded in field experience: every rule corresponds to an issue that recurred across multiple MVR review cycles in a real GMP QC environment.

---

## What this tool does

Given a `.docx` MVR file and a YAML rule catalog, it produces:

1. **`*_review.json`** — structured violation list (rule ID, severity, location, message, matched text)
2. **`*_review.md`** — human-readable Markdown report grouped by severity
3. **`*_reviewed.docx`** — original document with inline Word comments at violation locations

### Supported check types

| Type | Purpose | Example |
|---|---|---|
| `regex_replace` | Find a wrong expression to be replaced | `"허용 기준"` → `"허용기준"` |
| `regex_find` | Flag patterns that should/shouldn't appear | Forbid `"검교정 일자"` column |
| `table_header` | Validate specific table headers | Table 2 must not contain `"Test"` |
| `required_section` | Confirm mandatory sections exist | `Revision History`, signatures, abbreviation definitions |
| `format_check` | Validate formatting (font, spacing) | Body font must be 11pt |

### Current rule catalog

18 rules organized by ID prefix:

- `RULE-MVR-0XX` — terminology / spacing
- `RULE-MVR-1XX` — table headers
- `RULE-MVR-2XX` — required sections
- `RULE-MVR-3XX` — formatting
- `RULE-MVR-4XX` — cross-references (placeholder for v2)

See `MVR_rules.yaml` for the full catalog and `RULE_SCHEMA.md` for the rule schema.

---

## Usage

### Install

```bash
pip install -r requirements.txt
```

Dependencies: `pyyaml`, `python-docx`, `lxml`.

### Validate the rule catalog

```bash
python validate_rules.py
```

Confirms the rule YAML satisfies the schema and prints rule statistics.

### Generate a sample MVR with intentional violations

```bash
python build_fake_mvr.py
```

Produces `sample_MVR_with_violations.docx` containing all 12 rule violation patterns.

### Review a document

```python
from review_engine import review_document
from exporters import export_all

violations, _ = review_document("your_MVR.docx", "MVR_rules.yaml")
export_all(violations, "your_MVR.docx", "output/")
```

This produces three files in `output/`:
- `your_MVR_review.json`
- `your_MVR_review.md`
- `your_MVR_reviewed.docx` (open in Microsoft Word to see comments)

---

## Repository structure

```
mvr-reviewer/
├── MVR_rules.yaml              # Rule catalog (the central asset)
├── RULE_SCHEMA.md              # Rule schema documentation
├── review_engine.py            # Document parser + rule application
├── exporters.py                # JSON / Markdown / DOCX-with-comments output
├── validate_rules.py           # Rule catalog self-validation
├── build_fake_mvr.py           # Test fixture generator
├── samples/
│   └── sample_MVR_with_violations.docx
└── output/                     # Example outputs (committed for demo)
```

## Adding a new rule

1. Pick the next free ID in the appropriate range (e.g., `RULE-MVR-004` for terminology).
2. Add the rule to `MVR_rules.yaml` following the schema in `RULE_SCHEMA.md`.
3. Run `python validate_rules.py` to confirm the catalog is valid.
4. Run the engine on a sample document to confirm the rule fires correctly.

No code changes are required for `regex_replace`, `regex_find`, `table_header`, `required_section`, or `format_check` rules. New `check_type` values require adding a checker function to `review_engine.py`.

## Limitations and roadmap

**Current limitations**
- Format checks for `line_spacing`, `paragraph_spacing`, and `alignment` are flagged for manual confirmation rather than automated detection (extracting these reliably from `python-docx` is brittle).
- Cross-reference validation (e.g., "every table referenced in body actually exists") is not yet implemented.
- Comments are inserted at paragraph level; cell-level comments inside tables are not yet supported.

**Roadmap**
- v2: full format-check automation, cell-level table comments, cross-reference validation.
- Future: companion repositories for SOP, MVP, TTP, TTR, specifications, and stability documents, integrated under a common desktop shell.

## License

MIT. Rule content reflects internal conventions of one GMP site and is provided for illustrative and methodological purposes; adapt rules to your own organizational standards before production use.

## Citation

If this tool informs academic work, please cite the broader research program on Regulatory HCI (citation forthcoming).
