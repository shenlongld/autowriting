# AGENTS.md

## Cursor Cloud specific instructions

### Product overview

This repository is a **content-generation workspace** for academic writing (not a web application). It has two workflows:

1. **Digital-human report** — run `build_digital_human_report.py` to generate a formatted Word document on digital-human technology, ethics, and governance.
2. **Eco speech / presentation** — driven by `prompt.txt`, `background.txt`, and files under `writing_norm/` and `support/`; output is LLM-assisted and not automated by code in this repo.

There are no long-running services, databases, Docker containers, or HTTP servers.

### Prerequisites

- Python 3.12+ (available as `python3` on the VM)
- `python-docx` (installed via the VM update script)

### Running the report generator

```bash
cd /workspace
python3 build_digital_human_report.py
```

**Output:** `数字人领域的技术演进_伦理风险与治理路径研究_修订版.docx` in the repo root (~56 KB).

The script runs built-in validation after generation: it rejects banned transition words and subsections over 800 characters. A successful run with exit code 0 means validation passed.

### Lint and tests

- No ESLint, Ruff, pytest, or CI configuration is defined in this repo.
- Use `python3 -m py_compile build_digital_human_report.py` for a basic syntax check.
- The report script's `validate_text()` function is the primary automated quality gate.

### Environment variables

None required. No `.env` file or API keys are referenced in the codebase.

### Gotchas

- `大作业.docx` is defined as `BASE` in the script but is unused; the script creates a fresh `Document()` from scratch.
- `example.txt` (speech workflow output) is referenced in `prompt.txt` but is not present in the repo.
- Re-running the generator overwrites the output `.docx` in the repo root.
