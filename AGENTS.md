# AGENTS.md

## Project overview

This repository is a **content and prompt workspace** with two deliverable tracks:

| Track | Entry point | Output |
|-------|-------------|--------|
| **Grand official speech** (Qinghai green development) | `prompt.txt`, `background.txt`, `writing_norm/`, `support/` | `example.txt` (~10 min speech) |
| **Digital human ethics report** | `build_digital_human_report.py` | `数字人领域的技术演进_伦理风险与治理路径研究_修订版.docx` |

There is no web server, database, Docker stack, or Node toolchain.

## Standard workflows

### Speech

1. Read `prompt.txt`, `background.txt`, and `writing_norm/norm.txt`.
2. Optionally use `support/*.txt` without drifting from the background topic.
3. Write or revise `example.txt` per rhetorical rules in `prompt.txt`.
4. Self-check against `writing_norm/delivery.txt` dates and formats.

### Report generator

```bash
pip install -r requirements.txt
python build_digital_human_report.py
```

The script builds a Word document from scratch (no input template required on `main`), runs built-in validation (banned wording, subsection length ≤ 800 chars), and writes the output `.docx` in the repo root.

## Lint / test / build / run

| Action | Command |
|--------|---------|
| Install deps | `pip install -r requirements.txt` |
| Lint | Manual review against `prompt.txt` and `writing_norm/`; report script enforces `BANNED` words on build |
| Test (report) | `python build_digital_human_report.py` (exits non-zero if validation fails) |
| Test (speech) | Rehearse timing (~10 min); verify facts against `background.txt` |
| Build / run | No compiler or dev server |

## Cursor Cloud specific instructions

- **Dependencies:** `python-docx` only (see `requirements.txt`). Python 3.12+ is sufficient.
- **No long-running services** — no API, DB, or frontend to start.
- **Report E2E:** run `python build_digital_human_report.py` and confirm the output `.docx` exists (~55 KB, 100+ paragraphs).
- **Speech E2E:** produce or edit `example.txt` using source files under `writing_norm/` and `support/`.
- **Generated artifacts:** the report `.docx` is build output; do not treat it as a required committed file unless the user asks.
- **External research:** `prompt.txt` allows supplementing from the web; network is optional for basic setup.
- **Git:** no active pre-commit hooks in this repo (only default `.git/hooks/*.sample`).
