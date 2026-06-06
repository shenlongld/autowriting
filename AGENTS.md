# AGENTS.md

## Project overview

This repository is a **content and prompt workspace** for auto-writing course deliverables (grand official speech, PPT narrative, video script) about China's ecological progress, centered on Qinghai's **送绿水 / 送绿电 / 送绿算** story. There is **no application runtime**, package manager, or CI toolchain—only UTF-8 text sources.

| Path | Role |
|------|------|
| `prompt.txt` | Master instructions for the speech deliverable |
| `background.txt` | Primary source material (Xinhua article) |
| `writing_norm/` | Style, structure, and delivery deadlines |
| `support/` | Optional supplementary talking points |
| `example.txt` | Output: ~10-minute grand official speech |

## Standard workflow

1. Read `prompt.txt`, `background.txt`, and `writing_norm/norm.txt`.
2. Optionally use `support/*.txt` without drifting from the background topic.
3. Write or revise the deliverable in `example.txt` per rhetorical rules in `prompt.txt` (gravitas, anaphora/tricolons; no essay markers like “Firstly/Secondly”).
4. Self-check against `writing_norm/delivery.txt` dates and formats.

## Lint / test / build / run

| Action | Command / approach |
|--------|-------------------|
| Lint | Manual review against `prompt.txt` and `writing_norm/` |
| Test | Rehearse timing (~10 min); verify facts against `background.txt` |
| Build | N/A (no compiler) |
| Run | N/A (no dev server) |

There are **no** `npm`, `pip`, `docker compose`, or Makefile targets in this repo.

## Cursor Cloud specific instructions

- **No dependency install** is required on VM startup; the update script is a no-op sanity check only.
- **No services** to start (no API, DB, or frontend). “E2E” for this repo means producing a valid `example.txt`.
- **Primary agent task:** generate or edit `example.txt` using `background.txt` + `writing_norm/` + optional `support/`.
- **External research:** `prompt.txt` allows supplementing from the web (e.g. the Xinhua URL at the top of `background.txt`); network is optional, not required for basic setup.
- **Git:** normal `git add` / `commit` / `push` for content changes; no pre-commit hooks are configured in the repository.
