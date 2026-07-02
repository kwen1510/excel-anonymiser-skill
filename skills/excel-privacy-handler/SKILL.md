---
name: excel-privacy-handler
description: Create, install, verify, or use a locked local-first Excel privacy handler for .xlsx workbooks. Use when Codex must inspect only safe workbook metadata, anonymise selected Excel columns, deduplicate anonymised workbooks, restore with a private key, handle password-protected Excel files via a local popup, package the handler for reuse, or enforce a workflow where Codex must never read original cell values, privacy mappings, passwords, restored outputs, or real workbook contents.
---

# Excel Privacy Handler

## Core Rule

Never read, print, preview, summarize, log, or inspect original workbook cell values. Do not ask the user to paste workbook passwords into chat, terminal commands, configs, or source files.

Codex may see only safe metadata: workbook filename, sheet names, headers, row/column counts, anonymised IDs, manifests, and configs. Treat `privacy_handler` as a locked executable once `verify-self` passes against approved checksums.

## Installing The Template

To scaffold the handler into a project, run the bundled installer from this skill:

```bash
python3 <skill-dir>/scripts/install_template.py --target <project-dir>
```

Use `--force` only when intentionally replacing an existing scaffold:

```bash
python3 <skill-dir>/scripts/install_template.py --target <project-dir> --force
```

After installing dependencies, verify the scaffold:

```bash
python3 -m pip install -r requirements.txt
python3 -m pytest
python3 privacy_handler/excel_privacy_tool.py verify-self
```

## Safe Workflow

Use only documented commands:

```bash
python3 privacy_handler/excel_privacy_tool.py inspect original_data/original.xlsx --header-row 1
python3 privacy_handler/excel_privacy_tool.py validate-config configs/privacy_config.json
python3 privacy_handler/excel_privacy_tool.py anonymise original_data/original.xlsx --config configs/privacy_config.json
python3 privacy_handler/excel_privacy_tool.py dedupe safe_for_codex/processed.xlsx --config configs/dedupe_config.example.json
python3 privacy_handler/excel_privacy_tool.py restore safe_for_codex/processed.deduped.xlsx --config configs/restore_config.example.json
python3 privacy_handler/excel_privacy_tool.py checksum
python3 privacy_handler/excel_privacy_tool.py write-approved-checksums
python3 privacy_handler/excel_privacy_tool.py verify-self
```

For password-protected workbooks, let the handler show the local password popup. Do not capture or relay the password through Codex-visible channels.

## Locked-Mode Conduct

After approval:

- Do not edit `privacy_handler/*`.
- Do not open `original_data/*`.
- Do not open `private_keys/*`.
- Do not open `restored_outputs/*` unless explicitly allowed by the user.
- Do not write custom Python, notebooks, `python -c`, pandas, or openpyxl inspection scripts to read workbook values.
- Edit only `configs/*.json` and inspect only safe outputs under `safe_for_codex/`.
- Run `verify-self` before `anonymise`, `dedupe`, or `restore`.

## GitHub Packaging Checks

Before push, run:

```bash
python3 -m pytest
python3 privacy_handler/excel_privacy_tool.py verify-self
python3 /Users/etdadmin/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/excel-privacy-handler
git status --short
```

Do not commit real files from `original_data/`, `private_keys/`, `safe_for_codex/`, or `restored_outputs/`; only their `.gitkeep` placeholders belong in Git.
