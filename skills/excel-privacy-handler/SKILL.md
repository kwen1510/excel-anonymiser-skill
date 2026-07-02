---
name: excel-privacy-handler
description: Install or use a simple local Excel masking workflow for Codex. Use when a user wants Codex to inspect only Excel headers/metadata, ask which columns to mask, create a masked .xlsx using random non-reversible IDs, keep a private encrypted CSV key for later unmasking, handle password-protected workbooks with a local popup, and prevent Codex from reading original cell values, private keys, restored outputs, or workbook passwords.
---

# Excel Privacy Handler

## Core Rule

Never read, print, preview, summarize, log, or inspect original workbook cell values. Do not ask the user to paste workbook passwords into chat, terminal commands, configs, or source files.

Codex may inspect only safe metadata and files under `safe_for_codex/`.

## Simple Workflow

1. Put the real workbook in `original_data/`.
2. Run `inspect` to get only sheet names, headers, row counts, and column counts:

```bash
python3 privacy_handler/excel_privacy_tool.py inspect original_data/original.xlsx --header-row 1
```

3. Ask the user which columns to mask, using the headers from inspect output.
4. Run `mask`:

```bash
python3 privacy_handler/excel_privacy_tool.py mask original_data/original.xlsx --sheet Responses --columns Name,Email,Class
```

5. Work only on the masked workbook in `safe_for_codex/`.
6. When the user asks to unmask, run:

```bash
python3 privacy_handler/excel_privacy_tool.py unmask safe_for_codex/original.masked.xlsx --key-file private_keys/original.mask_key.csv --secret-key-file private_keys/original.mask_secret.key --output restored_outputs/original.unmasked.xlsx
```

## What Masking Creates

`mask` creates:

- `safe_for_codex/<name>.masked.xlsx`
- `private_keys/<name>.mask_key.csv`
- `private_keys/<name>.mask_secret.key`

Masked values look like `MASK_8f4a91c2e3b7`. They are random IDs, not hashes. The private CSV key maps masked IDs back to encrypted originals, and the secret key decrypts them during `unmask`.

If rows are deleted or deduplicated in the masked workbook, `unmask` still replaces any remaining masked IDs it finds.

## Password-Protected Workbooks

If the workbook is password protected, the handler shows a local password popup. Let the user type into that popup. Never request the password in chat or command arguments.

## Installing The Template

To scaffold the handler into a project, run:

```bash
python3 <skill-dir>/scripts/install_template.py --target <project-dir>
```

Use `--force` only when intentionally replacing an existing scaffold.

After installing dependencies:

```bash
python3 -m pip install -r requirements.txt
python3 -m pytest
python3 privacy_handler/excel_privacy_tool.py verify-self
```

## Locked Conduct

- Do not edit `privacy_handler/*` unless the user explicitly asks to change the tool.
- Do not open `original_data/*`.
- Do not open `private_keys/*`.
- Do not open `restored_outputs/*` unless explicitly allowed.
- Do not write custom Python, notebooks, `python -c`, pandas, or openpyxl inspection scripts to read workbook values.
- Run `verify-self` before `mask`, `unmask`, `anonymise`, `dedupe`, or `restore`.
