# Excel Anonymiser Skill

This repository contains a locked, local-first Excel privacy handler and a Codex skill package for installing and using it safely.

The handler lets Codex work with Excel files without reading original workbook cell values. Codex may inspect only workbook metadata, configs, anonymised IDs, and safe manifest files.

## Privacy Boundary

Codex may see:
- `privacy_handler` source before approval
- `configs/*.json`
- inspect output
- `safe_for_codex/*.anonymised.xlsx`
- `safe_for_codex/*.privacy_manifest.json`

Codex must not see:
- `original_data/*.xlsx`
- `private_keys/*.privacy_key.xlsx`
- `private_keys/*.secret.key`
- `restored_outputs/*.xlsx`
- workbook passwords
- restored workbook contents
- sample rows or raw cell values

## Install Dependencies

```bash
python3 -m pip install -r requirements.txt
```

## Verify

```bash
python3 -m pytest
python3 privacy_handler/excel_privacy_tool.py verify-self
python3 /Users/etdadmin/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/excel-privacy-handler
```

## Use The Handler

Place the original workbook in `original_data/`, then inspect safe metadata:

```bash
python3 privacy_handler/excel_privacy_tool.py inspect original_data/original.xlsx --header-row 1
```

Edit `configs/privacy_config.json` using only sheet names, column headers, modes, and output paths.

Validate and anonymise:

```bash
python3 privacy_handler/excel_privacy_tool.py validate-config configs/privacy_config.json
python3 privacy_handler/excel_privacy_tool.py anonymise original_data/original.xlsx --config configs/privacy_config.json
```

Codex should then work only with:

```text
safe_for_codex/original.anonymised.xlsx
safe_for_codex/original.privacy_manifest.json
```

Deduplicate an anonymised workbook:

```bash
python3 privacy_handler/excel_privacy_tool.py dedupe safe_for_codex/processed.xlsx --config configs/dedupe_config.example.json
```

Restore after processing:

```bash
python3 privacy_handler/excel_privacy_tool.py restore safe_for_codex/processed.deduped.xlsx --config configs/restore_config.example.json
```

## Password-Protected Workbooks

If the workbook is password protected, the handler shows a local password popup. Do not paste passwords into Codex chat, terminal commands, config files, or source files.

Password-protected workbook support uses `msoffcrypto-tool`; the password is held only in memory for local decryption.

## Codex Skill

The reusable skill package is in:

```text
skills/excel-privacy-handler/
```

To install the handler template into another project:

```bash
python3 skills/excel-privacy-handler/scripts/install_template.py --target /path/to/project
```

Use `--force` only when intentionally replacing an existing scaffold.

## Locked Mode Instruction

After approval:

```text
Do not edit privacy_handler/*.
Do not open original_data/*.
Do not open private_keys/*.
Do not open restored_outputs/* unless explicitly allowed.
Do not ask the user to paste workbook passwords into chat, terminal commands, configs, or source files.
Only run documented privacy_handler commands and edit configs/*.json.
Treat privacy_handler as a locked executable.
```
