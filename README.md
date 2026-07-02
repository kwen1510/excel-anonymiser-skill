# Excel Anonymiser Skill

## Install

Clone the repository:

```bash
git clone https://github.com/kwen1510/excel-anonymiser-skill.git
cd excel-anonymiser-skill
```

Install Python dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Verify the handler and skill package:

```bash
python3 -m pytest
python3 privacy_handler/excel_privacy_tool.py verify-self
python3 /Users/etdadmin/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/excel-privacy-handler
```

Optional: install the bundled handler template into another project:

```bash
python3 skills/excel-privacy-handler/scripts/install_template.py --target /path/to/project
```

Use `--force` only when intentionally replacing an existing scaffold:

```bash
python3 skills/excel-privacy-handler/scripts/install_template.py --target /path/to/project --force
```

## What This Is

This repository contains two pieces:

- A locked local Excel privacy handler in `privacy_handler/`.
- A reusable Codex skill package in `skills/excel-privacy-handler/`.

The handler lets Codex help with Excel workflows without reading the real workbook values. It does this by creating a safe anonymised workbook for Codex and a separate private key file that only the local user should control.

## How It Works

1. Put the real workbook in `original_data/`.
2. Run `inspect`. This reads only workbook structure and the configured header row, then prints safe metadata: filename, sheet names, row counts, column counts, and headers.
3. Edit `configs/privacy_config.json`. The config names sheets, column headers, anonymisation modes, and output folders. It must not contain real cell values or passwords.
4. Run `anonymise`. The handler replaces selected column values with safe IDs such as `NAME_550e8400`, writes an anonymised workbook to `safe_for_codex/`, writes a safe manifest to `safe_for_codex/`, and writes the private mapping key to `private_keys/`.
5. Codex works only on the anonymised workbook and manifest. It can edit, process, or deduplicate the safe workbook without seeing original values.
6. Run `restore` locally when finished. The handler uses the private key file to put original values back into a restored workbook under `restored_outputs/`.

The important separation is:

- `safe_for_codex/` contains files Codex may inspect.
- `private_keys/`, `original_data/`, and `restored_outputs/` contain files Codex must not inspect.

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

The handler must never print original cell values, mapping pairs, restored workbook contents, sample rows, unique original values, or frequency counts of original values.

## Basic Workflow

Inspect safe workbook metadata:

```bash
python3 privacy_handler/excel_privacy_tool.py inspect original_data/original.xlsx --header-row 1
```

Validate the config:

```bash
python3 privacy_handler/excel_privacy_tool.py validate-config configs/privacy_config.json
```

Create the anonymised workbook:

```bash
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

## Anonymisation Modes

The config supports these column modes:

- `keep`: keep the value unchanged in the anonymised workbook.
- `uuid`: replace each unique value with a random prefixed ID.
- `stable_hash`: replace each unique value with a salted deterministic hash.
- `encrypt`: replace each value with an encrypted token.
- `blank`: remove the value in the anonymised workbook.

For most identifying columns, use `uuid`.

## Password-Protected Workbooks

If the workbook is password protected, the handler shows a local password popup. Do not paste passwords into Codex chat, terminal commands, config files, or source files.

Password-protected workbook support uses `msoffcrypto-tool`. The password is held only in memory for local decryption and is not written to logs, configs, manifests, key files, or terminal output.

## Commands

```bash
python3 privacy_handler/excel_privacy_tool.py inspect <input.xlsx> [--header-row 1]
python3 privacy_handler/excel_privacy_tool.py anonymise <input.xlsx> --config <config.json>
python3 privacy_handler/excel_privacy_tool.py dedupe <input.xlsx> --config <config.json>
python3 privacy_handler/excel_privacy_tool.py restore <input.xlsx> --config <config.json>
python3 privacy_handler/excel_privacy_tool.py checksum
python3 privacy_handler/excel_privacy_tool.py write-approved-checksums
python3 privacy_handler/excel_privacy_tool.py verify-self
python3 privacy_handler/excel_privacy_tool.py validate-config <config.json>
```

## Codex Skill

The reusable skill package is in:

```text
skills/excel-privacy-handler/
```

The skill contains:

- `SKILL.md`: instructions Codex should follow.
- `scripts/install_template.py`: installer for copying the handler scaffold into another project.
- `assets/template/`: a bundled copy of the handler scaffold.

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
