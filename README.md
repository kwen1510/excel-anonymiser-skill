# Excel Anonymiser Skill

```text
Use this GitHub repository as a Codex skill and local Excel privacy tool: https://github.com/kwen1510/excel-anonymiser-skill. Clone it into the current project folder, install the requirements, run the tests and verify-self check, then help me inspect an Excel workbook safely. First run the inspect command to list only sheet names, headers, row counts, and column counts. Then ask me which columns I want to mask. After I answer, run the mask command to create a masked workbook in safe_for_codex/ plus a private encrypted CSV key and secret key in private_keys/. Work only on files in safe_for_codex/. Never read, print, summarize, preview, or inspect original cell values, private key CSV contents, secret keys, restored outputs, or workbook passwords. If the workbook is password protected, let the local popup ask me for the password; do not ask me to paste it into chat, config, or terminal commands. When I ask to unmask, run the unmask command with the private key files and write the result to restored_outputs/.
```

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

## What This Is

This is a deliberately simple local Excel masking tool for Codex workflows.

The idea is:

1. Codex can inspect the workbook structure safely.
2. The user chooses which columns to mask.
3. The tool replaces those values with random gibberish IDs.
4. The tool writes a private encrypted CSV key that can map the gibberish IDs back later.
5. Codex works only on the masked workbook.
6. The user can unmask the workbook at the end.

## How It Works

Put the real workbook in `original_data/`.

Run `inspect`:

```bash
python3 privacy_handler/excel_privacy_tool.py inspect original_data/original.xlsx --header-row 1
```

This prints only safe metadata:

- workbook filename
- sheet names
- headers
- row counts
- column counts

It does not print sample rows or cell values.

After seeing the headers, choose the columns to mask. Then run:

```bash
python3 privacy_handler/excel_privacy_tool.py mask original_data/original.xlsx --sheet Responses --columns Name,Email,Class
```

This creates:

```text
safe_for_codex/original.masked.xlsx
private_keys/original.mask_key.csv
private_keys/original.mask_secret.key
```

The masked workbook contains random IDs like:

```text
MASK_8f4a91c2e3b7
```

These IDs are random and not reversible by themselves. The private CSV key stores the mapping, and the original values inside that CSV are encrypted with the secret key file.

Codex should then work only with:

```text
safe_for_codex/original.masked.xlsx
```

When the work is finished, unmask:

```bash
python3 privacy_handler/excel_privacy_tool.py unmask safe_for_codex/original.masked.xlsx --key-file private_keys/original.mask_key.csv --secret-key-file private_keys/original.mask_secret.key --output restored_outputs/original.unmasked.xlsx
```

If rows were deleted or deduplicated while working on the masked workbook, unmasking still works for the remaining masked IDs.

## Password-Protected Workbooks

If the workbook is not password protected, the commands just run.

If the workbook is password protected, the tool shows a local password popup. The password is held only in memory for local decryption.

Do not paste workbook passwords into:

- Codex chat
- terminal commands
- config files
- source files

## Privacy Boundary

Codex may see:

- inspect output
- `safe_for_codex/*.masked.xlsx`
- `safe_for_codex/*.privacy_manifest.json`
- config files that contain only sheet names, headers, modes, and paths

Codex must not see:

- `original_data/*.xlsx`
- `private_keys/*.mask_key.csv`
- `private_keys/*.mask_secret.key`
- `private_keys/*.privacy_key.xlsx`
- `restored_outputs/*.xlsx`
- workbook passwords
- sample rows or raw cell values

## Commands

Simple workflow:

```bash
python3 privacy_handler/excel_privacy_tool.py inspect <input.xlsx> [--header-row 1]
python3 privacy_handler/excel_privacy_tool.py mask <input.xlsx> --columns Name,Email [--sheet Responses]
python3 privacy_handler/excel_privacy_tool.py unmask <masked.xlsx> --key-file <mask_key.csv> --secret-key-file <mask_secret.key> --output <output.xlsx>
```

Advanced workflow remains available:

```bash
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
Do not edit privacy_handler/* unless the user explicitly asks to change the tool.
Do not open original_data/*.
Do not open private_keys/*.
Do not open restored_outputs/* unless explicitly allowed.
Do not ask the user to paste workbook passwords into chat, terminal commands, configs, or source files.
Only run documented privacy_handler commands and edit configs/*.json.
Treat privacy_handler as a locked executable.
```
