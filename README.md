# Excel Anonymiser Skill

Please paste the following into Codex:

```text
You are helping me use the Excel Anonymiser Skill from this GitHub repository:
https://github.com/kwen1510/excel-anonymiser-skill

Goal:
Set up a local Excel privacy workflow where you can help me work with Excel files without ever reading the original cell values.

Do this step by step:

1. Clone or use the repository in the current project folder.
   If it is not already present, run:
   git clone https://github.com/kwen1510/excel-anonymiser-skill.git

2. Enter the repository folder and install requirements:
   python3 -m pip install -r requirements.txt

3. Verify the tool:
   python3 -m pytest
   python3 privacy_handler/excel_privacy_tool.py verify-self

4. Tell me to put my real Excel workbook inside:
   original_data/

5. Once the workbook is there, inspect it safely with:
   python3 privacy_handler/excel_privacy_tool.py inspect original_data/<file.xlsx> --header-row 1

Important:
The inspect command may show only:
- workbook filename
- sheet names
- headers
- row counts
- column counts

It must not show sample rows or original cell values.

6. After inspect, show me the sheet names and headers from the safe output, then ask:
   "Which columns do you want to mask?"

7. After I choose the columns, run the simple mask command:
   python3 privacy_handler/excel_privacy_tool.py mask original_data/<file.xlsx> --sheet <sheet_name> --columns Column1,Column2,Column3

This creates:
- safe_for_codex/<file>.masked.xlsx
- private_keys/<file>.mask_key.csv
- private_keys/<file>.mask_secret.key

The masked workbook contains random gibberish IDs such as MASK_8f4a91c2e3b7. These are not reversible by themselves. The private CSV key plus secret key are needed to unmask later.

8. From this point onward, work only with:
   safe_for_codex/<file>.masked.xlsx

Never open or inspect:
- original_data/*.xlsx
- private_keys/*.mask_key.csv
- private_keys/*.mask_secret.key
- restored_outputs/*.xlsx

Never read, print, summarize, preview, log, or infer:
- original cell values
- private mapping key contents
- restored workbook contents
- workbook passwords
- sample rows
- unique original values
- frequency counts of original values

9. If the workbook is password protected, let the local popup ask me for the password.
Do not ask me to paste the password into Codex chat, terminal commands, config files, or source files.

10. When I ask to unmask, run:
    python3 privacy_handler/excel_privacy_tool.py unmask safe_for_codex/<file>.masked.xlsx --key-file private_keys/<file>.mask_key.csv --secret-key-file private_keys/<file>.mask_secret.key --output restored_outputs/<file>.unmasked.xlsx

At the end, tell me where the masked workbook is, where the private key files are, and how to run unmask. Keep the workflow simple.
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
