# Locked Local Excel Privacy Handler

This project provides a local-first Excel privacy handler. It lets Codex work with anonymised Excel files and safe metadata while keeping the original workbook values, privacy mappings, restored outputs, and passwords private.

## Privacy Model

Codex may see:
- privacy_handler source code before approval
- configs/*.json
- inspect output
- safe_for_codex/*.anonymised.xlsx
- safe_for_codex/*.privacy_manifest.json

Codex must not see:
- original_data/*.xlsx
- private_keys/*.privacy_key.xlsx
- private_keys/*.secret.key
- restored_outputs/*.xlsx

The handler never prints original cell values, privacy mappings, restored workbook contents, workbook passwords, sample rows, unique original values, or frequency counts of original values.

## Password-Protected Workbooks

If a workbook is password protected, the handler prompts the local user with a native password popup. The password must not be pasted into Codex chat, command arguments, config files, terminal scripts, or source files.

The password is held only in memory long enough to decrypt the workbook locally. Decrypted bytes are passed directly to openpyxl through an in-memory stream. The password is not logged, printed, written to manifests, written to key files, or stored in configs.

Password-protected workbook support requires `msoffcrypto-tool`.

## Recommended Workflow

```bash
python privacy_handler/excel_privacy_tool.py inspect original_data/original.xlsx --header-row 1
python privacy_handler/excel_privacy_tool.py mask original_data/original.xlsx --sheet Responses --columns Name,Email,Class
```

Then Codex works only on:

```text
safe_for_codex/original.masked.xlsx
```

After processing:

```bash
python privacy_handler/excel_privacy_tool.py unmask safe_for_codex/original.masked.xlsx --key-file private_keys/original.mask_key.csv --secret-key-file private_keys/original.mask_secret.key --output restored_outputs/original.unmasked.xlsx
```

## Locking Workflow

Before approval:

```bash
python privacy_handler/excel_privacy_tool.py checksum
python privacy_handler/excel_privacy_tool.py write-approved-checksums
python privacy_handler/excel_privacy_tool.py verify-self
```

After approval, treat `privacy_handler` as a locked executable.

Future Codex instruction:

Do not edit privacy_handler/*.
Do not open original_data/*.
Do not open private_keys/*.
Do not open restored_outputs/* unless explicitly allowed.
Do not ask the user to paste workbook passwords into chat, terminal commands, configs, or source files.
Only run documented privacy_handler commands and edit configs/*.json.
Treat privacy_handler as a locked executable.

## Commands

```bash
python privacy_handler/excel_privacy_tool.py inspect <input.xlsx> [--header-row 1]
python privacy_handler/excel_privacy_tool.py mask <input.xlsx> --columns <headers> [--sheet <sheet>]
python privacy_handler/excel_privacy_tool.py unmask <masked.xlsx> --key-file <mask_key.csv> --secret-key-file <mask_secret.key> --output <output.xlsx>
python privacy_handler/excel_privacy_tool.py anonymise <input.xlsx> --config <config.json>
python privacy_handler/excel_privacy_tool.py dedupe <input.xlsx> --config <config.json>
python privacy_handler/excel_privacy_tool.py restore <input.xlsx> --config <config.json>
python privacy_handler/excel_privacy_tool.py checksum
python privacy_handler/excel_privacy_tool.py write-approved-checksums
python privacy_handler/excel_privacy_tool.py verify-self
python privacy_handler/excel_privacy_tool.py validate-config <config.json>
```

## Modes

- `keep`: keep the value as-is in the anonymised workbook.
- `uuid`: replace each unique value with a random replacement ID.
- `stable_hash`: replace each unique value with a salted deterministic hash.
- `encrypt`: replace each value with an encrypted token.
- `blank`: remove the value in the anonymised workbook.

For most cases, use `uuid`.
