# RMA Management GAS

This repository is managed with clasp and pushes to the Apps Script project below.

## Target IDs

- Spreadsheet ID: `1uGWkEmfg0B00mGANWSVLLMgMUWKEiKIbHncDZhRzBzs`
- Apps Script project ID: `1RWF1L7jGGXs28r32oP6Ae65ziOwTglOC0v5hSBoktSy5vHx3UyQtp6h9`
- RMA Registration Form template document ID: `1CSSxjyQbFN4HYw4uHJnPS9NqOTRvIsc9`

## Sheets

- `__init`: Common entry information. Column A is `Field`, column B is `Value`.
- `NewEntry`: Returned goods input only. Five rows are prepared for `Goods 1` through `Goods 5`.
- `RMA Ledger`: Generated entry ledger.

Common information now lives in `__init`, including applicant, return address, delivery note, invoice, project, reason, and remarks.

`NewEntry` uses this table layout:

```text
Goods No | TLV Code | Part Number | Description | Serial Number (batch No. + Serial No.) | Quantity | Error description
Goods 1  |          |             |             |                                        |          |
Goods 2  |          |             |             |                                        |          |
Goods 3  |          |             |             |                                        |          |
Goods 4  |          |             |             |                                        |          |
Goods 5  |          |             |             |                                        |          |
```

## Workflow

1. Run `clasp push` from this folder.
2. Reload the target spreadsheet.
3. Run `Jitbit > Setup` once to create or refresh the sheets.
4. Fill common values in `__init`.
5. Fill returned goods rows in `NewEntry`.
6. Run `Jitbit > Generate form`.

The script issues an entry ID, appends a row to `RMA Ledger`, copies the RMA Registration Form template, fills known fields, appends generated entry data, and stores the generated document URL in the ledger.

## Commands

```powershell
clasp status
clasp push
```

## Notes

- Re-running setup preserves existing `__init` values.
- Re-running setup clears and rebuilds `NewEntry`.
- Generated documents are stored in the Google Drive folder `RMA Generated Forms`.
- If Google asks for authorization on first run, review the requested permissions and allow the Apps Script project.
