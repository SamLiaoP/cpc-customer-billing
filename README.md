# cpc-customer-billing

Generate monthly customer billing statements (請款對帳單) for a CPC gas station from the
credit-account exports (簽帳客戶對帳單) of the CPC system. Excel in, Excel out.

## Usage

```bash
python3 點兩下產生請款單.py              # 2026-01 plus previous, current and next month
python3 點兩下產生請款單.py 2026-01      # one or more specific months
python3 點兩下產生請款單.py 2026-01 --no-overwrite   # keep an existing 01月.xlsx
python3 點兩下產生請款單.py --no-pull    # skip the automatic git pull
python3 點兩下產生請款單.py --no-pause   # Windows: do not wait for Enter before closing
```

Every run first does `git pull` so the script is always the latest version; if the script
itself was updated it restarts with the new version. On Windows the console waits for Enter
before closing so the messages stay visible.

## Folder layout

```
cpc-customer-billing/
├── 點兩下產生請款單.py
├── 客戶折讓表.xlsx            discount rule per customer (not committed)
└── 2026-01/                   one folder per month (not committed)
    ├── 輸入/A002-1.xlsx       CPC exports, file name starts with the customer code
    └── 01月.xlsx              generated output
```

## Output workbook

- `本月總計` – per customer: litres / amount / discount by product (92, 95, diesel, urea, other)
- `簽收單` – sign-off sheet
- `<code><name>-1/-2/-3` – statement sorted by time, grouped by plate, grouped by product
  (a grouping sheet is skipped when there is only one plate or one product)

Customers without an input file are not generated. Reversal pairs (a sale and its negative
cancellation) are removed automatically. Customers whose discount note contains 「結餘款」 get a
balance block (尚餘 / 匯入 / 油款 / 餘額) to fill in by hand.

## Requirements

Python 3.9+ and `openpyxl` (`pip install openpyxl`).
