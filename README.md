# Purchase Order Workbook

Generate monthly spending, recurring purchase, and task order summaries from an Excel purchase order export. The tool runs locally and preserves the source worksheets in the output.

## Usage

Install Python 3.10 or later and run:

```shell
python -m pip install openpyxl
python purchasing_orders_report.py orders.xlsx --output purchasing_analysis.xlsx
```

The first worksheet must contain the columns `Request #`, `Date Created`, and `Total $`. Include `Vendor`, `Line Item Summary`, and `Status` to populate the detail views. Set `Account Numbers` to values like `TO-12` or `TO-12, TO-14` to group orders by task order. When multiple task orders appear on one row, the report estimates equal shares.

Open the output in a spreadsheet application to calculate the summary formulas.
