"""Sample project: Excel report automation (pandas + openpyxl).

Takes a messy raw sales export (messy_sales.xlsx) and produces a clean,
formatted monthly report (clean_report.xlsx):
  - normalizes inconsistent date formats
  - strips whitespace / fixes casing in text columns
  - drops exact duplicates, fills missing regions with 'Unknown'
  - adds a Total column, a summary sheet, styled headers and totals row
"""
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "messy_sales.xlsx"
DST = "clean_report.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F4E5F")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TOTAL_FILL = PatternFill("solid", fgColor="E8F1F5")
TOTAL_FONT = Font(bold=True, size=11)
THIN = Border(*[Side(style="thin", color="B0BEC5")] * 4)


def clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # normalize text columns
    for col in ["Product", "Region", "Salesperson"]:
        df[col] = df[col].astype(str).str.strip().str.title()
    df["Region"] = df["Region"].replace({"Nan": "Unknown", "None": "Unknown", "": "Unknown"})
    # normalize the two date formats used in the raw export
    df["Date"] = pd.to_datetime(df["Date"], format="mixed", dayfirst=False)
    # drop exact duplicates, keep first
    df = df.drop_duplicates()
    # money columns
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0).astype(int)
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="coerce").fillna(0)
    df["Total"] = (df["Qty"] * df["UnitPrice"]).round(2)
    return df.sort_values("Date").reset_index(drop=True)


def style_sheet(ws, nrows, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(1, nrows + 2):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = THIN
    for c in range(1, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16
    # totals row
    tr = nrows + 2
    ws.cell(row=tr, column=1, value="TOTAL").font = TOTAL_FONT
    for c in range(1, ncols + 1):
        ws.cell(row=tr, column=c).fill = TOTAL_FILL
        ws.cell(row=tr, column=c).font = TOTAL_FONT
        ws.cell(row=tr, column=c).border = THIN


def main():
    raw = pd.read_excel(SRC)
    print(f"Raw rows: {len(raw)}")
    df = clean(raw)
    print(f"Clean rows: {len(df)} (removed {len(raw) - len(df)} duplicates/bad rows)")

    summary = (
        df.groupby("Region", as_index=False)
        .agg(Orders=("Total", "count"), Revenue=("Total", "sum"))
        .sort_values("Revenue", ascending=False)
    )

    with pd.ExcelWriter(DST, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Sales", index=False)
        summary.to_excel(w, sheet_name="Summary", index=False)
        style_sheet(w.sheets["Sales"], len(df), len(df.columns))
        style_sheet(w.sheets["Summary"], len(summary), len(summary.columns))
        # revenue total formula-free value
        w.sheets["Sales"].cell(row=len(df) + 2, column=len(df.columns), value=df["Total"].sum())
        w.sheets["Summary"].cell(row=len(summary) + 2, column=3, value=summary["Revenue"].sum())

    print(f"Saved -> {DST}  |  total revenue: ${df['Total'].sum():,.2f}")


if __name__ == "__main__":
    main()
