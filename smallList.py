import datetime as dt
from pathlib import Path

import pandas as pd
from xlsxwriter import Workbook
from xlsxwriter.utility import xl_rowcol_to_cell

from member import getMemberInfo


def makeBlancDataFrame() -> pd.DataFrame:
    """Makes a blanc dataframe, ready to be exported to excel."""
    memberinfo = getMemberInfo()
    commissielist = memberinfo[["clublidnummer", "voornaam", "tussenvoegsel", "achternaam"]]
    withEmpty = commissielist.assign(Totaal="", Bedrag_a=0, Bedrag_b=0, Bedrag_c=0)
    return withEmpty

def exportBlancListToExcel(df: pd.DataFrame = makeBlancDataFrame()):
    """Exports a blanc list to excel in the same directory."""
    # Setup filename
    x = dt.datetime.now()
    name = f"Empty Incasso List [{x.strftime('%a %d %b, %H-%M-%S')}].xlsx"
    writer = pd.ExcelWriter(name, engine='xlsxwriter')
    sheetName = "Lege incassolijst"
    workbook = writer.book
    if workbook is None:
        return
    
    # Export to excel
    df.to_excel(writer, index=False, startrow=1, sheet_name=sheetName)
    sheet: Workbook.worksheet_class = writer.sheets[sheetName]

    # Write headers
    sheet.write("A1", f"Created at {x.strftime('%a %d %b, %H:%M:%S')}")
    totalColumn = 4
    sheet.write(0, totalColumn, "Totaal")

    # Write header sums
    for i in range(totalColumn, len(df.columns)):
        formula = f"=SUM({xl_rowcol_to_cell(2, i)}:{xl_rowcol_to_cell(1000, i)})"
        sheet.write_formula(0, i, formula=formula)

    # Write sums per row
    for i in range(2, len(df.index) + 2):
        formula = f"=SUM({xl_rowcol_to_cell(i, totalColumn + 1)}:{xl_rowcol_to_cell(i, len(df.columns)-1)})"
        sheet.write_formula(i, totalColumn, formula=formula)

    writer.close()

def readSmallExcel(filename: Path) -> pd.DataFrame:
    df = pd.read_excel(filename, header=1)
    df = df[df.Totaal != 0]
    df = df.drop(columns=["voornaam", "tussenvoegsel", "achternaam", "Totaal"])
    df = df[df.columns[df.sum()!=0]]
    return df

# exportBlancListToExcel(makeBlancDataFrame())