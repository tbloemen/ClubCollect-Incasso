import datetime as dt
import json
from pathlib import Path

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

from schemas import CommitteeExcel

STARTING_ROW = 3


class SmallList:
    def __init__(self, commissielist: pd.DataFrame) -> None:
        """Makes a blanc dataframe, ready to be exported to excel."""
        commissielist[CommitteeExcel.total] = ""
        commissielist[CommitteeExcel.bedrag_a] = 0
        commissielist[CommitteeExcel.bedrag_b] = 0
        commissielist[CommitteeExcel.bedrag_c] = 0
        self.committeelist = commissielist

    def exportToExcel(self):
        """Exports a blanc list to excel in the same directory."""
        with open(Path.cwd().joinpath("config.json"), "r") as config:
            data = json.load(config)

        # Setup filename
        x = dt.datetime.now()
        date_string = x.strftime('%a %d %b, %H-%M-%S')
        name = f"{data['filename']['small_incassolist']} [{date_string}].xlsx"
        writer = pd.ExcelWriter(name, engine='xlsxwriter')
        sheetName = data['filename']['small_incassolist']

        # Export to excel
        self.committeelist.to_excel(writer, index=False, startrow=STARTING_ROW -
                                    1, sheet_name=sheetName)
        sheet: xlsxwriter.Workbook.worksheet_class = writer.sheets[sheetName]

        # Write
        created_format: xlsxwriter.workbook.Format = writer.book.add_format()  # type: ignore
        created_format.set_align("justify")
        created_format.set_italic()
        sheet.merge_range(
            0, 0, 0, 2, f"Created at {date_string}", created_format)
        merged_format: xlsxwriter.workbook.Format = writer.book.add_format()  # type: ignore
        merged_format.set_text_wrap()
        sheet.merge_range(1, 0, 1, 3, data["instructions"], merged_format)
        totalColumn = 4
        totalformat: xlsxwriter.workbook.Format = writer.book.add_format()  # type: ignore
        totalformat.set_align("justify")
        sheet.write(STARTING_ROW-3, totalColumn-1, data["total"], totalformat)
        sheet.autofit()
        sheet.set_column_pixels(0, 0, 40)
        sheet.set_row_pixels(1, 80)

        # Write header sums
        for i in range(totalColumn, len(self.committeelist.columns)):
            formula = f"=SUM({xl_rowcol_to_cell(STARTING_ROW, i)}:{xl_rowcol_to_cell(STARTING_ROW + len(self.committeelist.index)-1, i)})"
            sheet.write_formula(0, i, formula=formula)

        # Write sums per row
        for i in range(STARTING_ROW, len(self.committeelist.index) + STARTING_ROW):
            formula = f"=SUM({xl_rowcol_to_cell(i, totalColumn + 1)}:{xl_rowcol_to_cell(i, len(self.committeelist.columns)-1)})"
            sheet.write_formula(i, totalColumn, formula=formula)

        writer.close()

    def readSmallExcel(self, filename: Path) -> pd.DataFrame:
        df = pd.read_excel(filename, skiprows=STARTING_ROW-1)
        df = df[df[CommitteeExcel.total] != 0]
        df = df.drop(columns=[CommitteeExcel.fname, CommitteeExcel.infix,
                     CommitteeExcel.lname, CommitteeExcel.total])
        df = df[df.columns[df.sum() != 0]]
        df = df.set_index(CommitteeExcel.id)
        return df

    def extractSmallLists(self, dir: str) -> pd.DataFrame:
        # try:
        df = pd.DataFrame()
        for x in Path.cwd().joinpath(dir).glob("*.xlsx"):
            new_df = self.readSmallExcel(x)
            if df.empty:
                df = new_df
                continue
            merged = pd.merge(df, new_df, how="outer",
                              left_index=True, right_index=True)
            df = merged
        return df
        # except:
        #     raise Exception(
        #         "There has gone something wrong with combining the small lists into one dataframe.")

    def combineSmallLists(self, dir1: str, dir2: str) -> pd.DataFrame:
        smallLists = self.extractSmallLists(dir1)
        if smallLists is None:
            raise FileNotFoundError(
                "All committee excels have been processed already.")
        processedSmallLists = self.extractSmallLists(dir2)
        if not processedSmallLists.empty:
            smallLists = pd.merge(
                smallLists, processedSmallLists, on=CommitteeExcel.id)
        return smallLists
