import datetime as dt

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

from constants import (FILENAMESMALLLIST, INSTRUCTIONS, STARTING_ROW,
                       TOTALSTRING)
from readers import Format, Reader
from schemas import ClubCollectSchema, CommitteeExcel


class Exporter:
    def export_to_excel(self, information_df: pd.DataFrame) -> None:
        raise NotImplementedError(
            "This should be implemented by inherited classes")

    def factory_method(self, format: Format) -> "Exporter":
        if format == Format.CLUBCOLLECT:
            return ClubCollectExporter()
        elif format == Format.COMMITTEE_EXCEL:
            return CommitteeExcelExporter()
        elif format == Format.EXCEL:
            return ExcelExporter()
        else:
            raise ValueError(format)


class ClubCollectExporter(Exporter):
    swaps_out = {
        ClubCollectSchema.id: "club_membership_number",
        ClubCollectSchema.fname: "firstname",
        ClubCollectSchema.infix: "infix",
        ClubCollectSchema.lname: "lastname",
        ClubCollectSchema.email: "email",
        ClubCollectSchema.phone: "phone",
        ClubCollectSchema.iban: "bank_iban",
        ClubCollectSchema.countrycode: "countrycode"
    }

    def export_to_excel(self, clubcollect_schema_df: pd.DataFrame) -> None:
        smallLists = Reader().factory_method(
            Format.COMMITTEE_EXCEL).read_data()
        df = pd.merge(clubcollect_schema_df, self.amount_description_helper(
            smallLists), on=ClubCollectSchema.id, how="inner")
        df.insert(0, ClubCollectSchema.id, df.pop(ClubCollectSchema.id))
        df = df.rename(columns=self.swaps_out)

        x = dt.datetime.now()
        writer = pd.ExcelWriter(
            f"Incasso [Processed at {x.strftime('%a %d %b, %H-%M-%S')}].xlsx", engine="xlsxwriter")
        df.to_excel(writer, index=False)

        writer.sheets["Sheet1"].autofit()
        writer.close()
        return

    def amount_description_helper(self, smallList: pd.DataFrame):
        # transform col with name and amounts to amount-n and description-n columns
        smallList = smallList.replace(0, np.nan)
        for i, col in enumerate(smallList):
            if col == ClubCollectSchema.id:
                continue
            amount = f"amount-{i + 1}"
            description = f"description-{i + 1}"
            smallList = smallList.rename(columns={col: amount})
            arr = np.where(np.isnan(smallList[amount]),
                           [""]*len(smallList), [col]*len(smallList))
            smallList.insert(smallList.columns.get_loc(
                amount)+1, description, arr)
        return smallList


class CommitteeExcelExporter(Exporter):
    swaps_out = {
        CommitteeExcel.id: "ID",
        CommitteeExcel.fname: "First name",
        CommitteeExcel.infix: "Infix",
        CommitteeExcel.lname: "Last name",
        CommitteeExcel.total: "Total",
        CommitteeExcel.bedrag_a: "Bedrag A",
        CommitteeExcel.bedrag_b: "Bedrag B",
        CommitteeExcel.bedrag_c: "Bedrag C"
    }

    def export_to_excel(self, smallschema_df: pd.DataFrame) -> None:
        """Exports a blanc list to excel in the same directory."""

        # Setup dataframe
        smallschema_df[CommitteeExcel.total] = ""
        smallschema_df[CommitteeExcel.bedrag_a] = 0
        smallschema_df[CommitteeExcel.bedrag_b] = 0
        smallschema_df[CommitteeExcel.bedrag_c] = 0
        smallschema_df = smallschema_df.rename(columns=self.swaps_out)

        # Setup filename
        x = dt.datetime.now()
        date_string = x.strftime('%a %d %b, %H-%M-%S')
        name = f"{FILENAMESMALLLIST} [{date_string}].xlsx"
        writer = pd.ExcelWriter(name, engine='xlsxwriter')
        sheetName = FILENAMESMALLLIST

        # Export to excel
        smallschema_df.to_excel(writer, index=False,
                                startrow=STARTING_ROW-1, sheet_name=sheetName)
        sheet: xlsxwriter.Workbook.worksheet_class = writer.sheets[sheetName]

        # Write
        created_format: xlsxwriter.workbook.Format = writer.book.add_format()  # type: ignore
        created_format.set_align("justify")
        created_format.set_italic()
        sheet.merge_range(
            0, 0, 0, 2, f"Created at {date_string}", created_format)
        merged_format: xlsxwriter.workbook.Format = writer.book.add_format()  # type: ignore
        merged_format.set_text_wrap()
        sheet.merge_range(1, 0, 1, 3, INSTRUCTIONS, merged_format)
        totalColumn = 4
        totalformat: xlsxwriter.workbook.Format = writer.book.add_format()  # type: ignore
        totalformat.set_align("justify")
        sheet.write(STARTING_ROW-3, totalColumn-1, TOTALSTRING, totalformat)
        sheet.autofit()
        sheet.set_column_pixels(0, 0, 40)
        sheet.set_row_pixels(1, 80)

        # Write header sums
        for i in range(totalColumn, len(smallschema_df.columns)):
            formula = f"=SUM({xl_rowcol_to_cell(STARTING_ROW, i)}:{xl_rowcol_to_cell(STARTING_ROW + len(smallschema_df.index)-1, i)})"
            sheet.write_formula(0, i, formula=formula)

        # Write sums per row
        for i in range(STARTING_ROW, len(smallschema_df.index) + STARTING_ROW):
            formula = f"=SUM({xl_rowcol_to_cell(i, totalColumn + 1)}:{xl_rowcol_to_cell(i, len(smallschema_df.columns)-1)})"
            sheet.write_formula(i, totalColumn, formula=formula)

        writer.close()


class ExcelExporter(Exporter):
    def export_to_excel(self, information_df: pd.DataFrame) -> None:
        return super().export_to_excel(information_df)
