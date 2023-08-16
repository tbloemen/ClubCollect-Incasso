import datetime as dt
import queue
from enum import Enum
from pathlib import Path

import numpy as np
import pandas as pd

from schemas import ClubCollectSchema, SmallSchema
from smallList import SmallList


class Format(Enum):
    CLUBCOLLECT = "ClubCollect"
    EXCEL = "Excel"


class MemberReader:
    def __init__(self, memberdata_dir: str) -> None:
        myqueue: queue.PriorityQueue[tuple[float,
                                           Path]] = queue.PriorityQueue()
        for foo in Path.cwd().joinpath(memberdata_dir).glob("*"):
            if foo.name.endswith((".csv", ".xlsx")):
                myqueue.put((-foo.stat().st_mtime, foo))
        if myqueue.empty():
            raise FileNotFoundError(
                f"No member data file is supplied in the folder {memberdata_dir}.")
        result = myqueue.get()[1]

        if result.name.endswith("csv"):
            self.raw = pd.read_csv(result, sep=";")
            self.format = Format.CLUBCOLLECT
        else:
            self.raw = pd.read_excel(result, "Huidige_leden")
            self.format = Format.EXCEL
        self.dir = memberdata_dir

    def convertToBigDataFrame(self) -> pd.DataFrame:
        raise NotImplementedError(
            "This method should only be called by inherited classes")

    def convertToSmallDataFrame(self) -> pd.DataFrame:
        big_df = self.convertToBigDataFrame()
        small_df = big_df[[SmallSchema.id, SmallSchema.fname,
                           SmallSchema.infix, SmallSchema.lname]]
        return small_df

    def exportToExcel(self) -> None:
        raise NotImplementedError(
            "This method should only be called by inherited classes")

    def factory_method(self):
        if self.format == Format.CLUBCOLLECT:
            return ClubCollectReader(self.dir)
        elif self.format == Format.EXCEL:
            raise NotImplementedError("Excel parsing is yet to be implemented")
        else:
            raise ValueError(format)


class ClubCollectReader(MemberReader):
    swaps_in = {
        "clublidnummer": ClubCollectSchema.id,
        "voornaam": ClubCollectSchema.fname,
        "tussenvoegsel": ClubCollectSchema.infix,
        "achternaam": ClubCollectSchema.lname,
        "E-mailadres voor facturatie": ClubCollectSchema.email,
        "telefoonnummer 1": ClubCollectSchema.phone,
        "iban": ClubCollectSchema.iban}

    swaps_out = {
        ClubCollectSchema.id: "club_membership_number",
        ClubCollectSchema.fname: "firstname",
        ClubCollectSchema.infix: "infix",
        ClubCollectSchema.lname: "lastname",
        ClubCollectSchema.email: "email",
        ClubCollectSchema.email: "phone",
        ClubCollectSchema.iban: "bank_iban",
        ClubCollectSchema.countrycode: "countrycode"
    }

    def convertToBigDataFrame(self) -> pd.DataFrame:
        headers = list(self.swaps_in.keys())
        stripped_member_info: pd.DataFrame = self.raw[headers]
        df = stripped_member_info.rename(columns=self.swaps_in)
        df[ClubCollectSchema.countrycode] = df.apply(
            lambda row: str(row[ClubCollectSchema.iban])[:2], axis=1)

        return df

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

    def exportToExcel(self, unprocessed_dir: str, processed_dir: str) -> None:
        smallList = SmallList(self.convertToSmallDataFrame())
        smallLists = smallList.combineSmallLists(
            unprocessed_dir, processed_dir)
        df = pd.merge(self.convertToBigDataFrame(), self.amount_description_helper(
            smallLists), on=ClubCollectSchema.id)
        df.insert(0, ClubCollectSchema.id, df.pop(ClubCollectSchema.id))
        df = df.rename(columns=self.swaps_out)

        x = dt.datetime.now()
        writer = pd.ExcelWriter(
            f"Incasso [Processed at {x.strftime('%a %d %b, %H-%M-%S')}].xlsx", engine="xlsxwriter")
        df.to_excel(writer, index=False)

        writer.sheets["Sheet1"].autofit()
        writer.close()
        return
