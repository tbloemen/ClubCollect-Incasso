import queue
from enum import Enum
from pathlib import Path

import pandas as pd

from constants import (MEMBERDATA_DIR, PROCESSED_DIR, STARTING_ROW,
                       UNPROCESSED_DIR)
from schemas import ClubCollectSchema, CommitteeExcel, SmallSchema


class Format(Enum):
    CLUBCOLLECT = "ClubCollect"
    EXCEL = "Excel"
    COMMITTEE_EXCEL = "Committee Excel"


class Reader:
    def read_data(self) -> pd.DataFrame:
        raise NotImplementedError(
            "This method should only be called by inherited classes")

    def factory_method(self, format: Format) -> "Reader":
        if format == Format.CLUBCOLLECT or format == Format.EXCEL:
            return MemberReader(MEMBERDATA_DIR).factory_method()
        elif format == Format.COMMITTEE_EXCEL:
            return CommitteeExcelReader()
        else:
            raise ValueError()


class MemberReader(Reader):
    format: Format

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
            self.raw = pd.read_excel(result)
            self.format = Format.EXCEL
        self.dir = memberdata_dir

    def get_small_schema_df(self) -> pd.DataFrame:
        big_df = self.read_data()
        small_df = big_df[[SmallSchema.id, SmallSchema.fname,
                           SmallSchema.infix, SmallSchema.lname]]
        return small_df

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

    def read_data(self) -> pd.DataFrame:
        headers = list(self.swaps_in.keys())
        stripped_member_info: pd.DataFrame = self.raw[headers]
        df = stripped_member_info.rename(columns=self.swaps_in)
        df[ClubCollectSchema.countrycode] = df.apply(
            lambda row: str(row[ClubCollectSchema.iban])[:2], axis=1)

        return df


class CommitteeExcelReader(Reader):
    def read_data(self) -> pd.DataFrame:
        smallLists = self.extractSmallLists(UNPROCESSED_DIR)
        if smallLists is None:
            raise FileNotFoundError(
                "All committee excels have been processed already.")
        print(smallLists)
        processedSmallLists = self.extractSmallLists(PROCESSED_DIR)
        if not processedSmallLists.empty:
            smallLists = pd.merge(
                smallLists, processedSmallLists, left_index=True, right_index=True, how="outer")
        print(processedSmallLists)
        return smallLists

    def readSmallExcel(self, filename: Path) -> pd.DataFrame:
        df = pd.read_excel(filename, skiprows=STARTING_ROW-1)
        df = df[df[CommitteeExcel.total] != 0]
        df = df.drop(columns=[CommitteeExcel.fname, CommitteeExcel.infix,
                     CommitteeExcel.lname, CommitteeExcel.total])
        df = df[df.columns[df.sum() != 0]]
        df = df.set_index(CommitteeExcel.id)
        return df

    def extractSmallLists(self, dir: str) -> pd.DataFrame:
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
