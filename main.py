import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd

from member import getBigListInfo, getMemberInfo
from smallList import exportBlancListToExcel, readSmallExcel

UNPROCESSED_DIR = "Unprocessed committee lists"
PROCESSED_DIR = "Processed committee lists"
ID = "clublidnummer"
BIGLIST_HEADERS = ["firstname", "infix", "lastname",
                   "email", "phone", "countrycode", "bank_iban"]


def extractSmallLists(dir: str) -> pd.DataFrame | None:
    df = None
    for x in Path.cwd().joinpath(dir).glob("*.xlsx"):
        if df is None:
            df = readSmallExcel(x)
        else:
            df = pd.merge(df, readSmallExcel(x), on=ID)
    return df


def setupDirectory() -> None:
    Path.cwd().joinpath(UNPROCESSED_DIR).mkdir(parents=True, exist_ok=True)
    Path.cwd().joinpath(PROCESSED_DIR).mkdir(parents=True, exist_ok=True)


def moveProcessedSmallListsToProcessed() -> None:
    for x in Path.cwd().joinpath(UNPROCESSED_DIR).glob("*.xlsx"):
        x.rename(Path.cwd().joinpath(PROCESSED_DIR).joinpath(x.name))


def transformSmallformattingToBigFormatting(smallList: pd.DataFrame):
    # transform col with name and amounts to amount-n and description-n columns
    count = 1
    for col in smallList:
        if col == ID:
            continue
        amount = f"amount-{count}"
        description = f"description-{count}"
        smallList = smallList.rename(columns={col: amount})
        arr = np.where(smallList[amount] != 0, col, np.nan)
        smallList.insert(smallList.columns.get_loc(amount)+1, description, arr)
        count += 1

    return smallList


def createBigList() -> None:
    smallLists = extractSmallLists(UNPROCESSED_DIR)
    if smallLists is None:
        return
    processedSmallLists = extractSmallLists(UNPROCESSED_DIR)
    if processedSmallLists is not None:
        smallLists = pd.merge(smallLists, processedSmallLists, on=ID)

    df = pd.merge(getBigListInfo(),
                  transformSmallformattingToBigFormatting(smallLists), on=ID)
    df.insert(0, ID, df.pop(ID))
    df = df.rename(columns={ID: "club_membership_number"})

    x = dt.datetime.now()
    writer = pd.ExcelWriter(
        f"Incasso [Processed at {x.strftime('%a %d %b, %H-%M-%S')}].xlsx", engine="xlsxwriter")
    df.to_excel(writer, index=False)
    writer.close()
    return


def main():
    setupDirectory()
    answer = str(
        input("Do you want to create a new empty committee list? (y/n) "))
    if answer == "y":
        exportBlancListToExcel()

    answer = str(
        input("Do you want to fill in the committee lists into a final sheet? (y/n) "))
    if answer == "y":
        createBigList()
        moveProcessedSmallListsToProcessed()


if __name__ == "__main__":
    main()
