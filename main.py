from pathlib import Path

import numpy as np
import pandas as pd

from member import getBigListInfo, getMemberInfo
from smallList import readSmallExcel

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


def setupDirectory():
    Path.cwd().joinpath(UNPROCESSED_DIR).mkdir(parents=True, exist_ok=True)
    Path.cwd().joinpath(PROCESSED_DIR).mkdir(parents=True, exist_ok=True)


def moveProcessedSmallListsToProcessed() -> pd.DataFrame | None:
    # Extract small unprocessed lists
    smallListExtraction = extractSmallLists(UNPROCESSED_DIR)
    if smallListExtraction is None:
        return

    # TODO move lists to processed map
    for x in Path.cwd().joinpath(UNPROCESSED_DIR).glob("*.xlsx"):
        x.rename(Path.cwd().joinpath(PROCESSED_DIR).joinpath(x.name))

    return smallListExtraction


def parseBigList(bigList: pd.DataFrame) -> pd.DataFrame:
    df = bigList.drop(columns=BIGLIST_HEADERS)
    df = df.rename(columns={"club_membership_number": ID})
    nums = []
    for col in df:
        if str(col).__contains__("amount-"):
            num = str(col)[7:]
            nums.append(num)
    for num in nums:
        colName = f"description-{num}"
        values = df[~df[colName].isna()].pop(colName)
        name = values.iloc[0]
        df = df.rename(columns={f"amount-{num}": name})
        df = df.drop(columns=colName)
    return df


def createBigList(smallList: pd.DataFrame, bigList: pd.DataFrame | None = None):
    memberInfo = getBigListInfo()

    if bigList is not None:
        smallList = pd.merge(smallList, parseBigList(bigList), on=ID)

    # transform col with name and amounts to amount-n and description-n columns
    count = 1
    for col in smallList:
        if col == ID:
            continue
        amount = f"amount-{count}"
        description = f"description-{count}"
        smallList = smallList.rename(columns={col: amount})
        # smallList[description] = np.where(
        #     smallList[amount] != 0, col, np.nan)
        arr = np.where(smallList[amount] != 0, col, np.nan)
        smallList.insert(smallList.columns.get_loc(amount)+1, description, arr)
        count += 1

    df = pd.merge(memberInfo, smallList, on=ID)
    df.insert(0, ID, df.pop(ID))

    df = df.rename(columns={ID: "club_membership_number"})
    print(df)

    writer = pd.ExcelWriter("testing.xlsx", engine="xlsxwriter")
    df.to_excel(writer, index=False)
    writer.close()
    return


bigList = parseBigList(pd.read_excel(
    Path.cwd() / "Mei clubcollect incasso.xlsx"))
smallList = extractSmallLists(UNPROCESSED_DIR)
moveProcessedSmallListsToProcessed()
if smallList is not None:
    createBigList(smallList)
