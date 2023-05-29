from pathlib import Path

import pandas as pd

MEMBER_HEADERS = ["voornaam", "tussenvoegsel", "achternaam",
                  "E-mailadres voor facturatie", "telefoonnummer 1", "clublidnummer", "iban"]


def getMemberInfo() -> pd.DataFrame:
    """Gets the most recent ClubCollect member info csv in the current directory. \n
    The start of the file should be named Export_* as it is by default."""
    result = None
    for foo in Path.cwd().glob("Export_*"):
        if result is None or foo.stat().st_mtime > result.stat().st_mtime:
            result = foo
    if result is None:
        raise Exception(
            "No member info csv is supplied in the current directory")
    return pd.read_csv(result, sep=";")


def getBigListInfo(memberInfo: pd.DataFrame = getMemberInfo()) -> pd.DataFrame:
    memberInfo = memberInfo[MEMBER_HEADERS]
    swaps = {
        "voornaam": "firstname",
        "tussenvoegsel": "infix",
        "achternaam": "lastname",
        "E-mailadres voor facturatie": "email",
        "telefoonnummer 1": "phone",
        "iban": "bank_iban"
    }
    df = memberInfo.rename(columns=swaps)
    df["countrycode"] = df.apply(lambda row: str(row["bank_iban"])[:2], axis=1)
    return df
