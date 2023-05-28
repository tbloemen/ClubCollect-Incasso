import datetime as dt
from pathlib import Path

import pandas as pd


def getMemberInfo() -> pd.DataFrame:
    """Gets the most recent ClubCollect member info csv in the current directory. \n
    The start of the file should be named Export_* as it is by default."""
    result = None
    for foo in Path.cwd().glob("Export_*"):
        if result is None or foo.stat().st_mtime > result.stat().st_mtime:
            result = foo
    if result is None:
        raise Exception("No member info csv is supplied in the current directory")
    return pd.read_csv(result, sep=";")
    

def getBlancList() -> pd.DataFrame:
    memberinfo = getMemberInfo()
    commissielist = memberinfo[["clublidnummer", "voornaam", "tussenvoegsel", "achternaam"]]
    withEmpty = commissielist.assign(Bedrag_a=0, Bedrag_b=0, Bedrag_c=0)
    return withEmpty

def exportBlancListToExcel(df: pd.DataFrame):
    x = dt.datetime.now()
    name = f"Empty Incasso List [{x.strftime('%a %d %b, %H-%M-%S')}].xlsx"
    df.to_excel(Path.cwd() / name, index=False)

print(exportBlancListToExcel(getBlancList()))