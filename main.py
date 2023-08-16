import os
from pathlib import Path

from readers import MemberReader
from smallList import SmallList

UNPROCESSED_DIR = "Unprocessed committee lists"
PROCESSED_DIR = "Processed committee lists"
MEMBERDATA_DIR = "Member data"


def setupDirectory() -> None:
    Path.cwd().joinpath(UNPROCESSED_DIR).mkdir(parents=True, exist_ok=True)
    Path.cwd().joinpath(PROCESSED_DIR).mkdir(parents=True, exist_ok=True)
    Path.cwd().joinpath(MEMBERDATA_DIR).mkdir(parents=True, exist_ok=True)


def moveProcessedSmallListsToProcessed() -> None:
    for x in Path.cwd().joinpath(UNPROCESSED_DIR).glob("*.xlsx"):
        x.rename(Path.cwd().joinpath(PROCESSED_DIR).joinpath(x.name))


def main():
    try:
        setupDirectory()
        reader = MemberReader(MEMBERDATA_DIR).factory_method()
        smallList = SmallList(reader.convertToSmallDataFrame())
        answer = str(
            input("Do you want to create a new empty committee list? (y/n) "))
        if answer == "y":
            smallList.exportToExcel()

        answer = str(
            input("Do you want to fill in the committee lists into a final sheet? (y/n) "))
        if answer == "y":
            reader.exportToExcel(UNPROCESSED_DIR, PROCESSED_DIR)
            moveProcessedSmallListsToProcessed()
    except Exception as e:
        print(e)
        os.system('pause')


if __name__ == "__main__":
    main()
