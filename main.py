import os
from pathlib import Path

from constants import MEMBERDATA_DIR, PROCESSED_DIR, UNPROCESSED_DIR
from exporters import Exporter
from readers import Format, MemberReader


def setup_directory() -> None:
    Path.cwd().joinpath(UNPROCESSED_DIR).mkdir(parents=True, exist_ok=True)
    Path.cwd().joinpath(PROCESSED_DIR).mkdir(parents=True, exist_ok=True)
    Path.cwd().joinpath(MEMBERDATA_DIR).mkdir(parents=True, exist_ok=True)


def move_files() -> None:
    for x in Path.cwd().joinpath(UNPROCESSED_DIR).glob("*.xlsx"):
        x.rename(Path.cwd().joinpath(PROCESSED_DIR).joinpath(x.name))


def main():
    try:
        setup_directory()
        member_reader = MemberReader(MEMBERDATA_DIR).factory_method()
        committee_excel_exporter = Exporter().factory_method(Format.COMMITTEE_EXCEL)
        answer = str(
            input("Do you want to create a new empty committee list? (y/n) "))
        if answer == "y":
            committee_excel_exporter.export_to_excel(
                member_reader.get_small_schema_df())

        answer = str(
            input("Do you want to fill in the committee lists into a final sheet? (y/n) "))
        if answer == "y":
            Exporter().factory_method(member_reader.format).export_to_excel(
                member_reader.read_data())
            move_files()
    except Exception as e:
        print(e)
        os.system('pause')


if __name__ == "__main__":
    main()
