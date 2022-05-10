import os
import json
from typing import Iterator

digit = lambda s: s.isdigit() and int(s) or s

def parse_excel(table) -> list:
    """
        Parse a D2R excel text file and return a array of the rows.
    """
    split_table = table.split("\n")
    columns = split_table[0].split("\t")
    rows = []
    for line in split_table[1:]:
        row = line.split("\t")
        rows.append({columns[n]: digit(row[n]) for n in range(len(row))})
    return rows


def create_json_data_file(file: str, out: str):
    """
        Create a json file from the given excel file.
    """
    with open(f"{file}", "r") as f:
        table = f.read()
        with open(f"{out.replace('.json', '')}.json", "w") as f:
            f.write(json.dumps(parse_excel(table), indent=4))


def create_python_data_file(file: str, out: str):
    """
        Create a python file from the given excel file.
    """
    with open(f"{file}", "r") as f:
        table = f.read()
        out = out.replace('.py', '')
        with open(f"{out}.py", "w") as f:
            out = out.split("/")[-1]
            f.write(f"{out.replace('.py', '')} = {json.dumps(parse_excel(table), indent=4)}")
            

def create_json_data_files(outdir="json") -> Iterator[str]:
    """
        Create json files from the excel files at the given directory, also yields the file names.
    """
    os.makedirs(outdir, exist_ok=True)
    for file in os.listdir("excel"):
        if file.endswith(".txt"):
            create_json_data_file(f"excel/{file}", f"{outdir}/{file[:-4]}.json")
            yield file


def create_python_data_files(outdir="python") -> Iterator[str]:
    """
        Create python files from the excel files at the given directory, also yields the file names.
    """
    os.makedirs(outdir, exist_ok=True)
    for file in os.listdir("excel"):
        if file.endswith(".txt"):
            create_python_data_file(f"excel/{file}", f"{outdir}/{file[:-4]}.py")
            yield file
        


if __name__ == "__main__":
    for f in create_json_data_files():
        print(f"Created {f}.json")
    print('\n')
    for f in create_python_data_files():
        print(f"Created {f}.py")