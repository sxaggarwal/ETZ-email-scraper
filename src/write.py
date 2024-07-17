# Write data into MT - only need prices so might be easier using just pyodbc

from read import Data


def write_to_mt(d: Data) -> None:
    d.part_number
