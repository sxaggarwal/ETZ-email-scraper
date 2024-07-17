from pydantic import BaseModel, ConfigDict, Field, field_validator
from typing import Optional, Annotated
from datetime import datetime
import pandas as pd
import openpyxl


FILEPATH = r"C:\PythonProjects\email-excel-scraper\configs"


class Data(BaseModel):
    model_config = ConfigDict(extra='forbid', populate_by_name=True)
    part_number: str = Field(alias="Part Number")
    description: Optional[str] = Field(alias="Description")
    qty: float = Field(alias="Qty")
    thickness: float = Field(alias="Thickness")
    width: Optional[float] = Field(alias="Part Width")
    length: Optional[float] = Field(alias="Part Length")
    process: Optional[str] = Field(alias="Process")
    price: Optional[float] = Field(alias="Price")
    ship_charge: Optional[float] = Field(alias="Shipping Charge")
    lead_time: Optional[datetime] = Field(alias="LeadTime")
    good_until: Optional[datetime] = Field(alias="GoodUntil (MM/DD/YYYY)")
    currency_code: Optional[str] = Field(alias="Currency Code")
    freight: Optional[str] = Field(alias="Freight")  # NOTE: not sure about type
    nre: Optional[float] = Field(alias="NRE (non recording pricing)")
    moq: Optional[int] = Field(alias="MOQ")
    item_description: Optional[str] = Field(alias="Item Description")
    itar: Optional[bool] = Field(alias="ITAR or Not")
    cert_req: Optional[bool] = Field(alias="Certification required")  # NOTE: not sure about type
    dpas_rating: Optional[str] = Field(alias="DPAS rating")  # NOTE: not sure about type
    need_date: Optional[datetime] = Field(alias="Need Date")
    other_charges: Optional[float] = Field(alias="Other Charges")
    rev: Optional[str] = Field(alias="Rev")
    comments: Optional[str] = Field(alias="Comments")


def excel_to_dict(excel_path: str) -> dict[str: str]:
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb['Sheet1']
    headers = [cell.value for cell in sheet[1]]
    row_2 = [cell.value for cell in sheet[3]]
    return {headers[i]: row_2[i] for i in range(len(headers))}


if __name__ == "__main__":
    d = excel_to_dict(r"C:\PythonProjects\email-excel-scraper\configs\fin.xlsx")
    model = Data(**d)
    print(model.model_dump())
