__author__ = "David Katz-Wigmore"
__version__ = "rc4"

import pandas as pd

from datetime import date
from dateutil.relativedelta import relativedelta
from tkinter.filedialog import asksaveasfilename

class FindReviewDates:
    def __init__(self, file, start_month, start_year):
        self.year = start_year
        self.month = start_month
        self.raw_data = pd.read_excel(
            file,
            sheetname="Report 1",
            skiprows=2,
            index_col=None,
            names=[
                "CT ID",
                "Client First Name",
                "Client Last Name",
                "Entry",
                "Exit",
                "Service",
                "Program",
                "CM",
                "CM-End",
                "Review Date",
                "Interim or Follow-Up",
                "Review Type"
            ]
        )
        self.start_date = date(year=start_year, month=start_month, day=1)
        end_date = self.start_date + relativedelta(months=+1)

    def highlight_cells(self, data_frame):
        """
        Highlights cells yellow when their value matches the criteria

        :param data_frame:
        :return:
        """
        (
                   data_frame["Exit"].month == self.month + relativedelta(month=+1)
               ) | (
            data_frame["Service"] < (self.start_date + relativedelta(month=-1))
        )
        return ["background-color: yellow"]

    def find_initials(self,data_frame):
        """
        Mask for initials using regex

        :param data_frame:
        :return:
        """
        return (data_frame["CM"].str.extract("([A-Z][A-Z])"))

    def process(self):
        """
        Actually processes the document.

        Currently has lots of errors.  Needs serious trouble shooting.  Grrr dates == stupid.

        :return:
        """
        cols = [
            "CM",
            "CT ID",
            "Client First Name",
            "Client Last Name",
            "Entry",
            "Exit",
            "Service",
            "Program",
            "CM-End",
            "Review Date",
            "Interim or Follow-Up",
            "Review Type"
        ]
        data = self.raw_data
        data["Entry Year"] = [pd.to_datetime(data["Entry"].tolist()).year]
        data["Entry Month"] = [pd.to_datetime(data["Entry"].tolist()).month]
        mask =(data["Entry Year"] != self.year ) & (data["Entry Month"] == self.month)
        data_masked = data[mask]
        data_masked["CM"].str.replace("[A-Za-z/s]*",self.find_initials(data_masked), case=False)
        data_masked.style.apply(self.highlight_cells(data_masked))
        data_masked.drop("Entry Year")
        data_masked.drop("Entry Month")
        data_out = data_masked[cols]
        data_out.drop("CM-End")
        return data_out

    def write_to_excel(self):
        """
        Saves the data frame to an excel document.  Calls the process method so you don't need to.

        :return:
        """
        data_frame = self.process()
        writer = pd.ExcelWriter(asksaveasfilename(defaultextension=".xlsx"))
        data_frame.to_excel(writer, sheet_name="Processed")
        writer.save()
