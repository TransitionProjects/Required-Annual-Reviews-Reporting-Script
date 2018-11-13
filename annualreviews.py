__author__ = "David Marienburg"
__version__ = "rc5"

#  This needs to be run prior to the beginning of the new reporting month
#  If it is not, you will need to make an alternate version of the 'today' variable
#  where 'today' == the last day of the proceeding month

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from datetime import date
from dateutil.relativedelta import relativedelta

import pandas as pd


def process_reviews():
    # establish the dates and date ranges required for processing
    # use this next row if you run the report on or after the first of the month
    today = date.today() + relativedelta(days=-2)
    # otherwise leave the today varaible as is
    # today =  date.today()
    review_month = (today + relativedelta(months=+2)).month
    review_year = (today + relativedelta(months=+2)).year
    first_of_next_month = today.replace(day=+1) + relativedelta(months=+1)
    last_of_next_month = first_of_next_month + relativedelta(months=+2) + relativedelta(days=-1)
    last_day_of_current_month = first_of_next_month + relativedelta(days=-1)
    first_day_of_current_month = first_of_next_month + relativedelta(months=-1)

    # open the excel output by ServicePoint and convert it to a pandas data frame
    file = askopenfilename()
    original_data = pd.read_excel(file, sheet_name="Report 1", header=2) # .drop("CM-End", axis=1)

    # add a couple of columns to the DataFrame to make the deletion of irrelevant rows
    original_data["Entry Month"] = original_data["Entry"].dt.month
    original_data["Entry Year"] = original_data["Entry"].dt.year
    original_data["Current Month"] = review_month
    original_data["Current Year"] = review_year

    # delete rows that where the entry date is not equal to current month + 2 in previous years
    cleaned = original_data[
        ((original_data["Entry Month"] == original_data["Current Month"]) & (
            original_data["Entry Year"] < original_data["Current Year"]))
    ].drop(["Entry Month", "Entry Year", "Current Month", "Current Year"], axis=1)

    # re-organize the columns so that CM name is the first column
    sorted = cleaned[[
        "CM",
        "CT ID",
        "Client First Name",
        "Client Last Name",
        "Entry",
        "Exit",
        "Service",
        "Program",
        "Review Date",
        "Interim or Follow-Up",
        "Review Type"
    ]].drop_duplicates(subset="CT ID", keep="first")

    # create the xlsx sheet to allow for formatting
    writer = pd.ExcelWriter(asksaveasfilename(), engine="xlsxwriter", date_format="mm/dd/yyyy")
    sorted.to_excel(writer, sheet_name="Processed", index=False)
    original_data.to_excel(writer, sheet_name="Raw Data", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Processed"]

    # establish the format for highlighting cells
    error_format = workbook.add_format({
        "bold": True,
        "bg_color": "yellow",
        "fg_color": "yellow"
    })

    # loop through the values in the exit date column highlighting cells where exit date is during the coming month
    for row in sorted.index:
        # these just simplified the typing of the if statements
        exit_date = pd.to_datetime(sorted.ix[row, "Exit"]).date()
        service_date = pd.to_datetime(sorted.ix[row, "Service"]).date()

        # highlighting cells where exit date is during the coming month
        if pd.isnull(sorted.ix[row, "Exit"]):
            pass
        elif ((first_of_next_month <= exit_date) and (exit_date <= last_of_next_month)):
            worksheet.write(row, 5, sorted.ix(row, "Exit"), error_format)
        else:
            pass

        # highlighting cells where service is prior to the last 30 days
        if pd.isnull(service_date):
            worksheet.write(row, 7, sorted.ix[row, "Service"], error_format)
        elif first_day_of_current_month > service_date:
            worksheet.write(row, 7, sorted.ix[row, "Service"], error_format)
        else:
            pass

    writer.save()

if __name__ == "__main__":
    process_reviews()
