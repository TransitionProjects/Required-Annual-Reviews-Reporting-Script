__author__ = "David Marienburg"
__version__ = "rc6"
"""
This script should be used to process the monthy Required Annual Reviews
report.
"""

__author__ = "David Marienburg"
__maintainer__ = "David Marienburg"
__version__ = "2.0rc1"

import pandas as pd

from datetime import date
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta as rd
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class AnnualReviewReport:
    def __init__(self, file_path, provider_path):
        """
        Initialize the class and create two dataframes, review_df and cm_df,
        as well as the self.provider_listself.

        :file_path: an askopenfilename object
        :provider list: an askopenfilename object of providers to be used in the
        report
        """
        self.reviews_df = pd.read_excel(file_path, sheet_name="Reviews")
        self.cm_df = pd.read_excel(file_path, sheet_name="Case Managers")
        self.provider_list = pd.read_excel(provider_path)["Service Provide Provider"].tolist()

    def filter_entries(self):
        """
        Create a dataframe containing participants who need annual reviews.
        This will be determined using the following criteria:
        1) Participants must have been enrolled in the provider for a least
           a full calendar year
        2) Participants will not have had a annual assessment within a month
           of their entry date during the current calendar year
        3) Entry provider must be present in the self.provider_list
        """
        # create a copy of the reviews_df containing only participats with a
        # provider in the self.provider_list and they entry year is prior to the
        # current year
        entries = self.reviews_df[
            self.reviews_df["Entry Exit Provider Id"].isin(self.provider_list) &
            (self.reviews_df["Entry Exit Entry Date"].dt.year < dt.today().year)
        ]

        # convert all the date columns into dt objects
        entries["Entry Exit Exit Date"] = entries["Entry Exit Exit Date"].dt.date
        entries["Entry Exit Review Date"] = entries["Entry Exit Review Date"].dt.date

        # create columns for each element of the start date: day, month, year
        entries["Entry Month"] = entries["Entry Exit Entry Date"].dt.month
        entries["Entry Year"] = entries["Entry Exit Entry Date"].dt.year
        entries["Entry Day"] = entries["Entry Exit Entry Date"].dt.day
        entries["Entry Exit Entry Date"] = entries["Entry Exit Entry Date"].dt.date

        # Define Review Period Start and Review Period End and fill with nan
        entries["Review Period Start"] = ""
        entries["Review Period End"] = ""

        # Create dates that the annual review must be between
        for row in entries.index:
            entries.loc[row, "Review Period Start"] = date(
                year=dt.today().year,
                month=entries.loc[row, "Entry Month"],
                day=entries.loc[row, "Entry Day"]
            ) + rd(months=-1)
            entries.loc[row, "Review Period End"] = date(
                year=dt.today().year,
                month=entries.loc[row, "Entry Month"],
                day=entries.loc[row, "Entry Day"]
            ) + rd(months=+1)

        # drop the now extraneous day, month, and year columns
        entries = entries[[
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Entry Exit Provider Id",
            "Entry Exit Entry Date",
            "Entry Exit Exit Date",
            "Review Period Start",
            "Entry Exit Review Date",
            "Review Period End",
            "Entry Exit Review Point in Time Type",
            "Entry Exit Review Type"
        ]]

        # filter the entries df dropping duplicates based on client id, entry
        # provider, and entry exit review date
        cleaned = entries.sort_values(
            by=["Entry Exit Review Date"],
            ascending=False
        ).drop_duplicates(subset="Client Uid", keep="first")

        # Create an output df where the most recent review date is non-existant,
        # the review date does not fall between the Review Period Start or
        # Review Period End dates, or the review type is something other than
        # Annual Assessment
        output = cleaned[
            (
                cleaned["Entry Exit Review Date"].isna() |
                (cleaned["Review Period Start"] > cleaned["Entry Exit Review Date"]) |
                (cleaned["Review Period End"] < cleaned["Entry Exit Review Date"]) |
                (cleaned["Entry Exit Review Type"] != "Annual Assessment")
            ) &
            (
                cleaned["Entry Exit Exit Date"].isna() |
                (cleaned["Entry Exit Exit Date"] > cleaned["Review Period End"])
            )
            ]

        return output

    def filter_cms(self):
        """
        Create a dataframe containing case managers of participants needing
        annual reviews.  This will be determined using the following
        criteria:
        1) Case manager's provider is not an SSVF provider if SSVF is not
           in self.provider.provider_list
        2) If there are multiple active case managers the newest one will
           be kept, all others will be discarded.
        """
        cms = self.cm_df[self.cm_df["Case Worker Provider"].isin(self.provider_list)]
        return cms.sort_values(
            by=["Client Uid", "Case Worker Date Started"],
            ascending=False
        ).drop_duplicates(subset="Client Uid", keep="first")

    def merge_entries_and_cms(self):
        """
        Merge the entries data frame from the filter entries method and the
        case manager dataframe from the filter cms method.  This merge will
        be a left merge.
        """
        reviews = self.filter_entries()
        cms = self.filter_cms()
        merged = reviews.merge(cms, on="Client Uid", how="left")

        return merged


    def save_df(self, required_df):
        """
        Save the merged and raw data frames to an excel spreadsheet with
        each data frame getting its own sheet within the larger workbook.
        """
        writer = pd.ExcelWriter(
            asksaveasfilename(title="Save As"),
            engine="xlsxwriter"
        )
        required_df[[
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Entry Exit Provider Id",
            "Entry Exit Entry Date",
            "Review Period Start",
            "Review Period End",
            "Case Worker Name"
        ]].to_excel(writer, sheet_name="Required Reviews", index=False)
        self.reviews_df.to_excel(writer, sheet_name="Raw Entry Data", index=False)
        self.cm_df.to_excel(writer, sheet_name="Raw CM Data", index=False)
        writer.save()

if __name__ == "__main__":
    report_path = askopenfilename(title="Open the AnnualReviews Report")
    provider_path = askopenfilename(title="Open the ProviderList.xlsx file")
    run = AnnualReviewReport(report_path, provider_path)
    reviews = run.merge_entries_and_cms()
    run.save_df(reviews)
