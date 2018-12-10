__author__ = "David Marienburg"
__version__ = "2.0"
"""
This script should be used to process the monthy Required Annual Reviews
report.
"""

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
        as well as the self.provider_list.

        :file_path: an askopenfilename object
        :provider list: an askopenfilename object of providers to be used in the
        report
        """
        self.entries_df = pd.read_excel(file_path, sheet_name="EntryData")
        self.reviews_df = pd.read_excel(file_path, sheet_name="ReviewData")
        self.cm_df = pd.read_excel(file_path, sheet_name="CMData")
        self.placements_df = pd.read_excel(file_path, sheet_name="PlacementData")
        self.provider_list = pd.read_excel(provider_path)["Service Provide Provider"].tolist()

    def filter_entries(self, hud_entries_df):
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
        entries = hud_entries_df[
            hud_entries_df["Entry Exit Provider Id"].isin(self.provider_list) &
            (hud_entries_df["Entry Exit Entry Date"].dt.year < dt.today().year)
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
            # set the month, day and year of the review range
            try:
                review_date = date(
                    year=dt.today().year,
                    month=entries.loc[row, "Entry Month"],
                    day=entries.loc[row, "Entry Day"]
                )
            except ValueError:
                review_date = date(
                    year=dt.today().year,
                    month=entries.loc[row, "Entry Month"],
                    day=entries.loc[row, "Entry Day"] - 1
                )

            # add values to each of the appliable columns using the relativedelta
            # method to create the high and low end of the acceptable review
            # dates
            entries.loc[row, "Review Period Start"] = review_date + rd(months=-1)
            entries.loc[row, "Review Period End"] = review_date + rd(months=+1)

        # Create an output df where the most recent review date is non-existant,
        # the review date does not fall between the Review Period Start or
        # Review Period End dates, or the review type is something other than
        # Annual Assessment
        # make a copy of the reviewable_df showing completed reviews or entry
        # exits that ended prior to the review period.
        good_reviews = entries[
            (
                entries["Entry Exit Exit Date"].isna() |
                (entries["Entry Exit Exit Date"] > entries["Review Period Start"])
            ) &
            (entries["Entry Exit Review Type"] == "Annual Assessment") &
            (
                (entries["Review Period Start"] < entries["Entry Exit Review Date"]) &
                (entries["Review Period End"] > entries["Entry Exit Review Date"])
            )
        ].drop_duplicates(
            subset=["Client Unique Id", "Entry Exit Uid"]
        )

        # finalize the shape of the output data frame
        output = entries[
            ~(entries["Entry Exit Uid"].isin(good_reviews["Entry Exit Uid"])) &
            (
                (entries["Entry Exit Exit Date"] > entries["Review Period Start"]) |
                entries["Entry Exit Exit Date"].isna()
            )
        ].drop_duplicates(
            subset=["Client Unique Id", "Entry Exit Uid"]
        )[[
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Entry Exit Provider Id",
            "Review Period Start",
            "Review Period End"
        ]]

        return output

    def filter_vets_entries(self, vets_entries_df, placements_df):
        """
        Create a dataframe containing participants who need annual reviews.
        This will be determined using the following criteria:
        1) Participants must have been enrolled in the provider for a least
           a full calendar year after they were placed into housing
        2) Participants will not have had a annual assessment within a month
           of their entry date during the current calendar year
        3) Entry provider must be present in the self.provider_list
        """
        # outer join the vets_entries_df and the placement_df on the client
        # unique id column
        joined_df = vets_entries_df.merge(
            placements_df,
            how="outer",
            on="Client Unique Id"
        )

        # convert all date columns into dt objects
        joined_df["Entry Exit Entry Date"] = joined_df["Entry Exit Entry Date"].dt.date
        joined_df["Entry Exit Exit Date"] = joined_df["Entry Exit Exit Date"].dt.date
        joined_df["Placement Date"] = joined_df["Placement Date(3072)"].dt.date
        joined_df["Housing Move-in Date(9160)"] = joined_df["Housing Move-in Date(9160)"].dt.date
        joined_df["Entry Exit Review Date"] = joined_df["Entry Exit Review Date"].dt.date

        # slice the data frame so that only rows where the placement date is
        # greater than the entry exit entry date, the placement date is not nan,
        # and the placement date is at least a calendar year from today.
        reviewable_df = joined_df[
            (
                joined_df["Placement Date(3072)"].notna() &
                (
                    (joined_df["Entry Exit Entry Date"] < joined_df["Placement Date"]) |
                    (joined_df["Entry Exit Entry Date"] == joined_df["Placement Date"])
                ) &
                (
                    joined_df["Placement Date(3072)"].dt.year < dt.today().year
                )
            )
        ]

        # Create columns for each element of the placement date: day, month, year
        reviewable_df["Placed Month"] = reviewable_df["Placement Date(3072)"].dt.month
        reviewable_df["Placed Year"] = reviewable_df["Placement Date(3072)"].dt.year
        reviewable_df["Placed Day"] = reviewable_df["Placement Date(3072)"].dt.day

        # Define Review Period Start and Review Period End columns and fill with
        # nan values
        reviewable_df["Review Period Start"] = ""
        reviewable_df["Review Period End"] = ""

        # Create Dates that the annual review must be between
        for row in reviewable_df.index:
            # set the month, day and year of the review range but if a value
            # error occurs, most likely due to a leap year, simply subtract one
            # from the day number that is passed to the datetime.date method
            try:
                review_date = date(
                    year=dt.today().year,
                    month=reviewable_df.loc[row, "Placed Month"],
                    day=reviewable_df.loc[row, "Placed Day"]
                )
            except ValueError:
                review_date = date(
                    year=dt.today().year,
                    month=reviewable_df.loc[row, "Placed Month"],
                    day=reviewable_df.loc[row, "Placed Day"] - 1
                )

            # add values to each of the appliable columns using the relativedelta
            # method to create the low and high ends of the acceptable date ranges
            reviewable_df.loc[row, "Review Period Start"] =  review_date + rd(months=-1)
            reviewable_df.loc[row, "Review Period End"] = review_date + rd(months=+1)

        # make a copy of the reviewable_df showing completed reviews or entry
        # exits that ended prior to the review period.
        good_reviews = reviewable_df[
            (
                reviewable_df["Entry Exit Exit Date"].isna() |
                (reviewable_df["Entry Exit Exit Date"] > reviewable_df["Review Period Start"])
            ) &
            (reviewable_df["Entry Exit Review Type"] == "Annual Assessment") &
            (
                (reviewable_df["Review Period Start"] < reviewable_df["Entry Exit Review Date"]) &
                (reviewable_df["Review Period End"] > reviewable_df["Entry Exit Review Date"])
            )
        ].drop_duplicates(
            subset=["Client Unique Id", "Entry Exit Uid"]
        )

        # finalize the shape of the output data frame
        output = reviewable_df[
            ~(reviewable_df["Entry Exit Uid"].isin(good_reviews["Entry Exit Uid"])) &
            (
                (reviewable_df["Entry Exit Exit Date"] > reviewable_df["Review Period Start"]) |
                reviewable_df["Entry Exit Exit Date"].isna()
            )
        ].drop_duplicates(
            subset=["Client Unique Id", "Entry Exit Uid"]
        )[[
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Entry Exit Provider Id",
            "Review Period Start",
            "Review Period End"
        ]]

        return output

    def filter_cms(self):
        """
        Create a dataframe containing case managers of participants needing
        annual reviews.  This will be determined using the following
        criteria:
        1) Case manager's provider is not an SSVF provider if SSVF is not
           in self.provider_list
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
        # merge the self.entries_df and the self.reviews_df
        merged = self.entries_df.merge(
            self.reviews_df,
            on="Entry Exit Uid",
            how="outer"
        ).dropna(subset=["Entry Exit Entry Date"])

        # create a HUD entries only version of the data frame since these are
        # reviewed annually from entry date
        hud_entries_df = merged[
            ~(merged["Entry Exit Provider Id"].str.contains("SSVF")) &
            ~(merged["Entry Exit Provider Id"].str.contains("Vets")) &
            ~(merged["Entry Exit Provider Id"].str.contains("Veterans"))
        ]

        # create a HUD entries only version of the data frame since these are
        # reviewed annually from placement date
        vets_entries_df = merged[
            (merged["Entry Exit Provider Id"].str.contains("OHA"))
        ]

        # call the filter_hud_entries method
        hud_reviews = self.filter_entries(hud_entries_df)

        # call the filter_vets_entries method
        vets_reviews = self.filter_vets_entries(vets_entries_df, self.placements_df)

        # call the filter_cms method
        cms = self.filter_cms()

        # concatenate the vets and hud reviews data frames the reindex the
        # resultant dataframe and call the output reviews
        reviews = pd.concat([hud_reviews, vets_reviews], ignore_index=True)

        # merge the reviews and the cms dataframes
        final = reviews.merge(cms, on="Client Uid", how="left")

        # return the merged data frame
        return final

    def save_df(self, required_df):
        """
        Save the merged and raw data frames to an excel spreadsheet with
        each data frame getting its own sheet within the larger workbook.
        """
        writer = pd.ExcelWriter(
            asksaveasfilename(title="Save As"),
            engine="xlsxwriter"
        )
        reviews = required_df[[
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Entry Exit Provider Id",
            "Review Period Start",
            "Review Period End",
            "Case Worker Name"
        ]]


        reviews[
            reviews["Entry Exit Provider Id"].str.contains("Vets") |
            reviews["Entry Exit Provider Id"].str.contains("Veterans")
        ].to_excel(writer, sheet_name="Vets Required Reviews", index=False)
        reviews[
            ~(reviews["Entry Exit Provider Id"].str.contains("Vets")) &
            ~(reviews["Entry Exit Provider Id"].str.contains("Veterans")) &
            ~(reviews["Entry Exit Provider Id"].str.contains("SSVF"))
        ].to_excel(writer, sheet_name="Ret Required Reviews", index=False)
        reviews.to_excel(writer, sheet_name="All Required Reviews", index=False)
        self.entries_df.to_excel(writer, sheet_name="Raw Entry Data", index=False)
        self.reviews_df.to_excel(writer, sheet_name="Raw Review Data", index=False)
        self.cm_df.to_excel(writer, sheet_name="Raw CM Data", index=False)
        writer.save()

if __name__ == "__main__":
    report_path = askopenfilename(title="Open the AnnualReviews Report")
    provider_path = askopenfilename(title="Open the ProviderList.xlsx file")
    run = AnnualReviewReport(report_path, provider_path)
    reviews = run.merge_entries_and_cms()
    run.save_df(reviews)
