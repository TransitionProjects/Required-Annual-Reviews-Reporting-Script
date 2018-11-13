__author__ = 'David-Katz-Wigmore'
__version__ = 'rc3'

#  This is not currently working and needs significantly more debugging.

import re

from tkinter.filedialog import asksaveasfilename
from xlrd import open_workbook
from xlrd import xldate_as_tuple
from xlutils.copy import copy
from xlwt import easyxf
from datetime import datetime
from dateutil.relativedelta import relativedelta

class FindReviewDates:
    def __init__(self, edit_month, edit_year, path):
        """

        :param edit_month: a 1 - 2 digit int
        :param edit_year: a 4 digit int
        :param path: a file path
        """
        self.working_date = datetime(year=edit_year, month=edit_month, day=1) + relativedelta(months=+2)
        self.working_month = self.working_date.month
        self.working_year = self.working_date.year

        self.read_book = open_workbook(path)
        self.sheet = self.read_book.sheet_by_index(0)
        self.workbook = copy(self.read_book)
        self.edit_sheet = self.workbook.add_sheet("Processed")

    def processing(self):
        """
        This calls all the methods needed to actually process the report.
        :return:
        """
        self.write_headers()
        self.write_body()
        self.save_sheet()

    def month_check(self, date_in_question):
        """

        :param date_in_question: a python date object
        :return: boolean
        """
        if datetime.date(date_in_question).month == self.working_month:
            return True
        else:
            return False

    def year_check(self, date_in_question):
        """

        :param date_in_question: a python date object
        :return: boolean
        """
        if datetime.date(date_in_question).year == self.working_year:
            return True
        else:
            return False

    def excel_to_python_date(self, excel_date_object):
        """

        :param excel_date_object: excel date float
        :return: python date object
        """
        print(excel_date_object)
        date_tuple = xldate_as_tuple(excel_date_object, self.read_book.datemode)

        python_datetime = datetime(
            date_tuple[0],
            date_tuple[1],
            date_tuple[2],
            date_tuple[3],
            date_tuple[4],
            date_tuple[5]
        )
        return python_datetime

    def write_headers(self):
        """

        :return:
        """
        header = "Required Annual Review Assessments- Review due by {} / {}".format(
            self.working_month,
            self.working_year
        )
        self.edit_sheet.write(0, 0, header, easyxf("font: height 400;"))
        for column in range(0, self.sheet.ncols):
            if column == 7:
                self.edit_sheet.write(2, 0, self.sheet.cell_value(rowx=2, colx=column))
            else:
                self.edit_sheet.write(2, column+1, self.sheet.cell_value(rowx=2, colx=column))

    def write_body(self):
        """

        :return:
        """
        write_row = 3
        for row in range(3, self.sheet.nrows):
            if self.month_check(self.excel_to_python_date(self.sheet.cell_value(rowx=row, colx=3))):
                col = 0
                for column in range(0, self.sheet.ncols):
                    if (column == 4) and (self.sheet.cell_value(rowx=row, colx=column)):
                        month = self.month_check(self.excel_to_python_date(self.sheet.cell_value(row, column)).month)
                        year = self.year_check(self.excel_to_python_date(self.sheet.cell_value(row, column)).year)
                        if year and month:
                            self.edit_sheet.write(
                                write_row,
                                column + 1,
                                self.sheet.cell_value(rowx=row, colx=column),
                                easyxf("pattern: pattern solid_fill, fore_colour yellow;")
                            )
                            col += 1
                        else:
                            self.edit_sheet.write(write_row, column + 1, self.sheet.cell_value(rowx=row, colx=column))
                    elif column == 5:
                        try:
                            datetime(self.excel_to_python_date(self.sheet.cell_value(rowx=row, colx=column)))
                        except ValueError:
                            self.edit_sheet.write(
                                row,
                                column + 1,
                                "No Data Entered",
                                easyxf("pattern: pattern solid_fill, fore_colour yellow;")
                            )
                            column += 1
                        else:
                            if (
                                        self.year_check(self.excel_to_python_date(
                                            self.sheet.cell_value(rowx=row, colx=column))) == False
                            ) and (
                                    self.month_check(
                                        self.excel_to_python_date(self.sheet.cell_value(rowx=row, colx=column)))
                            ):
                                self.edit_sheet.write(
                                    write_row,
                                    column + 1,
                                    self.sheet.cell_value(rowx=row, colx=column)
                                )
                                column += 1
                            else:
                                self.edit_sheet.write(
                                    write_row,
                                    column + 1,
                                    self.sheet.cell_value(rowx=row, colx=column)
                                )
                                column += 1
                    elif column == 7:
                        try:
                            regex_search = re.search('([A-Z][A-Z])',self.sheet.cell_value(rowx=row, colx=column))
                        except Exception:
                            self.edit_sheet.write(write_row, 0, self.sheet.cell_value(rowx=row, colx=column))
                        except:
                            raise
                        else:
                            self.edit_sheet.write(write_row, 0, regex_search.group())
                            column += 1
                    else:
                        self.edit_sheet.write(write_row, column + 1, self.sheet.cell_value(rowx=row, colx=column))
                        column += 1
                row += 1

    def save_sheet(self):
        """

        :return:
        """
        self.workbook.save(
            asksaveasfilename(defaultextension=".xls", initialfile="Required Annual Reviews Assessment (Processed).xls")
        )
