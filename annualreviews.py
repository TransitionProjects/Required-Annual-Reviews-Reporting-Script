__author__ = 'David Katz'
__version__ = 'rc2'

'''Currently this program throws and error when it encounters months Nov or Dec (as input by the user)  This will need
to be fixed to ensure that this program can be used without modifications in the future

Additionally this program needs some serious cleaning up for PEP-8 compliance as well as elimination of deeply nested
conditionals and use of global variables.  This is ugly code that just happens to work but will be hard to maintain
going forward.'''

import re

from tkinter.filedialog import asksaveasfilename
from xlrd import open_workbook
from xlrd import xldate_as_tuple
from xlutils.copy import copy
from xlwt import easyxf

wmonth = []
month = 0
wyear = []
year = 0

def processing(path, m, y):
    file_path = str(path)
    rb = open_workbook(file_path)
    sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    editsheet = wb.add_sheet('Processed')
    month_entry = int(m)
    year_entry = int(y)

    def working_month():
        """
        the workingMonth function advances the month by two respecting the 12 month year
        :return:
        """
        global wmonth
        global month
        if month_entry == 11:
            month = 1
            wmonth.append(1)
        elif month_entry == 12:
            month = 2
            wmonth.append(2)
        else:
            c = month_entry + 2
            month = c
            wmonth.append(c)

    def working_year():
        """
        the workingYear function advances they wyear variable by 1 in the event that the wmonth is Dec. or Nov.
        :return:
        """
        global year
        if (month == 11) or (month == 12):
            year = year_entry + 1
        else:
            year = year_entry

    def datecheck(d):
        """
        dateCheck will reduce a date tuple to only the month data then look at a given date and return 'true' if the
        date's month is equal to the working month
        :param d:
        :return:
        """
        global wmonth
        dlist = list(d)
        dlist.pop(0)
        dlist.pop(3)
        dlist.pop(2)
        dlist.pop(2)
        dlist.pop(1)
        print(dlist)
        print(wmonth)
        if dlist == wmonth:
            return True
        else:
            return False

    def yearcheck(y):
        dlist = list(y)
        dlist.pop(1)
        dlist.pop(1)
        dlist.pop(1)
        dlist.pop(1)
        dlist.pop(1)
        if dlist == wyear:
            return True
        else:
            return False

    def dateconvert(r, c):
        """
        dateConvert and yearConvert change excel date codes into date tuples so that python can read/process them
        :param r:
        :param c:
        :return:
        """
        print(xldate_as_tuple(sheet.cell(r, c).value, rb.datemode))
        date_value = xldate_as_tuple(sheet.cell(r, c).value, rb.datemode)
        return datecheck(date_value)

    def yearconvert(r, c):
        print("row: {}; coll: {}".format(r, c))
        date_value = xldate_as_tuple(sheet.cell(r, c).value, rb.datemode)
        return yearcheck(date_value)

    def writesheet(sheet, editsheet):
        """
        writeSheet writes data to the "processed" sheet
        :param sheet:
        :param editsheet:
        :return:
        """
        row = 3
        header = 'Required Annual Review Assessments- by %d / %d' % (month, year)
        editsheet.write(0, 0, header, easyxf('font: height 400;'))
        for c in range(0, sheet.ncols):
            if c == 7:
                editsheet.write(2, 0, sheet.cell_value(rowx=2, colx=c))
            else:
                editsheet.write(2, c+1, sheet.cell_value(rowx=2, colx=c))
        for r in range(3, sheet.nrows):
            if dateconvert(r, 3):
                col = 0
                for c in range(0, sheet.ncols):
                    if (c == 4) and (sheet.cell_value(rowx=r, colx=c) != ''):
                        if yearconvert(r, 4) and datecheck(dateconvert(r, 4)):
                            editsheet.write(row, c+1, sheet.cell_value(rowx=r, colx=c),
                                            easyxf('pattern: pattern solid_fill,fore_colour yellow;'))
                            col += 1
                        else:
                            editsheet.write(row, c+1, sheet.cell_value(rowx=r, colx=c))
                    elif c == 5:
                        try:
                            yearconvert(r, 5)
                            dateconvert(r, 5)
                        except ValueError:
                            editsheet.write(row, c+1, "No Data Entered",
                                            easyxf("pattern: pattern solid_fill, fore_colour yellow;"))
                            col += 1
                        else:
                            if (yearconvert(r, 5) is False) and (dateconvert(r, 5)):
                                editsheet.write(row, c+1, sheet.cell_value(rowx=r, colx=c),
                                                easyxf('pattern: pattern solid_fill, fore_colour yellow;'))
                                col += 1
                            else:
                                editsheet.write(row, c+1, sheet.cell_value(rowx=r, colx=c))
                    elif c == 7:
                        try:
                            g = re.compile("([A-Z][A-Z])")
                            p = re.search(g, sheet.cell_value(rowx=r, colx=c))
                        except Exception:
                            editsheet.write(row, 0, sheet.cell_value(rowx=r, colx=c))
                        except AttributeError:
                            editsheet.write(row, 0, sheet.cell_value(rowx=r, colx=c))
                        else:
                            g = re.compile("([A-Z][A-Z])")
                            p = re.search(g, sheet.cell_value(rowx=r, colx=c))
                            print(r)
                            print(c)
                            editsheet.write(row, 0, p.group())
                    else:
                        print("BONK!")
                        editsheet.write(row, c+1, sheet.cell_value(rowx=r, colx=c))
                        col += 1
                row += 1
            else:
                print("OOOF!")

    working_month()
    working_year()
    writesheet(sheet, editsheet)
    wb.save(
        asksaveasfilename(defaultextension='.xls', initialfile='Required Annual Reviews Assessments (Processed).xls')
    )
    print("Processing Complete")
