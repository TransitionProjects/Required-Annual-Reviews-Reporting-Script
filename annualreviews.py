__author__ = 'David Katz'
__version__ = 'rc1'

'''Currently this program throws and error when it encounters months Nov or Dec (as input by the user)  This will need
to be fixed to ensure that this program can be used without modifications in the future'''

from xlrd import open_workbook, xldate_as_tuple
from datetime import date, datetime, time
from xlutils.copy import copy
from xlwt import Workbook, easyxf
import re

rb = open_workbook(path)
sheet = rb.sheet_by_index(0)
wb = copy(rb)
editsheet = wb.add_sheet('Processed')
m = int(raw_input('Month(m):'))
y = int(raw_input('Year(yyyy):'))
month = 0
wmonth = []
year = 0
wyear = []


'''the workingMonth function advances the month by two respecting the 12 month year'''
def workingMonth():
    global wmonth
    global month
    if m == 11:
        month = 1
        wmonth.append(1)
    elif m == 12:
        month = 2
        wmonth.append(2)
    else:
        c = m + 2
        month = c
        wmonth.append(c)

'''the workingYear function advances they wyear variable by 1 in the event that the wmonth is Dec. or Nov.'''
def workingYear():
    global year
    if (month == 11) or (month == 12):
        year = y + 1
    else:
        year = y

'''dateCheck will reduce a date tuple to only the month data then look at a given date and return 'true' if the date's
month is equal to the working month'''
def dateCheck(d):
    global wmonth
    dlist = list(d)
    dlist.pop(0)
    dlist.pop(3)
    dlist.pop(2)
    dlist.pop(2)
    dlist.pop(1)
    #print dlist
    #print wmonth
    if dlist == wmonth:
        return True
    else:
        return False

def yearCheck(y):
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

'''dateConvert and yearConvery change excel date codes into date tuples so that python can read/process them'''
def dateConvert(r,c):
    #print xldate_as_tuple(sheet.cell(r,c).value, rb.datemode)
    date_value = xldate_as_tuple(sheet.cell(r,c).value, rb.datemode)
    return dateCheck(date_value)

def yearConvert(r,c):
    date_value = xldate_as_tuple(sheet.cell(r,c).value, rb.datemode)
    return yearCheck(date_value)

'''writeSheet writes data to the "processed" sheet'''
def writeSheet():
    row = 3
    header = 'Required Annual Review Assessments- by %d / %d' % (month, year)
    editsheet.write(0,0, header, easyxf('font: height 400;'))
    for c in range(0,sheet.ncols):
        if c == 7:
            editsheet.write(2, 0,sheet.cell_value(rowx = 2, colx=c))
        else:
            editsheet.write(2, c+1, sheet.cell_value(rowx = 2, colx=c))
    for r in range (3,sheet.nrows):
        if dateConvert(r,3) == True:
            print "POW!!!"
            col = 0
            for c in range(0,sheet.ncols):
                #print r
                #print c
                #print sheet.cell_value(rowx = r, colx = c)
                if (c == 4) and (sheet.cell_value(rowx = r, colx = 4) != ''):
                    if (yearConvert(r,4) == True) and (dateCheck(dateConvert(r,4)) == True):
                        #print "ping!"
                        editsheet.write(row,c+1,sheet.cell_value(rowx = r, colx = c),easyxf('pattern: pattern solid_fill,fore_colour yellow;' ))
                        col += 1
                    else:
                        editsheet.write(row,c+1,sheet.cell_value(rowx = r, colx = c))
                elif (c == 5):
                    if (yearConvert(r,5) == False) and (dateConvert(r,5) == True):
                        #print "pong!"
                        editsheet.write(row, c+1, sheet.cell_value(rowx = r, colx = c),easyxf('pattern: pattern solid_fill, fore_colour yellow;'))
                        col += 1
                    else:
                        editsheet.write(row,c+1,sheet.cell_value(rowx = r, colx = c))
                elif (c == 7):
                    try:
                        g = re.compile(ur'([A-Z][A-Z])')
                        p = re.search(g,sheet.cell_value(rowx = r, colx = c))
                        print p.group()
                    except Exception:
                        editsheet.write(row,0,sheet.cell_value(rowx = r, colx = c))
                    except AttributeError:
                        editsheet.write(row,0,sheet.cell_value(rowx = r, colx = c))
                    else:
                        g = re.compile(ur'([A-Z][A-Z])')
                        p = re.search(g,sheet.cell_value(rowx = r, colx = c))
                        print r
                        print c
                        editsheet.write(row,0,p.group())
                else:
                    print "BONK!"
                    editsheet.write(row,c+1,sheet.cell_value(rowx = r, colx = c))
                    col += 1
            row += 1
        else:
            print "OOOF!"

workingMonth()
workingYear()
writeSheet()
wb.save('Required Annual Reviews Assessments (Processed).xls')
