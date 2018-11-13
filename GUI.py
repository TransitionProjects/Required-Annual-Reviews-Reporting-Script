__author__ = "dkatz"
__version__ = "rc2"

from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfilename
from AnnualReviewsReportRunnerv2 import processing

root = Tk()
root.title("Annual Reviews Report Runner")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

path = StringVar()
month = IntVar()
year = IntVar()

def process():
    processing(path=path.get(), m=month.get(), y=year.get())

def open():
    files = askopenfilename()
    path.set(files)

ttk.Label(mainframe, text="Path:").grid(column=1, row=1,sticky=W)
ttk.Entry(mainframe, textvariable=path, width=100).grid(column=1, row=2, columnspan=4, sticky=W)
ttk.Label(mainframe, text="Enter the month and year the report was run, not the month they are to be sent out.").grid(
    column=1, row=3, columnspan=4)
ttk.Label(mainframe, text="Month(M):",).grid(column=1, row=4, sticky=E)
ttk.Entry(mainframe, textvariable=month).grid(column=2, row=4, sticky=W)
ttk.Label(mainframe, text="Year(YYYY):").grid(column=3, row=4, sticky=E)
ttk.Entry(mainframe, textvariable=year).grid(column=4, row=4, sticky=W)

ttk.Button(mainframe, text="Process", command=process).grid(column=3, row=5, sticky=E)
ttk.Button(mainframe, text="Open", command=open).grid(column=4, row=5, sticky=E)

root.mainloop()
