import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from pprint import pformat as pf
import config
import loadcsv
import auswertung
import logging
import traceback
import os

elements = 3
_h = 30
_px = 8
_py = 10
bh = 2
bw = 30
#size_x = 8 * bw + 2 * _px + 50
size_x = 260
size_y = elements*bh*_h + 2 * _py - 10
#size_y = 240

path = os.path.dirname(config.oildata_logfile)
if not os.path.exists(path):
    os.makedirs(path)

logging.basicConfig(filename=config.oildata_logfile, format='%(asctime)s %(levelname)s, %(message)s',level=logging.DEBUG)
logging.info('oildataCSV started.')
try:
    # create the root window
    root = tk.Tk()
    root.configure(bg='darkgrey')
    root.title('OildataCSV')
    root.resizable(False, False)
    root.geometry(f"{size_x}x{size_y}")

    #*************************************
    def select_file():
        filetypes = (('text files', '*.csv'), ('All files', '*.*'))
        filename = fd.askopenfilename(title='File öffnen', initialdir=config.initialdir, filetypes=filetypes)
        try:
            loadcsv.load(filename)
        except Exception as e:
            showinfo(title='Error',message=str(e))
            raise
    open_button = tk.Button(root, text='CSV File einlesen', command=select_file, height=bh, width=bw)
    #open_button.pack(side = tk.TOP, padx = _px, pady = _py)
    open_button.pack(side = tk.TOP, padx = _px, pady = _py)
    #*************************************

    #*************************************
    #oil = tk.StringVar(root, value='Jenbacher N Oil 40')
    def summary_select_file():
        filetypes = (('Excel files', '*.xlsx'),('Old Excel files', '*.xls'),('Alle files', '*.*'))
        filename = fd.askopenfilename(title='File öffnen', initialdir=config.initialdir, filetypes=filetypes)
        try:
            auswertung.summary(filename)
            #auswertung.summary(filename, oil.get())
        except Exception as e:
            showinfo(title='Fehler',message=str(e) + str(traceback.format_tb(e.__traceback__)))     
    # Unique Engines button
    #frame = tk.Frame(root, borderwidth=2, relief=tk.SUNKEN)
    #frame.pack()
    #oiltype_text = tk.Entry(frame, borderwidth=10, relief=tk.FLAT, textvariable=oil, justify='center', width=bw)
    #oiltype_text.pack(side = tk.TOP, padx = _px, pady = _py)
    unique_button = tk.Button(root, text="Zusammenfassung - Excel file", command=summary_select_file, height=bh, width=bw)
    unique_button.pack(side = tk.TOP, padx = _px, pady = _py)
    #*************************************

    #*************************************
    def quit():
        raise SystemExit
    # quit button
    quit_button = tk.Button(root, text='Ende', command=quit, height=bh, width=bw)
    quit_button.pack(side = tk.TOP, padx = _px, pady = _py)
    #*************************************

    # run the application
    root.mainloop()
except SystemExit as e:
    logging.info('OildataCSV exited.')
except Exception as e:
    logging.exception(str(e) + str(traceback.format_tb(e.__traceback__)))