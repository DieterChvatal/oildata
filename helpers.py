import pandas as pd
import os
import re
import config
import logging
from chardet.universaldetector import UniversalDetector
import csv


# regular Expressions for the corr function.
date_pattern = "^((\d{2}[-,/,\.]){2}\d{2,4}\s(\d{2}:){1,2}\d{2})(.*)" #matches all date constructs in spectro csv
date_pat = re.compile(date_pattern)
num_pattern = "^[<* ]{0,2}([+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?)[\?\*\/\+]*$" #matches all number like constructs in spectro csv
num_pat = re.compile(num_pattern)

# Umwandels der Datenwerte in Zahlen, z.B. Entfernen von '<' bei '<0.1'
def corr(s):
    # s nur dann behandeln, wenn es ein String ist.  
    if type(s) == str:
        # Leere Felder , Feldee mit '-' und schon vorher leere Felder als None zurücgeben.
        # das bedeutet, dass Excel ein leeres Feld an der Stelle anzeigt.
        if s in ['', '-', None]:
            return None
        if num_pat.match(s):
            try:
                s = float(num_pat.match(s).group(1))
                return s
            except TypeError as err:
                print(f"{s} raised {str(err)}")
                logging.info(f"applymap corr - regex number , {s} raises {str(err)}")
                raise
        if date_pat.match(s):
            try:
                s = pd.to_datetime(date_pat.match(s).group(1))
                return s
            except ValueError as err:
                logging.info(f"applymap corr - regex date , {s} raises {str(err)}")
                print(f"{s} raised {str(err)}")
                raise
    return s # s has already a different type than str.  

# die Spectro Datei in ein temporäres File kopieren
# dabei die Datei Kodierung auf utf-8 umstellen
# und Fehler beseitigen
# 1.) - Anführungszeichen vor und nach einer Zeile entfernen.
# ...
def checkfile(filename):
    if os.path.exists(config.tempfile):
        os.remove(config.tempfile)

    ae = analyze_file(filename)

    with open(filename, "r", encoding=ae['encoding']) as in_file, open(config.tempfile, 'w', encoding='utf-8') as out_file:
        lines = (l for l in in_file)
        try:
            for line in lines:
                oline = line.replace('\"','')
                out_file.write(oline)
        except Exception as err:
            logging.error(f"Checkfile Error: {str(err)}")
            raise
    return config.tempfile

def analyze_file(filename):

    detector = UniversalDetector()
    with open(filename, 'rb') as check_file:
        lines = (l for l in check_file)
        for line in lines:
            detector.feed(line)
            if detector.done: break
    detector.close()
    det_encoding = detector.result['encoding']

    with open(filename, 'r', newline="", encoding=det_encoding) as csvfile:
        dialect = csv.Sniffer().sniff(csvfile.read(1024), delimiters=";,")

    return {'encoding': det_encoding, 'delimiter': dialect.delimiter} 