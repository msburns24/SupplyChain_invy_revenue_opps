from PyQt5 import QtWidgets as qtw
from PyQt5 import QtGui as qtg
from PyQt5 import QtCore as qtc
import pandas as pd
import numpy as np

def import_file(filename):
	try:
		raw_df = pd.read_excel(filename)
		return raw_df
	except PermissionError:
		print("ERROR: Please close program before running script.")
		return 1
	except Exception as error:
		print(f"An upexpected error has occured: {error}")
		return 2