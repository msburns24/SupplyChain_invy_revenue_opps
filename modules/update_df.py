import pandas as pd
import numpy as np

def update_df(df, field):
	'''
	Removes all negative/zero values from the dataframe, then sorts decending.
	Takes 2 inputs:
		df - The dataframe to be updated
		field - The field to sort descending
	Returns the update dataframe.
	'''
	return df[df[field] > 0].sort_values(field, ascending=False)