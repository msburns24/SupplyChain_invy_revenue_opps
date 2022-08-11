import pandas as pd
import numpy as np

dev_mode = True

# Define update functions
def update_df(df, field):
	'''
	Removes all negative/zero values from the dataframe, then sorts decending.
	Takes 2 inputs:
		df - The dataframe to be updated
		field - The field to sort descending
	Returns the update dataframe.
	'''
	return df[df[field] > 0].sort_values(field, ascending=False)


# Establish DF
if dev_mode:
	raw_df = pd.read_excel("invy_bklg_example.xlsx")
else:
	raw_df = pd.read_excel("invy_bklg.xlsx")
raw_df["plant_material"] = raw_df.plant.map(str) + "_" + raw_df.material.map(str)
raw_df = raw_df.set_index("plant_material")

# Setup Needs/Excess DF's
raw_df["excess"] = raw_df["invy"] - raw_df["bklg_qty"]
raw_df["need"] = raw_df["bklg_qty"] - raw_df["invy"]
needs_df = raw_df[["material", "plant", "need"]]
excess_df = raw_df[["material", "plant", "excess"]]

# Update the needs & excess dataframes
needs_df = update_df(needs_df, "need")
excess_df = update_df(excess_df, "excess")

# Loop through every plant/PN combo under "needs"
shares = pd.DataFrame(columns = ["PN", "Needs Plant", "Excess Plant", "Share Qty"])

while True:
	# Check if needs DF is empty
	if needs_df.empty:
		break

	# Break out top PN/Plant/Qty from needs
	n_pn = needs_df.iloc[0]["material"]
	n_plant = needs_df.iloc[0]["plant"]
	n_qty = needs_df.iloc[0]["need"]
	excess_df_sel = excess_df[excess_df["material"] == n_pn]
	
	# Check if excess sub-df (for 1st PN) is empty
	if excess_df_sel.empty:
		needs_df = needs_df[needs_df["material"] != n_pn] # Remove 
		needs_df = update_df(needs_df, "need")
		continue
	else:
		# Get the plant/qty with the highest stock of the needed PN
		e_plant = excess_df_sel.iloc[0]["plant"]
		e_qty = excess_df_sel.iloc[0]["excess"]

		# If the excess plant has more than enough, share what's needed
		if e_qty > n_qty:
			# Make note of how much can be shared between plants
			share_qty = n_qty

			# Update values in needs/excess df's
			excess_df.at[(str(e_plant) + "_" + n_pn), "excess"] = int(e_qty) - int(n_qty) # Subtract shared qty from excess DF
			needs_df.at[str(n_plant) + "_" + n_pn, "need"] = 0 # Set needs qty to zero, as all needs have been sent

		# If the excess plant doesn't have enough, share everything
		else:
			# Make note of how much can be shared between plants
			share_qty = e_qty

			# Update values in needs/excess df's
			excess_df.at[(str(e_plant) + "_" + n_pn), "excess"] = 0 # Take all of the excess stock
			needs_df.at[str(n_plant) + "_" + n_pn, "need"] = int(n_qty) - int(e_qty) # Subtract needs qty from needs df

		# Make note of how much can be shared between plants
		n_plant_e_plant_pn = str(n_plant) + "-" + str(e_plant) + "-" + str(n_pn)
		new_row = [n_pn, str(n_plant), str(e_plant), share_qty]

		# shares = pd.concat([shares, pd.DataFrame(new_row)], ignore_index=True)
		shares.loc[len(shares.index)] = new_row

		# Update df's
		excess_df = update_df(excess_df, "excess")
		needs_df = update_df(needs_df, "need")



if dev_mode:
	print(shares)
else:
	shares.to_excel("output.xlsx")