import sys
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtGui as qtg
from PyQt5 import QtCore as qtc
import pandas as pd
import numpy as np


# To-Do
#  - Add dialog box to select custom file to import
#  - Refactor "run_program" function

class InvyCalculator(qtc.QObject):

    error_signal = qtc.pyqtSignal(str)
    file_imported_signal = qtc.pyqtSignal(str)
    program_run_signal = qtc.pyqtSignal()

    def __init__(self):
        super().__init__()

    def update_df(self, df, field):
        return df[df[field] > 0].sort_values(field, ascending=False)

    def import_file(self):
        filename = "invy_bklg_example.xlsx"
        # Need to run a dialog in frontend to grab filename, then pass it into this function

        try:
            self.raw_df = pd.read_excel(filename)
            self.file_imported_signal.emit(filename)
            return 0
        except PermissionError:
            self.error_signal.emit("ERROR: Please close program before running script.")
            return 1
        except Exception as error:
            self.error_signal.emit(f"An upexpected error has occured: {error}")
            return 2

    def run_program(self):
        # Check columns match template
        if list(self.raw_df.columns) != ['material', 'plant', 'invy', 'bklg_qty', 'bklg_val']:
            error = "ERROR: Columns do not match template. Please use correct template when uploading file."
            self.error_signal.emit(error)
            return 3

        # Clean DF
        self.raw_df["plant_material"] = self.raw_df.plant.map(str) + "_" + self.raw_df.material.map(str)
        self.raw_df = self.raw_df.set_index("plant_material")

        # Setup Needs/Excess DF's
        self.raw_df["excess"] = self.raw_df["invy"] - self.raw_df["bklg_qty"]
        self.raw_df["need"] = self.raw_df["bklg_qty"] - self.raw_df["invy"]
        needs_df = self.raw_df[["material", "plant", "need"]]
        excess_df = self.raw_df[["material", "plant", "excess"]]

        # Update the needs & excess dataframes
        needs_df = self.update_df(needs_df, "need")
        excess_df = self.update_df(excess_df, "excess")

        # Loop through every plant/PN combo under "needs"
        self.shares = pd.DataFrame(columns = ["PN", "Needs Plant", "Excess Plant", "Share Qty"])

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
                needs_df = self.update_df(needs_df, "need")
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

                # self.shares = pd.concat([self.shares, pd.DataFrame(new_row)], ignore_index=True)
                self.shares.loc[len(self.shares.index)] = new_row

                # Update df's
                excess_df = self.update_df(excess_df, "excess")
                needs_df = self.update_df(needs_df, "need")

        print(self.shares)
        # self.shares.to_excel("output.xlsx")
        self.program_run_signal.emit()

    def export_file(self):
        self.shares.to_excel("output.xlsx")


class MainWindow(qtw.QMainWindow):

    def __init__(self):
        """Main Window Constructor"""
        super().__init__()
        self.calculator = InvyCalculator()
        self.calculator.error_signal.connect(self.print_error)
        self.calculator.file_imported_signal.connect(self.program_ready)
        self.calculator.program_run_signal.connect(self.program_run)

        self.setup_UI()




    def setup_UI(self):
        self.setGeometry(400,400,400,300)
        self.setWindowTitle("Inventory & Revenue Calculator")

        # Create main widget
        main = qtw.QWidget()
        main.setLayout(qtw.QFormLayout())
        self.setCentralWidget(main)

        ###############
        # Add widgets #
        ###############

        # Import File Button
        self.import_btn = qtw.QPushButton("Import File")
        self.import_btn.clicked.connect(self.calculator.import_file)
        main.layout().addRow("1. Import file ", self.import_btn)

        # Run Program Button
        self.run_btn = qtw.QPushButton("Run Program")
        self.run_btn.setDisabled(True)
        self.run_btn.clicked.connect(self.calculator.run_program)
        main.layout().addRow("2. Run Program", self.run_btn)

        # Export File Button
        self.export_btn = qtw.QPushButton("Export File")
        self.export_btn.setDisabled(True)
        self.export_btn.clicked.connect(self.calculator.export_file)
        main.layout().addRow("3. Export File", self.export_btn)
        
        self.show()

    def print_error(self, error_str):
        print(error_str)

    def program_ready(self, filename):
        self.import_btn.setText(filename)
        self.run_btn.setDisabled(False)

    def program_run(self):
        self.export_btn.setDisabled(False)


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    mw = MainWindow()
    sys.exit(app.exec())