import sys
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtGui as qtg
from PyQt5 import QtCore as qtc
import pandas as pd
import numpy as np

# To-Do
#  - Refactor "run_program" function
#  - Add progress window when running program
#  - Add threading to show window


version = "0.4"

class InvyCalculator(qtc.QObject):

    error_signal = qtc.pyqtSignal(str, str)
    file_imported_signal = qtc.pyqtSignal(str)
    program_run_signal = qtc.pyqtSignal()
    template = pd.DataFrame(
        {
        'material': [],
        'plant': [],
        'invy': [],
        'bklg_qty': [],
        'bklg_val': []
        })


    def __init__(self):
        super().__init__()

    def update_df(self, df, field):
        return df[df[field] > 0].sort_values(field, ascending=False)

    def download_template(self):
        default_fn = str(qtc.QDir.currentPath()) + "/Inventory Revenue Template.xlsx"
        filename, _ = qtw.QFileDialog.getSaveFileName(
            None,
            "Save File",
            default_fn,
            'Microsoft Excel Workbook (*.xlsx);;Text Files (*.txt);;All Files (*)'
            )

        if filename:
            extension = filename.split(".")[-1]
            if extension not in ["xlsx", "xls"]:
                error_msg = "ERROR: Please choose a filename ending in (.xls, .xlsx)."
                self.error_signal.emit("Download Error", error_msg)
                return 1

            try:
                self.template.to_excel(filename, index=False)
            except PermissionError:
                error_msg = "ERROR: Please close file before saving."
                self.error_signal.emit("Permission Error", error_msg)
                return 2
            except Exception as error:
                self.error_signal.emit("Error", f"An upexpected error has occured: {error}")
                return 3



    def import_file(self):
        filename, _ = qtw.QFileDialog.getOpenFileName(
            None,
            "Select an excel file to open...",
            qtc.QDir.currentPath(),
            'All Files (*)'
            )

        if filename:
            # Check that file type is .xls or .xlsx
            extension = filename.split(".")[-1]
            if extension not in ["xls", "xlsx"]:
                error_msg = "ERROR: Please choose an excel file (*.xls, *.xlsx)."
                self.error_signal.emit("Import Error", error_msg)
                return 3

            try:
                self.raw_df = pd.read_excel(filename)
            except PermissionError:
                error_msg = "ERROR: Please close program before running script."
                self.error_signal.emit("Permission Error", error_msg)
                return 1
            except Exception as error:
                self.error_signal.emit("Error", f"An upexpected error has occured: {error}")
                return 2

            # Check columns match template
            if list(self.raw_df.columns) != ['material', 'plant', 'invy', 'bklg_qty', 'bklg_val']:
                error_msg = "ERROR: Columns do not match template. Please use correct template when uploading file."
                self.error_signal.emit("File Template Error", error_msg)
                return 3

            self.file_imported_signal.emit(filename)

    def run_program(self):
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

        # Signal that program has run successfully
        self.program_run_signal.emit()

    def export_file(self):
        default_fn = str(qtc.QDir.currentPath()) + "/Inventory Revenue Opportunities.xlsx"
        filename, _ = qtw.QFileDialog.getSaveFileName(
            None,
            "Save File",
            default_fn,
            'Microsoft Excel Workbook (*.xlsx);;Text Files (*.txt);;All Files (*)'
            )
        if filename:
            try:
                self.shares.to_excel(filename, index=False)
            except Exception as error:
                self.error_signal.emit("Error", error)


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
        self.setGeometry(400,400,600,300)
        self.setWindowTitle("Inventory & Revenue Calculator")

        # Create main widget
        self.main = qtw.QWidget()
        self.main.setLayout(qtw.QFormLayout())
        self.setCentralWidget(self.main)

        ###############
        # Add widgets #
        ###############

        # Program header
        self.header_label = qtw.QLabel("Inventory/Revenue Calculator")
        self.header_label.setAlignment(qtc.Qt.AlignCenter)
        self.main.layout().addRow(self.header_label)

        # Program Description
        description = """
        Calculates the inventory available to transfer to increase revenue and improve
        delivery. Takes an input file containing Part Number, Plant, Inventory, Backlog,
        and Backlog Value to calculate which plants have needs/excess compared to their
        inventory levels.
        """
        self.descr_label = qtw.QLabel(description)
        self.descr_label.setAlignment(qtc.Qt.AlignLeft)
        self.main.layout().addRow(self.descr_label)
        self.main.layout().addRow(qtw.QLabel()) # Blank row for separation

        # Steps header
        self.steps_label = qtw.QLabel("Instructions:")
        self.main.layout().addRow(self.steps_label)

        # Download Template Button
        self.dwnld_template_btn = qtw.QPushButton("Download Template")
        self.dwnld_template_btn.clicked.connect(self.calculator.download_template)
        self.main.layout().addRow("1. Download excel template ", self.dwnld_template_btn)

        # Import File Button
        self.import_btn = qtw.QPushButton("Import File")
        self.import_btn.clicked.connect(self.calculator.import_file)
        self.main.layout().addRow("2. Import populated template ", self.import_btn)

        # Run Program Button
        self.run_btn = qtw.QPushButton("Run Program")
        self.run_btn.setDisabled(True)
        self.run_btn.clicked.connect(self.calculator.run_program)
        self.main.layout().addRow("3. Run Program", self.run_btn)

        # Export File Button
        self.export_btn = qtw.QPushButton("Export File")
        self.export_btn.setDisabled(True)
        self.export_btn.clicked.connect(self.calculator.export_file)
        self.main.layout().addRow("4. Export File", self.export_btn)

        # Footer Label
        self.main.layout().addRow(qtw.QLabel())
        footer_msg = "Version 0.0.4. Maintained by Matt Burns (matthew.burns@te.com)."
        footer_msg += " Please contact for any reported bugs."
        self.footer_label = qtw.QLabel(footer_msg)
        self.footer_label.setAlignment(qtc.Qt.AlignCenter)
        self.main.layout().addRow(self.footer_label)
        
        self.setup_stylesheet()
        self.show()

    def setup_stylesheet(self):
        # Form layout
        self.main_stylesheet = """
            QWidget {
                background-color: #D0D0D0;
                font-size: 16px;
                font-family: segoe ui, sans;
                color: black;
                text-align: center;
            }
            
            QPushButton {
                background-color: #EEEEEE;
                font-size: 18px;
                height: 40px;
                border: 1px solid #A0A0A0;
                border-radius: 4px;
            }

            QPushButton:disabled {
                background-color: #DDDDDD;
                color: #B0B0B0;
            }

            QPushButton:hover {
                background-color: #DDEBF7;
                border: 1px solid #5B9BD5;
            }
        """
        self.main.setStyleSheet(self.main_stylesheet)

        # Header Label
        self.header_stylesheet = """
            font-size: 28px;
            font-family: segoe ui light, sans;
            text-align: middle;
        """
        self.header_label.setStyleSheet(self.header_stylesheet)

        # Steps Label
        self.steps_stylesheet = """
            font-weight: bold;
            font-size: 18px;
        """
        self.steps_label.setStyleSheet(self.steps_stylesheet)

        # Footer Label
        self.footer_stylesheet = """
            font-size: 12px;
        """
        self.footer_label.setStyleSheet(self.footer_stylesheet)



    def print_error(self, error_title, error_msg):
        qtw.QMessageBox.critical(self, error_title, error_msg)

    def program_ready(self, filename):
        filename_short = filename.split("/")[-1]
        self.import_btn.setText(filename_short)
        self.run_btn.setDisabled(False)

    def program_run(self):
        self.export_btn.setDisabled(False)


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    mw = MainWindow()
    sys.exit(app.exec())