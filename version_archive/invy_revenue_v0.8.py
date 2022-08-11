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

version = "0.5"

class InvyCalculator(qtc.QObject):

    error_signal = qtc.pyqtSignal(str, str)
    share_status = qtc.pyqtSignal(str)
    file_imported_signal = qtc.pyqtSignal(str, int)
    opps_remain = qtc.pyqtSignal(int, int)
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
        # Recalculate $ values
        field_val = field + "_val"
        if field_val in df.columns:
            df = df.drop([field_val], axis=1)
        df[field_val] = df[field] * df["piece_price"]

        # Sort & Return list
        return df[df[field] > 0].sort_values(field_val, ascending=False)

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

        self.share_status.emit("Template downloaded.")



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
            self.share_status.emit("Checking extension...")
            if extension not in ["xls", "xlsx"]:
                error_msg = "ERROR: Please choose an excel file (*.xls, *.xlsx)."
                self.error_signal.emit("Import Error", error_msg)
                return 3

            self.share_status.emit("Reading file...")
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
            self.share_status.emit("Checking template...")
            if list(self.raw_df.columns) != ['material', 'plant', 'invy', 'bklg_qty', 'bklg_val']:
                error_msg = "ERROR: Columns do not match template. Please use correct template when uploading file."
                self.error_signal.emit("File Template Error", error_msg)
                return 3

        # Clean DF
        self.share_status.emit("Cleaning dataframe...")
        self.raw_df["plant_material"] = self.raw_df.plant.map(str) + "_" + self.raw_df.material.map(str)
        self.raw_df = self.raw_df.set_index("plant_material")

        # Create piece value lookup
        self.piece_prices = self.calculate_piece_price(self.raw_df)

        # Setup Needs/Excess DF's
        self.share_status.emit("Creating needs and excess dataframes...")
        self.raw_df["excess"] = self.raw_df["invy"] - self.raw_df["bklg_qty"]
        self.raw_df["need"] = self.raw_df["bklg_qty"] - self.raw_df["invy"]
        self.needs_df = self.raw_df[["material", "plant", "need"]]
        self.needs_df = pd.merge(self.needs_df, self.piece_prices, on="material")
        self.needs_df["plant_material"] = self.needs_df.plant.map(str) + "_" + self.needs_df.material.map(str)
        self.needs_df = self.needs_df.set_index("plant_material")

        self.excess_df = self.raw_df[["material", "plant", "excess"]]
        self.excess_df = pd.merge(self.excess_df, self.piece_prices, on="material")
        self.excess_df["plant_material"] = self.excess_df.plant.map(str) + "_" + self.excess_df.material.map(str)
        self.excess_df = self.excess_df.set_index("plant_material")

        # Update the needs & excess dataframes
        self.share_status.emit("Updating dataframes...")
        self.needs_df = self.update_df(self.needs_df, "need")
        self.excess_df = self.update_df(self.excess_df, "excess")

        opportunities = len(self.needs_df.index)
        self.file_imported_signal.emit(filename, opportunities)
        self.share_status.emit("File imported successfully.")


    def calculate_piece_price(self, df):
        self.share_status.emit("Calculating piece prices...")
        pp_df = pd.pivot_table(df, index=["material"], values = ["bklg_qty", "bklg_val"], aggfunc=np.sum)
        pp_df["piece_price"] = pp_df["bklg_val"] / pp_df["bklg_qty"]
        pp_df = pp_df.drop(["bklg_qty", "bklg_val"], axis=1)
        pp_df["piece_price"] = pp_df["piece_price"].replace([np.inf], 0)
        pp_df.to_excel("pp_df.xlsx")
        return pp_df

    def run_program(self):

        # Loop through every plant/PN combo under "needs"
        self.shares = pd.DataFrame(columns = ["PN", "Needs Plant", "Excess Plant", "Share Qty", "Share Value"])
        self.share_status.emit("Checking records...")

        while True:
            # Check if needs DF is empty
            if self.needs_df.empty:
                break

            # Break out top PN/Plant/Qty from needs
            n_pn = self.needs_df.iloc[0]["material"]
            n_plant = self.needs_df.iloc[0]["plant"]
            n_qty = self.needs_df.iloc[0]["need"]
            self.excess_df_sel = self.excess_df[self.excess_df["material"] == n_pn]
            
            # Check if excess sub-df (for 1st PN) is empty
            if self.excess_df_sel.empty:
                self.needs_df = self.needs_df[self.needs_df["material"] != n_pn] # Remove 
                self.needs_df = self.update_df(self.needs_df, "need")

                needs_left = len(self.needs_df.index)
                revenue_gen = int(self.shares["Share Value"].sum())
                self.opps_remain.emit(needs_left, revenue_gen)
                continue
            else:
                # Get the plant/qty with the highest stock of the needed PN
                e_plant = self.excess_df_sel.iloc[0]["plant"]
                e_qty = self.excess_df_sel.iloc[0]["excess"]

                # If the excess plant has more than enough, share what's needed
                if e_qty > n_qty:
                    # Make note of how much can be shared between plants
                    share_qty = n_qty

                    # Update values in needs/excess df's
                    self.excess_df.at[(str(e_plant) + "_" + n_pn), "excess"] = int(e_qty) - int(n_qty) # Subtract shared qty from excess DF
                    self.needs_df.at[str(n_plant) + "_" + n_pn, "need"] = 0 # Set needs qty to zero, as all needs have been sent

                # If the excess plant doesn't have enough, share everything
                else:
                    # Make note of how much can be shared between plants
                    share_qty = e_qty

                    # Update values in needs/excess df's
                    self.excess_df.at[(str(e_plant) + "_" + n_pn), "excess"] = 0 # Take all of the excess stock
                    self.needs_df.at[str(n_plant) + "_" + n_pn, "need"] = int(n_qty) - int(e_qty) # Subtract needs qty from needs df

                # Make note of how much can be shared between plants
                share_val = share_qty * self.piece_prices["piece_price"][n_pn]
                n_plant_e_plant_pn = str(n_plant) + "-" + str(e_plant) + "-" + str(n_pn)
                new_row = [n_pn, str(n_plant), str(e_plant), share_qty, share_val]

                self.shares.loc[len(self.shares.index)] = new_row

                # Update df's
                self.excess_df = self.update_df(self.excess_df, "excess")
                self.needs_df = self.update_df(self.needs_df, "need")

                # Emit progress
                needs_left = len(self.needs_df.index)
                revenue_gen = int(self.shares["Share Value"].sum())
                self.opps_remain.emit(needs_left, revenue_gen)

        # Signal that program has run successfully
        self.share_status.emit("Program complete.")
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
        self.calculator.opps_remain.connect(self.update_progress)
        self.calculator.program_run_signal.connect(self.program_run)
        self.calculator.share_status.connect(self.update_status)

        # Create threading
        self.calc_thread = qtc.QThread()
        self.calculator.moveToThread(self.calc_thread)
        self.calc_thread.start()

        self.setup_UI()

    def setup_UI(self):
        self.setGeometry(400,400,600,300)
        self.setWindowTitle("Inventory & Revenue Calculator")
        self.setWindowIcon(qtg.QIcon("images/calculator.png"))

        # Create main widget
        self.main = qtw.QWidget()
        self.main.setLayout(qtw.QGridLayout())
        self.setCentralWidget(self.main)

        ###############
        # Add widgets #
        ###############

        # Program header
        self.header_label = qtw.QLabel("Inventory/Revenue Calculator")
        self.header_label.setAlignment(qtc.Qt.AlignCenter)
        self.main.layout().addWidget(self.header_label, 0, 0, 1, 2)

        # Program Description
        description = """
        Calculates the inventory available to transfer to increase revenue and improve
        delivery. Takes an input file containing Part Number, Plant, Inventory, Backlog,
        and Backlog Value to calculate which plants have needs/excess compared to their
        inventory levels.
        """
        self.descr_label = qtw.QLabel(description)
        self.descr_label.setAlignment(qtc.Qt.AlignLeft)
        self.main.layout().addWidget(self.descr_label, 1, 0, 1, 2)

        # Steps header
        self.steps_label = qtw.QLabel("Instructions:")
        self.main.layout().addWidget(self.steps_label, 2, 0, 1, 2)

        # Download Template 
        self.dwnld_template_lbl = qtw.QLabel("1. Download excel template")
        self.main.layout().addWidget(self.dwnld_template_lbl, 3, 0)
        self.dwnld_template_btn = qtw.QPushButton("Download Template")
        self.dwnld_template_btn.clicked.connect(self.calculator.download_template)
        self.dwnld_template_btn.setIcon(qtg.QIcon("images/download.png"))
        self.dwnld_template_btn.setIconSize(qtc.QSize(30,30))
        self.main.layout().addWidget(self.dwnld_template_btn, 3, 1)

        # Import File
        self.import_lbl = qtw.QLabel("2. Import populated template ")
        self.main.layout().addWidget(self.import_lbl, 4, 0)
        self.import_btn = qtw.QPushButton("Import File")
        self.import_btn.clicked.connect(self.calculator.import_file)
        self.import_btn.setIcon(qtg.QIcon("images/upload.png"))
        self.import_btn.setIconSize(qtc.QSize(30,30))
        self.main.layout().addWidget(self.import_btn, 4, 1)

        # Run Program Button
        self.run_lbl = qtw.QLabel("3. Run Program")
        self.main.layout().addWidget(self.run_lbl, 5, 0)
        self.run_btn = qtw.QPushButton("Run Program")
        self.run_btn.setDisabled(True)
        self.run_btn.clicked.connect(self.calculator.run_program)
        self.run_btn.setIcon(qtg.QIcon("images/calculator.png"))
        self.run_btn.setIconSize(qtc.QSize(30,30))
        self.main.layout().addWidget(self.run_btn, 5, 1)

        # Export File Button
        self.export_lbl = qtw.QLabel("4. Export File")
        self.main.layout().addWidget(self.export_lbl, 6, 0)
        self.export_btn = qtw.QPushButton("Export File")
        self.export_btn.setDisabled(True)
        self.export_btn.clicked.connect(self.calculator.export_file)
        self.export_btn.setIcon(qtg.QIcon("images/download.png"))
        self.export_btn.setIconSize(qtc.QSize(30,30))
        self.main.layout().addWidget(self.export_btn, 6, 1)

        # Progress Label
        self.progress_label = qtw.QLabel("Opportunities Searched: 0 / 0")
        self.main.layout().addWidget(self.progress_label, 7, 0, 1, 2)

        # Revenue Label
        self.revenue_label = qtw.QLabel("Revenue Found: $0")
        self.main.layout().addWidget(self.revenue_label, 8, 0, 1, 2)
        
        # Progress Bar
        self.progress_bar = qtw.QProgressBar()
        self.progress_bar.setValue(0)
        self.main.layout().addWidget(self.progress_bar, 9, 0, 1, 2)

        # Footer Label
        footer_msg = "Version 0.0.4. Maintained by Matt Burns (matthew.burns@te.com)."
        footer_msg += " Please contact for any reported bugs."
        self.footer_label = qtw.QLabel(footer_msg)
        self.footer_label.setAlignment(qtc.Qt.AlignCenter)
        self.main.layout().addWidget(self.footer_label, 10, 0, 1, 2)

        # Status bar
        self.status_bar = qtw.QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Inventory & Revenue Calculator")
        
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

        # Progress Label
        self.progress_label_stylesheet = """
            font-size: 18px;
            color: #A0A0A0;
        """
        self.progress_label.setStyleSheet(self.progress_label_stylesheet)
        self.revenue_label.setStyleSheet(self.progress_label_stylesheet)

        # Footer Label
        self.footer_stylesheet = """
            font-size: 12px;
        """
        self.footer_label.setStyleSheet(self.footer_stylesheet)



    def print_error(self, error_title, error_msg):
        qtw.QMessageBox.critical(self, error_title, error_msg)

    def program_ready(self, filename, opps):
        filename_short = filename.split("/")[-1]
        self.import_btn.setText(filename_short)
        self.run_btn.setDisabled(False)
        self.opportunities = opps
        self.progress_label.setText("Opportunities Searched: 0 / " + "{:,}".format(opps))
        self.progress_label_stylesheet = """
            font-size: 18px;
            color: black;
        """
        self.progress_label.setStyleSheet(self.progress_label_stylesheet)
        self.revenue_label.setStyleSheet(self.progress_label_stylesheet)

    def update_progress(self, n_remain, rev_gen):
        # Update progress through opp's
        n_done = self.opportunities - n_remain
        perc_done = int(100 * (n_done / self.opportunities))
        progress_str = "Opportunities Searched: "
        progress_str += "{:,}".format(n_done) + " / "
        progress_str += "{:,}".format(self.opportunities)
        self.progress_label.setText(progress_str)
        self.progress_bar.setValue(perc_done)

        # Update revenue generated
        rev_gen_str = "${:,}".format(rev_gen)
        rev_gen_msg = "Revenue Found: " + rev_gen_str
        self.revenue_label.setText(rev_gen_msg)

    def update_status(self, msg):
        self.status_bar.showMessage(msg)

    def program_run(self):
        self.export_btn.setDisabled(False)


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    mw = MainWindow()
    sys.exit(app.exec())