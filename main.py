from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QPushButton, QWidget, QFrame, QFileDialog, QLineEdit, QHBoxLayout
from PyQt6 import QtCore
import pandas as pd

import subprocess
import copy
import sys
import os



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()


    def initUI(self):
        self.setWindowTitle("CIS-DEIB")
        self.setGeometry(200, 200, 400, 300)

        # Main widget
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)

        # Layout
        self.layout = QVBoxLayout(main_widget)

        # Upload panel
        self.upload = QPushButton("Upload Excel")
        self.upload.setMaximumWidth(200)
        self.upload.setMinimumWidth(100)
        self.upload.clicked.connect(self.upload_file)

        self.upload_panel = QWidget()
        self.upload_layout = QVBoxLayout(self.upload_panel)
        self.upload_layout.addWidget(self.upload)
        self.upload_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignHCenter)

        self.layout.addWidget(self.upload_panel)

        # Horizontal line
        line1 = QFrame()
        line1.setFrameShape(QFrame.Shape.HLine)
        line1.setFrameShadow(QFrame.Shadow.Sunken)
        self.layout.addWidget(line1)

        # Cohort panel
        self.cohort_panel = QWidget()
        self.cohort_layout = QVBoxLayout(self.cohort_panel)
        self.layout.addWidget(self.cohort_panel)

        # Horizontal line
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.HLine)
        line2.setFrameShadow(QFrame.Shadow.Sunken)
        self.layout.addWidget(line2)

        # Semester panel
        self.semester_panel = QWidget()
        self.semester_layout = QVBoxLayout(self.semester_panel)
        self.layout.addWidget(self.semester_panel)

        # Horizontal line
        line3 = QFrame()
        line3.setFrameShape(QFrame.Shape.HLine)
        line3.setFrameShadow(QFrame.Shadow.Sunken)
        self.layout.addWidget(line3)

        # Option panel
        self.option_panel = QWidget()
        self.option_layout = QVBoxLayout(self.option_panel)
        self.layout.addWidget(self.option_panel)
    

    def clear_panels(self):
        # Clear all panels 
        self.semester = None
        for i in reversed(range(self.cohort_layout.count())):
            widget = self.cohort_layout.itemAt(i).widget()
            widget.setParent(None)
        for i in reversed(range(self.semester_layout.count())):
            widget = self.semester_layout.itemAt(i).widget()
            widget.setParent(None)
        for i in reversed(range(self.option_layout.count())):
            widget = self.option_layout.itemAt(i).widget()
            widget.setParent(None)


    def upload_file(self):
        # handle file dialog
        file, _ = QFileDialog.getOpenFileName(self, "Upload Excel", "", "Excel Files (*.xlsx *.xls)")

        try:
            if file and file.endswith(('.xlsx', '.xls')):
                self.process_file(file)
            else: 
                raise Exception("Not an Excel file")

        except Exception as e:
            # Print exceptions in panel
            self.clear_panels()
            error_message = f"{type(e).__name__}: {str(e)}"
            err = QLineEdit(error_message)
            err.setReadOnly(True)
            self.cohort_layout.addWidget(err)


    def process_file(self, file):
        self.clear_panels()

        # Load in Excel
        self.file_path = file
        self.file = pd.ExcelFile(file)

        assert len(self.file.sheet_names) > 1

        # Set cohort buttons
        if len(self.file.sheet_names) > 1: 
            for i in range(1, len(self.file.sheet_names)):
                button = QPushButton("COHORT " + self.file.sheet_names[i])
                button.clicked.connect(self.cohort_download)
                self.cohort_layout.addWidget(button)

        # Set semester buttons
        if len(self.file.sheet_names) > 1: 
            for i in range(1, len(self.file.sheet_names)):
                button = QPushButton(self.file.sheet_names[i])
                button.setCheckable(True)
                button.clicked.connect(lambda _, b=button: self.semester_button_click(b))
                self.semester_layout.addWidget(button)

        # Set option buttons
        option_widget = QWidget()
        option_layouts = QHBoxLayout(option_widget)
        for option in ["First-Year", "Retention", "Changed Major", "Graduates"]:
            button = QPushButton(option)
            button.clicked.connect(self.option_download)
            option_layouts.addWidget(button)
        self.option_layout.addWidget(option_widget)

        # Set window title
        self.setWindowTitle("FILE: " + file)


    def semester_button_click(self, button):
        # Select semester and unselect other semester buttons
        for i in range(self.semester_layout.count()):
            widget = self.semester_layout.itemAt(i).widget()
            if isinstance(widget, QPushButton):
                widget.setChecked(False)
        button.setChecked(True)
        self.semester = button.text()


    def option_download(self):
        try:
            # Read option and semester
            option = self.sender().text()
            if not self.semester:
                return
            semester = self.file.sheet_names.index(self.semester)
            assert semester > 0

            # Get output path
            file_name = os.path.splitext(os.path.basename(self.file_path))[0]
            downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            output_path = os.path.join(downloads_dir, f"{file_name}_{self.semester}_{option}.xlsx".replace(" ", "_"))
            
            # filter non-citizen, non-undergrad, rename 'Academic Plan' to 'Major'
            dfs = [pd.read_excel(self.file, sheet_name=s) for s in self.file.sheet_names]
            for idx, df in enumerate(dfs):
                if 'US Citizen' in df.columns:
                    df = df[df['US Citizen'].fillna('').eq('Citizen')]
                if 'Class/Level' in df.columns:
                    df = df[~df['Class/Level'].str.contains('grad', case=False, na=False)]
                if 'Academic Plan' in df.columns:
                    df = df.rename(columns={'Academic Plan': 'Major'})
                dfs[idx] = df

            # get only relevant dfs
            prev_df = dfs[semester - 1]
            curr_df = dfs[semester]

            # see what option is cho
            match option:
                case "First-Year":
                    # get netids of students from previous semester with CS major
                    prev_ids = prev_df[prev_df['Major'].str.contains('Computer Science', case=False, na=False)]['Netid']
                    df = curr_df[
                        # currently in CS major
                        (curr_df['Major'].str.contains('Computer Science', case=False, na=False)) & 
                        # was not in CS major in the previous data
                        (~curr_df['Netid'].isin(prev_ids))
                    ]

                case "Retention":
                    # get netids of students from previous semester with CS major
                    prev_ids = prev_df[prev_df['Major'].str.contains('Computer Science', case=False, na=False)]['Netid']
                    df = curr_df[
                        # currently in CS major
                        (curr_df['Major'].str.contains('Computer Science', case=False, na=False)) & 
                        # was also in CS major in the previous data
                        (curr_df['Netid'].isin(prev_ids))
                    ]

                case "Changed Major":
                    # get netids of students from previous semester with CS major
                    prev_ids = prev_df[prev_df['Major'].str.contains('Computer Science', case=False, na=False)]['Netid']
                    df = curr_df[
                        # currently not in CS major
                        (~curr_df['Major'].str.contains('Computer Science', case=False, na=False)) & 
                        # was in CS major in the previous data
                        (curr_df['Netid'].isin(prev_ids))
                    ]

                case "Graduates":
                    # get netids of current students
                    curr_ids = curr_df['Netid']
                    df = prev_df[
                        # was in CS major in the previous data 
                        (prev_df['Major'].str.contains('Computer Science', case=False, na=False)) & 
                        # no longer a student this data cycle
                        (~prev_df['Netid'].isin(curr_ids))
                    ]

            # write to output
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                overview = df.groupby(["Ethnic CU", "Gender"]).size().unstack(fill_value=0)
                overview.to_excel(writer, sheet_name="Overview")

                df.to_excel(writer, sheet_name="All", index=False)

                for group in df["Ethnic CU"].unique():
                    group_df = df[df["Ethnic CU"] == group]
                    group_df.to_excel(writer, sheet_name=f"Group - {group}", index=False)

            subprocess.run(["osascript", "-e", f'tell application "Finder" to reveal POSIX file "{os.path.expanduser(output_path)}"'])
        
        except Exception as e:
            # Print exceptions in middle panel
            self.clear_panels()
            error_message = f"{type(e).__name__}: {str(e)}"
            err = QLineEdit(error_message)
            err.setReadOnly(True)
            self.cohort_layout.addWidget(err)

    def cohort_download(self):
        try:
            # Read option and semester
            cohort = self.sender().text()
            semester = " ".join(cohort.split(" ")[1:])
            semester = self.file.sheet_names.index(semester)

            # Get output path
            file_name = os.path.splitext(os.path.basename(self.file_path))[0]
            downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            output_path = os.path.join(downloads_dir, f"{file_name}_{cohort}.xlsx".replace(" ", "_"))
            
            # filter non-citizen, non-undergrad, rename 'Academic Plan' to 'Major'
            dfs = [pd.read_excel(self.file, sheet_name=s) for s in self.file.sheet_names]
            for idx, df in enumerate(dfs):
                if 'US Citizen' in df.columns:
                    df = df[df['US Citizen'].fillna('').eq('Citizen')]
                if 'Class/Level' in df.columns:
                    df = df[~df['Class/Level'].str.contains('grad', case=False, na=False)]
                if 'Academic Plan' in df.columns:
                    df = df.rename(columns={'Academic Plan': 'Major'})
                dfs[idx] = df
            
            save_dfs = []
            save_dfs_names = []

            # get first semesters
            prev_df = dfs[semester - 1]
            curr_df = dfs[semester]
            
            # get first-years
            prev_ids = prev_df[prev_df['Major'].str.contains('Computer Science', case=False, na=False)]['Netid']
            df = curr_df[
                # currently in CS major
                (curr_df['Major'].str.contains('Computer Science', case=False, na=False)) & 
                # was not in CS major in the previous data
                (~curr_df['Netid'].isin(prev_ids))
            ]

            # save a copy of the dataframe
            save_dfs.append(copy.deepcopy(df))
            save_dfs_names.append("FirstYear " + self.file.sheet_names[semester])

            # make sure all dfs only have the current first years
            first_ids = set(df['Netid']) if 'Netid' in df.columns else {}
            dfs = [df[df["Netid"].isin(first_ids)] for df in dfs]

            while(semester + 1 < len(self.file.sheet_names)):
                semester += 1

                prev_df = dfs[semester - 1]
                curr_df = dfs[semester]

                # Retention
                df = curr_df[
                    # still in CS major
                    (curr_df['Major'].str.contains('Computer Science', case=False, na=False))
                ]
                save_dfs.append(copy.deepcopy(df))
                save_dfs_names.append("Retention " + self.file.sheet_names[semester])

                # ChangeMajor
                prev_ids = prev_df[prev_df['Major'].str.contains('Computer Science', case=False, na=False)]['Netid']
                df = curr_df[
                    (~curr_df['Major'].str.contains('Computer Science', case=False, na=False)) &
                    # was in CS major in the previous data
                    (curr_df['Netid'].isin(prev_ids))
                ]
                save_dfs.append(copy.deepcopy(df))
                save_dfs_names.append("ChangedMajor " + self.file.sheet_names[semester])

                # Gradutes
                curr_ids = curr_df['Netid']
                df = prev_df[
                    (prev_df['Major'].str.contains('Computer Science', case=False, na=False)) & 
                    # no longer a student this data cycle
                    (~prev_df['Netid'].isin(curr_ids))
                ]
                save_dfs.append(copy.deepcopy(df))
                save_dfs_names.append("Gradutes " + self.file.sheet_names[semester])

            # anomoly detector
            # detect the following: 
            #   1. student not following typical path
            #   Example: retained --> graduated --> retained
            #   2. student appearing in multiple sheet per semester
            #   Example: appeared as both retained and graduated in one semester
            anomoly=[]
            for netids in first_ids:
                sheets = []
                for idx, save_df in enumerate(save_dfs):
                    if save_df['Netid'].eq(netids).any():
                        sheets.append(save_dfs_names[idx])

                options = [i.split(" ")[0] for i in sheets]
                semesters = [" ".join(i.split(" ")[1:]) for i in sheets]

                if len(semesters) != len(set(semesters)):
                    anomoly.append(f"{netids}: Student appears in multiple sheets in the same semester/year.")
                    
                state = 0  # 0: FirstYear, 1: Retention, 2: final (ChangedMajor or Gradutes)

                for o in options:
                    if state == 0:
                        if o == "FirstYear":
                            continue
                        elif o == "Retention":
                            state = 1
                        elif o in ("ChangedMajor", "Gradutes"):
                            state = 2
                        else:
                            anomoly.append(f"{netids}: [Error] Unknown sheet name encountered.")
                            break
                    elif state == 1:
                        if o == "Retention":
                            continue
                        elif o in ("ChangedMajor", "Gradutes"):
                            state = 2
                        else:
                            anomoly.append(f"{netids}: [Error] Invalid sequence encountered.")
                            break
                    elif state == 2:
                        anomoly.append(f"{netids}: [Error] Student was considered graduated but came back as undergrad or reappeared.")
                        break

            save_dfs.insert(1, pd.DataFrame(anomoly, columns=['Anomoly']))
            save_dfs_names.insert(1, 'Anomoly')

            # write to output
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                for i in range(len(save_dfs)):
                    sdf = save_dfs[i]

                    if "Ethnic CU" in sdf.columns and "Gender" in sdf.columns:
                        overview = sdf.groupby(["Ethnic CU", "Gender"]).size().unstack(fill_value=0)
                        overview.to_excel(writer, sheet_name="OV " + save_dfs_names[i])
                        sdf.to_excel(writer, sheet_name=save_dfs_names[i], index=False)
                    else:
                        sdf.to_excel(writer, sheet_name=save_dfs_names[i], index=False)

            subprocess.run(["osascript", "-e", f'tell application "Finder" to reveal POSIX file "{os.path.expanduser(output_path)}"'])

        except Exception as e:
            # Print exceptions in middle panel
            self.clear_panels()
            error_message = f"{type(e).__name__}: {str(e)}"
            err = QLineEdit(error_message)
            err.setReadOnly(True)
            self.cohort_layout.addWidget(err)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
