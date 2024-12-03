# (MacOS) Download the App and Run
To download the App, you must open the Terminal and enter the following:
```
curl -L https://github.com/CornellCAC/CIS-DEIB/releases/download/Release/main.zip --output ~/Desktop/main.zip
```
A file called `main.zip` will appear on your desktop. Double-click to unzip; you should see `main` on your desktop.


# Terminology

* First-Year: A First-Year CS major is a student who was not affiliated with the CS major in the previous semester but is now affiliated. For example, First-Year students in FA2021 were unaffiliated in FA2021 but affiliated in FA2022.

* Retention: A retained CS major is a student who was affiliated with the CS major in the previous semester and is still affiliated. For example, retained students in FA2023 were affiliated with both FA2022 and FA2023. 

* Changed Major: A student is considered to have changed major if they were affiliated with the CS major in the immediate previous semester but no longer in the CS major in the selected semester. For example, a student who was in CS in FA2022 but is in FA2023 is a math major and is considered to have changed majors in FA2023. 

* Graduates: A student is considered a graduate if they appeared as a CS major in the previous semester but are no longer a student in the current semester. For example, a student who graduated in Winter 2022 or Spring 2023 is considered a graduate in  FA2023.
  
* Cohort: A cohort is a group of students affiliated in the same semester. 


# Extra Notes
1. The Excel sheet must have `Netid`, `Gender`, `Ethnic CU`, `Class/Level`, and either `Academic Plan` or `Major` (not both) in each worksheet.

1. Worksheets should be ordered in chronological order. 

1. If there is a 'US Citizen' column, Only students with 'Citizen' are tracked and counted.

1. Students should have consistent NetIDs between semesters.


# How to Build the Application from Scratch

## Install Python
This App was built with Python 3.13.0 and pip 24.2, but older versions of Python should also work.

## Install Packages
### MacOS/Windows
With Python installed, type the following in the a terminal, 
```
pip install -r requirements.txt
```
Windows should may substitute `pip` with `py3 -m pip` if needed.

## Build the Application
### MacOS/Windows
Type the following
```
pyinstaller --clean --noconfirm --onefile --windowed main.py
```
The app will be in the `dist` folder. 
### MacOS: Distributing the app
If you want to share the app, you must codesign the application. An example is shown in `build.sh`. You may change the signature with your own, and do:
```
chmod 700 build.sh
./build.sh
```
A main.zip will appear in the current directory, which contains the App.