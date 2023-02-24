# Excel Data Copier
This is a Python GUI program which allows you to copy data from one Excel sheet to another based on user-defined criteria.

### Prerequisite:
  - Python 3.x installed
  - `openpyxl` library to manage the excel processing. Install it using `pip install openpyxl`.
  - `threading` library to use multiple process threads.
  - `time` library to calculate the elapsed time by program.
  - `os` library to manage the files directory.
  - `tkinter` library to create GUI.
  - `datetime` library used in logging system.
  - `warnings` library to ignore the warnings.

### Installation
1. Make sure you have Python 3.x installed. If you don't have it installed, you can download it from the official website [here]('https://www.python.org/downloads/').
2. Clone this GitHub repository to your local machine or download the ZIP file and extract it to your desired location.
3. Open a terminal or command prompt and navigate to the directory where you cloned or extracted the repository.
4. Once the dependencies are installed, you can run the program by executing the following command:
  > `python Excel_Copier.py`

This will launch the GUI for the program.

### Usage
1. Select the source Excel file by clicking the `...` button next to the `Select the source Excel file` input field and browse to the file location.
2. Select the target Excel file by clicking the `...` button next to the `Select the target Excel file` input field and browse to the file location.
3. Enter the name of the sheet from the source Excel file in the `Enter the name of sheet from source Excel file` input field.
4. Enter the name of the sheet from the target Excel file in the `Enter the name of sheet from target Excel file` input field.
5. Enter the index of the source reference column (e.g. 1 for column A, 2 for column B, and so on) in the `Enter the source reference column index` input field.
6. Enter the index of the target reference column (e.g. 1 for column A, 2 for column B, and so on) in the `Enter the target reference column index` input field.
7. Enter the index of the source columns from which data should be copied (e.g. 1,2 for columns A and B) in the `Enter the index of source columns from which data to be copied` input field.
8. Enter the index of the target columns in which data should be copied (e.g. 1,2 for columns A and B) in the `Enter the index of target columns in which data to be copied` input field.
9. Click on the `Submit` button to initiate the data copy operation.
10. Once the operation is complete, a log file will be generated with information about the selected files, sheets, columns, copied data and elapsed time.

### Notes
- The program uses the openpyxl library to read and write Excel files. Please make sure that the library is installed before running the program.
- The source and target reference columns must contain the same type of data (e.g. both columns must contain numbers or text) in order for the copy operation to work correctly.
- The program assumes that the Excel files have the .xlsx extension. It may not work correctly with excel extensions .xlsm, .xlsb, .xltx, .xltm and .xlt
- Let the program finish completely and do not close it while it is on progress, otherwise TARGET EXCEL FILE WILL CORRUPT.

### Contributions
Contributions to this repo are welcome. If you find a bug or have a suggestion for improvement, please open an issue on the repository. If you would like to make changes to the code, feel free to submit a pull request.

### Acknowledgments
This program was created as a part of a programming challenge. Special thanks to the challenge organizers for the inspiration.
