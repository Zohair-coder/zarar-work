# Zarar's Excel Program

Use this program to copy specific cells of one excel sheet to another.

## Installation

Make sure you have [Python](https://www.python.org/downloads/) installed.
Download this program by clicking the green `Code` button and selecting `Download ZIP` from the dropdown menu. Extract the zip file into a folder. Inside the folder, open up a terminal. For Windows, follow [this](https://superuser.com/questions/339997/how-to-open-a-terminal-quickly-from-a-file-explorer-at-a-folder-in-windows-7#:~:text=You%20can%20also,open%20the%20shell.) guide to open the terminal. Install all the necessary Python packages by typing:
```
pip install -r requirements.txt
```

Once that's done, you can proceed to the Usage section.

## Usage

To use the program, copy the read and write Excel sheets to the folder you downloaded. Open up `main.py` with Notepad or any other text editor and change the values of the READ_FILE, WRITE_FILE, READ_FILE_SHEET and WRITE_FILE_SHEET to the respective names of the files and their sheet names.
Then, save the file and close it. Open up a terminal in the directory again and run the command:
```
python main.py
```
You should see a new directory called `output` with all the new files.
