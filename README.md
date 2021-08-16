# Zarar's Excel Program

Use this program to copy specific cells of one excel sheet to another.

## Installation

Make sure you have [Python](https://www.python.org/downloads/) and [Git](https://git-scm.com/downloads) installed.
Download this program by going to a directory in your file explorer and opening a terminal there. For Windows, follow [this](https://superuser.com/questions/339997/how-to-open-a-terminal-quickly-from-a-file-explorer-at-a-folder-in-windows-7#:~:text=You%20can%20also,open%20the%20shell.) guide to open the terminal.
Inside the terminal, run this command:
```
git clone https://github.com/Zohair-coder/zarar-work.git
```
A new directory with all the necessary files should be created. Navigate into the directory by typing in:
```
cd zarar-work
```
and then install all the necessary Python packages by typing:
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
