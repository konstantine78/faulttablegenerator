# faulttablegenerator
Converts contents of a *.csv file into a table inserted into a *.docx file format.
The main file is faulttablegenerator.py.  Pyinstaller command is used to create the faulttablegenerator.exe.  When faulttablegenerator is executed a prompt will come up for the user to enter in the name of the CSV file to read in.  The *.csv file must be of a specific configuration.  It assumes a table that has 3 columns (FAULT, FAULT TEXT, REMEDY).  These serve as headers.  Using the 'csv' python module, the contents of the *.csv file are read in and stored to a list of lists.  Using the docx package, a word document is defined, along w/ a table within that file that will be inundated with the contents of the stored list.  

# Files
.example.csv - This is an example CSV file that one could use for debugging/running the program.
.faulttablegenerator.py - Main file
.faulttablegenerator.exe - Executable created via command prompt 'pyinstall --onefile faulttablegenerator.py'.  Note that the "-w" flag is not used since the terminal is needed during the run of the script.  The '-w' flag will hide the terminal/command prompt.

References:
1) https://github.com/RadiantCoding/Code (Search for "Word Doc - Simple Table")
