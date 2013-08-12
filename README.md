Make a Statement
================

This will be a simple application for coverting spreadsheets into SQL CRUD statements. The generated statements shoud be either: a) formatted into an HTML file and immediately loaded or b) inserted into the user's clipboard for easy copy/paste into a database administration program. 

![Make a Statement Screenshot](http://i.imgur.com/HoFGlQj.png)

Currently a work in progress, at the moment an XLSX file can be loaded with its sheet-names and values displayed in th GUI. Once processing is completely working for XLSX files, support will be added for XLS and CSV files.

###Next steps
1. The UPDATE ListBox needs to show fields for the selected 'table' in the sheets ListBox.
2. Need to find a way to set variables for referring to the columns for the WHERE-CLAUSE builder.

Dependencies
------------
	sudo apt-get install python-tk
	sudo apt-get install python-setuptools
	sudo easy_install openpyxl
