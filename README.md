# What do we have here?
Scripts, usually in Python, written to automate administrative tasks! 

# Duplicates.py - What does it do?
This file takes a contact list on an excel document, and uses the emails listed to check if there are any duplicate contacts.

If none are found, the script prints a message for the user saying so, and the program is exited.
If the program finds that there are duplicate emails, the script creates a new excel file (in the same directory) with contact information for each contact listed in the original file, duplicates excluded.

# What do I need to run this?
This script is written in Python 3, so you'll need that installed in order for it to be able to run.

Additionally you'll need 2 libraries: xlrd, to be able to read the original Excel file, and xlsxwriter, to be able to create a new file without duplicates.
    To install these, enter this onto your Terminal window:
    $ pip install xlrd
    $ pip install xlsxwriter
    
# How do I use it?
This script was written for my user specifically, and assumes that the excel document is in the "Downloads" folder.

To run the script, type the following onto your Terminal window:
    $ python duplicates.py

The program will then ask you to name the file you want to use. You should enter the file name, without the file extension (the program assumes it's an excel sheet).

That's it!

# What's on the new excel file created?
The new file will be found in the same directory as your original file, within your "Downloads" folder, and will have the same general name as your original file, except with "NEW" added to the beginning of the name.

It'll use the same header, if there is one, as your original excel file, and will contain all of the contact information provided in the original excel sheet, excluding any duplicate contacts.
