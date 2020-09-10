import os
import xlrd, xlsxwriter  # lets you handle excel files (rd - read, other - write)
from sys import exit

#Path in which we'll be working (where we'll find our original file and
#   possibly making a new file later)
starting_path = "/Users/danielavalades/Downloads"
real_path = os.path.realpath(starting_path)
os.chdir(real_path)
print()


#ask user for file name
# MAYBE IT SHOULD ASK THE USER IF THERE'S A HEADER****
def askforfile():
    global file, filename
    file = input("Give name of file: ")
    file_ending = ".xlsx"
    filename = file + file_ending
askforfile()

#check if file exists, if it doesn't, ask again
while True:
    if file == '':
        print("\nWhen you press ENTER, you exit out of the program. Goodbye!\n")
        exit()
    elif os.path.exists(filename) == True:
        break
    else:
        print("""Oops, couldn't find that file. Try again without the ending (ex:'.xlsx', '.txt')
    OR press ENTER to exit program.\n""")
        askforfile()

# Open the excel file and read contents
wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0) # looks at first sheet

# SCAN FOR DUPLICATES
# Add all elements (rows) to a list
#   This program uses emails to see if any contacts are duplicated
emails = []
for contact in range(1, sheet.nrows): #starting index at 1 bc we're skipping the headers
    email = sheet.cell_value(contact, 130) #emails are stored in column 130
    emails.append(email) #make list of all emails

# define function that checks for duplicates
def duplicate_check(contacts):
    ''' Checks if a given list contains any duplicates '''
    if len(contacts) == len(set(contacts)):
        print("\nNo duplicates! We're done here. Bye!\n")
        exit()
    else:
        pass

duplicate_check(emails)

# MAKE A LIST OF ALL ELEMENTS (ROW#, EMAIL) WITHOUT DUPLICATES
# function that deletes all duplicated elements from above list
def distinctcontacts(contacts):
    ''' Creates 2 lists: (1) "repeats", a list of repeated elements whose elements
        are tuples made up of the row number where the contact is located,
        and the corresponding email, and (2) "email_by_row", a similar list
        that contains all distinct contacts by row number and email.'''
    global email_by_row, repeats
    # create a list of tuples whose first element is the row number for the contact
    #   and second element is the corresponding email
    email_by_row = list(enumerate(contacts, 2))
    repeated = []
    repeats = []
    size = len(email_by_row)
    #finds duplicates ****BUT FOR SOME REASON RETURNS THEM WITH DUPLICATES****
    for i in range(size):
        for k in range(i+1, size):
            if email_by_row[i][1] == email_by_row[k][1] and email_by_row[k][1] not in repeated:
                repeated.append(email_by_row[k])
    #THIS IS THE WORKING LIST OF THE ROWS WITH DUPLICATED EMAIL ADDRESSES
    for i in repeated:
        if i not in repeats:
            repeats.append(i)
    #delete all duplicated elements from working list (email_by_row)
    for file in repeats:
        email_by_row.remove(file)

distinctcontacts(emails)


#make list with all of the info from each non-duplicated contact
amt = len(email_by_row)
header = sheet.row_values(0)
index = [] # will become list of row number of contacts that aren't repeated
contacts_norepeats = [] # will become list of full contact info for indexed contacts
for i in range(amt):
    index.append(email_by_row[i][0])
#fills contact info for each non-repeated entry
for i in index:
    contacts_norepeats.append(sheet.row_values(i - 1))


# MAKE A FILE WITHOUT duplicates (using email_by_row)
newfilename = 'NEW' + filename
workbook = xlsxwriter.Workbook(newfilename)
worksheet = workbook.add_worksheet()
amtofcolumns = int(sheet.ncols) + 1 #amt of columns in originial file with contact info on them plus 1


row = 1 #row on which we want to start writing contact info
col = 0

# iterate over data in email_by_row and write it out row by row
for i in header:
    worksheet.write(0, col, i)
    col += 1

col = 0
for contact in contacts_norepeats:
    for item in contact:
        worksheet.write(row, col, item)
        col += 1
    row += 1
    col = 0

workbook.close()
print("\n I just created a new excel file with only distinct contacts.\n \
Check your 'Downloads' folder! Bye!\n")
