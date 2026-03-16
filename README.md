# Password Manager
An excel based Password Manager

# Disclaimer
This excel is not meant to be a replacement for any of the safe and secure Password Managers in the market. This is just an alternate tool, one you can use if you have restrictions to download or install or run a 3rd party Password Manager in your machine or something that you can use if you are planning to store your passwords in a plain text file.

I'd highly recommend you to use a proper Password Manager to store your passwords (if looking for a free one, check out KeePass Password Safe).

If you're still considering of using this tool, ensure to have 

1. workbook and sheet level protection 
2. a file level encryption (File > Info > Protect Workbook > Encrypt with Password)
3. using a long and complex password (20+ characters)
4. following all the safety considerations for keeping your password safe

(details of how, can be easily found online).

I do not take any reposnsibility for data loss or misuse by using this tool.

## How to use this excel

### 1. How to view the formulas and contents in the excel?

- All cells in the 'Main' sheet are locked, except the 'Search', 'Group' and result range. This limits users to interact with only the unlocked cells
- 'Data' sheet is hidden and the Workbook is protected, which limits the user to view only the visible 'Main' sheet
- Password to unprotect the sheet and workbook is - My.Excel.Password.Manager.v0

### 2. How to set up the data?

- Add all relevant records in the 'Data' sheet
- Columns G and H in sheet 'Data' are formulas and are automatically generated

### 3. How does the search functionality work?

- Column G in Sheet 'Data' holds the signature of the entry, which is basically a concatenation of the key fields on which the search is based on. Note: Password is not part of the signature.
- Column H in Sheet 'Data' holds the logic that matches the entry against the search criteria from the 'Main' sheet
  - 'Search' text in 'Main'!C2 is checked against the signature of the entries in 'Data' sheet
  - 'Group' text in 'Main'!C4 is used to filter for groups in 'Data'!A:A for a match
  - 'Search' text can be a partial or complete text of either of the fields - Group, Entry, Username, URL, Notes fields in the 'Data' sheet - to find a match
  - 'Group' text has to be an exact Group name in column A of the 'Data' sheet - to return entries of a matching group
- Column B6 in 'Main' sheet uses a LET function, that defines
  - source of data to check, i.e. the data in 'Data' sheet
  - the last row of the data in 'Data' sheet
  - the criteria range, based on the defined data and last row of data
  - output, which is the data filtered by the criteria
  - limits the results to 6 rows (inc. header)
- Cells 'Main'!E7:E11 are formatted so that the text is white in color, thereby making the passwords invisible, yet be copied using a Ctrl+C

### Features

- 'Search' & 'Group' can be used separately or in combination to filter to the exact entry
- A keyword of 'All' can be used in the 'Group' search to return all entries (limted to 5)
- When records matching the search criteria are displayed, the Password field is set to be not visible, yet can be copeid using a Ctrl+C
- When there is a match, the entries are limited to 5, so that any one viewing your screen do not get an idea on how many entries are present in your database
- 'Search' field can work with a part or whole of a keyword to return the matching results
- 'Group' field only works with a valid complete Group name from the Data sheet, this is to limit the usage to only those who knows what is in the database

