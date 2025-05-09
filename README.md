# distributed-translate
Translates a Google Sheets page between user-selected languages without altering the original layout. Creates a duplicate sheet with translated content using Google Translate.

# Features
Automates translation of entire Google Sheets
Preserves original layout and structure
Uses Google Sheets' built-in GOOGLETRANSLATE() function
Useful for translating structured content like CVs or data tables
Easily adjustable for different sheet sizes and structures

# How to Use
Open the script editor in your Google Sheets:
Extensions > Scripts > Manage scripts > (⋮ three-dot menu) > Edit script.
Copy and paste the script from the sample file into your own sheet.
Save and close the script editor.
Back in Google Sheets, go to:
Extensions > Macros and select the macro to start.
You will be prompted to enter:
Source language (e.g., en for English)
Target language (e.g., tr for Turkish)
The script will:
Duplicate the current sheet
Insert translation formulas in the new sheet
Keep your formatting unchanged

# How to Modify
To customize which cells are translated, open the script again:
Extensions > Macros > Manage macros > (⋮) > Edit script
Then make the following changes in the code:
Line 28:
Change the number 1 to your last column's index.
(e.g., A = 1, B = 2, C = 3, ...)
Line 61:
Change the number 13 to the number of rows you want to translate.
Line 62:
Change the number 2 to define which row the translation should start from.
These modifications let you fine-tune the translation area for more flexibility.
