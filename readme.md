#SplitXLSX

Hello!  This  is a very simple module and is intended to help you split large Excel files into more manageable ones.
To install it, just run 'npm install splitxlsx -g'

To use it, just enter your command from your preferred directory.

splitxlsx -S yoursourceFile.xlsx -N 1000 -O outputsFolder -W myWorkSheetTitle

If you need to convert a date, then you should specify that with a settings file. The settings file should be a JSON file with the following format, and referenced with the optional -J parameter.  In the following example, we convert the 5th column of the Student worksheet and the 2nd/3rd column of the Subscriptions worksheet to the excel date format.

{
"forceTypes": [
{
"type": "date",
"index": 4,
"sheetName": "Student"
},
{
"type": "date",
"index": 2,
"sheetName": "Subscription",
"startIndex": 1
},
{
"type": "date",
"index": 1,
"sheetName": "Subscription",
"startIndex": 1
}
]
}
