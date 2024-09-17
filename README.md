# OFM_Project_Analyser
OFM project analyser VBA

No OFM license? No problem! Find out how to work with an OFM project without it.

Every now and then, you may need to work with an OFM project created by someone else, or from an asset or company you are unfamiliar with. Before using the data, it’s good to ensure you don’t misinterpret the different variables in the project, or the units and multipliers to be applied to the data stored in the database. Through different OFM versions, that data has been stored in various ways, and in recent years, these settings have been written within the '.ofm' file (the project file), which is an XML file.

An XML file is a text file that stores data in a structured way. This means you can open it with a text editor and read it to understand how the project is set up. However, there are many things in the XML that may not be relevant, making it inconvenient to work with. Enter the ‘OFM Project Analyser’.

The ‘OFM Project Analyser’ consists of a few VBA procedures using the MSXML library. You need to run the ‘CheckDotOFMfile’ macro (press ALT+F8 to bring up the macro dialog box), which gets the name of the '.ofm' file and calls the procedures to retrieve the data for tables, user functions, and calculated variables. There’s more that could be extracted, but I limited it to what I needed for my purposes.

The animation illustrates how you use the workbook:
1. Run the macro.
2. Specify the project file.
3. A new workbook is created to store the data extracted from the project - remember to save it.

A new workbook is created with the data to avoid storing your macros and data in the same workbook. Doing it this way also allows you to copy the procedures from the spreadsheet to the one you use to store your macros with minimal or no changes.


