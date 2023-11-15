# Excel_Worksheet_Lister
A simple macro that iterates through all sheets in an Excel workbook then lists them on a new tab

**Overview**
This repository contains a VBA (Visual Basic for Applications) macro designed for use in Microsoft Excel. The ListSheets macro automates the process of listing all the worksheet names in the active workbook. It creates a new worksheet named "Sheet_List" (or uses an existing one) and populates it with the names of all other worksheets.

**Features**
Automatic Sheet Listing: Compiles a list of all worksheets in the active workbook.
Dynamic Sheet Creation: Creates a new worksheet for listing or uses an existing one if already present.
Performance Optimization: Utilizes Application.ScreenUpdating to enhance performance, especially useful in workbooks with a large number of sheets.

**Prerequisites**
Microsoft Excel with macro support.
Basic familiarity with Excel and VBA.

**Installation**
Open the Excel workbook where you want to use the macro.
Press Alt + F11 to open the VBA editor.
Create a new module and paste the ListSheets macro code into it.

**Usage**
Ensure that the Excel workbook where you want to list the sheet names is open.
Run the ListSheets macro.
A new worksheet named "Sheet_List" will be created (or an existing one will be used), and the names of all worksheets will be listed starting from cell A2.

**Customization**
You can modify the range within the macro if you wish to list the sheets in a different location or format.
Adjust the column width or add additional formatting as per your requirements.

**Limitations**
The macro lists all types of sheets present in the workbook. If you want to list specific types of sheets (e.g., exclude chart sheets), additional modifications to the code are needed.

**Author**
BluePhoenix
