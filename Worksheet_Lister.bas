Attribute VB_Name = "Worksheet_Lister"
Option Explicit
Sub ListSheets()
    ' Declared Variables
    Dim ws As Worksheet
    Dim x As Integer
    Dim y As Integer
    Dim Sheet_List As Worksheet
    
    Application.ScreenUpdating = False
    
    ' Check if worksheet exists, if not add a new worksheet and set the reference to Sheet_List
    On Error Resume Next ' Ignore error in the next line
    Set Sheet_List = ThisWorkbook.Sheets("Sheet_List")
    On Error GoTo 0 ' Resume normal error handling
    
    ' If Sheet_List does not exist then create it
    If Sheet_List Is Nothing Then
        Set Sheet_List = ThisWorkbook.Sheets.Add
        Sheet_List.Name = "Sheet_List"
    End If
    
    ' Clear the range before execution
    Sheet_List.Range("A:A").Clear
    
    ' Set up iterable variables
    x = 1
    y = 1
    
    ' Add and format the column title
    Sheet_List.Cells(x, y).Value = "SHEETS"
    Sheet_List.Cells(x, y).Font.Bold = True
    Sheet_List.Cells(x, y).Font.Underline = True
    
    ' Iterate through each sheet adding name to a new column
    x = 2
    For Each ws In Worksheets
         Sheet_List.Cells(x, y).Value = ws.Name
         x = x + 1
    Next ws
    
    Application.ScreenUpdating = True

End Sub


