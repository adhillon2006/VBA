Sub Criteria_MultiSheet_Split()


Dim duplicatex As Long
Dim lastrow As Long
Dim i As Integer
Dim mysheet As Variant
Dim mycriteria As Variant
Dim myValue As String
Dim x As Long
Dim y As Long
Dim ErrHandler As Error
Dim ErrHandlerblank As Error

Application.ScreenUpdating = False
mysheet = ActiveSheet.Name

'Create sheet for duplicates
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "Working"

'Select filtered sheet
Sheets(mysheet).Select

'choose criteria
mycriteria = InputBox("Choose Criteria")

On Error GoTo ErrHandler:
'Locates the first row using the 3rd column
    y = Application.WorksheetFunction.Match("*", Range(Cells(1, 3), Cells(20, 3)), 0)
'Region is the criteria and locates row
    x = Application.WorksheetFunction.Match(mycriteria, Range(Cells(y, 1), Cells(y, 1000)), 0)
On Error GoTo 0

lastrow = Sheets(mysheet).Rows(Rows.Count).End(xlUp).Row
duplicatex = Sheets("Working").Rows(Rows.Count).End(xlUp).Row

'Copy duplicates to Workingsheet
Sheets(mysheet).Select
Sheets(mysheet).Range(Cells(y, x), Cells(lastrow, x)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Working").Range("A1"), Unique:=True

'Count number of deplicates for loop
duplicatex = Sheets("Working").Rows(Rows.Count).End(xlUp).Row

'loop
On Error GoTo ErrHandlerblank:
If duplicatex - 1 < 50 Then
For i = 1 To duplicatex - 1

myValue = Sheets("Working").Cells(i + 1, 1)
'Filter
Range(Cells(x, y), Cells(x, y)).AutoFilter Field:=x, Criteria1:=myValue
'Copies Information
Range(Cells(x, y), Cells(x, y)).Select
ActiveSheet.AutoFilter.Range.Copy
Worksheets.Add.Paste
ActiveSheet.Name = Left(myValue, 31)
'Formats columns
Columns("A:CX").EntireColumn.AutoFit
Range("A1:AS100").EntireRow.AutoFit
' Unfilter
Worksheets(mysheet).ShowAllData
Sheets(mysheet).Select
Next i

End If

'Disable Delete warning
Application.DisplayAlerts = False
Sheets("Working").Delete
Application.DisplayAlerts = True

Application.ScreenUpdating = True

Exit Sub
ErrHandler:
MsgBox "Criteria must be text"

Exit Sub

ErrHandlerblank:
If myValue = "" Then
ActiveSheet.Name = "Blank"
End If
Resume Next

End Sub
