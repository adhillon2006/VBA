# VBA

Sub Criteria_Region_NewSheet()

Dim x As Long

Dim myValue As Variant

Dim mysheet As Variant

Application.ScreenUpdating = False

mysheet = ActiveSheet.Name

On Error Resume Next

'Locates the first row using the 3rd column
  
  y = Application.WorksheetFunction.Match("*", Range(Cells(1, 3), Cells(20, 3)), 0)

'Region is the criteria
  
  x = Application.WorksheetFunction.Match("Region", Range(Cells(y, 1), Cells(y, 1000)), 0)


On Error GoTo 0

'Input any value number, string, or date

myValue = InputBox("Filter by Region")

'Filter

Range(Cells(x, y), Cells(x, y)).AutoFilter Field:=x, Criteria1:=myValue

'Copies Information

Range(Cells(x, y), Cells(x, y)).Select

ActiveSheet.AutoFilter.Range.Copy

Worksheets.Add.Paste


'Formats columns

Columns("A:CX").EntireColumn.AutoFit

Range("A1:AS100").EntireRow.AutoFit


' Unfilter

Worksheets(mysheet).ShowAllData


Application.ScreenUpdating = True

End Sub
