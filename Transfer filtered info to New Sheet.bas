# VBA

Sub Criteria_Region_NewSheet()

Dim mysheet As Variant

Dim mycriteria As Variant

Dim myValue As Variant

Dim x As Long

Dim y As Long


Application.ScreenUpdating = False


mysheet = ActiveSheet.Name


mycriteria = InputBox("Choose Criteria")

On Error Resume Next

'Locates the first row using the 3rd column
  
  y = Application.WorksheetFunction.Match("*", Range(Cells(1, 3), Cells(20, 3)), 0)

'Region is the criteria
  
    x = Application.WorksheetFunction.Match(mycriteria, Range(Cells(y, 1), Cells(y, 1000)), 0)

On Error GoTo 0

'Input any value number, string, or date

myValue = InputBox("Filter by " & mycriteria & "")

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

Exit Sub

On Error GoTo ErrHandler:
MsgBox "Criteria or Filter not Found "

End Sub
