''' ActiveCell.FormulaR1C1 = "=HYPERLINK(""[Workbook.xlsx]Sheet1!A1"",""CLICK HERE"")"


''something liike the following will loop through column A in the Control sheet and turn the values in the cells into Hyperlinks.
''Not something I've had to do before so please excuse bugs:

Sub CreateHyperlinks()

Dim mySheet As String
Dim myRange As Excel.Range
Dim cell As Excel.Range
Set myRange = Excel.ThisWorkbook.Sheets("Control").Range("A1:A5") '<<adjust range to suit

For Each cell In myRange
    Excel.ThisWorkbook.Sheets("Control").Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:=cell.Value & "!A1" '<<from recorded macro
Next cell

End Sub
