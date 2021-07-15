# VBA-Advance-Filter-and-Criteria-Range

Sub Macro()

Dim filePath As String
Dim mR, cR As Range
Dim mR1, cR1 As Range
Dim mR2, cR2 As Range
Dim mR3, cR3 As Range
Dim masterBook, otherBook As Workbook

'Setup your 2 workbooks


Set masterBook = ActiveWorkbook
With masterBook.Sheets("Load")


'Specify data file path location in master copy

filePath = .Range("B7").Value 

End With

Set otherBook = Workbooks.Open(filePath)

'Setup your 2 ranges
Set mR = masterBook.Sheets("Load").Range("C6")
Set mR1 = masterBook.Sheets("Load").Range("D6")
Set mR2 = masterBook.Sheets("Load").Range("C5")
Set mR3 = masterBook.Sheets("Load").Range("D5")

'If you want a specific name for a worksheet 
'If your sheets do not change order, input as .Sheets(1,2,3,4 etc)

Set cR = otherBook.Sheets(1).Range("A16")
Set cR1 = otherBook.Sheets(1).Range("B16")
Set cR2 = otherBook.Sheets(1).Range("A15")
Set cR3 = otherBook.Sheets(1).Range("B15")

'Set the range in your masterbook equal to the range in your otherbook

mR.Value = cR.Value
mR1.Value = cR1.Value
mR2.Value = cR2.Value
mR3.Value = cR3.Value

'Close your otherbook without saving anything=
otherBook.Close False


End Sub
