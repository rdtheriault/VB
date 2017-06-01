'Author- Robert Theriault
'Code- Searches through a set range to check if the cell is a date or not. Ignores blank cells
'Notes- Add to excel macro module and change range below

Sub CheckDate()


Dim rng As Range, cell As Range
Dim msg As String

Set rng = Range("D2:AB611") 'change range here

For Each cell In rng
  If Not IsDate(cell.Value) And Not IsEmpty(cell.Value) Then
    msg = msg + " " + cell.Address    
  End If
Next cell

MsgBox msg


End Sub

