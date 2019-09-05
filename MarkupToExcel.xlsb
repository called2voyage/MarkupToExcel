Sub MarkupToColumns()

Dim markup As Range
Set markup = Worksheets("Sheet1").Range("A1", Range("A1").End(xlDown))

Dim text() As String
Dim num As Integer
Dim cur As Range
Dim newVal As String

For i = 1 To markup.Rows.Count
   Set cur = markup.Cells(i, 1)
   text() = Split(cur.Value, "|")
   col = 1
   For j = LBound(text) To UBound(text)
      If j <> UBound(text) Or text(j) <> "" Then
         newVal = text(j)
         Worksheets("Sheet1").Cells(i, col).Value = newVal
      End If
      col = col + 1
   Next j
Next i

End Sub 

Sub ColumnsToMarkup()

Dim markup As String
markup = ""

For i = 1 To Selection.Rows.Count
   For j = 1 To Selection.Columns.Count
      If j = 1 Then
         markup = markup & Worksheets("Sheet1").Cells(i, j).Value
      Else
         markup = markup & "|" & Worksheets("Sheet1").Cells(i, j).Value
      End If
   Next j
   If i <> Selection.Rows.Count Then
      markup = markup & vbNewLine
   End If
Next i

Selection.Clear
Worksheets("Sheet1").Range("A1").Value = markup

End Sub
