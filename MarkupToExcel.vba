' Copyright 2019 called2voyage
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.

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

Dim markup() As String
ReDim markup(Selection.Rows.Count)

For i = 1 To Selection.Rows.Count
   For j = 1 To Selection.Columns.Count
      If j = 1 Then
         markup(i) = Worksheets("Sheet1").Cells(i, j).Value
      Else
         markup(i) = markup(i) & "|" & Worksheets("Sheet1").Cells(i, j).Value
      End If
   Next j
Next i

Selection.Clear
For i = 1 To Selection.Rows.Count
    Worksheets("Sheet1").Cells(i, 1).Value = markup(i)
Next i

End Sub
