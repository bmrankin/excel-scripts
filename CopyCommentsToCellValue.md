## Copy Comments to a cell value
https://www.extendoffice.com/documents/excel/765-excel-convert-comments-to-cells.html
```
Function GetComments(pRng As Range) As String
'Updateby20140509
If Not pRng.Comment Is Nothing Then
    GetComments = pRng.Comment.Text
End If
End Function
```
