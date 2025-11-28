
Sub SuperscriptNumbers_AllSheets_ColumnsDEF()

    Dim ws As Worksheet
    Dim cols As Variant
    Dim col As Variant
    Dim lastRow As Long
    Dim c As Range
    Dim i As Long
    Dim txt As String
    
    ' Columns to superscript
    cols = Array("D", "E", "F")
    
    ' Loop through every worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        For Each col In cols
            
            lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
            
            For Each c In ws.Range(col & "1:" & col & lastRow)
                
                txt = CStr(c.Value)
                
                ' Apply superscript only to numeric characters
                For i = 1 To Len(txt)
                    If Mid(txt, i, 1) Like "[0-9]" Then
                        c.Characters(i, 1).Font.Superscript = True
                    End If
                Next i
                
            Next c
            
        Next col
        
    Next ws

End Sub
