Attribute VB_Name = "Module2"
Sub Reset()

    For Each ws In Worksheets
    
        LastSummaryRow = ws.Cells(Rows.Count, 12).End(xlUp).Row

        
        ws.Range("L1:M" & LastSummaryRow) = " "
    
    Next ws
    
    
End Sub

