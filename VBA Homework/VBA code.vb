Sub ticker():

Dim tickerGroup, currentValue, nextValue As String
Dim lastrow As Long
Dim i As Long
Dim totalVolume As Long
Dim summarytablerow As Integer
Dim volume As Long

    Sheets.Add.Name = "Combined_Data"
    Sheets("Combined_Data").Move Before:=Sheets(1)
    
For Each ws In Worksheets
        
    worksheetName = ws.Name

totalVolume = 0
summarytablerow = 1
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
        volume = Cells(i, 7).Value
        currentValue = Cells(i, 1).Value
        nextValue = Cells(i + 1, 1).Value
        
        If currentValue <> nextValue Then
        tickerGroup = currentValue
        totalVolume = totalVolume + volume
        
        summarytablerow = summarytablerow + 1
        
        Range("J" & summarytablerow).Value = tickerGroup
        Range("K" & summarytablerow).Value = totalVolume
        
        totalVolume = 0
        
        
        End If
    
    Next i
        
    Next ws
End Sub
