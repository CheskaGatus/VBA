Attribute VB_Name = "Module1"
Sub Button1_Click()

For Each ws In Worksheets

    Dim Header1 As String
    Dim Header2 As String
    Header1 = "Ticker"
    Header2 = "Total Stock Volume"
    
    ws.Range("I1").Value = Header1
    ws.Range("J1").Value = Header2

    Dim ticker As String
    Dim totalstockvolume As Double
    Dim summaryrow As Integer
    
    totalstockvolume = 0
    summaryrow = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        ws.Range("I" & summaryrow).Value = ticker
        ws.Range("J" & summaryrow).Value = totalstockvolume
        summaryrow = summaryrow + 1
        totalstockvolume = 0
        
        Else
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        
        End If
        
    Next i

Next

End Sub
