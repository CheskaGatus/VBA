Attribute VB_Name = "Module1"
Sub Button1()

For Each ws In Worksheets

    Dim Header1 As String
    Dim Header2 As String
    Dim Header3 As String
    Dim Header4 As String
    
    Header1 = "Ticker"
    Header2 = "Yearly Change"
    Header3 = "Percent Change"
    Header4 = "Total Stock Volume"
    
    ws.Range("I1").Value = Header1
    ws.Range("J1").Value = Header2
    ws.Range("K1").Value = Header3
    ws.Range("L1").Value = Header4

    Dim ticker As String
    Dim totalstockvolume As Double
    Dim summaryrow As Integer
    Dim change As Double
    Dim percentChange As Double
    Dim Start As Double
        
    totalstockvolume = 0
    summaryrow = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    Start = ws.Cells(2, 3).Value
    
    For i = 2 To LastRow
   
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        change = (ws.Cells(i, 6).Value - Start)
        
        
            If change = 0 Then
            percentChange = 0
            
            Else
                percentChange = 1
                If Start <> 0 Then
                    percentChange = change / Start
                End If
                
            End If
            
                ws.Range("J" & summaryrow).Interior.ColorIndex = 3
                
                If change > 0 Then
                    ws.Range("J" & summaryrow).Interior.ColorIndex = 4
                
                End If
                
        ws.Range("I" & summaryrow).Value = ticker
        ws.Range("J" & summaryrow).Value = change
        ws.Range("K" & summaryrow).Value = percentChange
        ws.Range("L" & summaryrow).Value = totalstockvolume
        
        summaryrow = summaryrow + 1
        totalstockvolume = 0
        Start = ws.Cells(i + 1, 3).Value
        Else
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        
        End If
        
    Next i

Next

End Sub

