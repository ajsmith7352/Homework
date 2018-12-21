Sub Stock_Data ()

Dim ws As Worksheet
Dim LastR As Long
Dim TotalVolume As Double
Dim Ticker As String
Dim Summary As Integer

For Each ws In Worksheets
ws.Activate

Summary = 2
TotalVolume = 0
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Volume"

    LastR = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastR
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Ticker = Cells(i, 1).Value
            
                TotalVolume = TotalVolume + Cells(i, 7).Value
            
                Range("I" & Summary).Value = Ticker
            
                Range("J" & Summary).Value = TotalVolume
            
                Summary = Summary + 1
            
                TotalVolume = 0
            
            Else
            
                TotalVolume = TotalVolume + Cells(i, 7).Value
            
            End If
        
        Next i

    Next ws


End Sub
