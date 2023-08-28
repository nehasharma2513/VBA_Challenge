Attribute VB_Name = "Module1"
Sub Ticker():
    Dim ws As Worksheet
    ' Iterating through each Worksheet
    For Each ws In Worksheets
        Dim rowNum As Long
        Dim rowTotal As Long
        Dim TotVol As LongLong
        Dim CloseVal As Double
        Dim OpenVal As Double
        Dim YearlyChange As Double
        Dim PerChange As Double
        ws.Activate
        ' Assigning column headers for each value
        Range("K1").Value = "Ticker"
        Range("L1").Value = "YearlyChange"
        Range("M1").Value = "PercentChange"
        Range("N1").Value = "Total Stock Volume"
        Range("R2").Value = "Greatest % Increase"
        Range("R3").Value = "Greatest % Decrease"
        Range("R4").Value = "Greatest Total Volume"
        Range("S1").Value = "Ticker"
        Range("T1").Value = "Value"
        TotVol = 0
        rowNum = 2
        rowTotal = Range("A1").End(xlDown).Row
        OpenVal = Range("C3").Value
        ' Part 1 - Calculating YearlyChange, PercentageChange and TotalStockVolume for each unique Ticker
        For i = 2 To rowTotal
            TotVol = TotVol + Cells(i, 7).Value
            If (Cells(i, 1).Value <> Cells((i + 1), 1).Value) Then
                Name = Cells(i, 1).Value
                Range("K" & rowNum).Value = Name
                Range("N" & rowNum).Value = TotVol
                CloseVal = Cells(i, 6).Value
                rowNum = rowNum + 1
                YearlyChange = CloseVal - OpenVal
                PerChange = YearlyChange / OpenVal
                Range("L" & rowNum - 1).Value = YearlyChange
                'Conditional Formatting(YearlyChange)  : highlighting positive change in green and negative change in red.
                If (YearlyChange <= 0) Then
                    Range("L" & rowNum - 1).Interior.ColorIndex = 3
                Else
                    Range("L" & rowNum - 1).Interior.ColorIndex = 4
                End If
                Range("M" & rowNum - 1).Value = PerChange
                Range("M2:M" & rowNum).NumberFormat = "0.00%"
                'Conditional Formatting(PercentageChange) : highlighting positive change in green and negative change in red.
                If (PerChange <= 0) Then
                    Range("M" & rowNum - 1).Interior.ColorIndex = 3
                Else
                    Range("M" & rowNum - 1).Interior.ColorIndex = 4
                End If
                OpenVal = Cells(i + 1, 3).Value
                TotVol = 0
            End If
    
        Next i
     
        Dim GrtVol As LongLong
        Dim GrtVolTckr As String
        Dim RngCount As Long
        Dim GrtInc As Double
        Dim GrtIncTckr As String
        Dim GrtDecTckr As String
        Dim GrtDec As Double
        RngCount = Range("K1").End(xlDown).Row
        GrtVol = Range("N2").Value
        GrtInc = Range("M2").Value
        GrtDec = Range("M2").Value
      ' Part 2 - Calculating Greatest Total Volume, Greatest % Increase and Greatest % Decrease from the analysis done in Part 1
        For i = 2 To RngCount
                If (GrtVol < Cells((i + 1), 14).Value) Then
                   GrtVol = Cells(i + 1, 14).Value
                    GrtVolTckr = Cells(i + 1, 11).Value
               End If
                If (GrtInc < Cells((i + 1), 13).Value) Then
                   GrtInc = Cells(i + 1, 13).Value
                    GrtIncTckr = Cells(i + 1, 11).Value
               End If
                  If (GrtDec > Cells((i + 1), 13).Value) Then
                   GrtDec = Cells(i + 1, 13).Value
                    GrtDecTckr = Cells(i + 1, 11).Value
               End If
        Next i
        ' Displaying values obtained from Part 2
        Range("S4").Value = GrtVolTckr
        Range("T4").Value = GrtVol
        Range("T2").Value = GrtInc
        Range("T2").NumberFormat = "0.00%"
        Range("S2").Value = GrtIncTckr
        Range("T3").Value = GrtDec
        Range("T3").NumberFormat = "0.00%"
        Range("S3").Value = GrtDecTckr
    Next ws
 
End Sub
