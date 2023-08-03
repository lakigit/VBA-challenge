Attribute VB_Name = "Module1"

'To run click on first sub and press F5

Sub ForAllSheets():
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call StockData
    Next
    Application.ScreenUpdating = True
End Sub


Sub StockData():


Range("I1").EntireColumn.Insert
Cells(1, 9).Value = "Ticker"

Range("J1").EntireColumn.Insert
Cells(1, 10).Value = "Yearly Change"

Range("K1").EntireColumn.Insert
Cells(1, 11).Value = "Percent Change"

Range("L1").EntireColumn.Insert
Cells(1, 12).Value = "Total Stock Volume"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"




Dim YChange As Double
Dim Pchange As Double
Dim TotalVolume As Double
Dim SummaryTableRow As Integer
Dim Tickertype As String
SummaryTableRow = 2
TotalVolume = 0
Dim openvalue As Double
Dim closevalue As Double
priceflag = True
Dim maxpercent As Double

 lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            
            Tickertype = Cells(i, 1).Value
            Range("I" & SummaryTableRow) = Tickertype
   
            closevalue = Cells(i, 6).Value
            YChange = closevalue - openvalue
            Range("J" & SummaryTableRow) = YChange
            
            
        If YChange < 0 Then
            Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            
        ElseIf YChange > 0 Then
            Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
        End If
        
            
        If YChange = 0 Or openvalue = 0 Then
            Cells(SummaryTableRow, 11).Value = 0
        Else
            Cells(SummaryTableRow, 11).Value = FormatPercent((YChange / openvalue))
        End If
        
        
        If Cells(SummaryTableRow, 11).Value < 0 Then
            Cells(SummaryTableRow, 11).Interior.ColorIndex = 3
        ElseIf Cells(SummaryTableRow, 11).Value > 0 Then
            Cells(SummaryTableRow, 11).Interior.ColorIndex = 4
        End If
        
            Volume = Cells(i, 7).Value
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Range("L" & SummaryTableRow) = TotalVolume
            
            maxpercent = WorksheetFunction.Max(Range("K2:K100000"))
            
            Range("Q2").Value = FormatPercent(maxpercent)
            If Cells(SummaryTableRow, 11).Value = maxpercent Then
            Range("P2").Value = Cells(SummaryTableRow, 9)
            End If
            
            minpercent = WorksheetFunction.Min(Range("K2:K100000"))
            
            Range("Q3").Value = FormatPercent(minpercent)
            If Cells(SummaryTableRow, 11).Value = minpercent Then
            Range("P3").Value = Cells(SummaryTableRow, 9)
            End If
            
            maxvolume = WorksheetFunction.Max(Range("L2:L100000"))
            
            Range("Q4").Value = maxvolume
            If Cells(SummaryTableRow, 12).Value = maxvolume Then
            Range("P4").Value = Cells(SummaryTableRow, 9)
            End If
            
            
            'reset variables and go to next ticker symbol
            
            SummaryTableRow = SummaryTableRow + 1
            TotalVolume = 0
            priceflag = True
            
            
            Else
            
        If priceflag Then
            openvalue = Cells(i, 3).Value
            priceflag = False
            
        End If
        
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            End If
            
        Next i

End Sub



