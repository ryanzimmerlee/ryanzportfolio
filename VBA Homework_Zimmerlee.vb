Sub RunMacrosOnWorkbook()
    
    Dim xSheet As Worksheet
    Application.ScreenUpdating = False
    For Each xSheet In Worksheets
        xSheet.Select
        Call TickerVolume
        Call ConditionalFormat
        Call CreateSummaryValueTable
    Next
    Application.ScreenUpdating = True
    
End Sub

Sub TickerVolume()

    Dim VolumeTotal As Variant
    Dim row As Long
    Dim BeginRow As Long
    Dim TickerName As String
    Dim SummaryRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim ValueChg As Double
    Dim PercChg As Double
    Dim lrow As Long
    
    'Find last row in the sheet based on column A data
    lrow = Cells(Rows.Count, "A").End(xlUp).row
    
    'Giving summary table column headers
    Cells(1, 9).Value = "Ticker Name"
    Cells(1, 10).Value = "Volume Total"
    Cells(1, 11).Value = "Yearly Change ($)"
    Cells(1, 12).Value = "Yearly Change (%)"
    Cells(1, 18).Value = "Year Opening Price"
    Cells(1, 19).Value = "Year Closing Price"
    Cells(1, 20).Value = "Begin Sum Range Row"
    Cells(1, 21).Value = "End Sum Range Row"
    
    'Volume total is set to zero
    VolumeTotal = 0
    
    'Defining beginning range for summary table
    SummaryRow = 2
    
    'Used for setting lower range limit
    BeginRow = 2
    
        For row = 2 To lrow
        
            If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
                
                'Grab beginning row value before resetting at the end of this IF statement
                BeginRow = BeginRow
                
                'Grab current ticker name
                TickerName = Cells(row, 1).Value
                
                'Grab current ticker total volume from "ELSE" loop
                VolumeTotal = VolumeTotal + Cells(row, 7).Value
                
                'Find current ticker year opening price and year closing price
                OpeningPrice = Cells(BeginRow, 3).Value
                ClosingPrice = Cells(row, 6).Value
                
                'Calculate value and percentage change for the currenct ticker between year end and year open
                ValueChg = ClosingPrice - OpeningPrice
                
                'There are zero values on one sheet for some of the stock prices...need to tell VBA what to do when this error occurs
                '"On error, go to error label at the end of this IF loop"
                On Error GoTo Err1
                PercChg = ValueChg / ClosingPrice
                
                'Place values for the currect ticker in the created summary table within current sheet
                Range("I" & SummaryRow).Value = TickerName
                Range("J" & SummaryRow).Value = VolumeTotal
                Range("K" & SummaryRow).Value = ValueChg
                Range("L" & SummaryRow).Value = PercChg
                
                'Test Year Opening Price and Year Closing Price
                Range("R" & SummaryRow).Value = OpeningPrice
                Range("S" & SummaryRow).Value = ClosingPrice
                
                'Test Beginning Row and End Row
                Range("T" & SummaryRow).Value = BeginRow
                Range("U" & SummaryRow).Value = row
                
                'Adjusting summary table by 1 for next ticker data
                SummaryRow = SummaryRow + 1
                
                'Reset total volume for next ticker
                VolumeTotal = 0
                
                'Set new beginning range row to current ticker row + 1
                BeginRow = row + 1
                
'Label - workaround for Div/0 error
Err1:
            ' Else loop for all ticker symbols that are the SAME...continue to summarize them
            Else
                VolumeTotal = VolumeTotal + Cells(row, 7).Value
                   
            End If
            
        Next row

End Sub


Sub ConditionalFormat()

    'Defining Conditional Formatting Range Variable
    Dim formatrng As Range
    
    'Defining formating range
    Set formatrng = Range("L2", Range("L2").End(xlDown))
    
    'Delete existing format conditions
    formatrng.FormatConditions.Delete
    
    'Defining conditional formatting variables
    Dim cond1 As FormatCondition
    Dim cond2 As FormatCondition
    
    'Equations for when/how to format the named range
    Set cond1 = formatrng.FormatConditions.Add(xlCellValue, xlGreater, "0")
    Set cond2 = formatrng.FormatConditions.Add(xlCellValue, xlLess, "0")

    'Over 0 format as green
    With cond1
        .Interior.Color = vbGreen
        .Font.Color = vbBlack
        .NumberFormat = "0.00%"
    End With
 
    'Under 0 format as red
    With cond2
        .Interior.Color = vbRed
        .Font.Color = vbBlack
        .NumberFormat = "0.00%"
    End With

End Sub


Sub CreateSummaryValueTable()

    Dim lrow As Integer
    Dim row As Integer
    Dim max As Double
    Dim maxticker As String
    Dim max2 As Variant
    Dim maxvolume
    Dim min As Double
    Dim minticker As String

    lrow = Cells(Rows.Count, "I").End(xlUp).row

    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    min = 0
    max = 0
    max2 = 0
    
    For row = 2 To lrow
        
        'Find Greatest Positive % Change ...if the current value is greater than the stored max, replace the max.
        If (Cells(row, 12).Value >= max) Then
            max = Cells(row, 12).Value
            maxticker = Cells(row, 9).Value
            
        End If
        
        ' Once iterated all the through last row, place the max value and adjacent ticker
        Range("P2").Value = max
        Range("P2").NumberFormat = "0.00%"
        Range("O2").Value = maxticker
        
        'Find Greatest Negative % Change ...if the current value is greater than the stored min, replace the min.
        If (Cells(row, 12).Value <= min) Then
            min = Cells(row, 12).Value
            minticker = Cells(row, 9).Value
            
        End If
        
        ' Once iterated all the through last row, place the min value and adjacent ticker
        Range("P3").Value = min
        Range("P3").NumberFormat = "0.00%"
        Range("O3").Value = minticker
        
        'Find Greatest Ticker Volume ...if the current value is greater than the stored max, replace the max.
        If (Cells(row, 10).Value >= max2) Then
            max2 = Cells(row, 10).Value
            maxvolume = Cells(row, 9).Value

        End If
        
        ' Once iterated all the through last row, place the max volume and adjacent ticker
        Range("P4").Value = max2
        Range("O4").Value = maxvolume

    Next row

End Sub
