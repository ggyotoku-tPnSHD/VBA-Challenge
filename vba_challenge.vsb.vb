'find unique tickers
'add stock volumes every time it sees unique ticker
'count number of rows; reset at different ticker
'total count number - current row for first row of the unique ticker

Sub stockInfo()
    
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim i           As Long
    Dim rrows       As Long
    Dim totalv      As LongLong
    Dim rowcount    As Long
    Dim yrlchange   As Double
    Dim prchange    As Double
    Dim sheet1      As Worksheet
    Dim ticker      As String
    Dim ticker_increase      As String
    Dim ticker_decrease      As String
    Dim max         As LongLong
    Dim maxper      As Double
    Dim minper      As Double
    
    totalv = 0
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    lRow2 = Cells(Rows.Count, 12).End(xlUp).Row
    rrows = 2
    rowcount = 3
    yrlchange = 0
    prchange = 0
    max = 0
    maxper = 0
    minper = 0
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    Application.ScreenUpdating = FALSE
    
    For Each sheet1 In Worksheets
        
        sheet1.Select
        
        For i = 2 To lRow
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Cells(rrows, 9).Value = Cells(i, 1).Value
                totalv = totalv + Cells(i, 7).Value
                Cells(rrows, 12).Value = totalv
                rrows = rrows + 1
                totalv = 0
                rowcount = rowcount + 1
                yrlchange = Cells(i, 6).Value - Cells(rowcount - i, 3).Value
                Cells(rrows - 1, 10).Value = yrlchange
                If yrlchange >= 0 Then
                    Cells(rrows - 1, 10).Interior.ColorIndex = 4
                Else
                    Cells(rrows - 1, 10).Interior.ColorIndex = 3
                    Cells(rrows - 1, 10).Font.Color = vbWhite
                End If
                yrlchange = 0
                prchange = (Cells(i, 6).Value - Cells(rowcount - i, 3).Value) / Cells(rowcount - i, 3).Value
                Roundv = Round(prchange, 2)
                Cells(rrows - 1, 11).Value = FormatPercent(Roundv)
                prchange = 0
                
            Else
                
                totalv = totalv + Cells(i, 7).Value
                rowcount = rowcount + 1
                
            End If
            
        Next i
        
        For e = 2 To lRow2
            If Cells(e, 12).Value > max Then
                max = Cells(e, 12)
                ticker = Cells(e, 9)
            End If
            If Cells(e, 11) > maxper Then
                maxper = Cells(e, 11)
                ticker_increase = Cells(e, 9)
            End If
            If Cells(e, 11) < minper Then
                minper = Cells(e, 11)
                ticker_decrease = Cells(e, 9)
            End If
            
        Next e
        
        Cells(4, 16).Value = FormatCurrency(max)
        Cells(3, 16).Value = FormatPercent(minper)
        Cells(2, 16).Value = FormatPercent(maxper)
        Cells(4, 15).Value = ticker
        Cells(2, 15).Value = ticker_increase
        Cells(3, 15).Value = ticker_decrease
        
        totalv = 0
        lRow = Cells(Rows.Count, 1).End(xlUp).Row
        lRow2 = Cells(Rows.Count, 12).End(xlUp).Row
        rrows = 2
        rowcount = 3
        yrlchange = 0
        prchange = 0
        max = 0
        maxper = 0
        minper = 0
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
    Next
    
    Application.ScreenUpdating = TRUE
    
End Sub