Attribute VB_Name = "Module1"
Sub WSLoop()


' Variables declaration
Dim LastRow As Double

Dim Ticker As String
Dim Volume As Double

Dim i As Double

Dim ReportRow As Double

Dim YearOpen As Double
Dim YearClose As Double
Dim YearChange As Double
Dim PercentChange As String

Dim GreatestIncreseTicker As String
Dim GreatestIncreasePercent As Double

Dim GreatestDecreateTicker As String
Dim GreatestDecreasePercent As Double

Dim GreatestVolumeTicker As String
Dim GreatestVolumeNumber As Double

'Loop thru each worksheet
For Each ws In Worksheets
    
    'Get Last row on worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Reset Counters
    Volume = 0
    ReportRow = 2
    
    'Loop thru ticker column
    For i = 2 To LastRow
        'Get ticker, running total for volume
        Ticker = Sheets(ws.Name).Cells(i, 1).Value
        Volume = Volume + Sheets(ws.Name).Cells(i, 7).Value
               
        'If opening price is zero, we need to use next value for opening price
        If YearOpen = 0 Then
           YearOpen = Sheets(ws.Name).Cells(i, 3).Value
        End If
        
        'Get opening price for the year for this ticker
        'Current ticker is not equal to last ticker, we're at a new stock
        If Sheets(ws.Name).Cells(i, 1).Value <> Sheets(ws.Name).Cells(i - 1, 1).Value Then
            YearOpen = Sheets(ws.Name).Cells(i, 3).Value
        End If
            
        'New symbol next line, do final tally for this ticker
        If Sheets(ws.Name).Cells(i + 1, 1).Value <> Sheets(ws.Name).Cells(i, 1).Value Then
            YearClose = Sheets(ws.Name).Cells(i, 6)
            YearChange = YearClose - YearOpen

            'Populate Ticker
            Sheets(ws.Name).Cells(ReportRow, 10) = Ticker
            
            'format YearChange green/red depending on value
            If YearChange >= 0 Then
                Sheets(ws.Name).Cells(ReportRow, 11).Interior.ColorIndex = 4
            Else
                Sheets(ws.Name).Cells(ReportRow, 11).Interior.ColorIndex = 3
            End If
            
            'Populate YearChange
            Sheets(ws.Name).Cells(ReportRow, 11) = YearChange
            
            'Populate PercenChange
            If (YearOpen <> 0 And YearClose <> 0) Then
                PercentChange = FormatPercent((YearClose - YearOpen) / YearOpen)
            ElseIf (YearOpen = 0 Or YearClose = 0) Then
                PercentChange = FormatPercent(0)
            End If
            Sheets(ws.Name).Cells(ReportRow, 12) = PercentChange
            
            'Populate Total Volume
            Sheets(ws.Name).Cells(ReportRow, 13) = Volume

            'Reset variables for next ticker
            Volume = 0
            YearOpen = 0
            YearClose = 0
            PercentChange = 0
            
            'Move to next row for the report output
            ReportRow = ReportRow + 1
        End If
    
    Next i

'Now loop thru and find:
'   Greatest % Increase
'   Greatest % Decrease
'   Greatest Total Volume

    'Preset each variable to the first row of data
    GreatestIncreaseTicker = Sheets(ws.Name).Cells(2, 10).Value
    GreatestIncreasePercent = Sheets(ws.Name).Cells(2, 12).Value
    GreatestDecreaseTicker = Sheets(ws.Name).Cells(2, 10).Value
    GreatestDecreasePercent = Sheets(ws.Name).Cells(2, 12).Value
    GreatestVolumeTicker = Sheets(ws.Name).Cells(2, 10).Value
    GreatestVolumeNumber = Sheets(ws.Name).Cells(2, 13).Value
    
    
    'Loop thru ticker column.  Start at 3, since used the first row as our base set.
    For i = 3 To LastRow

        'Find Greatest % Increase
        If Sheets(ws.Name).Cells(i, 12).Value > GreatestIncreasePercent Then
            GreatestIncreaseTicker = Sheets(ws.Name).Cells(i, 10).Value
            GreatestIncreasePercent = Sheets(ws.Name).Cells(i, 12).Value
        End If

        'Find Greatest % Decrease
        If Sheets(ws.Name).Cells(i, 12).Value < GreatestDecreasePercent Then
            GreatestDecreaseTicker = Sheets(ws.Name).Cells(i, 10).Value
            GreatestDecreasePercent = Sheets(ws.Name).Cells(i, 12).Value
        End If

       'Find Greatest Total Volume
        If Sheets(ws.Name).Cells(i, 13).Value > GreatestVolumeNumber Then
            GreatestVolumeTicker = Sheets(ws.Name).Cells(i, 10).Value
            GreatestVolumeNumber = Sheets(ws.Name).Cells(i, 13).Value
       End If
    Next i

    'Now populate greatest increase/decrease/volume
    Sheets(ws.Name).Cells(2, 16).Value = GreatestIncreaseTicker
    Sheets(ws.Name).Cells(2, 17).Value = FormatPercent(GreatestIncreasePercent)
    
    Sheets(ws.Name).Cells(3, 16).Value = GreatestDecreaseTicker
    Sheets(ws.Name).Cells(3, 17).Value = FormatPercent(GreatestDecreasePercent)

    Sheets(ws.Name).Cells(4, 16).Value = GreatestVolumeTicker
    Sheets(ws.Name).Cells(4, 17).Value = GreatestVolumeNumber

Next ws

End Sub

