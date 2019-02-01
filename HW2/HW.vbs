Sub wallStreet()
    'Easy
    'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
    'Ticker to coincide with volume

    'defaults
    Dim totalVolume As Double
    Dim i As Long
    Dim j As Long
    Dim titleCount As Long
    Dim openValue As Double
    Dim closeValue As Double
    Dim openPlace As Long
    Dim closePlace As Long
    Dim valueChange As Double
    Dim percentChange As Double
    
    Dim gInc As Double
    Dim gDec As Double
    Dim gTotVol As Double

    Dim gIncTicker As String
    Dim gDecTicker As String
    Dim gTotVolTicker As String
    

    titleCount = 2
    totalVolume = 0

    Dim ws As Worksheet
    'loop per year/worksheet
    
    For Each ws In Worksheets
        
        
        
        'FIX: has to be a better way than hardcoding
        openPlace = 2
        openValue = ws.Cells(openPlace, 3).Value
        
        'Placement of Values under Header
        titleCount = 2
        
        'Set Row Length for "for loop"
        endOfSheet = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        For i = 2 To endOfSheet
            'Add up volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'Conditional to test whether the stock symbol is the same
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            
            'Open value (first row first column) to close value (end row end column)
            closeValue = ws.Cells(i, 6).Value
            
            valueChange = closeValue - openValue
            ws.Cells(titleCount, 10).Value = valueChange
        

            'Yearly Change
                'Make sure not divided by zero
                If (openValue > 0) Then
                    '%change is (close-open)/ open
                    percentChange = (closeValue - openValue) / openValue
                Else
                    percentChange = 0
                End If
                
                'Make sure data type is percent based
                ws.Cells(titleCount, 11).Value = Format(percentChange, "#.##%")
                
                'Change Color of Cell
                If (percentChange > 0) Then
                ws.Cells(titleCount, 10).Interior.ColorIndex = 4
                ElseIf (percentChange = 0) Then
                ws.Cells(titleCount, 10).Interior.ColorIndex = 0
                ws.Cells(titleCount, 11).Value = ""
                Else
                ws.Cells(titleCount, 10).Interior.ColorIndex = 3
                End If
                
            'Summary of each Ticker with its corresponding value
            'Each Ticker name
            ws.Cells(titleCount, 9).Value = ws.Cells(i, 1).Value
            'Each Ticker Balance
            ws.Cells(titleCount, 12).Value = totalVolume
            
            
            
            'clear volume for next variable
            totalVolume = 0
            
            'increase row count for Next Ticker
            titleCount = titleCount + 1
            
            'make sure the next openValue starts at the top of the ticker list
            openValue = ws.Cells(i + 1, 3).Value
            End If
            
        Next i

        'New For Loop.  Tried iterating in "i loop" as it was occurring but had problems
        For j = 2 To endOfSheet
        
        'Since we are looking for greatest % inc, we can test that here by comparing one % with next and only swapping when it is.
            If (ws.Cells(j, 11).Value > gInc) Then
                gInc = ws.Cells(j, 11).Value
                gIncTicker = ws.Cells(j, 9).Value
            End If
                    
                              
        'Greatest Decrease (similar here)
            If (ws.Cells(j, 11).Value < gDec) Then
                gDec = ws.Cells(j, 11).Value
                gDecTicker = ws.Cells(j, 9).Value
            End If
                    
        'Calculate for Greatest Volume
            If (ws.Cells(j, 12).Value > gTotVol) Then
                gTotVol = ws.Cells(j, 12).Value
                gTotVolTicker = ws.Cells(j, 9).Value
            End If
            
        ws.Cells(2, 17).Value = Format(gInc, "#.##%")
        ws.Cells(2, 16).Value = gIncTicker
        ws.Cells(3, 17).Value = Format(gDec, "#.##%")
        ws.Cells(3, 16).Value = gDecTicker
        ws.Cells(4, 17).Value = gTotVol
        ws.Cells(4, 16).Value = gTotVolTicker
        Next j
            'Default back to zero for next sheet
            gInc = 0
            gDec = 0
            gTotVol = 0
    Next ws

    'Hard
    '* Your solution will include everything from the moderate challenge.
    '* Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
    '* Solution will look as follows.
    
    'Challenge?
End Sub



