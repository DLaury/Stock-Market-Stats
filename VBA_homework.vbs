Sub WorksheetLoop()

    ' Declare Current as a worksheet object variable.
    Dim currentWS As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each currentWS In Worksheets
        
        ' Create variables
        Dim ticker As String
        Dim vol As Double
        Dim stockOpen As Double
        Dim stockClose As Double
        Dim yearChange As Double
        Dim perChange As Double
        Dim hiPercent As Double
        Dim lowPercent As Double
        Dim lastRow As Double
        Dim changeRow As Double
        Dim base As Double
        Dim origin As Double
        Dim hiPName As String
        Dim lowPName As String
        Dim hiVol As Double
        Dim hiVolName As String
        
        ' Count to last row
        lastRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' set initial values
        vol = 0
        base = 2
        origin = 2
        currentWS.Cells(1, 9).Value = "Ticker"
        currentWS.Cells(1, 10).Value = "Total Stock Volume"
        currentWS.Cells(1, 11).Value = "Yearly Change"
        currentWS.Cells(1, 12).Value = "Percent Change"
        currentWS.Cells(1, 16).Value = "Ticker"
        currentWS.Cells(1, 17).Value = "Value"
        currentWS.Cells(2, 15).Value = "Greatest % increase"
        currentWS.Cells(3, 15).Value = "Greatest % decrease"
        currentWS.Cells(4, 15).Value = "Greatest total volume"
        
        ' Iterate through the last row
        For i = 2 To lastRow
            
            ' If the current cell equals the next cell add on to the total volume
            If currentWS.Cells(i, 1).Value = currentWS.Cells(i + 1, 1).Value Then
                vol = vol + currentWS.Cells(i, 7).Value
                
            ' Otherwise add the final value to total volume and write the ticker and volume into columns
            Else
                vol = vol + currentWS.Cells(i, 7).Value
                
                ' Find year close and open
                stockOpen = currentWS.Cells(origin, 3).Value
                stockClose = currentWS.Cells(i, 6).Value
                
                ' calculate year change and percent change dealing with some stocks not opening until mid year
                ' Run only if stockOpen is = 0
                If stockOpen = 0 Then
                    
                    ' From the beginning of the stock to the end check for this
                    For m = origin To i
                    
                        ' If the current cell is not equal to the next one do this
                        If currentWS.Cells(m, 3).Value <> currentWS.Cells(m + 1, 3).Value Then
                            
                            ' input stockOpen value and calculate year change and stock close
                            stockOpen = currentWS.Cells(m + 1, 3).Value
                            yearChange = stockClose - stockOpen
                            perChange = yearChange / stockOpen
                            
                            ' exit loop
                            m = i
                        End If
                    Next m
                Else
                    
                    ' Calculate year change and percent change for when a stock opened at the beginning of the year
                    yearChange = stockClose - stockOpen
                    perChange = yearChange / stockOpen
                End If
                
                If stockClose = 0 Then
                    
                    ' From the end of the stock to the beginning check for this
                    For n = i To origin Step -1
                    
                        ' If the current cell is not equal to the previous one do this
                        If currentWS.Cells(n, 3).Value <> currentWS.Cells(n - 1, 3).Value Then
                            
                            ' input stockClose value and calculate year change and stock close
                            stockClose = currentWS.Cells(n - 1, 3).Value
                            yearChange = stockClose - stockOpen
                            perChange = yearChange / stockOpen
                            
                            ' exit loop
                            n = origin
                        End If
                    Next n
                Else
                    
                    ' Calculate year change and percent change for when a stock opened at the beginning of the year
                    yearChange = stockClose - stockOpen
                    perChange = yearChange / stockOpen
                End If

                
                'set ticker value
                ticker = currentWS.Cells(i, 1).Value
            
                ' output values
                currentWS.Cells(base, 9).Value = ticker
                currentWS.Cells(base, 10).Value = vol
                currentWS.Cells(base, 11).Value = yearChange
                currentWS.Cells(base, 12).Value = perChange
                currentWS.Cells(base, 12).Style = "Percent"
                currentWS.Cells(base, 12).NumberFormat = "0.00%"
                
                ' reset volume for next ticker
                vol = 0
                
                'move base down so that the next ticker will be one row down
                base = base + 1
                
                ' set origin to the first row of the year
                origin = i + 1
                
            End If
           
        Next i
    
        ' find last row of yearly change
        changeRow = currentWS.Cells(Rows.Count, 11).End(xlUp).Row
            
        ' iterate down the row to conditionally change the cell colors
        For j = 2 To changeRow
            
            ' if the current value is greater than or equal to 0 make the box green
            If currentWS.Cells(j, 11).Value >= 0 Then
                currentWS.Cells(j, 11).Interior.ColorIndex = 4
            
            ' if the current value is less than 0 make the box red
            Else
                currentWS.Cells(j, 11).Interior.ColorIndex = 3
            End If
                
        Next j
        
        
        ' set original value of percent
        hiPercent = 0
        lowPercent = 0
        
        ' iterate through the row of percentages
        For k = 2 To changeRow
        
            ' check if the value is higher than what is currently in hiPercent.
            If currentWS.Cells(k, 12).Value > hiPercent Then
                
                ' Add to variable if it is higher and putting the ticker name into hiPName
                hiPercent = currentWS.Cells(k, 12).Value
                hiPName = currentWS.Cells(k, 9).Value
            End If
            
            '  check if the value is lower than what is currently in lowPercent
            If currentWS.Cells(k, 12).Value < lowPercent Then
            
                ' Add to variable if it is lower and putting the ticker name into lowPName
                lowPercent = currentWS.Cells(k, 12).Value
                lowPName = currentWS.Cells(k, 9).Value
            End If
            
         Next k
         
         ' write the data to the spreadsheet
         currentWS.Cells(2, 17).Value = hiPercent
         currentWS.Cells(2, 17).Style = "Percent"
         currentWS.Cells(2, 17).NumberFormat = "0.00%"
         currentWS.Cells(3, 17).Value = lowPercent
         currentWS.Cells(3, 17).Style = "Percent"
         currentWS.Cells(3, 17).NumberFormat = "0.00%"
         currentWS.Cells(2, 16).Value = hiPName
         currentWS.Cells(3, 16).Value = lowPName
         
         ' setting inital value to 0
         hiVol = 0
         
         ' iterate through the column of total volume
         For l = 2 To changeRow
         
            ' if value is greater than what is already in hiVol add it to hiVol and putting the ticker name in hiVolName
            If currentWS.Cells(l, 10).Value > hiVol Then
                hiVol = currentWS.Cells(l, 10).Value
                hiVolName = currentWS.Cells(l, 9).Value
            End If
         
          Next l
            
          ' put data into spreadsheet
          currentWS.Cells(4, 17).Value = hiVol
          currentWS.Cells(4, 16).Value = hiVolName
           
          ' auto fitting values
          currentWS.Columns("A:Q").AutoFit
          
    Next

End Sub