' DABC VBA Stocks homework for Jim Comas
' Part 1: Easy challenge
' Last modified: 10/20/2019
' Notes
'   this is refactor of my oringal attempt from Friday 10/18/2019
'   this adapts DABC vba day 3 class activities for cleaner iterating
'
' loop through all the stocks for one year
' for each run and take the following information.
  '* The ticker symbol.
  '* Yearly change from opening price at the beginning of a given year
  '**                to the closing price at the end of that year.
  '* The percent change from opening price at the beginning of a given year
  '**                     to the closing price at the end of that year.
  '* The total stock volume of the stock.

Sub vbaStocks_easy_main()
    ' Declare and set column locations for sheet structure
    Dim summaryRow As Integer      ' current row of entering summary table
    Dim sumTickerCol As String     ' alpha char of summary column for ticker symbol
    Dim sumPercentCol As String    ' alpha char of summary column for percent change
    Dim sumTotalVolCol As String   ' alpha char of summary total volume
    Dim sumYearlyChgCol As String  ' alpha char summary column for yearly
    
    Dim sumOpenValCol As String     ' alpha char of summary yearl open val column (for debug)
    Dim sumCloseValCol As String    ' alpha char of summary yearl open val column (for debug)
    
    Dim nameCol As Integer      ' Column # of ticker name data
    Dim openCol As Integer      ' Column # of open value data
    Dim closeCol As Integer     ' Column # of close value data
    Dim volCol As Integer       ' Column # of volume data

    ' Grab the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define column headers and locations of data and summary
    Call defTableColsByRef(ws, _
                    sumTickerCol, sumPercentCol, sumTotalVolCol, sumYearlyChgCol, _
                    nameCol, openCol, closeCol, volCol)
    Call defOpenCloseColsByRef(ws, sumOpenValCol, sumCloseValCol)
    
    ' initialize location of 1st row of summary data to fill
    summaryRow = 2

    'debug message
    'MsgBox "sumTickerCol: " & sumTickerCol & ", sumPercentCol: " & sumPercentCol _
    '        & ", sumTotalVolCol: " & sumTotalVolCol _
    '        & vbCrLf _
    '        & "[nameCol: " & nameCol & ", openCol: " & openCol _
    '        & ", closeCol: " & closeCol & ", volumeCol: " & volCol & "]"
 
    ' Declare and initialize variables
    Dim newStock As Boolean         ' flag for when new stock is detected in the run
    Dim tickerName As String        ' current Ticker Name
    Dim openValue As Double         ' stock value at opening of year
    Dim closeValue As Double        ' stock value at close of year
    Dim runningVolume As Double     ' running total for volume
    Dim yearlyChange As Double      ' yearly change of stock value
    Dim percentChange As Double     ' %-change of stock value for the year

    'find the last row of values in the sheet
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' first row of data is a new stock.  initialize the flag
    newStock = True

    ' Loop through all stock daily rows. 
    For i = 2 To lastRow

        ' if new stock detected, set the tracking values
        If (newStock) Then
            ' Reset the summary & tracking values
            newStock = False
            runningVolume = 0
            yearlyChange = 0
            percentChange = 0
            closeValue = 0

            ' intialize curent stock name and opening value
            tickerName = ws.Cells(i, nameCol).Value    ' current stock ticker name
            openValue = ws.Cells(i, openCol).Value  ' current stock year opening value
        End If

        ' Check if next row is still for current stock ticker,
        If ws.Cells(i + 1, nameCol).Value = ws.Cells(i, nameCol).Value Then
            ' add to runningVolume and continue the loop
            runningVolume = runningVolume + ws.Cells(i, volCol).Value

        Else ' if this is final row for current stock ticker
        
            ' Calculate summary values for current stock (runningValue, percentChange)
            ' Add final row volume to the Running Volume Total
            runningVolume = runningVolume + ws.Cells(i, volCol).Value
            ' Calculate % change in volume
            closeValue = ws.Cells(i, closeCol)
            yearlyChange = closeValue - openValue

            'check for a div/0 error with open value
            If (openValue <> 0) Then
              percentChange = yearlyChange / openValue
            Else
               ' MsgBox "opening value is 0 for stock [" & tickerName & "]"
               percentChange = 0
            End If
             
            'debug check for closeValue = 0
            'debug If (closeValue = 0) Then
            'debug    MsgBox "closing value is 0 for stock [" & tickerName & "]"
            'debug End If
            
            ' set the TickerName in the Summary Table
            ws.Range(sumTickerCol & summaryRow).Value = tickerName
            ' set the Total Volume to the Summary Table
            ws.Range(sumTotalVolCol & summaryRow).Value = runningVolume
            ' set the yearly change to the suammry table
            ws.Range(sumYearlyChgCol & summaryRow).Value = yearlyChange
            
            ' format the yearly change cell. If > 0 green else if < 0 red.
            If yearlyChange > 0 Then
                ws.Range(sumYearlyChgCol & summaryRow).Interior.ColorIndex = 4
                'ws.Range(sumOpenValCol & summaryRow).Value = "Green"
            Else
                ws.Range(sumYearlyChgCol & summaryRow).Interior.ColorIndex = 3
                'ws.Range(sumOpenValCol & summaryRow).Value = "Red"
            End If
                        
            ' set the percent volume change to the Summary Table
            ws.Range(sumPercentCol & summaryRow).Value = percentChange
            ws.Range(sumPercentCol & summaryRow).NumberFormat = "0.00%"
          
            'debug set summary yearly open value and close value data
            'ws.Range(sumOpenValCol & summaryRow).Value = openValue
            'ws.Range(sumCloseValCol & summaryRow).Value = closeValue

           ' Add one to the summary table row
            summaryRow = summaryRow + 1
            
            ' detected next row as new stock. set newStock flag to TRUE
            newStock = True
        End If

    Next i

    ' Autofit to display data
    ws.Columns("A:N").AutoFit

End Sub

' Set the locations of elements of the sheet and summary tabel
' these values are passed ByRef.  So setting in this subroutine...
' should set the values in variables in the calling (e.g main) Sub
Sub defTableColsByRef(ws As Worksheet, _
        sumTickerCol As String, sumPercentCol As String, sumTotalVolCol As String, sumYearlyChgCol As String, _
        nameCol As Integer, openCol As Integer, closeCol As Integer, volumeCol As Integer)
    
    ' define the columns
    sumTickerCol = "I"
    sumYearlyChgCol = "J"
    sumPercentCol = "K"
    sumTotalVolCol = "L"

    nameCol = 1
    openCol = 3
    closeCol = 6
    volumeCol = 7
    
    ' set the headers
    ws.Range(sumTickerCol & "1").Value = "Ticker"
    ws.Range(sumTotalVolCol & "1").Value = "Total Stock Volume"
    ws.Range(sumPercentCol & "1").Value = "% change"
    ws.Range(sumYearlyChgCol & "1").Value = "Yearly Change"
    

End Sub

Sub defOpenCloseColsByRef(ws As Worksheet, sumOpenValCol As String, sumCloseValCol As String)
    sumOpenValCol = "M"
    sumCloseValCol = "N"
    
    'ws.Range(sumOpenValCol & "1").Value = "Year Open Value"
    'ws.Range(sumCloseValCol & "1").Value = "Year Close Value"
    
End Sub

Sub myDebugMsg1(ws As Worksheet, ByVal prevRow As Long, ByVal curRow As Long)

    'MsgBox "This is how" & "to get a new line" _
    '   & vbCrLf _
    '   & " & use '_' to extend vba code across multiple lines"
    MsgBox "enter myDebugMsg1 " & ws.Name & ", " & Str(prevRow) & ", " & curRow
    
    If (prevRow > 1) Then
        MsgBox "Next stock in sheet [" & ws.Name & "]" _
                & vbCrLf _
                & "previous row: " & Str(prevRow) _
                & " Cells: " & ws.Cells(prevRow, 1) _
                & " | " & Str(ws.Cells(prevRow, 2)) _
                & " | " & Str(ws.Cells(prevRow, 3)) _
                & " | " & Str(ws.Cells(prevRow, 4)) _
                & " | " & Str(ws.Cells(prevRow, 5)) _
                & " | " & Str(ws.Cells(prevRow, 6)) _
                & " | " & Str(ws.Cells(prevRow, 7)) _
                & vbCrLf _
                & "current row: " & Str(curRow) _
                & " Cells: " & ws.Cells(curRow, 1) _
                & " | " & Str(ws.Cells(curRow, 2)) _
                & " | " & Str(ws.Cells(curRow, 3)) _
                & " | " & Str(ws.Cells(curRow, 4)) _
                & " | " & Str(ws.Cells(curRow, 5)) _
                & " | " & Str(ws.Cells(curRow, 6)) _
                & " | " & Str(ws.Cells(curRow, 7))
    Else
        MsgBox "This is first stock in sheet [" & ws.Name & "]" _
                & vbCrLf _
                & "current row: " & Str(curRow) _
                & " Cells: " & ws.Cells(curRow, 1) _
                & " | " & Str(ws.Cells(curRow, 2)) _
                & " | " & Str(ws.Cells(curRow, 3)) _
                & " | " & Str(ws.Cells(curRow, 4)) _
                & " | " & Str(ws.Cells(curRow, 5)) _
                & " | " & Str(ws.Cells(curRow, 6)) _
                & " | " & Str(ws.Cells(curRow, 7))
    End If
End Sub


