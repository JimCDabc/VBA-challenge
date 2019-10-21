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
    Dim summaryRow As Integer       ' current row for entering summary table
    Dim sumTickerCol As String     ' column of summary ticker range
    Dim sumPercentCol As String    ' column of summary percent range
    Dim sumTotalVolCol As String   ' column of summary total volume Range
    Dim nameCol As Integer
    Dim openCol As Integer
    Dim closeCol As Integer
    Dim volCol As Integer

    ' Grab the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Call defTableColsByRef(ws, sumTickerCol, sumPercentCol, sumTotalVolCol, nameCol, openCol, closeCol, volCol)
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
    Dim percentChange As Double
    
    newStock = True

    'find the last row of values in the sheet
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox "Last Row: [" & lastRow & "]"

    ' set first opening value
    'openValue = ws.Cells(2, openCol)

    ' Loop through all credit card purchases
    For i = 2 To lastRow

        ' if new stock detected, set the tracking values
        If (newStock) Then
            ' debug message to print precious and current row info
            ' Call myDebugMsg1(ws, (i - 1), i)
            
            
            ' Reset the summary & tracking values
            newStock = False
            runningVolume = 0
            percentChange = 0
            'closeValue = 0

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
            'closeValue = ws.Cells(i, closeCol)
            'If (openValue <> 0) Then
            '    percentChange = (closeValue - openValue) / openValue
            'Else
            '    percentChange = 0
            'End If
            
            ' set the TickerName in the Summary Table
            ws.Range(sumTickerCol & summaryRow).Value = tickerName
            ' set the Total Volume to the Summary Table
            ws.Range(sumTotalVolCol & summaryRow).Value = runningVolume
            ' set the percent volume change to the Summary Table
            'ws.Range(sumPercentCol & summaryRow).Value = percentChange
            'ws.Range(sumPercentCol & summaryRow).NumberFormat = "0.00%"
            ' Add one to the summary table row
            summaryRow = summaryRow + 1

            ' detected next row as new stock. set newStock flag to TRUE
            newStock = True
        End If

    Next i

End Sub

' Set the locations of elements of the sheet and summary tabel
' these values are passed ByRef.  So setting in this subroutine...
' should set the values in variables in the calling (e.g main) Sub
Sub defTableColsByRef(ws As Worksheet, sumTickerCol As String, sumPercentCol As String, sumTotalVolCol As String, _
                nameCol As Integer, openCol As Integer, closeCol As Integer, volumeCol As Integer)
    
    ' define the columns
    sumTickerCol = "I"
    sumTotalVolCol = "J"
    'sumPercentCol = "K"
    nameCol = 1
    openCol = 3
    closeCol = 6
    volumeCol = 7
    
    ' set the headers
    ws.Range(sumTickerCol & "1").Value = "Ticker"
    ws.Range(sumTotalVolCol & "1").Value = "Total Stock Volume"
    'ws.Range(sumPercentCol & "1").Value = "% change"

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





