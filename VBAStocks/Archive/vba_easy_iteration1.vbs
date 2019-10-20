' loop through all the stocks for one year
' for each run and take the following information.
  '* The ticker symbol.
  '* Yearly change from opening price at the beginning of a given year
  '**                to the closing price at the end of that year.
  '* The percent change from opening price at the beginning of a given year
  '**                     to the closing price at the end of that year.
  '* The total stock volume of the stock.

Sub vbaStocks_easy_main()
    'declare variables
    Dim curTicker As String
    Dim prevTicker As String
    Dim curOpen As Double
    Dim curClose As Double
    Dim curVolume As Double
    Dim curRow As Range
    Dim prevRow As Range
    Dim prevRowIndex As Long
    Dim summaryRow As Integer
    Dim tickerCol As Integer
    Dim percentCol As Integer
    Dim totalVolCol As Integer
    
   
    Dim curWS As Worksheet
    
    'set the current worksheet
    Set curWS = Application.ActiveSheet
    
    'Create header in sheet
    tickerCol = 9
    percentCol = 10
    totalVolCol = 11
    summaryRow = 2
    curWS.Cells(1, tickerCol).Value = "Ticker"
    curWS.Cells(1, percentCol).Value = "% change"
    curWS.Cells(1, totalVolCol).Value = "Total Stock Volume"

    'loop through all stocks rows in the sheet
    Set prevRow = curWS.Rows(2)
    prevRowIndex = 2
    prevTicker = prevRow.Cells(1, 1)
    curOpen = prevRow.Cells(1, 3)
    
    For i = 2 To curWS.Rows.Count
        Set curRow = curWS.Rows(i)
        curTicker = curRow.Cells(1, 1)
        
        ' check to see if <ticker> has changed
        If (curTicker <> prevTicker) Then
            ' Debug
           'MsgBox "Current Ticker [" & curTicker & "]" & vbCrLf _
            '    & "has changed from Previous Ticker [" + prevTicker + "]" & vbCrLf _
            '    & "at row [" + Str(i) + "]"
            
            ' Debug msgbox
           ' Call myDebug(prevRow, prevRowIndex, curRow, i)
            
            'calculate and set end of ticker values
            curClose = prevRow.Cells(1, 6)
            curWS.Cells(summaryRow, tickerCol) = prevTicker
            curWS.Cells(summaryRow, percentCol) = ((curClose - curOpen) / curOpen)
            curWS.Cells(summaryRow, percentCol).NumberFormat = "##########0.00%"
            curWS.Cells(summaryRow, totalVolCol) = curVolume
            summaryRow = summaryRow + 1
            
            'reset values for next ticker
            prevTicker = curTicker
            curVolume = 0
            curOpen = curRow.Cells(1, 3)

        End If
        
        curVolume = curVolume + curRow.Cells(1, 7)
        Set prevRow = curRow
        prevRowIndex = i

    Next i

End Sub

Sub myDebug(prevRow As Range, ByVal prevIndex As Long, curRow As Range, ByVal curIndex As Long)
    'MsgBox "This is how" & "to get a new line" _
    '   & vbCrLf _
    '   & " & extend code across multiple lines"
    
    MsgBox "previous row: " & Str(prevIndex) _
            & " Cells: " & prevRow.Cells(1, 1) _
            & " | " & Str(prevRow.Cells(1, 2)) _
            & " | " & Str(prevRow.Cells(1, 3)) _
            & " | " & Str(prevRow.Cells(1, 4)) _
            & " | " & Str(prevRow.Cells(1, 5)) _
            & " | " & Str(prevRow.Cells(1, 6)) _
            & " | " & Str(prevRow.Cells(1, 7)) _
            & vbCrLf _
            & "current row: " & Str(curIndex) _
            & " Cells: " & curRow.Cells(1, 1) _
            & " | " & Str(curRow.Cells(1, 2)) _
            & " | " & Str(curRow.Cells(1, 3)) _
            & " | " & Str(curRow.Cells(1, 4)) _
            & " | " & Str(curRow.Cells(1, 5)) _
            & " | " & Str(curRow.Cells(1, 6)) _
            & " | " & Str(curRow.Cells(1, 7))

End Sub

       
        'detect beginning of stock ticker
        'loop throu all stock entries for that stock
        'while not end of stock
        'detect end of ticker

        'calculate
