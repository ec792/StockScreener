Attribute VB_Name = "functions"
Option Explicit

Function get_price(tickerName As String, date_i As Date, Optional priceType = "Adj Close") As Double
    Dim stockData As Range
    Dim lastNo As Integer
    Dim rowNo As Integer
    Dim cell_i As Range
    Dim colNo As Integer
    
    'Set the range by finding the worksheet with the corresponding tickerName
    Set stockData = ThisWorkbook.Worksheets(tickerName).Range("A1")
    
    'find the last row of the dataset
    lastNo = stockData.End(xlDown).Row
    
    'initialize the rowNo so that we can find the corresponding row that matches the given date
    For Each cell_i In ThisWorkbook.Worksheets(tickerName).Range("A1:A" & lastNo)
        If (cell_i.Value = date_i) Then
            'here, we note that A1.Row = 1, but for the sake of using offset A1 is 0, we have to decrease rowNo by 1
            rowNo = cell_i.Row - 1
        End If
    Next
    
    'if the date is not found, then we notify the user and exit the sub
    If rowNo = 0 Then
        MsgBox "The date is not found!"
        Exit Function
    End If
    
    'choose which column to find the price based on the priceType given
    If (priceType = "Adj Close") Then
        colNo = 5
    ElseIf (priceType = "Open") Then
        colNo = 1
    ElseIf (priceType = "High") Then
        colNo = 2
    ElseIf (priceType = "Low") Then
        colNo = 3
    ElseIf (priceType = "Close") Then
        colNo = 4
    Else
        MsgBox "The price type is not correct!"
        Exit Function
    End If
    
    'return the price with the rowNo and the columnNo
    get_price = stockData.Offset(rowNo, colNo).Value
End Function

Function get_return(tickerName As String, date_i As Date, date_j As Date) As Double
    Dim stockData As Range
    Dim lastNo As Integer
    Dim rowNo_i As Integer
    Dim rowNo_j As Integer
    Dim cell_i As Range
    Dim price_i As Double
    Dim price_j As Double

    'Set the range by finding the worksheet with the corresponding tickerName
    Set stockData = ThisWorkbook.Worksheets(tickerName).Range("A1")
    
    'find the last row of the dataset
    lastNo = stockData.End(xlDown).Row
    
    'initialize two variables to find the row corresponding to i and j
    For Each cell_i In ThisWorkbook.Worksheets(tickerName).Range("A1:A" & lastNo)
        'here, we note that A1.Row = 1, but for the sake of using offset A1 is 0, we have to decrease rowNo by 1
        If (cell_i.Value = date_i) Then
            rowNo_i = cell_i.Row - 1
        End If
        If (cell_i.Value = date_j) Then
            rowNo_j = cell_i.Row - 1
        End If
    Next
    
    If rowNo_i = 0 Then
        MsgBox ("The first date is incorrect!")
        Exit Function
    End If
    
    If rowNo_j = 0 Then
        MsgBox ("The second date is incorrect!")
        Exit Function
    End If

    'now we can get the price
    price_i = stockData.Offset(rowNo_i, 5).Value
    price_j = stockData.Offset(rowNo_j, 5).Value
    
    If price_i = 0 Then
        MsgBox ("The stock " & tickerName & " is missing a value on " & date_i & "!")
        Exit Function
    End If
    
    'finally we can calculate the return and return it
    get_return = (price_j - price_i) / price_i
End Function

Sub portfolioValue()
    Dim portfolioValue As Double
    Dim cell_i As Range
    Dim portfolioDate As Date
    Dim price As Double
    
    portfolioDate = InputBox("Enter a date in MM/DD/YYYY format:")
    portfolioValue = 0
    
    For Each cell_i In ThisWorkbook.Worksheets("Portfolio").Range("E4:E8")
        If portfolioDate > #12/31/2021# Then
            MsgBox "This date is outside of range!"
            Exit Sub
        ElseIf portfolioDate >= cell_i.Value And portfolioDate < cell_i.Offset(0, 2).Value Then
            price = get_price(cell_i.Offset(0, -2).Value, portfolioDate, "Adj Close")
            portfolioValue = portfolioValue + (price * cell_i.Offset(0, -1).Value)
        ElseIf portfolioDate >= cell_i.Value And cell_i.Offset(0, 2).Value = "" Then
            price = get_price(cell_i.Offset(0, -2).Value, portfolioDate, "Adj Close")
            portfolioValue = portfolioValue + (price * cell_i.Offset(0, -1).Value)
        End If
    Next
    
    MsgBox "The portfolio value is $" & Format(VBA.Round(portfolioValue, 2), "#,##0.00") & " at " & CStr(portfolioDate) & "."
End Sub

Sub clearPortfolio()
    ThisWorkbook.Worksheets("Portfolio").Range("C4:J8").ClearContents
End Sub

