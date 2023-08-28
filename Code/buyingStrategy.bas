Attribute VB_Name = "buyingStrategy"
Private Sub loadFiles()
    Dim lastRow As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileName As String
    Dim cell_i As Range
    
    lastRow = ThisWorkbook.Worksheets("StockList").Range("B1").End(xlDown).Row
    
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For Each cell_i In ThisWorkbook.Worksheets("StockList").Range("B2:B" & lastRow)
        fileName = cell_i.Value
        Set wb = Workbooks.Open("/Users/ethanchen/Desktop/AEM2830/StockPrices/" & fileName & ".csv")
        wb.Worksheets(fileName).Activate
        Cells().Copy
            
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = fileName
        ws.Activate
        Range("A1").PasteSpecial
            
        wb.Close
    Next
End Sub

'this sub is used as a means to start our project, it will be revised such that it can be used as a button
Public Sub trackPercentageDrop()
    Dim cell_i As Range
    Dim stockDate As Range
    Dim lastRowSL As Integer
    Dim rowNoDT As Integer
    Dim percDrop As Double
    Dim openEarlier As Double
    Dim closeLater As Double
    Dim ws As Worksheet
    Dim i As Integer
    Dim j As Integer
    Dim min As Double
    Dim minRow As Integer
    Dim rowPort As Integer
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Create a Worksheet to track all the percentage drops by Stock, Date, and the magnitude of drop
    Set ws = ThisWorkbook.Worksheets("Drop Tracker")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Stock"
    ws.Range("B1").Value = "Date"
    ws.Range("C1").Value = "Price"
    ws.Range("D1").Value = "% Drop of Previous 5 Days"
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Borders(xlEdgeBottom).Color = vbBlack
    
    'initialize the row Indices for looping
    lastRowSL = ThisWorkbook.Worksheets("StockList").Range("B1").End(xlDown).Row
    rowNoDT = 2
    
    'iterate over each stock
    For Each cell_i In ThisWorkbook.Worksheets("StockList").Range("B2:B" & lastRowSL)
        'can only observe dates of year 2016 and 2017
        Set stockDate = ThisWorkbook.Worksheets(cell_i.Value).Range("A6")
        Do While (Year(stockDate) < 2018)
            'define percDrop to be the return of the (close of the later date - open of the earlier date) / open of earlier
            closeLater = get_price(cell_i.Value, stockDate.Value, "Close")
            openEarlier = get_price(cell_i.Value, stockDate.Offset(-4, 0).Value, "Open")
            percDrop = 0
            If openEarlier = 0 Then
                MsgBox ("The stock " & cell_i.Value & " is missing a value on " & stockDate.Offset(-4, 0).Value & "!")
            Else
                percDrop = (closeLater - openEarlier) / openEarlier
            End If
            'if the percentage drop is less than -0.20, we keep track of that
            If (percDrop < -0.2) Then
                ws.Range("A" & rowNoDT).Value = cell_i.Value
                'log the day after so that we can buy it
                ws.Range("B" & rowNoDT).Value = stockDate.Offset(1, 0).Value
                ws.Range("C" & rowNoDT).Value = get_price(cell_i.Value, stockDate.Offset(1, 0).Value)
                ws.Range("D" & rowNoDT).Value = percDrop
                rowNoDT = rowNoDT + 1
                'we exit the loop and buy the stock (if it is part of the first five)
                Exit Do
            End If
            Set stockDate = stockDate.Offset(1, 0)
        Loop
    Next
    
    'this is a code that get the first earliest occurence and copies it over to portfolio
    For i = 1 To 5
        min = ws.Range("B2").Value
        minRow = 2
        rowNoDT = rowNoDT - 1
        For j = 2 To rowNoDT
            If (min > ws.Range("B" & j).Value) Then
                min = ws.Range("B" & j).Value
                minRow = j
            End If
        Next
        'copy over to portfolio
        rowPort = i + 3
        ThisWorkbook.Worksheets("Portfolio").Range("C" & rowPort).Value = ws.Range("A" & minRow).Value
        ThisWorkbook.Worksheets("Portfolio").Range("D" & rowPort).Value = 1000000 / ws.Range("C" & minRow).Value
        ThisWorkbook.Worksheets("Portfolio").Range("E" & rowPort).Value = ws.Range("B" & minRow).Value
        ThisWorkbook.Worksheets("Portfolio").Range("F" & rowPort).Value = ws.Range("C" & minRow).Value
        
        'clear the row and repeat
        ws.Range("A" & minRow & ":D" & minRow).Delete
    Next
End Sub

