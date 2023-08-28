Attribute VB_Name = "sellingStrategy"
Option Explicit


'now we have our five stocks: AMD, FCX, MU, NOW, TSLA which are listed in the portfolio sheet

Sub cleanSheets()
    Dim ws As Worksheet
    
    'In the selling stage we can first hide all other stocks to make workbook more user-friendly
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Report" And _
        ws.Name <> "Portfolio" And _
        ws.Name <> "StockList" And _
        ws.Name <> "Stock Details >>" And _
        ws.Name <> "Graphs" And _
        ws.Name <> ThisWorkbook.Worksheets("portfolio").Range("C4").Value And _
        ws.Name <> ThisWorkbook.Worksheets("portfolio").Range("C5").Value And _
        ws.Name <> ThisWorkbook.Worksheets("portfolio").Range("C6").Value And _
        ws.Name <> ThisWorkbook.Worksheets("portfolio").Range("C7").Value And _
        ws.Name <> ThisWorkbook.Worksheets("portfolio").Range("C8").Value Then
            ws.Visible = False
        End If
    Next
    
End Sub

Sub makeCopy(stockTicker As String)
    Dim counter As Long
    
    'Make a copy of the stock to add analysis columns
    
     ThisWorkbook.Worksheets(stockTicker).Copy After:=ThisWorkbook.Worksheets("Stock Details >>")
     ThisWorkbook.Worksheets(stockTicker & " (2)").Name = stockTicker & "_MA"
     
     
     'Add columns for percChange and Moving Average
     ThisWorkbook.Worksheets(stockTicker & "_MA").Range("H1").Value = "Percentage Change"
     ThisWorkbook.Worksheets(stockTicker & "_MA").Range("I1").Value = "Moving Average"
     
     'Keep only 01/01/2020 data onwards
    counter = 2
    Do Until ThisWorkbook.Worksheets(stockTicker & "_MA").Range("A" & counter).Value = ""
        If Year(Range("A" & counter)) < 2020 Then
            Rows(counter).Delete Shift:=xlUp
        Else
            counter = counter + 1
        End If
    Loop
          
End Sub


Sub SimplePercChange_MA(stockTicker As String)
    Dim openPrice As Double
    Dim highPrice As Double
    Dim percChange As Double
    Dim cell_i As Range
    Dim lastStockRow As Long
    Dim firstDay As String
    Dim secondDay As String
    Dim thirdDay As String
    Dim fourthDay As String
    Dim fifthDay As String
    
    lastStockRow = ThisWorkbook.Worksheets(stockTicker).Range("A1").End(xlDown).Row
    
    For Each cell_i In ThisWorkbook.Worksheets(stockTicker).Range("A2:A" & lastStockRow)
        openPrice = ThisWorkbook.Worksheets(stockTicker).Range("B" & cell_i.Row).Value
        highPrice = ThisWorkbook.Worksheets(stockTicker).Range("C" & cell_i.Row).Value
            If openPrice > 0 Then
                percChange = (highPrice - openPrice) / openPrice
                ThisWorkbook.Worksheets(stockTicker).Range("H" & cell_i.Row).Value = percChange
            End If
    Next
    
    For Each cell_i In ThisWorkbook.Worksheets(stockTicker).Range("H6:H" & lastStockRow)
        firstDay = ThisWorkbook.Worksheets(stockTicker).Range("H" & cell_i.Row - 4).Value
        secondDay = ThisWorkbook.Worksheets(stockTicker).Range("H" & cell_i.Row - 3).Value
        thirdDay = ThisWorkbook.Worksheets(stockTicker).Range("H" & cell_i.Row - 2).Value
        fourthDay = ThisWorkbook.Worksheets(stockTicker).Range("H" & cell_i.Row - 1).Value
        ThisWorkbook.Worksheets(stockTicker).Range("I" & cell_i.Row) = Application.WorksheetFunction.Average(firstDay, secondDay, thirdDay, fourthDay, cell_i.Value)
    Next
End Sub

Sub SellingTracker(stockTicker As String)
    Dim lastRow As Long
    Dim cell_i As Range
    Dim trackerTargetRow As Integer
    
    lastRow = ThisWorkbook.Worksheets(stockTicker).Range("A1").End(xlDown).Row
    
    'move any stock that has MA change > 0.05
    For Each cell_i In ThisWorkbook.Worksheets(stockTicker).Range("I6:I" & lastRow)
        If cell_i.Value > 0.05 Then
            If ThisWorkbook.Worksheets("Selling Tracker").Range("A2") <> "" Then
                        trackerTargetRow = ThisWorkbook.Worksheets("Selling Tracker").Range("A1").End(xlDown).Row + 1
                        'MsgBox trackerTargetRow
                        ThisWorkbook.Worksheets("Selling Tracker").Range("A" & trackerTargetRow).Value = VBA.Left(stockTicker, VBA.Len(stockTicker) - 3)
                        ThisWorkbook.Worksheets("Selling Tracker").Range("B" & trackerTargetRow).Value = cell_i.Offset(0, -8).Value
                        ThisWorkbook.Worksheets("Selling Tracker").Range("C" & trackerTargetRow).Value = cell_i.Value
                        
                    
                    Else
                        trackerTargetRow = 2
                        'MsgBox trackerTargetRow
                        ThisWorkbook.Worksheets("Selling Tracker").Range("A" & trackerTargetRow).Value = VBA.Left(stockTicker, VBA.Len(stockTicker) - 3)
                        ThisWorkbook.Worksheets("Selling Tracker").Range("B" & trackerTargetRow).Value = cell_i.Offset(0, -8).Value
                        ThisWorkbook.Worksheets("Selling Tracker").Range("C" & trackerTargetRow).Value = cell_i.Value
                    End If
        End If
    Next
    
End Sub


Sub readyToSell2()
    Dim lastRow As Integer
    Dim cell_i As Range
    Dim cell_j As Range
    Dim cell_k As Range
    Dim sellDate As Date
    Dim sellPrice As Double
    Dim lastStockRow As Long
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Do Until ThisWorkbook.Worksheets("Selling Tracker").Range("A2").Value = ""
        ThisWorkbook.Worksheets("Selling Tracker").Rows(2).EntireRow.Delete
    Loop
    
    For Each cell_i In ThisWorkbook.Worksheets("Portfolio").Range("C4:C8")
        Call makeCopy(cell_i.Value)
        Call SimplePercChange_MA(cell_i.Value & "_MA")
        Call SellingTracker(cell_i.Value & "_MA")
            
        lastRow = ThisWorkbook.Worksheets("Selling Tracker").Range("A1").End(xlDown).Row
        For Each cell_j In ThisWorkbook.Worksheets("Selling Tracker").Range("A1:A" & lastRow)
            If cell_j.Value = cell_i.Value Then
                sellDate = DateAdd("d", 1, cell_j.Offset(0, 1).Value)
                ThisWorkbook.Worksheets("Portfolio").Range("G" & cell_i.Row).Value = sellDate
                Exit For
            End If
        Next
        
        lastStockRow = ThisWorkbook.Worksheets(cell_i.Value).Range("A1").End(xlDown).Row
        For Each cell_k In ThisWorkbook.Worksheets(cell_i.Value).Range("A1:A" & lastStockRow)
            If cell_k.Value = sellDate Then
                ThisWorkbook.Worksheets("Portfolio").Range("H" & cell_i.Row).Value = ThisWorkbook.Worksheets(cell_i.Value).Range("A" & cell_k.Row).Offset(0, 5).Value
            Exit For
            End If
        Next
        
    Next
    
    'Delete all MA sheets
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets("Portfolio").Range("C4").Value & "_MA").Delete
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets("Portfolio").Range("C5").Value & "_MA").Delete
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets("Portfolio").Range("C6").Value & "_MA").Delete
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets("Portfolio").Range("C7").Value & "_MA").Delete
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets("Portfolio").Range("C8").Value & "_MA").Delete
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'User should end on Portfolio sheet
    ThisWorkbook.Worksheets("Portfolio").Activate
    
End Sub

Sub calculateReturn()
    Dim cell_i As Range
    
    For Each cell_i In ThisWorkbook.Worksheets("Portfolio").Range("I4:I8")
        If ThisWorkbook.Worksheets("Portfolio").Range("F" & cell_i.Row).Value <> "" And ThisWorkbook.Worksheets("Portfolio").Range("G" & cell_i.Row).Value <> "" Then
            cell_i.Value = get_return(cell_i.Offset(0, -6).Value, cell_i.Offset(0, -4).Value, cell_i.Offset(0, -2).Value)
            cell_i.Offset(0, 1).Value = (cell_i.Offset(0, -1).Value - cell_i.Offset(0, -3).Value) * cell_i.Offset(0, -5).Value
        End If
    Next
    
End Sub

