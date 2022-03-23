Attribute VB_Name = "Module1"
Sub Stock_Analysis()
  Dim ShName As String
'setting loop to go through all worksheets
For Each ws In Worksheets
    'ShName = ws.Name
    'MsgBox ShName
    Worksheets(ws.Name).Activate
    
    
    'organize cells for data process
    Range("A1").CurrentRegion.Sort key1:=Range("A1"), order1:=xlAscending, Header:=xlYes
    
    
    'setting Variables
    Dim ticker As String
    Dim youngest As Long
    Dim oldest As Long
    Dim stock_total As Single
    Dim perc_change As Double
    Dim open_price As Double
    Dim close_price As Double
    'Dim yearly_change As Double
    Dim k As Long
    Dim lastrow As Long
    
    'setting columns to save Data
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % decrease"
    Range("O4").Value = "Greatest total volume"
    
    'Conditional Formatting for Change
    Range("J1").EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Range("J1").EntireColumn.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    Range("J1").EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    Range("J1").EntireColumn.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
    Range("J1").FormatConditions.Delete
    
    'Formatting for Percent Fields
    Range("K2").EntireColumn.NumberFormat = "0.00%"
    Range("Q2", "Q3").NumberFormat = "0.00%"
    
    'Formatting Stock VOlume Data
    Range("G1").EntireColumn.NumberFormat = "0"
    Range("L1").EntireColumn.NumberFormat = "0"
    Range("Q4").NumberFormat = "0"
    
    
    'yearly change formula
    'yearly_change = close_price - open_price
    
    'percent change formula
    'perc_change = ((close_price - open_price) / open_price) * 100
    
    'Mark first row
    ticker = "First row"
    
    'declare results row variable
    k = 2
    
    'declaring lastrow Variable
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'starting loop to find ticker
    For i = 2 To lastrow
        
        'Check to see if the ticker has changed
        If ticker <> Cells(i, 1).Value Then
            'Check to see if we need to save the calculations
            If ticker <> "First row" Then
                Cells(k, 9).Value = ticker
                Cells(k, 10).Value = close_price - open_price
                Cells(k, 11).Value = ((close_price - open_price) / open_price) '* 100
                'Cells(k, 12).Value = Cells(i, 7).Value
                
                'Pulling highest, lowest and greatest total
                If Cells(k, 11).Value > Range("Q2").Value Then
                    Range("Q2").Value = Cells(k, 11).Value
                    Range("P2").Value = Cells(k, 9).Value
                End If
                If Cells(k, 11).Value < Range("Q3").Value Then
                    Range("Q3").Value = Cells(k, 11).Value
                    Range("P3").Value = Cells(k, 9).Value
                End If
                If Cells(k, 12).Value > Range("Q4").Value Then
                    Range("Q4").Value = Cells(k, 12).Value
                    Range("P4").Value = Cells(k, 9).Value
                End If
                
                k = k + 1
            End If
            
            
            
            'Set new start values
            ticker = Cells(i, 1).Value
            youngest = Cells(i, 2).Value
            oldest = Cells(i, 2).Value
            open_price = Cells(i, 3).Value
            close_price = Cells(i, 6).Value
            Cells(k, 12).Value = Cells(i, 7).Value
    
        Else
            'data compile for ticker
            Cells(k, 12).Value = Cells(k, 12).Value + Cells(i, 7).Value
            If youngest > Cells(i, 2).Value Then
                open_price = Cells(i, 3).Value
                youngest = Cells(i, 2).Value
            End If
            If oldest < Cells(i, 2).Value Then
                close_price = Cells(i, 6).Value
                oldest = Cells(i, 2).Value
            End If
            
        End If
             
    Next i
        
    'After Data is done compiling adjust column width to fit data
    ActiveSheet.UsedRange.EntireColumn.AutoFit
               
'next worksheet to run data
Next ws
               
    
End Sub


