Attribute VB_Name = "Module1"
' 1. The ticker symbol
' 2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 4. The total stock volume of the stock. The result should match the following image:
' 5. Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
'----------------------------------------------
Sub Mod2Challenge()

'Steps 1-4
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'Create Variables
Dim SummaryTable As String
SummaryTable = 2
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim totalVoume As Double
Dim volume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Create Summary Table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To lastRow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ' Set the Ticker
        Ticker = Cells(i, 1).Value
        ' Print Ticker in Summary Table
        Range("I" & SummaryTable).Value = Ticker
        'Add Up The Open & Close
        OpenPrice = Cells(i, 3).Value + OpenPrice
        ClosePrice = Cells(i, 6).Value + ClosePrice
        ' Create Yearly Change
        YearlyChange = ClosePrice - OpenPrice
        'Print Yearly Change in Summary Table
        Range("J" & SummaryTable).Value = YearlyChange
            
            ' Change Colors
            If YearlyChange > 0 Then
                Range("J" & SummaryTable).Interior.ColorIndex = 4
            ElseIf YearlyChange < 0 Then
                Range("J" & SummaryTable).Interior.ColorIndex = 3
            Else
                Range("J" & SummaryTable).Interior.ColorIndex = 0
            End If
           
         ' PercentChange Calculations
          PercentChange = ((ClosePrice - OpenPrice) / OpenPrice) * 100
          ' Print PercentChange
          Range("K" & SummaryTable).Value = PercentChange
         ' Show as Percentage
          
          Range("K:K").NumberFormat = "0.00%"
        'Total Stock Volume
        
        ' Print in Summary Table
        Range("L" & SummaryTable).Value = totalVolume
        
         
        'Add 1 to Summary Table Row
        SummaryTable = SummaryTable + 1
        ' Reset YearlyChange
        YearlyChange = 0
        ' Rest PercentChange
         PercentChange = 0
        ' Reset totalVolume
        totalVolume = 0
           
           
    Else
        OpenPrice = Cells(i, 3).Value + OpenPrice
        ClosePrice = Cells(i, 6).Value + ClosePrice
        
        totalVolume = Cells(i, 7).Value + totalVolume
    End If
    Next i
    
    
    ' Step 5
    ' Naming
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    
    ' Print Values onto Greatest
    Cells(2, 16).Value = Application.WorksheetFunction.Max(Range("K:K"))
    Cells(3, 16).Value = Application.WorksheetFunction.Min(Range("K:K"))
    Cells(4, 16).Value = Application.WorksheetFunction.Max(Range("L:L"))
    
    ' Print Ticker for each of the Above
    
    Cells(2, 15).Value = Cells(WorksheetFunction.Match(Range("P2").Value, Range("K:K"), 0), 9).Value
    Cells(3, 15).Value = Cells(WorksheetFunction.Match(Range("P3").Value, Range("K:K"), 0), 9).Value
    Cells(4, 15).Value = Cells(WorksheetFunction.Match(Range("P4").Value, Range("L:L"), 0), 9).Value
    Next ws
    
    End Sub
    
