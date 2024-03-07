Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
'define the variables

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryTableRow As Integer
    
    ' create a loop so that it will work on all three sheets
    
    For Each ws In ThisWorkbook.Worksheets

'populate the values and pull the row

     
        SummaryTableRow = 2 '
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row '
        
'create headers

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
     
  'start the loop and add the calcilations
  
  
     
        For i = 2 To LastRow
         
            If ws.Range("A" & (i + 1)).Value <> ws.Range("A" & i).Value Then
                Ticker = ws.Range("A" & i).Value
                ClosingPrice = ws.Range("F" & i).Value
               YearlyChange = ClosingPrice - OpeningPrice
                
'calculate the percent changes and add the values to our summary table


                If OpeningPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpeningPrice
                End If
                
        
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
                
    'next section to apply conditional formatting positive = green, negative = red, ref: https://learn.microsoft.com/en-us/office/vba/api/excel.interior.color
    'runtime error encountered (6) overflow encountered due to the variable is too large for the data type. Use Total volume to handle large numbers as Double or long
    
    
                If YearlyChange > 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.Color = RGB(255, 0, 0)
                End If
                
              
                TotalVolume = 0
                SummaryTableRow = SummaryTableRow + 1
            End If
            
         
            TotalVolume = TotalVolume + ws.Range("G" & i).Value
            If OpeningPrice = 0 Then
                OpeningPrice = ws.Range("C" & i).Value
            End If
        Next i
        
      'adjust coloumn width ref: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
      
      
        ws.Columns("I:L").AutoFit
    Next ws
End Sub

