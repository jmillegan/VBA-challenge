Attribute VB_Name = "Module1"
Sub stockanalysis()

'To run in each worksheet in the workbook

Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate



'name variable for stocks details
Dim ticker As String
Dim stock_volume As Double
        stock_volume = 0
Dim openprice As Double
       openprice = Cells(2, 3).Value
Dim closeprice As Double
Dim yearly_change As Double
Dim percent_change As Double
       percent_change = 0
' Summary_table that keep information for each ticker
Dim summary_row As Integer
       summary_row = 2
       
    
    
'Loop through all daily entries for yearly change determination
 For i = 2 To 800000
'check that we are within the same ticker
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        Range("I" & summary_row).Value = ticker
        closeprice = Cells(i, 6).Value
        
        yearly_change = closeprice - openprice
        Range("J" & summary_row).Value = yearly_change
             If yearly_change < 0 Then
             Range("J" & summary_row).Interior.ColorIndex = 3
        Else
             Range("J" & summary_row).Interior.ColorIndex = 4
        End If
        
        stock_volume = stock_volume + Cells(i, 7).Value
        Range("L" & summary_row).Value = stock_volume
        percent_change = yearly_change / openprice
        Range("K" & summary_row).Value = FormatPercent(percent_change)
     
     openprice = Cells(i + 1, 3).Value
     summary_row = summary_row + 1
     stock_volume = 0
     percent_change = 0
      'if same keep keep totaling
      Else
         stock_volume = stock_volume + Cells(i, 7).Value
      End If
      
      Next i
      
   'For t = 2 To LastRow
    
        Cells(2, 14) = "Greatest % Increase"
        Cells(3, 14) = "Greatest % Decrease"
        Cells(4, 14) = "Greatest Total Volume"
        
    'look at column J and bring over column I match
    
      ' If Range("J") = Cells(2, 16).Value Then
       'Cells(2, 15).Value = Cells("I").Value
        'Cells(3, 15) =
        'Cells(4, 15) =
        'End If
   'Next
     
        Cells(2, 16).Value = FormatPercent(WorksheetFunction.Max(Range("K:K")))
        Cells(3, 16).Value = FormatPercent(WorksheetFunction.Min(Range("K:K")))
        Cells(4, 16).Value = WorksheetFunction.Max(Range("L:L"))
        
    Next ws
End Sub
