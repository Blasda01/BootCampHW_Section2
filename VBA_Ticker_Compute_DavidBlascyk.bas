Attribute VB_Name = "Module1"
Sub RunforAllWorksheets()

   Dim ws As Worksheet
   
    For Each ws In Worksheets
        ws.Activate
        Call ClearFields
        Call RunCode
    Next
    
End Sub
Sub RunCode()

'-----This program assumes the worksheets will be compiled in such a way that the data is grouped by ticker code,
'-----and within each ticker it is sortedby ascending date.

'Create Output labels: Ticker, Yearly Change, Percent Change, Total Stock Volume

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("I1:L1").Font.Bold = True
Range("I:I,J:J,K:K,L:L").EntireColumn.AutoFit

'Define variables

Dim EndofData As Double
Dim rowNumber As Integer
Dim RowCount As Double

EndofData = Range("A2").End(xlDown).Row
rowNumber = 2

Dim OpenPrice As Single
Dim ClosePrice As Single
Dim FirstVolume As Double
Dim LastVolume As Double

Cells(rowNumber, 9).Value = Cells(2, 1).Value '<--First Ticker code in the list
OpenPrice = Cells(2, 3).Value   '<--- OpenPrice for first ticker code
FirstVolume = Cells(2, 7).Row   '<--- Destination for first ticker code's starting volume


'--------Create Loop to return some output values for current ticker code and store values for next ticker code-----

For RowCount = 2 To EndofData

'Compare adjacent rows' ticker value. If they are differnt store and compute the new data

    If Cells(RowCount, 1).Value <> Cells(RowCount + 1, 1) Then

        ' Close price for current ticker code
        ClosePrice = Cells(RowCount, 6).Value
    
        'Yearly Change for current ticker code
        Cells(rowNumber, 10).Value = ClosePrice - OpenPrice
    
        'Conditional Red or Green fill for Yearly Change
        If Cells(rowNumber, 10).Value > 0 Then
            Cells(rowNumber, 10).Interior.ColorIndex = 4
        Else
            Cells(rowNumber, 10).Interior.ColorIndex = 3
        End If
    
        'Percent Change for current ticker code
        If OpenPrice <> 0 Then
            Cells(rowNumber, 11).Value = Cells(rowNumber, 10).Value / OpenPrice
            Cells(rowNumber, 11).NumberFormat = "0.00%"
        Else
            Cells(rowNumber, 11).Value = 0
        End If
      
        'Total Stock Volume for current ticker code (sum of all volumes from FirstVolume to LastVolume)
        LastVolume = Cells(RowCount, 7).Row
    
        Cells(rowNumber, 12).Value = Application.Sum(Range(Cells(FirstVolume, 7), Cells(LastVolume, 7)))
        
        'Save OpenPrice value for next ticker code,
        OpenPrice = Cells(RowCount + 1, 3)
    
        'Save FirstVolume for next ticker code,
        FirstVolume = Cells(RowCount + 1, 7).Row
    
        'Returns next ticker code to the next output row,
        rowNumber = rowNumber + 1
        Cells(rowNumber, 9).Value = Cells(RowCount + 1, 1).Value
    
    End If
    
Next RowCount
    
'Create labels for Greatest % Increase, Greatest % Decrease, and Greatest Volume

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
    
Range("O2:O4,P1,Q1").Font.Bold = True

    
'Define variables

Dim dblMax As Double
Dim dblMin As Double
Dim dblVMax As Double

'Worksheet function MAX returns the largest value in a range

dblMax = Application.WorksheetFunction.Max(Range("K2:K500")) '<---How to define end range in dynamic fashion?

Range("Q2").Value = dblMax
Range("Q2").NumberFormat = "0.00%"

dblMin = Application.WorksheetFunction.Min(Range("K2:K500")) '<---How to define end range in dynamic fashion?

Range("Q3").Value = dblMin
Range("Q3").NumberFormat = "0.00%"

dblVMax = Application.WorksheetFunction.Max(Range("L2:L500")) '<---How to define end range in dynamic fashion?

Range("Q4").Value = dblVMax
Range("O:O,P:P,Q:Q").EntireColumn.AutoFit

'Define variables for Greatest % ticker code

Dim OutputEndofData As Double

OutputEndofData = Range("I2").End(xlDown).Row


'Create for loop to match Max/Min values with ticker code

For OutputRowNumber = 2 To OutputEndofData

    If Cells(OutputRowNumber, 11).Value = dblMax Then
        Range("P2").Value = Cells(OutputRowNumber, 9).Value
    
    ElseIf Cells(OutputRowNumber, 11).Value = dblMin Then
        Range("P3").Value = Cells(OutputRowNumber, 9).Value
        
    End If
    
    If Cells(OutputRowNumber, 12).Value = dblVMax Then
        Range("P4").Value = Cells(OutputRowNumber, 9).Value
    End If
    
 Next OutputRowNumber

End Sub



Sub ClearFields()
'
' ClearFields Macro
'

'
    Columns("I:T").Select
    Selection.Clear
End Sub

