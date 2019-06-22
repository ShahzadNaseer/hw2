Attribute VB_Name = "Module3"
Sub Button3_Click()

'-----------------------------
'       C H A L L E N G E
'-----------------------------
    Dim totalRow As Long
    Dim lastTicker As String
    Dim lastTickerVolTotal As Double
    Dim summaryTableRow As Integer
    Dim openingPrice As Single
    Dim closingPrice As Single
    Dim changePercent As Single
    Dim rng As Range, fnd As Range
    Dim ws As Worksheet
    
        
    For Each ws In Worksheets

        ws.Activate
        MsgBox ("Sheet Name: " + ws.Name)
        
        summaryTableRow = 2
        Range("J" & summaryTableRow - 1).Value = "Ticker"
        Range("K" & summaryTableRow - 1).Value = "Yearly Change"
        Range("L" & summaryTableRow - 1).Value = "Percent Change"
        Range("M" & summaryTableRow - 1).Value = "Total Stock Volume"
        
        totalRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        openingPrice = -99
        
        For i = 2 To totalRow
        
            ' Check if we are still within the same Stock Ticker
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
              ' Add to the lastTickerVolTotal
               
              lastTickerVolTotal = lastTickerVolTotal + Cells(i, 7).Value
              
              If openingPrice = -99 Then
              
                openingPrice = Cells(i, 3).Value

                
              End If
              

            Else
              ' Set the Stock Ticker
              lastTicker = Cells(i, 1).Value
              closingPrice = Cells(i, 6).Value
                
              lastTickerVolTotal = lastTickerVolTotal + Cells(i, 7).Value
        
              Range("J" & summaryTableRow).Value = lastTicker
              
              Range("K" & summaryTableRow).Value = (closingPrice - openingPrice)
                 
              If openingPrice <> 0 Then
              
                changePercent = ((closingPrice - openingPrice) / openingPrice)
              
                Range("L" & summaryTableRow).Value = FormatPercent(changePercent, 2, vbFalse, vbFalse, vbTrue)
                
              End If
               
              If (closingPrice - openingPrice) < 0 Then
              
                Range("K" & summaryTableRow).Interior.ColorIndex = 3
                
              Else
                Range("K" & summaryTableRow).Interior.ColorIndex = 4
              End If
              
             
        
              Range("M" & summaryTableRow).Value = lastTickerVolTotal
        
              ' Add one to the summary table row
              summaryTableRow = summaryTableRow + 1
              
              ' Reset the ticker total
              lastTickerVolTotal = 0
              openingPrice = -99
        
            End If
        
          Next i

        
        ' Summary (Hard)
        Set rng = Range("L:L")
        Range("P2").Value = "Greatest % Increase"
        Range("P3").Value = "Greatest % Decrease"
        Range("P4").Value = "Greatest Total Volume"
        
        Range("Q1").Value = "Ticker"
        Range("R1").Value = "Value"
        
        
       ' find max change and its address
        mx = WorksheetFunction.Max(rng)
        Range("R2").Value = FormatPercent(mx, 2, vbFalse, vbFalse, vbTrue)
        rowAddress = WorksheetFunction.Index(rng, WorksheetFunction.Match(mx, rng, 0)).Row
        Range("Q2").Value = Range("J" & rowAddress).Value
        
        ' Min change and its address
        mn = WorksheetFunction.Min(rng)
        Range("R3").Value = FormatPercent(mn, 2, vbFalse, vbFalse, vbTrue)
        rowAddress = WorksheetFunction.Index(rng, WorksheetFunction.Match(mn, rng, 0)).Row
        Range("Q3").Value = Range("J" & rowAddress).Value
        
        ' Max Volume and its address
        Set rng = Range("M:M")
        mxV = WorksheetFunction.Max(rng)
        Range("R4").Value = mxV
        rowAddress = WorksheetFunction.Index(rng, WorksheetFunction.Match(mxV, rng, 0)).Row
        Range("Q4").Value = Range("J" & rowAddress).Value
        
    Next ws

End Sub


