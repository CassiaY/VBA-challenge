Attribute VB_Name = "Module1"
Sub yearly():
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate
    
    Dim lastRow As Double
    Dim i As Double
    Dim total As Double
    Dim ticker As String
    Dim tableRow As Double
    Dim openingP As Double
    Dim closingP As Double
    Dim maxPC As Double
    Dim minPC As Double
    Dim maxVL As Double
    Dim maxPC_tickr As String
    Dim minPC_tickr As String
    Dim maxVL_tickr As String

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    total = 0
    tableRow = 2
    maxPC = 0
    minPC = 0
    maxVL = 0

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("K2:K" & lastRow).NumberFormat = "0.00%"

    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"

    Range("P2", "P3").NumberFormat = "0.00%"
    Range("P4").NumberFormat = "##0.00E+00"

    Columns("J").ColumnWidth = 12
    Columns("K").ColumnWidth = 13
    Columns("L").ColumnWidth = 16
    Columns("N").ColumnWidth = 18

        For i = 2 To lastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                total = total + Cells(i, 7).Value
                closingP = Cells(i, 6).Value
                Cells(tableRow, 9).Value = ticker
                Cells(tableRow, 12).Value = total
                Cells(tableRow, 10).Value = closingP - openingP
                If closingP - openingP < 0 Then
                    Cells(tableRow, 10).Interior.ColorIndex = 3
                    Cells(tableRow, 11).Interior.ColorIndex = 3
                Else
                    Cells(tableRow, 10).Interior.ColorIndex = 4
                    Cells(tableRow, 11).Interior.ColorIndex = 4
                End If
                Cells(tableRow, 11).Value = (closingP - openingP) / openingP
                tableRow = tableRow + 1
                total = 0
            ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                openingP = Cells(i, 3).Value
            Else
                total = total + Cells(i, 7).Value
            End If
        Next i
    
        For i = 2 To lastRow
            If Cells(i, 11).Value > maxPC Then
                maxPC = Cells(i, 11).Value
                maxPC_tickr = Cells(i, 9).Value
            ElseIf Cells(i, 11).Value < minPC Then
                minPC = Cells(i, 11).Value
                minPC_tickr = Cells(i, 9).Value
            ElseIf Cells(i, 12).Value > maxVL Then
                maxVL = Cells(i, 12).Value
                maxVL_tickr = Cells(i, 9).Value
            End If
        Next i
    
    Range("O2").Value = maxPC_tickr
    Range("P2").Value = maxPC
    Range("O3").Value = minPC_tickr
    Range("P3").Value = minPC
    Range("O4").Value = maxVL_tickr
    Range("P4").Value = maxVL

Next ws

End Sub

