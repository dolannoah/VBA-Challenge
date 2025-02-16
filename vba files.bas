Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim quarterStartPrice As Double
    Dim quarterEndPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim currentQuarter As Integer
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables for each worksheet
        outputRow = 2
        totalVolume = 0
        currentQuarter = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""
        ticker = "" ' Initialize ticker
        
        ' Find the last row with data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Add headers for output
        ws.Cells(1, 7).Value = "Ticker"
        ws.Cells(1, 8).Value = "Quarterly Change"
        ws.Cells(1, 9).Value = "Percentage Change"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 12).Value = "Greatest % Increase"
        ws.Cells(1, 13).Value = "Greatest % Decrease"
        ws.Cells(1, 14).Value = "Greatest Total Volume"
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check if the date is valid
            If IsDate(ws.Cells(i, 2).Value) Then
                currentQuarter = GetQuarter(ws.Cells(i, 2).Value)
            Else
                ' Skip this row if the date is invalid
                Debug.Print "Skipping Row " & i & " - Invalid Date: " & ws.Cells(i, 2).Value
                GoTo NextRow
            End If
            
            ' Check if the ticker symbol has changed or if it's a new quarter
            If ws.Cells(i, 1).Value <> ticker Or currentQuarter <> GetQuarter(ws.Cells(i, 2).Value) Then
                ' Output the results for the previous ticker and quarter
                If ticker <> "" Then
                    ws.Cells(outputRow, 7).Value = ticker
                    ws.Cells(outputRow, 8).Value = quarterlyChange
                    ws.Cells(outputRow, 9).Value = percentChange
                    ws.Cells(outputRow, 10).Value = totalVolume
                    
                    ' Apply conditional formatting
                    If quarterlyChange > 0 Then
                        ws.Cells(outputRow, 8).Interior.Color = RGB(0, 255, 0) ' Green
                    ElseIf quarterlyChange < 0 Then
                        ws.Cells(outputRow, 8).Interior.Color = RGB(255, 0, 0) ' Red
                    End If
                    
                    ' Check for greatest % increase, % decrease, and total volume
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ticker
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ticker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = ticker
                    End If
                    
                    outputRow = outputRow + 1
                End If
                
                ' Reset variables for the new ticker and quarter
                ticker = ws.Cells(i, 1).Value
                quarterStartPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' Accumulate the total volume
            totalVolume = totalVolume + ws.Cells(i, 5).Value
            
            ' Update the quarter end price
            quarterEndPrice = ws.Cells(i, 4).Value
            
            ' Calculate quarterly change and percentage change
            quarterlyChange = quarterEndPrice - quarterStartPrice
            If quarterStartPrice <> 0 Then
                percentChange = (quarterlyChange / quarterStartPrice) * 100
            Else
                percentChange = 0
            End If
            
NextRow:
        Next i
        
        ' Output the last ticker and quarter
        If ticker <> "" Then
            ws.Cells(outputRow, 7).Value = ticker
            ws.Cells(outputRow, 8).Value = quarterlyChange
            ws.Cells(outputRow, 9).Value = percentChange
            ws.Cells(outputRow, 10).Value = totalVolume
            
            ' Apply conditional formatting
            If quarterlyChange > 0 Then
                ws.Cells(outputRow, 8).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf quarterlyChange < 0 Then
                ws.Cells(outputRow, 8).Interior.Color = RGB(255, 0, 0) ' Red
            End If
            
            ' Check for greatest % increase, % decrease, and total volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
        End If
        
        ' Output greatest % increase, % decrease, and total volume
        ws.Cells(2, 12).Value = greatestIncreaseTicker
        ws.Cells(2, 13).Value = greatestDecreaseTicker
        ws.Cells(2, 14).Value = greatestVolumeTicker
        ws.Cells(3, 12).Value = greatestIncrease
        ws.Cells(3, 13).Value = greatestDecrease
        ws.Cells(3, 14).Value = greatestVolume
    Next ws
End Sub

Function GetQuarter(dateValue As Date) As Integer
    Dim monthValue As Integer
    monthValue = month(dateValue)
    
    Select Case monthValue
        Case 1 To 3
            GetQuarter = 1
        Case 4 To 6
            GetQuarter = 2
        Case 7 To 9
            GetQuarter = 3
        Case 10 To 12
            GetQuarter = 4
    End Select
End Function
