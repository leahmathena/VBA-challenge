Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim greatestIncreaseTicker As String, greatestDecreaseTicker As String, greatestVolumeTicker As String
    Dim i As Long

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Skip if worksheet name is not in the list
        If ws.Name Like "Q*" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            outputRow = 2
            
            ' Set up headers
            ws.Cells(1, 9).Resize(1, 4).Value = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            
            ' Initialize variables
            ticker = ""
            totalVolume = 0
            greatestIncrease = -1E+30
            greatestDecrease = 1E+30
            greatestVolume = 0
            
            ' Loop through rows
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value <> ticker Then
                    ' Process the previous ticker data
                    If ticker <> "" Then
                        ' Calculate the quarterly change and percent change
                        quarterlyChange = closePrice - openPrice
                        percentChange = (quarterlyChange / openPrice) * 100
                        
                        ' Output results
                        ws.Cells(outputRow, 9).Value = ticker
                        ws.Cells(outputRow, 10).Value = quarterlyChange
                        ws.Cells(outputRow, 11).Value = percentChange / 100 ' Correct the percentage calculation
                        ws.Cells(outputRow, 12).Value = totalVolume
                        
                        ' Update greatest values
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

                        ' Move to the next row for the next ticker
                        outputRow = outputRow + 1
                    End If

                    ' Update current ticker
                    ticker = ws.Cells(i, 1).Value
                    openPrice = ws.Cells(i, 3).Value
                    closePrice = ws.Cells(i, 6).Value
                    totalVolume = ws.Cells(i, 7).Value
                Else
                    ' Update closing price and total volume
                    closePrice = ws.Cells(i, 6).Value
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                End If
            Next i
            
            ' Process the last ticker in the sheet
            If ticker <> "" Then
                quarterlyChange = closePrice - openPrice
                percentChange = (quarterlyChange / openPrice) * 100
                
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentChange / 100 ' Correct the percentage calculation
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Update greatest values
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
            
            ' Apply conditional formatting
            Dim rng As Range
            Set rng = ws.Range("J2:J" & outputRow)
            With rng
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red
            End With
            
            Set rng = ws.Range("K2:K" & outputRow)
            With rng
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red
            End With
            
            ' Format percentage columns
            ws.Range("K2:K" & outputRow).NumberFormat = "0.00%" ' Percentage format with 2 decimal places
            
            ' Format greatest values output
            ws.Cells(2, 16).NumberFormat = "0.00%" ' Percentage format
            ws.Cells(3, 16).NumberFormat = "0.00%" ' Percentage format
            
            ' Output greatest values
            ws.Cells(2, 15).Value = greatestIncreaseTicker
            ws.Cells(2, 16).Value = greatestIncrease / 100 ' Correct the percentage calculation
            ws.Cells(3, 15).Value = greatestDecreaseTicker
            ws.Cells(3, 16).Value = greatestDecrease / 100 ' Correct the percentage calculation
            ws.Cells(4, 15).Value = greatestVolumeTicker
            ws.Cells(4, 16).Value = greatestVolume
        End If
    Next ws
End Sub

