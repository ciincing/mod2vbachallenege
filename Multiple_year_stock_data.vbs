Attribute VB_Name = "Module1"
Sub stocksorter():

    For Each ws In Worksheets

        Dim ticker As String
        Dim yearlychange As Single
        Dim percentchange As Double
        Dim totalvolume As LongLong
        Dim openingprice As Single
        Dim closingprice As Single
        Dim greatestticker As String
        Dim value
        Dim greatestinc As Single
        Dim greatestdec As Single
        Dim greatesttotal
        Dim pricerow As Long

        totalvolume = 0
        pricerow = 2
        sumrow = 2
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        tablelastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ws.Range("I1").value = "Ticker"
        ws.Range("J1").value = "Yearly Change"
        ws.Range("K1").value = "Percent Change"
        ws.Range("L1").value = "Total Stock Volume"
        ws.Range("P1").value = "Ticker"
        ws.Range("Q1").value = "Value"
        ws.Range("O2").value = "Greatest % Increase"
        ws.Range("O3").value = "Greatest % Decrease"
        ws.Range("O4").value = "Greatest Total Volume"

        For i = 2 To lastrow:

            If ws.Cells(i + 1, 1).value = ws.Cells(i, 1).value Then

                ticker = ws.Cells(i, 1).value
                ws.Range("I" & sumrow).value = ticker

                totalvolume = totalvolume + ws.Cells(i, 7).value

                openingprice = ws.Range("C" & pricerow).value
                closingprice = ws.Cells(i, 6).value

               

            Else
                
                totalvolume = totalvolume + ws.Cells(i, 7).value
                ws.Range("L" & sumrow).value = totalvolume
                totalvolume = 0
                
                closingprice = ws.Cells(i, 6).value
                yearlychange = closingprice - openingprice
                percentchange = yearlychange / openingprice
                ws.Range("J" & sumrow).value = yearlychange
                ws.Range("J" & sumrow).NumberFormat = "0.00"
                ws.Range("K" & sumrow).value = percentchange
                ws.Range("K" & sumrow).NumberFormat = "0.00%"
                
                If ws.Range("J" & sumrow).value > 0 Then
                    ws.Range("J" & sumrow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & sumrow).Interior.ColorIndex = 3
                End If
                
                sumrow = sumrow + 1
                pricerow = i + 1
                 
                
            End If
        Next i
        
        For i = 2 To tablelastrow:
            If ws.Range("K" & i + 1).value > greatestinc Then
                greatestinc = ws.Range("K" & i + 1).value
                ticker = ws.Range("I" & i + 1).value
                ws.Range("P2").value = ticker
                ws.Range("Q2").value = greatestinc
                ws.Range("Q2").NumberFormat = "0.00%"
            ElseIf ws.Range("K" & i + 1).value < greatestdec Then
                greatestdec = ws.Range("K" & i + 1).value
                ticker = ws.Range("I" & i + 1).value
                ws.Range("P3").value = ticker
                ws.Range("Q3").value = greatestdec
                ws.Range("Q3").NumberFormat = "0.00%"
            ElseIf ws.Range("L" & i + 1).value > greatesttotal Then
                greatesttotal = ws.Range("L" & i + 1).value
                ticker = ws.Range("I" & i + 1).value
                ws.Range("P4").value = ticker
                ws.Range("Q4").value = greatesttotal
            End If
        Next i
                
    Next ws
End Sub
