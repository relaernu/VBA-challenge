Sub FillSheets()
    For i = 1 To Worksheets.Count
        FillSheet (i)
    Next i
End Sub

Sub FillSheet(sheet_idx As Integer)
    Dim sht As Worksheet
    Dim rows, result_row As Long
    Dim current_ticker As String
    Dim open_value, close_value, total_volumn As Double
    Dim curent_date As Long
    Dim difference, pectdiff As Double
    Set sht = Worksheets(sheet_idx)

    rows = Application.WorksheetFunction.CountA(sht.Range("A:A"))
    Set sort_range = sht.Range("A1:G" & rows)
    sort_range.Sort key1:=sht.Range("A1"), Order1:=xlAscending, key2:=sht.Range("B1"), Order2:=xlAscending, Header:=xlYes
    result_row = 2
    current_ticker = ""
    open_value = sht.Cells(2, 3).Value
    total_volumn = sht.Cells(2, 7).Value
    For i = 2 To rows

        current_ticker = sht.Cells(i, 1).Value
        If current_ticker <> sht.Cells(i + 1, 1).Value Then
            difference = close_value - open_value
            If open_value <> 0 Then
                pectdiff = difference / open_value
            Else
                pectdiff = 0
            End If
            sht.Cells(result_row, 9).Value = current_ticker
            sht.Cells(result_row, 10).Value = difference
            sht.Cells(result_row, 11).Value = pectdiff
            sht.Cells(result_row, 11).NumberFormat = "0.00%"
            If difference >= 0 Then
                sht.Cells(result_row, 10).Interior.Color = VBA.ColorConstants.vbGreen
            Else
                sht.Cells(result_row, 10).Interior.Color = VBA.ColorConstants.vbRed
            End If
            sht.Cells(result_row, 12).Value = total_volumn
            result_row = result_row + 1
            open_value = sht.Cells(i + 1, 3).Value
            total_volumn = sht.Cells(i + 1, 7).Value
        Else
            close_value = sht.Cells(i + 1, 6).Value
            total_volumn = total_volumn + sht.Cells(i + 1, 7).Value
        End If
    Next i
    FillHeader (sheet_idx)
    
    Dim inc, dec, vol As Double
    Dim inc_ticker, dec_ticker, vol_ticker As String
    
    inc_ticker = sht.Cells(2, 9).Value
    dec_ticker = sht.Cells(2, 9).Value
    vol_ticker = sht.Cells(2, 9).Value
    
    inc = sht.Cells(2, 11).Value
    dec = sht.Cells(2, 11).Value
    vol = sht.Cells(2, 12).Value
    
    For i = 3 To result_row - 1
        current_ticker = sht.Cells(i, 9).Value
        pectdiff = sht.Cells(i, 11).Value
        total_volumn = sht.Cells(i, 12).Value
        
        If inc < pectdiff Then
            inc_ticker = current_ticker
            inc = pectdiff
        End If
        If dec > pectdiff Then
            dec_ticker = current_ticker
            dec = pectdiff
        End If
        If vol < total_volumn Then
            vol_ticker = current_ticker
            vol = total_volumn
        End If
    Next i
    
    sht.Cells(1, 15).Value = "Ticker"
    sht.Cells(1, 16).Value = "Value"
    
    sht.Cells(2, 14).Value = "Greatest % Increase"
    sht.Cells(2, 15).Value = inc_ticker
    sht.Cells(2, 16).NumberFormat = "0.00%"
    sht.Cells(2, 16).Value = inc
    
    sht.Cells(3, 14).Value = "Greatest % Decrease"
    sht.Cells(3, 15).Value = dec_ticker
    sht.Cells(3, 16).NumberFormat = "0.00%"
    sht.Cells(3, 16).Value = dec
    
    sht.Cells(4, 14).Value = "Greatest Total Volumn"
    sht.Cells(4, 15).Value = vol_ticker
    sht.Cells(4, 16).Value = vol
    
    sht.Range("I:P").Columns.AutoFit
    
End Sub

Sub FillHeader(sheet_idx As Integer)
    Dim sht As Worksheet
    Set sht = Worksheets(sheet_idx)
    sht.Range("I1").Value = "Ticker"
    sht.Range("J1").Value = "Yearly Change"
    sht.Range("K1").Value = "Percent Change"
    sht.Range("L1").Value = "Total Stock Volumn"
End Sub
