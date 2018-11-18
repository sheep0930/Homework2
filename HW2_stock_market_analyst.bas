Attribute VB_Name = "Module1"
Sub TotalStockVolume()
    For Each ws In ThisWorkbook.Worksheets
        ws.Select
        Call TotalStockVolumeSingleSheet
    Next
End Sub

Private Sub TotalStockVolumeSingleSheet()
    'Define the variables
    Dim ticker As String
    Dim volume As LongLong
    Dim openvalue As Double
    Dim closevalue As Double
    Dim yearlychange As Double
    Dim percentagechange As Double
    Dim cnt As Integer
    Dim maxincreaseticker As String
    Dim maxdecreaseticker As String
    Dim maxvolumeticker As String
    Dim maxtotalvolume As LongLong
    Dim percentincreasemax As Double
    Dim percentdecreasemax As Double
    
    'Name the titles of the results
    Cells(1, 9).Value = "Ticker"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest total volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    'Set the total stock volume value and the open value for the first ticker. Set the final ticker row number to be 1.
    volume = Cells(2, 7).Value
    openvalue = Cells(2, 3).Value
    cnt = 1
    percentincreasemax = 0
    percentdecreasemax = 0
    greattotalvolume = 0
    
    'A for loop to grab the total amount of volume each stock had over the year and fill the value and the corresponding ticker in the excel sheet.
    For i = 3 To Cells(Rows.Count, 1).End(xlDown).Row
    
        'Add the total stock volume value if the ticker stays the same
        If (Cells(i - 1, 1).Value = Cells(i, 1).Value) Then
            volume = volume + Cells(i, 7).Value
        
        'Assign the total stock volume value and the corresponding information if the ticker name changes and reset the values.
        Else
            cnt = cnt + 1
            ticker = Cells(i - 1, 1).Value
            Cells(cnt, 9).Value = ticker
            Cells(cnt, 12).Value = volume
            
            closevalue = Cells(i - 1, 6).Value
            yearlychange = closevalue - openvalue
            If openvalue <> 0 Then
                percentagechange = (closevalue - openvalue) / openvalue
            Else
                percentagechange = 0
            End If
            Cells(cnt, 10).Value = yearlychange
            If yearlychange >= 0 Then
                Cells(cnt, 10).Interior.ColorIndex = 4
            Else
                Cells(cnt, 10).Interior.ColorIndex = 3
            End If
            Cells(cnt, 11).Value = percentagechange
            Cells(cnt, 11).NumberFormatLocal = "0.00%"
            If percentagechange >= percentincreasemax Then
                percentincreasemax = percentagechange
                maxincreaseticker = ticker
            End If
            If percentagechange <= percentdecreasemax Then
                percentdecreasemax = percentagechange
                maxdecreaseticker = ticker
            End If
            If volume >= maxtotalvolume Then
                maxtotalvolume = volume
                maxvolumeticker = ticker
            End If
            openvalue = Cells(i, 3).Value
            Cells(2, 15).Value = maxincreaseticker
            Cells(3, 15).Value = maxdecreaseticker
            Cells(2, 16).Value = percentincreasemax
            Cells(3, 16).Value = percentdecreasemax
            Cells(2, 16).NumberFormatLocal = "0.00%"
            Cells(3, 16).NumberFormatLocal = "0.00%"
            Cells(4, 15).Value = maxvolumeticker
            Cells(4, 16).Value = maxtotalvolume
            volume = Cells(i, 7).Value
        End If
    Next i
End Sub
