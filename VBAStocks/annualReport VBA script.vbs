Sub annualReport()

    ' Declare variable
    Dim i As Long
    Dim last_row As Long
    Dim j As Long
    Dim result_row As Long
    Dim start_point As Long
    Dim ticker As String
    Dim vdate As String
    Dim vYear As String
    Dim vopen As Double
    Dim vclose As Double
    Dim vvolume As Double
    Dim annual_open As Double
    Dim annual_close As Double
    Dim annual_volume As Double
    Dim annual_change As Double
    Dim percent_change As Double
    Dim sht As Worksheet
    Dim num_ws As Integer
    Dim k As Integer
    
    num_ws = ThisWorkbook.Worksheets.Count
    
    For k = 1 To num_ws
        Set sht = ThisWorkbook.Worksheets(k)
        sht.Activate
        ' Sort the data
        With ActiveSheet.Sort
         .SortFields.Add Key:=sht.Range("A1"), Order:=xlAscending
         .SortFields.Add Key:=sht.Range("B1"), Order:=xlAscending
         .SetRange sht.Range("A1", sht.Range("G1").End(xlDown))
         .Header = xlYes
         .Apply
        End With
    
    
        annual_volume = 0
        annual_close = 0
        annual_open = 0
        start_point = 2
        ' Fetch number of the rows on the sheet
        last_row = sht.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Provide headers
        sht.Cells(1, 9).Value = "Ticker"
        sht.Cells(1, 10).Value = "Yearly Change"
        sht.Cells(1, 11).Value = "Percent Change"
        sht.Cells(1, 12).Value = "Total Stock Volume"
        ' Iterate through the rows
        For i = 2 To last_row
            ticker = sht.Cells(i, 1).Value
    
            vdate = sht.Cells(i, 2).Value
            vopen = sht.Cells(i, 3).Value
            vclose = sht.Cells(i, 6).Value
            vvolume = sht.Cells(i, 7).Value
            
            ' Fetch opening price
            If ticker <> sht.Cells(i - 1, 1).Value Then
                annual_open = vopen
            End If
            
            ' Keep summing the daily volume
            annual_volume = annual_volume + vvolume
    
    
            If ticker <> sht.Cells(i + 1, 1).Value Then
                'Output ticker
                sht.Cells(start_point, 9).Value = ticker
                ' Fetch closing price
                annual_close = vclose
                'Print the output before starting the next ticker
                annual_change = annual_close - annual_open
                sht.Cells(start_point, 10).Value = annual_change
                ' Conditional formating for positive or negative change
                If annual_change > 0 Then
                    sht.Cells(start_point, 10).Interior.ColorIndex = 4
                ElseIf annual_change < 0 Then
                    sht.Cells(start_point, 10).Interior.ColorIndex = 3
                Else
                    sht.Cells(start_point, 10).Interior.ColorIndex = xlNone
                End If
                ' Conditional block to handle zero values in denominator
                If annual_open = 0 Then
                    percent_change = 0
                Else
                    percent_change = (annual_change / annual_open)
                End If
    
                sht.Cells(start_point, 11).Value = percent_change
                sht.Cells(start_point, 11).NumberFormat = "0.00%" ' Display as percentage
                sht.Cells(start_point, 12).Value = annual_volume
                annual_volume = 0
                annual_close = 0
                annual_open = 0
                ticker = ""
                annual_change = 0
                percent_change = 0
                start_point = start_point + 1
            End If
    
    
        Next i
    
        Dim bigPlus As Double
        Dim bigPlus_ticker As String
        Dim bigMinus As Double
        Dim bigMinus_ticker As String
        Dim hugeVolume As Double
        Dim hugevolume_ticker As String
        result_row = Cells(Rows.Count, 9).End(xlUp).Row
        bigPlus = 0
        bigMinus = 0
        hugeVolume = 0
        
        For j = 2 To result_row
        'Find Greatest Increase
            If sht.Cells(j, 11) > bigPlus Then
            bigPlus = sht.Cells(j, 11).Value
            bigPlus_ticker = sht.Cells(j, 9).Value
            End If
    
        'Find Greatest Decrease
            If sht.Cells(j, 11) < bigMinus Then
            bigMinus = sht.Cells(j, 11).Value
            bigMinus_ticker = sht.Cells(j, 9).Value
            End If
    
        'Find Greatest Volume
            If sht.Cells(j, 12) > hugeVolume Then
            hugeVolume = sht.Cells(j, 12).Value
            hugevolume_ticker = sht.Cells(j, 9).Value
            End If
        Next j
    
        sht.Cells(1, 15).Value = "Ticker"
        sht.Cells(1, 16).Value = "Value"
        sht.Cells(2, 14).Value = "Greatest % Increase"
        sht.Cells(2, 15).Value = bigPlus_ticker
        sht.Cells(2, 16).Value = bigPlus
        sht.Cells(2, 16).NumberFormat = "0.00%"
        sht.Cells(3, 14).Value = "Greatest % Decrease"
        sht.Cells(3, 15).Value = bigMinus_ticker
        sht.Cells(3, 16).Value = bigMinus
        sht.Cells(3, 16).NumberFormat = "0.00%"
        sht.Cells(4, 14).Value = "Greatest Total Volume"
        sht.Cells(4, 15).Value = hugevolume_ticker
        sht.Cells(4, 16).Value = hugeVolume
        
    Next k
    
End Sub

