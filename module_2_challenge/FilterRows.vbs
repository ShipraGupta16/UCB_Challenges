Attribute VB_Name = "Module1"
Sub FilterRows():
    ' declare variables with data type
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim row_num As Integer
    Dim ws As Worksheet
    
    'for loop to read each worksheet in whole workbook
    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        ' Set the initial values
        total_volume = 0
        row_num = 2
            
        ' Find the first and last rows of date from second column
        date_first_row = Cells(2, 2).Value
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        date_last_row = Cells(last_row, 2).Value
        
        'MsgBox (date_first_row)
        'MsgBox (date_last_row)
        ' Add the headers to designated columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
            
            ' for loop to the last row to calculate total volume
            For i = 2 To last_row
                total_volume = total_volume + Cells(i, 7).Value
                
                ' if date matches the starting of the year date
                If Cells(i, 2).Value = date_first_row Then
                    ticker = Cells(i, 1).Value
                    open_price = Cells(i, 3).Value
                
                ' if date matches the ending of the year date
                ElseIf Cells(i, 2).Value = date_last_row Then
                    close_price = Cells(i, 6).Value
                    
                    ' Calculate the yearly change and percent change
                    yearly_change = close_price - open_price
                    percent_change = Round((yearly_change / open_price) * 100, 2)
                    
                    ' Assign the calculated value in respective columns
                    Cells(row_num, 9).Value = ticker
                    Cells(row_num, 10).Value = yearly_change
                    Cells(row_num, 11).Value = percent_change & "%"
                    Cells(row_num, 12).Value = total_volume
                    row_num = row_num + 1
                    ' reset the total volume
                    total_volume = 0
                End If
            Next i
            
            ' Find the last row of yearly change column and highlight with colors
            year_change_row = Cells(Rows.Count, 10).End(xlUp).Row
            
            
            ' Find the greatest (% increase, % decrease and total volume)
            ' Add the headers tio the designated columns and rows
            row_num = 2
            Cells(row_num, 15).Value = "Greatest % Increase"
            Cells(row_num + 1, 15).Value = "Greatest % Decrease"
            Cells(row_num + 2, 15).Value = "Greatest Total Volume"
            Cells(row_num - 1, 16).Value = "Ticker"
            Cells(row_num - 1, 17).Value = "Value"
            last_percent_row = Cells(Rows.Count, 11).End(xlUp).Row
            
            Dim max_percent As Double
            Dim min_percent As Double
            Dim max_volume As Double
            max_percent = 0
            max_ticker = ""
            last_percent_row = Cells(Rows.Count, 11).End(xlUp).Row
            min_percent = Cells(last_percent_row, 11).Value
            min_ticker = ""
            max_volume = 0
            max_volume_ticker = ""
            
            ' Find the greatest % increase, decrease and total volume
            For k = 2 To last_percent_row
                ' highlight the cell to green if greater than 0
                If Cells(k, 10).Value > 0 Then
                    Cells(k, 10).Interior.ColorIndex = 4
                ' highlight the cell to green if less than 0
                ElseIf Cells(k, 10).Value < 0 Then
                    Cells(k, 10).Interior.ColorIndex = 3
                ' else no color change
                Else
                    Cells(k, 10).Interior.ColorIndex = 0
                End If
                ' Find the greatest % increase
                If max_percent < Cells(k, 11).Value Then
                    max_percent = Cells(k, 11).Value
                    max_ticker = Cells(k, 9).Value
                End If
            
                ' Find the greatest % decrease
                If min_percent > Cells(k, 11).Value Then
                    min_percent = Cells(k, 11).Value
                    min_ticker = Cells(k, 9).Value
                End If
                
                ' Find the highest stock volume
                If max_volume < Cells(k, 12).Value Then
                    max_volume = Cells(k, 12).Value
                    max_volume_ticker = Cells(k, 9).Value
                End If
                ' Assign the values into appropriate cell value
            Next k
            Cells(2, 16).Value = max_ticker
            Cells(2, 17).Value = max_percent * 100 & "%"
            
            Cells(3, 16).Value = min_ticker
            Cells(3, 17).Value = min_percent * 100 & "%"
            
            Cells(4, 16).Value = max_volume_ticker
            Cells(4, 17).Value = max_volume
            
        Next ws
  End Sub


