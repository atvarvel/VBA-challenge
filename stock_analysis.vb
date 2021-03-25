Sub Stock_Analysis()

    'Loop through all worksheets
    For Each ws in Worksheets

        ' Set initial variable for holding yearly change
        Dim yearly_change As Double
        yearly_change = 0

        ' Set initial variable for holding percent change
        Dim percent_change As Double
        percent_change = 0

        ' Set initial variable for holding the total volume of each ticker
        Dim total_volume As Variant
        total_volume = 0

        ' Keep track of the location for each ticker in the output table
        Dim table_row As Integer
        table_row = 2

        ' Set initial variable for the open date of the current ticker
        Dim current_open As Double

        ' Set initial variable for the close date of the current ticker
        Dim current_close As Double

        ' Find the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through each row in the current sheet
        For i = 2 To last_row

            ' Check if we are still within the same ticker, if not...
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1)) Then

                ' Set the ticker in the output table
                ws.Cells(table_row, 9).Value = ws.Cells(i, 1).Value

                ' If it is the last one, set the close value in the current row as the current close value
                current_close = ws.Cells(i, 6).Value

                ' Calculate the yearly change
                yearly_change = current_close - current_open

                ' Add the yearly change to the output table
                ws.Cells(table_row, 10).Value = yearly_change

                ' If the yearly change is negative, then...
                If (yearly_change < 0) Then

                    ' Change cell color to red
                    ws.Cells(table_row, 10).Interior.ColorIndex = 3

                ' If the change is positive, then...
                Else

                    ' Change cell color to green
                    ws.Cells(table_row, 10).Interior.ColorIndex = 4

                End If

                ' If the first open value for this ticker is 0, then...
                If (current_open = 0) Then

                    ' Add n/a for the percent change in the output table
                    ws.Cells(table_row, 11).Value = "n/a"

                ' If it is not 0, then...
                Else
                
                    ' Calculate the percent change
                    percent_change = yearly_change / current_open
                
                    ' Add the percent change to the output table
                    ws.Cells(table_row, 11).Value = percent_change

                End if
                
                ' Add the volume in the current row to the total volume
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                ' Add the final total volume to the output table
                ws.Cells(table_row, 12).Value = total_volume
                
                ' Go to the next row in the output table for the next ticker
                table_row = table_row + 1
                
                ' Reset yearly change
                yearly_change = 0
                
                ' Reset percent change
                percent_change = 0
                
                ' Reset total volume
                total_volume = 0
                
            ' If we are still within the same ticker, then...
            Else

                ' If the current row is the first day in the year, then...
                If (ws.Cells(i, 2).Value Like "*0101*") Then

                    ' Set the open value of the current row as the current open value
                    current_open = ws.Cells(i, 3).Value

                End If

                ' Add the volume in the current row to the total volume
                total_volume = total_volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i


        ' BONUS


        ' Set variables for finding the greatest percent increase and decrease and greatest total volume
        Dim percent_range As Range
        Dim total_volume_range As Range
        Dim percent_max As Double
        Dim percent_min As Double
        Dim total_volume_max As Variant

        ' Find the last row of the output table
        output_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' Column K is the column we need to search in for the maximum and minimum percent values
        Set percent_range = ws.Range("K2:K" & output_last_row)

        ' Column L is the column we need to search in for the maximum total volume
        Set total_volume_range = ws.Range("L2:L" & output_last_row)

        ' Calculate maximum percent
        percent_max = Application.WorksheetFunction.Max(percent_range)

        'Calculate minimum percent
        percent_min = Application.WorksheetFunction.Min(percent_range)

        ' Calculate maximum total volume
        total_volume_max = Application.WorksheetFunction.Max(total_volume_range)

        ' Loop through each row in the output table
        For i = 2 To output_last_row

            ' If the current row has the same value as the maximum percent, then...
            If (ws.Cells(i, 11).Value = percent_max) Then

                ' Add the ticker value to the new table
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

                ' Add max percent to the table
                ws.Cells(2, 17).Value = percent_max

            ' If the current row has the same value as the minimum percent, then...
            ElseIf (ws.Cells(i, 11).Value = percent_min) Then

                ' Add the ticker value to the new table
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

                ' Add min percent value to the table
                ws.Cells(3, 17).Value = percent_min

            ' If the current row has the same value as the maximum total volume, then...
            ElseIf (ws.Cells(i, 12).Value = total_volume_max) Then

                ' Add the ticker value to the new table
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

                ' Add max total volume to the table
                ws.Cells(4, 17).Value = total_volume_max

            End If

        Next i
    
    Next ws

End Sub