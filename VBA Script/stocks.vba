Attribute VB_Name = "stocks"
Sub stocks()
    For Each ws In Worksheets
        
        ' Declaring variables
        Dim ticker As String
        Dim ticker_counter As Integer
        Dim total_volume As Double
        Dim volume As Double
        Dim worksheet_name As String
        Dim percent_change As Double
        Dim greatest_total_volume As Double
        Dim greatest_percent_increase As Double
        Dim greatest_percent_decrease As Double
        
        worksheet_name = ws.Name


        ' Initialising Row value With ticker_counter = 2 and setting up variable j = 2
        ticker_counter = 2
        j = 2

        ' Initialising total volume = 0
        total_volume = 0

        ' Initialising first_open
        ' first_open = Cells(2, 3).Value

        ' Inserting Headers into New cells Via Ranges
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Loop through all rows
        lastrow_A = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow_A

            ' giving ticker its initial value
            ticker = Cells(i, 1).Value
            volume = Cells(i, 7).Value

            If Cells(i + 1, 1).Value <> ticker Then

                ' writing the different values of ticker in the Column 'I'
                ws.Range("I" & ticker_counter).Value = ticker                               ' MsgBox (Ticker) --> Do Not Do this. It has ~1500 unique values. #Learnt from experience :D

                ' Adding Volume one last time To Total Volume
                total_volume = total_volume + volume

                ' writing value of total_volume To Column 'L'
                ws.Range("L" & ticker_counter).Value = total_volume

                ' writing quarterly change value To Column 'J'
                ws.Range("J" & ticker_counter).Value = ws.Cells(ticker_counter, 10).Value
                ws.Cells(ticker_counter, 10).Value = ws.Cells(i, 6) - ws.Cells(j, 3).Value

                ' Percent Change
                percent_change = ((ws.Cells(i, 6) - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                ws.Cells(ticker_counter, 11).Value = Format(percent_change, "Percent")

                ' Conditional formatting
                        If ws.Cells(ticker_counter, 10).Value < 0 Then
                           ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                        Else
                           ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
                        End If

                ' incrementing ticker counter
                ticker_counter = ticker_counter + 1

                ' resetting total_volume
                total_volume = 0
                
                j = i + 1

            Else
                ' Total Volume
                total_volume = total_volume + volume
            
            End If
            
        Next i

        ' Find last non-blank cell in column I
        lastrow_I = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Summary of Stocks
        greatest_total_volume = ws.Cells(2, 12).Value
        greatest_percent_increase = ws.Cells(2, 11).Value
        greatest_percent_decrease = ws.Cells(2, 11).Value

            ' Loop for summary of stocks
            For i = 2 To lastrow_I
                
            If ws.Cells(i, 12).Value > greatest_total_volume Then
                greatest_total_volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
            Else
                greatest_total_volume = greatest_total_volume
                
            End If
            
            
            If ws.Cells(i, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
            Else
                greatest_percent_increase = greatest_percent_increase
            
            End If
            
            
            If ws.Cells(i, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
            Else
                greatest_percent_decrease = greatest_percent_decrease
            
            End If

            ' Formatting the summary values
            ws.Cells(2, 17).Value = Format(greatest_percent_increase, "Percent")
            ws.Cells(3, 17).Value = Format(greatest_percent_decrease, "Percent")
            ws.Cells(4, 17).Value = Format(greatest_total_volume, "Scientific")
            
            Next i

        Worksheets(worksheet_name).Columns("A:Z").AutoFit
    
    Next ws

End Sub




