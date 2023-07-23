Sub MYSD()
    
    'ws as Worksheet - in order to conduct a loop for each workseet
    Dim ws As Worksheet
    
    'Looping each of the data through all the worksheets
    For Each ws In Worksheets
    
    'Labelling and defining Each Column
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
'Legend for each of the values below:-
'.......................................

'TIC is Ticker - (Range {I1})
'LRA is Last Row for Column A -  (Range {A2:759002 for Eg: {2018 worksheet} or {A2:last row of data}
'LRK is Last Row for Column K -  (Range {K2:?} or {?:last row of data, when all the stock tickers have accumulated}
'TSV is Total Stock Volume - (Range {L1})
'SUMM is considered to be the Summary of the data
'OP is the Opening Price of the data
'CP is the Closing Price of the data
'YC is Yearly Change - (Range {J1})
'PR is the Previous Record
'PC is Percent Change - (Range {K1})
'GI is Greatest % Increase - Range {O2})
'GD is Greatest % Decrease - Range {O3})
'LRV is considered to the Last Row Value
'GTV is Greatest Total Volume - Range {O4})

  
    'Stating all the variables below
    Dim TIC As String
    Dim LRA As Long
    Dim LRK As Long
    Dim TSV As Double
    TSV = 0
    Dim SUMM As Long
    SUMM = 2
    Dim OP As Double
    Dim CP As Double
    Dim YC As Double
    Dim PR As Long
    PR = 2
    Dim PC As Double
    Dim GI As Double
    GI = 0
    Dim GD As Double
    GD = 0
    Dim LRV As Long
    Dim GTV As Double
    GTV = 0

    'Formula for the last row of column A is determined using this formula
    LRA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
           
    'Starting a loop for all the rows in the data
    For i = 2 To LRA

        'Looping Column G or Column # 7 to obtain the total stock volume
        TSV = TSV + ws.Cells(i, 7).Value
    
        'Iterate and see if the ticker name is the same as the ticker info before
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Assign a ticker value in the 1st column
            TIC = ws.Cells(i, 1).Value
                
            'Printing the remaining Ticker names under the SUMM Table in Column I
            ws.Range("I" & SUMM).Value = TIC
                
            'Print the remaining Ticker Stock Volume Info under the SUMM Table in Column L
            ws.Range("L" & SUMM).Value = TSV
               
            'Reset the counter for Total Stock Volume
            TSV = 0

            'Now looping for each value on the opening price, and taking the previous record
            OP = ws.Range("C" & PR)
                
            'Peform the same concept for the closing price
            CP = ws.Range("F" & i)
                
            'Criteria Set for Yearly Change
            YC = CP - OP
            ws.Range("J" & SUMM).Value = YC
                
            ''Ensure the value of YC only comes to 2 decimal places, as the value is in dollar. (as per documentation)
            ws.Range("J" & SUMM).NumberFormat = "$0.00"

            'Create a formula to determine % change, OP is 0 and if the PC is 0
            If OP = 0 Then
                PC = 0
                    
                'Create a new variable called as Yearly Open Guide (YOG), that iterates through the opening price and previous record
                Else
                YOG = ws.Range("C" & PR)
                PC = YC / OP
                        
            End If
                
            'PPopulate the decimal #'s in column K
            ws.Range("K" & SUMM).Value = PC
                
            'Convert the Percent Change Column to a % instead of decimal #'s
            ws.Range("K" & SUMM).NumberFormat = "0.00%"

            'Proceed with creating Conditional Formatting criteria for column J, (Yearly Change); color is green if the value is greater then 0
            If ws.Range("J" & SUMM).Value >= 0 Then
            ws.Range("J" & SUMM).Interior.ColorIndex = 4
                    
                Else
                'Conditional Formatting, color is red, if the value is less then 0
                ws.Range("J" & SUMM).Interior.ColorIndex = 3
                
            End If
            
            ''Proceed with creating Conditional Formatting criteria for column K as well, (Percent Change Column); color is green if the % is greater then 0 and positive
            If ws.Range("K" & SUMM).Value >= 0 Then
            ws.Range("K" & SUMM).Interior.ColorIndex = 4
                    
                Else
                ''Conditional Formatting, color is red, if the % is less then 0 and negative
                ws.Range("K" & SUMM).Interior.ColorIndex = 3
                
            End If
            
            'Now Create and an addtional value to the Summary
            SUMM = SUMM + 1
              
            'The Previous Record is then iterated
            PR = i + 1
                
        End If
                
        'Going over the next iteration of data
        Next i

        'Formula for the last row of column A is determined using this formula
        LRK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Create another loop to determine the following results
        For i = 2 To LRK
            
            'Finding the Greatest % Increase in Value
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

            'Finding the Greatest % Decrease in Value
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            'Finding the Greatest Total Volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        'Change the % Values for Q2 & Q3 to % with 2 decimal places
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub


