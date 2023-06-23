Attribute VB_Name = "Module1"
Sub VBA_Challenge()
    
    'Defining variables that will be used later on in the code
    Dim LastRow, Ticker_Row, Ticker_Vol, Yearly_Per, Vol_Max As Variant
    Dim Open_Val, Close_Val, Yearly_Ch, Per_Max, Per_Min As Double
    Dim Ticker_Name, TickerMax, TickerMin, TickerVol As String
    Dim ws As Worksheet




    For Each ws In Worksheets
    
        'Setting some variables to an intial value
        Ticker_Row = 2
        Ticker_Vol = 0
        Open_Val = 0
        Close_Val = 0
        Yearly_Ch = 0
    
        'Setting up my Headers and my columns in my other tables
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
    
            'variable for selecting the last row of the current worksheet the code is running on
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

                'Set up a Loop to run through Rows 2 to Last Row
                For i = 2 To LastRow
                
                    'Set a Condition that if the values in Column A from Row 2 onward don't match then perform the following code.
                    If ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value Then

                        'Declare the Variable Ticker_Name to store the old Value from the If statement above
                        Ticker_Name = ws.Range("A" & i).Value

                        'Set the Variable to store the Vol amount for each iteration of the loop from the Else function then add the last Ticker Volume for the Ticker to that Volume
                        Ticker_Vol = Ticker_Vol + ws.Range("G" & i).Value

                        Close_Val = ws.Range("F" & i).Value

                        'Function to show Yearly change
                        Yearly_Ch = (Close_Val - Open_Val)
 
                        'Function for the Yearly Percentage Change
                        Yearly_Per = (Yearly_Ch / Open_Val)

                        'Print the Value that is stored in the Variable Ticker_Name and put it in the Column J starting at Row 2.
                        ws.Range("J" & Ticker_Row).Value = Ticker_Name
                    
                        'Print the value that is stored in the Ticker_Vol in the second row of column H
                        ws.Range("M" & Ticker_Row).Value = Ticker_Vol
                       
                        'Print the Yearly_Ch value in Row 2 of Column K
                        ws.Range("K" & Ticker_Row).Value = Yearly_Ch

                        'Print my Yearly Percent in Row L
                        ws.Range("L" & Ticker_Row).Value = Yearly_Per
                    
                        'Change my Yearly_Per format to a percent with two decimal places
                        ws.Range("L" & Ticker_Row).NumberFormat = "0.00%"
                        ws.Range("R2:R3").NumberFormat = "0.00%"
                    
                            'This conditional sets up which cells will be colored.
                            If Yearly_Ch > 0 Then
                            
                                'Colors the interior cells fun green for the cells that are greater than zero
                                ws.Range("K" & Ticker_Row).Interior.Color = RGB(0, 153, 102)
                            
                            Else
                            
                                'Colors the interior cells fun Red for the cells that are less than zero
                                ws.Range("K" & Ticker_Row).Interior.Color = RGB(255, 0, 102)
                            
                            End If
                        
                            'The following condition finds the Maximum/Minimum Percents and also the maximum volume.
                            If ws.Range("L" & Ticker_Row).Value > Per_Max Then
                            
                                Per_Max = ws.Range("L" & Ticker_Row).Value
                                TickerMax = ws.Range("J" & Ticker_Row).Value
                            
                            ElseIf ws.Range("L" & Ticker_Row).Value < Per_Min Then
                        
                                Per_Min = ws.Range("L" & Ticker_Row).Value
                                TickerMin = ws.Range("J" & Ticker_Row).Value
                            
                            End If

                            If ws.Range("M" & Ticker_Row).Value > Vol_Max Then
                            
                                Vol_Max = ws.Range("M" & Ticker_Row).Value
                                TickerVol = ws.Range("J" & Ticker_Row).Value
                            
                            End If
                        
                        'This adds the values from the Per_Max/Per_Min/Vol.Max and prints them to their respective cells.
                        ws.Range("R2").Value = Per_Max
                        ws.Range("R3").Value = Per_Min
                        ws.Range("R4").Value = Vol_Max
                        ws.Range("Q2").Value = TickerMax
                        ws.Range("Q3").Value = TickerMin
                        ws.Range("Q4").Value = TickerVol
        
                        'Add one to the Ticker_Row Variable and then save that value to be reused when the loop cycles again.
                        Ticker_Row = Ticker_Row + 1
                    
                        'Offset the Open_Val to one row below, this will overwrite my existing Open_Val with the next Ticker names Open value
                        Open_Val = ws.Range("C" & i).Offset(1, 0).Value
                    


                        'I need to set my Ticker_Vol to zero or it will keep adding each Ticker Volume together.
                        Ticker_Vol = 0

                    Else
                    
                        'Adding each Same Ticker name Volume to each other
                        Ticker_Vol = Ticker_Vol + ws.Range("G" & i).Value
                    
                    '===================================================================================
                    'This nested function is just to grab the very first C2 column for my first Open_Val.
                    '===================================================================================
                        If i = 2 Then
                        
                            Open_Val = ws.Range("C" & i).Value

                        End If
                     
                End If
            
            Next i
        Next ws
End Sub

