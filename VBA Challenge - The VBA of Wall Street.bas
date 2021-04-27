Attribute VB_Name = "Module1"
Sub Alphabet_Test()

    'Setting data types
    Dim Stock_Name As String
    Dim Total_Stock As Double
    Dim Summary_Table_Row As Integer
    Dim Yearly_Change As Single
    Dim Percent_Change As Single
    Dim last_row As Long
    Dim j As Integer
    
    
    
    'Setting initial variables
    Total_Stock = 0
    Yearly_Change = 0
    Summary_Table_Row = 2
    j = 0
    
    
    'Setting cell titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
    'Getting row number of last row with data
    last_row = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through all stocks
    For i = 2 To last_row
    
        'Check if we are still within the same stock name. If not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Add to the total stock
            Total_Stock = Total_Stock + Cells(i, 7).Value
            
            'For when Total_Stock is zero
            If Total_Stock = 0 Then
            
                Range("I" & Summary_Table_Row).Value = Cells(i, 1).Value
                Range("J" & Summary_Table_Row).Value = 0
                Range("K" & Summary_Table_Row).Value = "%" & 0
                Range("L" & Summary_Table_Row).Value = 0
            
            Else
            
                'find first non-zero starting value
                If Cells(Summary_Table_Row, 3) = 0 Then
                    For Value = Summary_Table_Row To i
                        If Cells(Value, 3).Value <> 0 Then
                            Summary_Table_Row = Value
                            Exit For
                        End If
                    Next Value
                End If
            
                'Calculating yearly change
                Yearly_Change = Cells(i, 6) - Cells(Summary_Table_Row, 3)
                'Calculating percent change
                Percent_Change = Round((Yearly_Change / Cells(i, 3) * 100), 2)
            
                'Add one tot he summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'Print the summary table values
                Range("I" & Summary_Table_Row).Value = Cells(i, 1).Value
                Range("J" & Summary_Table_Row).Value = Round(Yearly_Change, 2)
                Range("K" & Summary_Table_Row).Value = "%" & Percent_Change
                Range("L" & Summary_Table_Row).Value = Total_Stock
                
                'Color pattern: green if positive, red if negative
                'Select Case Yearly_Change
                   ' Case Is > 0
                      '  Range("J" & 2 + j).Interior.ColorIndex = 4
                   ' Case Is < 0
                      '  Range("J" & 2 + j).Interior.ColorIndex = 3
                   ' Case Else
                      '  Range("J" & 2 + j).Interior.ColorIndex = 0
               ' End Select
            End If
            
            'Reset the stock total
            Total_Stock = 0
            Yearly_Change = 0
            j = j + 1
            
        Else
        
            'Add to the total stock
            Total_Stock = Total_Stock + Cells(i, 7).Value
        End If
    Next i
    
End Sub
