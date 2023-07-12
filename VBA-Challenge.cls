Attribute VB_Name = "Module1"
Sub stockmarket()
    Const IN_VOLUME_COL As Integer = 7
    Const IN_OPEN_COL As Integer = 3
    Const IN_CLOSE_COL As Integer = 6
    Const YEARLY_CHANGE_COL As Integer = 10
    Const FIRST_DATA_ROW As Integer = 2
    Const PERCENT_CHANGE_COL As Integer = 11
    Dim input_row As Long
    Dim output_row As Long
    Dim sum As LongLong
    Dim ticker As String
    Dim first_open As Double
    Dim last_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim start_row As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
        'prepare for first stock
        start_row = 2
        sum = 0
        output_row = FIRST_DATA_ROW
        first_open = Cells(FIRST_DATA_ROW, IN_OPEN_COL).Value
        For input_row = FIRST_DATA_ROW To Cells(Rows.Count, 1).End(xlUp).Row
                ticker = Cells(input_row, 1).Value
                sum = sum + Cells(input_row, IN_VOLUME_COL).Value
                'last row of current stock
                If Cells(input_row + 1, 1).Value <> ticker Then
                     last_close = Cells(input_row, IN_CLOSE_COL)
                    Range("K2:K3001").NumberFormat = "0.00%"
                    'calculations
                    yearly_change = last_close - first_open
                    percent_change = yearly_change / first_open
                    'input
                   If (yearly_change) > 0 Then
                        'Range("J" & start_row & ":J" & input_row).Interior.ColorIndex = 4
                        Cells(output_row, "J").Interior.ColorIndex = 4
                        Cells(output_row, "K").Interior.ColorIndex = 4
                   ElseIf (yearly_change) <= 0 Then
                        'Range("J" & start_row & ":J" & input_row).Interior.ColorIndex = 3
                        Cells(output_row, "J").Interior.ColorIndex = 3
                        Cells(output_row, "K").Interior.ColorIndex = 3
                    End If

                    'output
                    Cells(output_row, "I").Value = ticker
                    Cells(output_row, YEARLY_CHANGE_COL).Value = yearly_change
                    Cells(output_row, PERCENT_CHANGE_COL).Value = percent_change
                    Cells(output_row, "L").Value = sum
                    
                    'prepare for next stock
                    sum = 0
                    output_row = output_row + 1
                    first_open = Cells(input_row + 1, IN_OPEN_COL).Value
                    start_row = input_row + 1
               End If
                
            Next input_row
            
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("O4").Value = "Gretaest % Increase"
        Range("O5").Value = "Greatest % Decrease"
        Range("O6").Value = "Greatest Total Volume"
        Range("P3").Value = "Ticker"
        Range("Q3").Value = "Value"
        
    Next ws
    MsgBox ("done")
    
End Sub
