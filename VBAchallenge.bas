Attribute VB_Name = "VBAchallenge"
Sub stock_expert()

Dim ticker_name As String
Dim i As Double
Dim LastRow As Double
Dim ticker_column As String
Dim ticker_summary As Long
Dim total_sum As Double
Dim ws As Worksheet
Dim yearly_change As Double
Dim percent_change As Double
Dim ticker_counter As Long
Dim x As Double
Dim max As Long


For Each ws In Worksheets


ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

ws.range("K:K").NumberFormat = "0.00%"



ticker_summary = 2 'volume column I
total_sum = 0 'volume column L
ticker_counter = 0


last_row_1 = ws.Cells(Rows.count, 1).End(xlUp).Row


For i = 2 To last_row_1
    

ticker_name = ws.Cells(i, 1).Value 'ticker column I
total_sum = total_sum + ws.Cells(i, 7) 'volume column L
ticker_counter = ticker_counter + 1



If ws.Cells(i + 1, 1).Value <> ticker_name Then
    
    ws.Cells(ticker_summary, 9).Value = ticker_name 'ticker column I
    ws.Cells(ticker_summary, 12) = total_sum 'volume column L
    yearly_change = (ws.Cells(i, 6) - ws.Cells(i - ticker_counter + 1, 3)) 'year column J
    ws.Cells(ticker_summary, 10) = yearly_change 'year column J


    If ws.Cells(i - ticker_counter + 1, 3) = 0 And ws.Cells(ticker_counter, 6) <> 0 Then
    ws.Cells(i, 10) = " "
    Else
    
    percent_change = yearly_change / ws.Cells(i - ticker_counter + 1, 3) 'percent column K
    ws.Cells(ticker_summary, 11) = percent_change ' percent column K


        
    End If
    
    
    If ws.Cells(ticker_summary, 11) > 0 Then
        ws.Cells(ticker_summary, 11).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(ticker_summary, 11) < 0 Then
        ws.Cells(ticker_summary, 11).Interior.ColorIndex = 3
        
        
    
    End If
    

    ticker_summary = ticker_summary + 1  'ticker column I
    total_sum = 0 'volume column L
    ticker_counter = 0 'year column J
    
    

    
    
    
End If

Next i

Next ws

End Sub

