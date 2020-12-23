Attribute VB_Name = "Module1"
Sub stockTicker()

    ' Declare needed variables
    Dim currentTicker As String
    Dim i As Long
    Dim summaryRow As Integer
    Dim openPrice As Double
    Dim closingPrice As Double
    Dim runningSum As Long
    Dim firstDay As Boolean
    Dim tickerCount As Integer
    Dim lastRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Long
    
    
    ' Declare column constants
    Dim COL_A As Integer
    Dim COL_B As Integer
    Dim COL_C As Integer
    Dim COL_D As Integer
    Dim COL_E As Integer
    Dim COL_F As Integer
    Dim COL_G As Integer
    Dim COL_i As Integer
    Dim COL_J As Integer
    Dim COL_K As Integer
    Dim COL_L As Integer
    
    ' Assign variable values
    COL_A = 1
    COL_B = 2
    COL_C = 3
    COL_D = 4
    COL_E = 5
    COL_F = 6
    COL_G = 7
    COL_i = 9
    COL_J = 10
    COL_K = 11
    COL_L = 12

    'Loop through each worksheet
    For Each ws In Worksheets
    
        'Assign first ticker
        currentTicker = ws.Range("A2").Value 'Starting ticker
    
        'Assign Column Names
        ws.Range("I1").Value = "<ticker>"
        ws.Range("J1").Value = "<change over year>"
        ws.Range("K1").Value = "<percent change>"
        ws.Range("L1").Value = "<sum>"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
    
        'Set sheet last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Reset all iterative counters for new worksheet
        summaryRow = 2
        runningSum = 0
        firstDay = True
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        i = 2
    
       ' Begin vertical loop
       For i = 2 To lastRow + 1 'Plus one because we need one last loop to display last row
    
          ' Display values when loop moves to next ticker value and move to new ticker
          If ws.Cells(i, COL_A).Value <> currentTicker Then
                 
               'Display ticker in summary table
               ws.Cells(summaryRow, COL_i).Value = currentTicker
               
               'Display change in price over year in summary table
               ws.Cells(summaryRow, COL_J).Value = closingPrice - openPrice
               
               'Conditional formatting for percent column
               If ws.Cells(summaryRow, COL_J).Value > 0 Then
                   ws.Cells(summaryRow, COL_J).Interior.ColorIndex = 4
               Else
                   ws.Cells(summaryRow, COL_J).Interior.ColorIndex = 3
               End If
               
               'Calculate and display percent
               If openPrice <> 0 Then
                   ws.Cells(summaryRow, COL_K).Value = FormatPercent((closingPrice - openPrice) / openPrice)
                   
                   'Check for greatest increase/decrease and display in sub-summary table
                   If ws.Cells(summaryRow, COL_K).Value > greatestIncrease Then
                       greatestIncrease = ws.Cells(summaryRow, COL_K).Value
                       
                       ws.Range("O2").Value = currentTicker
                       ws.Range("P2").Value = FormatPercent(greatestIncrease)
                       
                   ElseIf ws.Cells(summaryRow, COL_K).Value < greatestDecrease Then
                       greatestDecrease = ws.Cells(summaryRow, COL_K).Value
                       
                       ws.Range("O3").Value = currentTicker
                       ws.Range("P3").Value = FormatPercent(greatestDecrease)
                   End If
               End If
               
               'Display running sum
               ws.Cells(summaryRow, COL_L).Value = CStr(runningSum) + "00" ' Overflow error on running sum, so added back 00 on the end
               
               'Check for largest sum and display if true
               If runningSum > greatestVolume Then
                   greatestVolume = ws.Cells(summaryRow, COL_L).Value / 100
                   
                   ws.Range("O4").Value = currentTicker
                   ws.Range("P4").Value = CStr(greatestVolume) + "00"
               End If
               
               'Set new summary table row
               summaryRow = summaryRow + 1
               
               ' Reset values
               openPrice = 0
               closingPrice = 0
               runningSum = 0
               firstDay = True
               
               ' Next ticker
                currentTicker = ws.Cells(i, COL_A).Value
                
           End If
               
          
          'Runs execution for current ticker if true
           If ws.Cells(i, COL_A).Value = currentTicker Then
                
               'Check if opening day and set price if true
               If firstDay = True Then
                   openPrice = ws.Cells(i, COL_C).Value
                   firstDay = False
                   
               'Check if closing day and set price if true
               ElseIf ws.Cells(i + 1, COL_A).Value <> currentTicker Then
                   closingPrice = ws.Cells(i, COL_F).Value
               
                End If
               
               'Add to running sum
               runningSum = runningSum + ws.Cells(i, COL_G).Value / 100 ' Due to overflow error, will add back in display
           End If
           
       Next i
        
        'Autofit columns for nice display
        ws.Columns("A:P").AutoFit
       
    Next ws
    
End Sub

