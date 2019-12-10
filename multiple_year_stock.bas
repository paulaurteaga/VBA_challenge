Attribute VB_Name = "Module1"
Sub homework()
'Looping thru every worksheey
For Each ws In Worksheets
'Creating the variables we are gonna need
Dim numberStock
Dim counter
Dim min, max, total, percentage, count
numberStock = 0
total = 0
min = 0
max = 0
percentage = 0
counter = 0
Dim greatestInc, greatestDec, greatestVol
greatestInc = 0
greatestDec = 0
greatestVol = 0
count = 0
'Create all the text its gonna appear in all worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Changed"
ws.Range("L1").Value = "Stocks"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest increase"
ws.Range("N3").Value = "Greatest Decrease"
ws.Range("N4").Value = "Total"
'This counter will be useful to know where to ride the info of every single type
'of ticker we find
'whenever we find a new ticker we add 1 to the counter
'the counter will be our row index for the new created columns
counter = 2
lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
'looping thru all the rows til the last one
For i = 2 To lastrow
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    numberStock = numberStock + ws.Cells(i, 7).Value
    'this counter will be useful for later on to calculate the opening price
    'wewill substract this number to the row index(i) to go back to the first ticker
    'of the same type
        count = count + 1
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    numberStock = numberStock + ws.Cells(i, 7).Value
    'closing and opening price
    max = ws.Cells(i, 6).Value
    min = ws.Cells(i - count, 3).Value
    total = max - min
    'gotta check the denominator is not equals 0
        If min <> 0 Then
            percentage = total / min
        Else
        percentage = 0
        End If
    'Pasting the yearly change
    ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(counter, 10).Value = total
    'formating the yearly change cells
        If total < 0 Then
        ws.Cells(counter, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(counter, 10).Interior.ColorIndex = 4
        End If
    'Pasting the percentage and total number of stocks
    ws.Cells(counter, 11).Value = percentage
    ws.Cells(counter, 12).Value = numberStock
    'the counter will add 1 so whenever we find a new ticker will be paste below
    counter = counter + 1
    'Gotta reiniciate this variables for the next ticker
    numberStock = 0
    count = 0
    End If
Next i
        'Looping thru the new columns
        lastrowNew = ws.Cells(Rows.count, 9).End(xlUp).Row
        'Maximun and minimum of the total increases
        greatestInc = WorksheetFunction.max(ws.Range("K:K"))
        greatestDec = WorksheetFunction.min(ws.Range("K:K"))
        'Maximum of total number of stocks
        greatestVol = WorksheetFunction.max(ws.Range("L:L"))
        
        For J = 2 To lastrowNew
        'If the maximum yearly increase and maximun total num of stocks match we paste both values
         If ws.Cells(J, 11).Value = greatestInc And ws.Cells(J, 12).Value = greatestVol Then
         ws.Range("O2").Value = ws.Cells(J, 9).Value
         ws.Range("P2").Value = greatestInc
         ws.Range("O4").Value = ws.Cells(J, 9).Value
         ws.Range("P4").Value = greatestVol
         'If the minimun yearly increase and maximun total num of stocks match we paste both values
         ElseIf ws.Cells(J, 11).Value = greatestDec And ws.Cells(J, 12).Value = greatestVol Then
         ws.Range("O3").Value = ws.Cells(J, 9).Value
         ws.Range("P3").Value = greatestDec
         ws.Range("O4").Value = ws.Cells(J, 9).Value
         ws.Range("P4").Value = greatestVol
         'Just if the maximun yearly change matches we paste the value
         ElseIf ws.Cells(J, 11).Value = greatestInc Then
         ws.Range("O2").Value = ws.Cells(J, 9).Value
         ws.Range("P2").Value = greatestInc
         'Just if the minimun yearly change matches we paste the value
         ElseIf ws.Cells(J, 11).Value = greatestDec Then
         ws.Range("O3").Value = ws.Cells(J, 9).Value
         ws.Range("P3").Value = greatestDec
         'Just if the maximun total number of stock matches we paste the value
         ElseIf ws.Cells(J, 12).Value = greatestVol Then
         ws.Range("O4").Value = ws.Cells(J, 9).Value
         ws.Range("P4").Value = greatestVol
         End If
        Next J
        'Formatting for the end alwayssss
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("P2:P3").NumberFormat = "0.00%"
Next ws
End Sub

 
