Sub easy1()

Application.ScreenUpdating = False

RowCount = Range("a1", Range("a1").End(xlDown)).Rows.Count

Dim stock
Set stock = CreateObject("scripting.dictionary")

'''testing theory
'stock.Add Cells(2, 1).Value, 44
'MsgBox (stock(Cells(2, 1).Value))
'stock(Cells(2, 1).Value) = stock(Cells(2, 1).Value) + 44
'MsgBox (stock(Cells(2, 1).Value))
'''

For i = 2 To RowCount
stock(Cells(i, 1).Value) = stock(Cells(i, 1).Value) + Cells(i, 7).Value
'for visual
Cells(i, 8) = stock(Cells(i, 1).Value)
Next i

'''write the table to the sheet
'MsgBox (stock("A"))
'For j = 0 To stock.Count
Dim kc
kc = 1
Cells(1, 10) = "Ticker"
Cells(1, 11) = "Volume"

For Each Key In stock.Keys()
    kc = kc + 1
    Cells(kc, 10) = Key
    Cells(kc, 11) = stock(Key)

Next Key


End Sub