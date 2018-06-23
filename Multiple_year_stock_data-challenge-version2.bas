Attribute VB_Name = "Module3"
Sub stock_summary_new()

Dim sheet1, sheet2, sheet3, sheet4 As Worksheet

Set sheet1 = Worksheets("2016")
Set sheet2 = Worksheets("2015")
Set sheet3 = Worksheets("2014")
Set sheet4 = Worksheets("summary")

sheet4.Cells(1, 2).Value = "Ticker 2016"
sheet4.Cells(1, 3).Value = "total stock volume 2016"
sheet4.Cells(1, 4).Value = "Ticker 2015"
sheet4.Cells(1, 5).Value = "total stock volume 2015"
sheet4.Cells(1, 6).Value = "Ticker 2016"
sheet4.Cells(1, 7).Value = "total stock volume 2014"

    

Dim ticker_4, ticker5, ticker6 As String
Dim vol_4, vol5, vol6 As Double
Dim m_4, m5, m6 As Integer

vol6 = 0
m6 = 2
vol5 = 0
m5 = 2
vol_4 = 0
m_4 = 2
    
lastrow1 = sheet1.Cells(Rows.Count, 1).End(xlUp).Row
lastrow2 = sheet2.Cells(Rows.Count, 1).End(xlUp).Row
lastrow3 = sheet3.Cells(Rows.Count, 1).End(xlUp).Row

    '2014 stock
    For i = 2 To lastrow3
        If sheet3.Cells(i + 1, 1).Value <> sheet3.Cells(i, 1).Value Then
            ticker_4 = sheet3.Cells(i, 1).Value
            vol_4 = vol_4 + sheet3.Cells(i, 7).Value
            sheet4.Range("F" & m_4).Value = ticker_4
            Range("G" & m_4).Value = vol_4
            
            m_4 = m_4 + 1
            
            vol_4 = 0
        Else
            vol_4 = vol_4 + sheet3.Cells(i, 7).Value
        End If
    Next i

    '2015 stock
    For i = 2 To lastrow2
        If sheet2.Cells(i + 1, 1).Value <> sheet2.Cells(i, 1).Value Then
            ticker5 = sheet2.Cells(i, 1).Value
            vol5 = vol5 + sheet2.Cells(i, 7).Value
            sheet4.Range("D" & m5).Value = ticker5
            sheet4.Range("E" & m5).Value = vol5
            
            m5 = m5 + 1
            
            vol5 = 0
        Else
            vol5 = vol5 + sheet2.Cells(i, 7).Value
        End If
    Next i
    
    '2016 stock
     For i = 2 To lastrow1
        If sheet1.Cells(i + 1, 1).Value <> sheet1.Cells(i, 1).Value Then
            ticker6 = sheet1.Cells(i, 1).Value
            vol6 = vol6 + sheet1.Cells(i, 7).Value
            sheet4.Range("b" & m6).Value = ticker6
            sheet4.Range("C" & m6).Value = vol6
            
            m6 = m6 + 1
            
            vol6 = 0
        Else
            vol6 = vol6 + sheet1.Cells(i, 7).Value
        End If
    Next i
    
End Sub




