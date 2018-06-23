Attribute VB_Name = "Module1"
Sub Summary_2014()

Dim ticker_4 As String
Dim vol_4 As Double
Dim m_4 As Integer

vol_4 = 0
m_4 = 2

    For i = 2 To 705714
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_4 = Cells(i, 1).Value
            vol_4 = vol_4 + Cells(i, 7).Value
            Range("I" & m_4).Value = ticker_4
            Range("J" & m_4).Value = vol_4
            
            m_4 = m_4 + 1
            
            vol_4 = 0
        Else
            vol_4 = vol_4 + Cells(i, 7).Value
        End If
    Next i

End Sub
Sub Summary_2015()

Dim ticker5 As String
Dim vol5 As Double
Dim m5 As Integer

vol5 = 0
m5 = 2

    For i = 2 To 760192
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker5 = Cells(i, 1).Value
            vol5 = vol5 + Cells(i, 7).Value
            Range("I" & m5).Value = ticker5
            Range("J" & m5).Value = vol5
            
            m5 = m5 + 1
            
            vol5 = 0
        Else
            vol5 = vol5 + Cells(i, 7).Value
        End If
    Next i

End Sub
Sub summary_2016()

Dim ticker6 As String
Dim vol6 As Double
Dim m6 As Integer

vol6 = 0
m6 = 2

    For i = 2 To 797711
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker6 = Cells(i, 1).Value
            vol6 = vol6 + Cells(i, 7).Value
            Range("I" & m6).Value = ticker6
            Range("J" & m6).Value = vol6
            
            m6 = m6 + 1
            
            vol6 = 0
        Else
            vol6 = vol6 + Cells(i, 7).Value
        End If
    Next i

End Sub

