Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Row <> 1 Or Selection.Column <> 12 Then
        If Selection.Row = Range("N1").Value And Selection.Column = Range("N2").Value Then
            Range("A11").Value = "정답"
            Range("K11").Value = True
        Else
            Range("A11").Value = "오답"
            Range("M11").Select
            Range("K11").Value = False
        End If
        If Range("K11").Value Then
            Range("J11").Value = Range("J11").Value + 1
            퍼즐생성
        Else
            Range("J11").Value = 0
        End If
    End If
End Sub


Sub 퍼즐생성()
    lvl = 11 - Range("L1").Value
    a = 60 - 5 * lvl + Int(Rnd * 195)
    Range("A1:J10").Interior.Color = RGB(a, a, a)
    dif1 = Int(Rnd * 7) + 1
    i = lvl + a \ (70 \ lvl)
    x = Int(Rnd * 10 + 1)
    y = Int(Rnd * 10 + 1)
    Range("N1").Value = x
    Range("N2").Value = y
    s = a + i
    If (s > 255) Or (s < 1) Then i = i * (-1)
    Select Case dif1
        Case 1
            Cells(x, y).Interior.Color = RGB(a + i, a, a)
        Case 2
            Cells(x, y).Interior.Color = RGB(a, a + i, a)
        Case 3
            Cells(x, y).Interior.Color = RGB(a, a, a + i)
        Case 4
            Cells(x, y).Interior.Color = RGB(a + i, a + i, a)
        Case 5
            Cells(x, y).Interior.Color = RGB(a + i, a, a + i)
        Case 6
            Cells(x, y).Interior.Color = RGB(a, a + i, a + i)
        Case 7
            Cells(x, y).Interior.Color = RGB(a + i, a + i, a + i)
    End Select
    Range("L1").Select
End Sub



Sub temp
	Range("A12:A1048576").Entirerow.HiddenTrue
	Range("M1:XFD").entirecolumn.hidden=True
End Sub
		
			