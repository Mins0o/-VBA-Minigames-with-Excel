Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Interior.Color <> RGB(255, 255, 255) Then
        cnt = 1
        Dim Pos(2)
        Dim tPos(2)
        For i = 2 To 51 Step 1
            For j = 2 To 51 Step 1
                If ((Worksheets(1).Cells(i, j).Value <> "trsr") And (Worksheets(1).Cells(i, j).Value <> "")) Then
                    cnt = cnt + 1
                End If
                If Worksheets(1).Cells(i, j).Value = "trsr" And tPos(1) <= 1 Then
                    tPos(1) = i
                    tPos(2) = j
                End If
            Next j
        Next i
        Pos(1) = Selection.Row
        Pos(2) = Selection.Column
        Rng = Round(((Pos(1) - tPos(1)) ^ 2 + (Pos(2) - tPos(2)) ^ 2) ^ 0.5, 0)
        If Selection.Value <> "trsr" Then
            Selection.Value = Rng
            For i = 2 To 51 Step 1
                For j = 2 To 51 Step 1
                    If Round(((Pos(1) - i) ^ 2 + (Pos(2) - j) ^ 2) ^ 0.5, 0) <= Rng Then
                        Worksheets(1).Cells(i, j).Interior.Color = RGB(255 * (686 / (37 * Rng + 649)), 0, 70 + 37 * Rng / 10)
                    End If
                Next j
            Next i
        Else
            MsgBox ("보물을 찾았습니다. 시행횟수" & cnt)
        End If
        Range("A1").Value = "거리" & Rng
        Range("A2").Value = cnt
    End If
End Sub



"------------------------------------------------------------------------------------------------"
Sub initialize()
    Range("B2:AY51").ClearContents
    Range("A1").Value = 0
    Range("B2:AY51").Interior.Color = RGB(70, 70, 255)
    trsr = Int(Rnd * 2500)
    Worksheets(1).Cells(2 + trsr \ 50, 2 + trsr - (trsr \ 50) * 50).Value = "trsr"
End Sub


