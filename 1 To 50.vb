Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Row = 1 And Selection.Column = 6 Then scramble
    If Target.Column <> 6 Then
        If Selection.Value <> "" Then
            v = Selection.Value
            wrong = 0
            For i = 1 To 5 Step 1
                For j = 1 To 5 Step 1
                    If v > Worksheets(1).Cells(i, j).Value Then
                        If Worksheets(1).Cells(i, j) <> "" Then wrong = 1
                    End If
                Next j
            Next i
            If wrong = 0 Then
                If v = 1 Then
                    Range("F5").Font.Color = RGB(101, 255, 101)
                    Range("F4").Font.Color = RGB(0, 0, 0)
                    Range("F3").Value = 1
                    Range("F7").Value = Range("E7").Value
                End If
                Range("F3").Value = Range("F3").Value + 1
                If Range("F3") = 51 Then Range("F3:F5").ClearContents
                If v < 26 Then
                    Selection.Value = v + 25
                ElseIf v < 50 Then
                    Selection.ClearContents
                End If
                
                If v = 50 Then
                    Range("G7").Value = Range("E7").Value
                    Range("H7").Value = Range("G7").Value - Range("F7").Value
                    Range("A6").Value = "경과시간: " & Round(Range("H7").Value * 86400, 2)
                    Range("F5").Font.Color = RGB(0, 0, 0)
                    Range("F4").Font.Color = RGB(255, 101, 101)
                End If
            End If
        ElseIf Selection.Column > 6 Then Cells(Selection.Row, 5).Selection
        ElseIf Selection.Row > 5 Then Cells(5, Selection.Column).Selection
        End If
    End If
End Sub
				

Sub scramble()
    Range("A6").Value = "1 to 50"
    Dim R(25)
    
    Range("A1:E5").Interior.Color = RGB(100 + Int(Rnd * 155), 100 + Int(Rnd * 155), 100 + Int(Rnd * 155))
    For i = 1 To 25 Step 1
        Do
        randint = 1 + Int(Rnd * 25)
        same = 0
        For j = 1 To i Step 1
            If randint = R(j) Then same = 1
        Next
        Loop While same = 1
        R(i) = randint
    Next
    For i = 0 To 4 Step 1
        For j = 1 To 5 Step 1
            Worksheets(1).Cells((i + 1), j).Value = R(i * 5 + j)
        Next
    Next
    Range("F1").Select
    Range("F3:F5").ClearContents
End Sub

