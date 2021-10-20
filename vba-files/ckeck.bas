Attribute VB_Name = "check"

Public  Sub check()
    Dim OA, Bonus
    Dim nu1, nu2, nu3, Rownu As Integer
    Set OADic = CreateObject("SCRIPTING.DICTIONARY")
    Set Bondic = CreateObject("SCRIPTING.DICTIONARY")
    OA = Sheets("核對").[A1].CurrentRegion
    Bon = Sheets("Bonus").[A1].CurrentRegion
    ReDim OA1(1 To UBound(OA), 1 To 2)
    ReDim Bon1(1 To UBound(Bon), 1 To 2)
    
    ReDim WrongBonus(1 To UBound(OA), 1 To 1)
    ReDim MultiOA(1 To UBound(OA), 1 To 1)
    ReDim MultiBonus(1 To UBound(Bon), 1 To 1)
    For i = 2 To UBound(OA)
        OADic(OA(i, 2) & OA(i, 6)) = OA(i, 5)
        OATotal = OATotal + OA(i, 5)
    Next i
    
    For i = 2 To UBound(Bon)
        Bondic(Bon(i, 1) & Bon(i, 5)) = Bon(i, 3)
        BonTotal = BonTotal + Bon(i, 3)
    Next i
    
    nu1 = 0
    nu2 = 0
    nu3 = 0
    For i = 2 To UBound(OA)
        OA1(i, 1) = Bondic(OA(i, 2) & OA(i, 6))
        OA1(i, 2) = Round(Bondic(OA(i, 2) & OA(i, 6)), 2) = Round(OA(i, 5), 2)
        If OA1(i, 2) = False And OA1(i, 1) <> "" Then
            nu1 = nu1 + 1 '錯誤
            WrongBonus(nu1, 1) = OA(1, 2) & ":" & OA(i, 2) & ",   " & OA(1, 3) & ":" & OA(i, 3) & ",   " & OA(1, 5) & ":" & OA(i, 5) & ",   " & OA(1, 7) & ":" & Round(Bondic(OA(i, 2) & OA(i, 6)), 2) & ",   " & OA(1, 6) & ":" & OA(i, 6)
        ElseIf OA1(i, 2) = False And OA1(i, 1) = "" Then
            nu2 = nu2 + 1 '多生成
            MultiOA(nu2, 1) = OA(1, 2) & ":" & OA(i, 2) & ",   " & OA(1, 3) & ":" & OA(i, 3) & ",   " & OA(1, 5) & ":" & OA(i, 5) & ",  " & OA(1, 6) & ":" & OA(i, 6)
        End If
    Next i
    
    For i = 2 To UBound(Bon)
        Bon1(i, 1) = OADic(Bon(i, 1) & Bon(i, 5))
        Bon1(i, 2) = Round(OADic(Bon(i, 1) & Bon(i, 5)), 2) = Round(Bon(i, 3), 2)
        If Bon1(i, 2) = False And Bon1(i, 1) = "" Then
            nu3 = nu3 + 1  '少生成
            MultiBonus(nu3, 1) = Bon(1, 1) & ":" & Bon(i, 1) & ",   " & "計算金額" & ":" & Bon(i, 3) & ",   " & Bon(1, 5) & ":" & Bon(i, 5)
        End If
    Next i
    OA1(1, 1) = "模板金額"
    OA1(1, 2) = "核對"
    Sheets("核對").Range("G1:H1").Resize(UBound(OA1)) = OA1
    Sheets("Bonus").Range("F1:G1").Resize(UBound(Bon1)) = Bon1
    
    
    
    Rownu = 11
    Select Case nu1
    Case Is > 0
    Cells(Rownu, 13) = "金額錯誤"
    For i = 1 To nu1
        Cells(Rownu + i, 13) = WrongBonus(i, 1)
    Next i
    Rownu = Rownu + nu1 + 1
    End Select ''

    Select Case nu2
    Case Is > 0
    Cells(Rownu, 13) = "領獎單多生成"
    For i = 1 To nu2

        Cells(Rownu + i, 13) = MultiOA(i, 1)
    Next i
    Rownu = Rownu + nu2 + 1
    End Select
   
    Select Case nu3
    Case Is > 0
    Cells(Rownu, 13) = "領獎單少生成"
    For i = 1 To nu3
        Cells(Rownu + i, 13) = MultiBonus(i, 1)
    Next i
    End Select
    

    Sheets("核對").[M1] = IIf(Sheets("核對").[D2] = "周返点差回赠", "周返点差回赠", Sheets("OA").Range("H" & Sheets("OA").Cells(1, 8).End(xlDown).Row + 1))
    Sheets("核對").[N3] = Round(OATotal, 2)
    Sheets("核對").[N4] = Round(BonTotal, 2)
    Sheets("核對").[N5] = UBound(OA) - 1
    Sheets("核對").[N6] = UBound(Bon) - 1
    Sheets("核對").[O3] = Round(OATotal, 2) = Round(BonTotal, 2)
    Sheets("核對").[O5] = UBound(OA) - 1 = UBound(Bon) - 1
    If Sheets("核對").[O3] = False Then Sheets("核對").[O3].Interior.Color = RGB(255, 0, 0)
    If Sheets("核對").[O3] = True Then Sheets("核對").[O3].Interior.ColorIndex = xlNone
    If Sheets("核對").[O5] = False Then Sheets("核對").[O5].Interior.Color = RGB(255, 0, 0)
    If Sheets("核對").[O5] = True Then Sheets("核對").[O5].Interior.ColorIndex = xlNone
    'End Select
End Sub