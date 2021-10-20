Attribute VB_Name = "CheckBonus"

Public  Sub oaBonus()
    Sheets("OA").Select
    [A1].Select
    ActiveSheet.PasteSpecial Format:="HTML", NoHTMLFormatting:=True
    vito = Cells(1, 2).End(xlDown).Row
    eve = Cells(Rows.Count, "B").End(xlUp).Row
    Dim OA()
    ReDim OA(1 To eve - vito, 1 To 6)
    i = vito + 1
    Do
    A = InStr(1, Cells(i, 13), "$") + 1
    B = InStr(1, Cells(i, 13), "美元")
    aa = Mid(Cells(i, 13), A, B - A) * 1
    nu = nu + 1
    OA(nu, 1) = Cells(i, 2)
    OA(nu, 2) = Cells(i, 4)
    OA(nu, 3) = Cells(i, 5)
    OA(nu, 4) = Cells(i, 10)
    OA(nu, 5) = aa * Cells(i, 14)
    OA(nu, 6) = Cells(i, 18)
    i = i + 1
    Loop Until Cells(i, 2) = ""
    Sheets("核對").Select
    Range("A2:F2").Resize(UBound(OA)) = OA
    Range("M1") = Sheets("OA").Cells(i, 7)
End Sub

Public  Sub tvt()
    Sheets("OA").Select
    [A1].Select
    ActiveSheet.PasteSpecial Format:="HTML", NoHTMLFormatting:=True
    On Error Resume Next
    Cells.Select
    Selection.NumberFormatLocal = "G/通用格式"
    On Error GoTo 0
End Sub

Public  Sub bonus()
    Sheets("Bonus").Select
    [A1].Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    
    If Sheets("核對").[D2] = "周返点差回赠" Then
        Dim Rebonus
        Rebonus = Sheets("Bonus").[A1].CurrentRegion
        ReDim NewR(1 To UBound(Rebonus), 1 To 5)
        Range("A1:E" & Range("A1").End(xlDown).Row).ClearContents
        For i = 1 To UBound(Rebonus)
            NewR(i, 1) = Rebonus(i, 1)
            NewR(i, 2) = "LIVE01"
            NewR(i, 3) = Rebonus(i, 2)
            NewR(i, 4) = "周返赠金"
            NewR(i, 5) = Right(Rebonus(i, 3), 6)
        Next i
        Range("A2:E2").Resize(UBound(NewR)) = NewR
        [A1] = "交易帳號"
        [B1] = "服务器"
        [C1] = "赠金模板金額"
        [D1] = "赠金类型"
        [E1] = "MetaTrader备注"
    End If
    Sheets("核對").Select
End Sub

Public  Sub record()
    Sheets("赠金紀錄").Select
    Range("A1").Select
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
End Sub

Public  Sub 手数模板()
    Sheets("手數模板").Select
    [A1].Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub