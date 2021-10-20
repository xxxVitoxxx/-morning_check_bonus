Attribute VB_Name = "acetopLottery"

Public  Sub 指定计算()
    Call 手数模板
    Sheets("A抽獎").Select
    Dim Lost, syarr, LOSTarr()
    Dim x As Integer
    x = Sheets("A抽獎").[N15]
    Set LostDic = CreateObject("SCRIPTING.DICTIONARY")
    Lost = Sheets("手數模板").[A1].CurrentRegion
    nia = Application.Match("活動品種", Sheets("A抽獎").Columns("N"), 0) + 1
    nin = Sheets("A抽獎").Cells(Rows.Count, "N").End(xlUp).Row
    syarr = Sheets("A抽獎").Range("N" & nia & ":N" & nin)

    For i = 2 To UBound(Lost) - 1
        For j = 1 To UBound(syarr)
        If Lost(i, 7) = syarr(j, 1) And Round(Lost(i, 4), 2) >= x Then
            LostDic(Lost(i, 1) & "," & Lost(i, 7)) = Lost(i, 4) & "," & Int(Round(Lost(i, 4), 2) / x) & "$"
        End If
        Next j
    Next i

    ReDim LOSTarr(1 To LostDic.Count, 1 To 4)
    Itemm = LostDic.Keys
    For i = 0 To UBound(LOSTarr) - 1
        aa = Split(Itemm(i), ",")
        Ans = LostDic(Itemm(i))
        
        Ans1 = Split(Ans, "$")
        Ans2 = Split(Ans1(0), ",")
        lnu = lnu + 1
        LOSTarr(lnu, 1) = aa(0)
        LOSTarr(lnu, 2) = Ans2(0)
        LOSTarr(lnu, 3) = aa(1)
        LOSTarr(lnu, 4) = Ans2(1)
    Next i
    
    Range("A2:D2").Resize(UBound(LOSTarr)) = LOSTarr
End Sub

Public  Sub delete()
    Sheets("OA").Cells.ClearContents
    Sheets("手數模板").Cells.ClearContents
    Sheets("A抽獎").Range("A2:I" & Range("A2").End(xlDown).Row).Select
    Selection.ClearContents
    Sheets("A抽獎").Range("K2:M2").ClearContents
    Sheets("A抽獎").Range("K4:M4").ClearContents
    [M2].Interior.ColorIndex = xlNone
    [M4].Interior.ColorIndex = xlNone
End Sub

Public  Sub acetopOa()
    Call TVT
    vito = Sheets("OA").Cells(1, 2).End(xlDown).Row
    eve = Sheets("OA").Cells(Rows.Count, "B").End(xlUp).Row
    Dim OA
    OA = Sheets("OA").Range("B" & vito & ":S" & eve)
    Sheets("A抽獎").Select
    Set OADic = CreateObject("SCRIPTING.DICTIONARY")
    Set MoneyDic = CreateObject("SCRIPTING.DICTIONARY")
    For i = 2 To UBound(OA)
        OADic(OA(i, 3) & OA(i, 17)) = OADic(OA(i, 3) & OA(i, 17)) + 1
        If OA(i, 10) = "赠金" Then
        MoneyDic(OA(i, 3) & OA(i, 17)) = MoneyDic(OA(i, 3) & OA(i, 17)) + OA(i, 13)
        End If
    Next i
    
    Dim nu As Integer
    nu = 2
    Do
    Cells(nu, 5) = OADic(Cells(nu, 1) & Cells(nu, 3))
    Cells(nu, 6) = Cells(nu, 4) = OADic(Cells(nu, 1) & Cells(nu, 3))
    Cells(nu, 7) = IIf(MoneyDic(Cells(nu, 1) & Cells(nu, 3)) = "", 0, MoneyDic(Cells(nu, 1) & Cells(nu, 3)))
    OTotal = OTotal + Cells(nu, 5)
    ATotal = ATotal + Cells(nu, 4)
    nu = nu + 1
    Loop Until Cells(nu, 1) = ""
    [K2] = OTotal
    [L2] = ATotal
    [M2] = OTotal = ATotal
    If [M2] = "False" Then [M2].Interior.Color = RGB(255, 0, 0)
    If [M2] = "True" Then [M2].Interior.ColorIndex = xlNone
End Sub

Public  Sub acetopRecord()
    Call RECORDDD
    Sheets("A抽獎").Select
    Dim rnu1, rnu2 As Integer, Bonus
    rnu1 = Sheets("赠金紀錄").Cells(1, 2).End(xlDown).Row
    rnu2 = Sheets("赠金紀錄").Cells(Rows.Count, 2).End(xlUp).Row
    Set BonusDic = CreateObject("SCRIPTING.DICTIONARY")
    Bonus = Sheets("赠金紀錄").Range("B" & rnu1 & ":W" & rnu2)
    For i = 1 To UBound(Bonus)
        BonusDic(Bonus(i, 8) & Bonus(i, 22)) = BonusDic(Bonus(i, 8) & Bonus(i, 22)) + Bonus(i, 11)
    Next i
    Dim nu As Integer
    nu = 2
    Do
    Cells(nu, 8) = IIf(BonusDic(Cells(nu, 1) & Cells(nu, 3)) = "", 0, BonusDic(Cells(nu, 1) & Cells(nu, 3)))
    Cells(nu, 9) = Cells(nu, 7) = BonusDic(Cells(nu, 1) & Cells(nu, 3))
    OTotal1 = OTotal1 + Cells(nu, 7)
    ATotal1 = ATotal1 + Cells(nu, 8)
    nu = nu + 1
    Loop Until Cells(nu, 1) = ""
    [K4] = OTotal1
    [L4] = ATotal1
    [M4] = OTotal1 = ATotal1
    If [M4] = "False" Then
        [M4].Interior.Color = RGB(255, 0, 0)
    Else
        [M4].Interior.ColorIndex = xlNone
    End If
End Sub

Public  Sub paste()
    Sheets("Bonus").Activate
    Sheets("Bonus").Range("L1").Activate
    ActiveSheet.Paste
End Sub