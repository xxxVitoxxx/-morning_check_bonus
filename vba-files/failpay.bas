Attribute VB_Name = "failPay"

Public  Sub failPay()
    Sheets("OA").Select
    Sheets("OA").[A1].Select
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    Sheets("注資不成功").Select
    Dim nu1, nu2, i, Tester As Integer, OAarr
    nu1 = Sheets("OA").Cells(1, 2).End(xlDown).Row
    nu2 = Sheets("OA").Cells(Rows.Count, 2).End(xlUp).Row
    mt4nu = Application.CountIf(Sheets("OA").Columns("K"), "MT4平台") 'MT4平台
    mt5nu = Application.CountIf(Sheets("OA").Columns("K"), "MT5平台") 'MT5平台
    OAarr = Sheets("OA").Range("B" & nu1 + 1 & ":AA" & nu2)
    Set mt4dic = CreateObject("SCRIPTING.DICTIONARY")
    Set mt5dic = CreateObject("SCRIPTING.DICTIONARY")
    Testarray = Array("test", "测试", "TEST", "林文娜", "财务", "黄志聚", "黄韦达", "客服", "?務", "陈大明")
    For i = 1 To UBound(OAarr)
     Tester1 = 0
     For j = 0 To UBound(Testarray)
        Tester = InStr(1, OAarr(i, 6), Testarray(j)) + InStr(1, OAarr(i, 7), Testarray(j)) + InStr(1, OAarr(i, 8), Testarray(j))
        Tester1 = Tester1 + Tester
     Next j
     If Tester1 = 0 Then
        If OAarr(i, 10) = "MT4平台" Then mt4dic(OAarr(i, 6) & "," & OAarr(i, 7) & "," & OAarr(i, 8)) = mt4dic(OAarr(i, 6) & "," & OAarr(i, 7) & "," & OAarr(i, 8)) + 1
        If OAarr(i, 10) = "MT5平台" Then mt5dic(OAarr(i, 6) & "," & OAarr(i, 7) & "," & OAarr(i, 8)) = mt5dic(OAarr(i, 6) & "," & OAarr(i, 7) & "," & OAarr(i, 8)) + 1
     End If
     
    Next i
    '----------------
    Dim D As Date
    D = Date
    If mt5nu <> 0 Then
        ReDim mt5arr(1 To mt5nu, 1 To 4)
        K = mt5dic.Keys
        Y = mt5dic.ITEMS
        For i = 0 To UBound(K)
        Nex = Split(K(i), ",")
            nu5 = nu5 + 1
            mt5arr(nu5, 1) = Nex(0)
            mt5arr(nu5, 2) = Nex(1)
            mt5arr(nu5, 3) = Nex(2)
            mt5arr(nu5, 4) = Y(i)
        Next i
        
        mt5dic.RemoveAll
        ReDim mt5rr(1 To UBound(K) + 1, 1 To 4)
        For i = 1 To UBound(K) + 1
            If Not mt5dic.exists(mt5arr(i, 2) & "," & mt5arr(i, 3)) Then '如果關鍵字不存在
                Rownu = Rownu + 1
                mt5dic(mt5arr(i, 2) & "," & mt5arr(i, 3)) = Rownu 
                mt5rr(Rownu, 1) = mt5arr(i, 1)
                mt5rr(Rownu, 2) = mt5arr(i, 2)
                mt5rr(Rownu, 3) = mt5arr(i, 3)
                mt5rr(Rownu, 4) = mt5arr(i, 4)
            Else
                r = mt5dic(mt5arr(i, 2) & "," & mt5arr(i, 3))
                If mt5arr(i, 1) > mt5rr(r, 1) Then mt5rr(r, 1) = mt5arr(i, 1)
                mt5rr(r, 4) = mt5rr(r, 4) + mt5arr(i, 4)
            End If
        Next i
        Erase mt5arr
        idxMin = LBound(mt5rr)
        idxMax = UBound(mt5rr)
        For i = idxMin To idxMax - 1
            For j = i + 1 To idxMax
                If mt5rr(i, 3) = mt5rr(j, 3) And mt5rr(i, 1) < mt5rr(j, 1) Then mt5rr(j, 1) = mt5rr(j, 1) '彆頗埜梛瘍珨欴
                If mt5rr(i, 4) < mt5rr(j, 4) Or mt5rr(i, 4) = "" Then
                    temp = mt5rr(i, 4)
                    mt5rr(i, 4) = mt5rr(j, 4)
                    mt5rr(j, 4) = temp
                    
                    temp = mt5rr(i, 3)
                    mt5rr(i, 3) = mt5rr(j, 3)
                    mt5rr(j, 3) = temp
                    
                    temp = mt5rr(i, 2)
                    mt5rr(i, 2) = mt5rr(j, 2)
                    mt5rr(j, 2) = temp
                    
                    temp = mt5rr(i, 1)
                    mt5rr(i, 1) = mt5rr(j, 1)
                    mt5rr(j, 1) = temp
                End If
            Next j
        Next i
        ReDim mt5arr(1 To UBound(mt5rr), 1 To 4)
        For i = 1 To UBound(mt5rr)
            Select Case mt5rr(i, 4)
            Case Is >= 3
             fnu = fnu + 1
             mt5arr(fnu, 1) = mt5rr(i, 1)
             If mt5rr(i, 1) = "" Then mt5arr(fnu, 1) = 0
             mt5arr(fnu, 2) = mt5rr(i, 2)
             mt5arr(fnu, 3) = mt5rr(i, 3)
             mt5arr(fnu, 4) = mt5rr(i, 4)
            End Select
        Next i
        '--------------------------在上面起泡排序無法完整
        '--------------------------
        Sheets("注資不成功").[A1] = "日期"
        Sheets("注資不成功").[B1] = D
        Sheets("注資不成功").[A2] = "交易帳號"
        Sheets("注資不成功").[B2] = "姓名"
        Sheets("注資不成功").[C2] = "會員帳號"
        Sheets("注資不成功").[D2] = "失敗次數"
        Sheets("注資不成功").Range("A3:D3").Resize(UBound(mt5arr)) = mt5arr '注資不成功
    End If
    
    If mt4nu <> 0 Then
        ReDim mt4arr(1 To mt4nu, 1 To 4)
        K = mt4dic.Keys
        Y = mt4dic.ITEMS
        For i = 0 To UBound(K)
        Nex = Split(K(i), ",")
            nu4 = nu4 + 1
            mt4arr(nu4, 1) = Nex(0)
            mt4arr(nu4, 2) = Nex(1)
            mt4arr(nu4, 3) = Nex(2)
            mt4arr(nu4, 4) = Y(i)
        Next i
        mt4dic.RemoveAll
        '...
        ReDim mt4rr(1 To UBound(K) + 1, 1 To 4)
        For i = 1 To UBound(K) + 1
            If Not mt4dic.exists(mt4arr(i, 2) & "," & mt4arr(i, 3)) Then '如果關鍵字不存在
                Rownu1 = Rownu1 + 1
                mt4dic(mt4arr(i, 2) & "," & mt4arr(i, 3)) = Rownu1 '字典
                mt4rr(Rownu1, 1) = mt4arr(i, 1)
                mt4rr(Rownu1, 2) = mt4arr(i, 2)
                mt4rr(Rownu1, 3) = mt4arr(i, 3)
                mt4rr(Rownu1, 4) = mt4arr(i, 4)
            Else
                r = mt4dic(mt4arr(i, 2) & "," & mt4arr(i, 3))
                If mt4arr(i, 1) > mt4rr(r, 1) Then mt4rr(r, 1) = mt4arr(i, 1)
                mt4rr(r, 4) = mt4rr(r, 4) + mt4arr(i, 4)
            End If
        Next i
        Erase mt4arr
        idxMin = LBound(mt4rr)
        idxMax = UBound(mt4rr)
        For i = idxMin To idxMax - 1
            For j = i + 1 To idxMax
                If mt4rr(i, 3) = mt4rr(j, 3) And mt4rr(i, 1) < mt4rr(j, 1) Then
                    mt4rr(j, 1) = mt4rr(j, 1)
                End If
                If mt4rr(i, 4) < mt4rr(j, 4) Or mt4rr(i, 4) = "" Then
                    temp = mt4rr(i, 4)
                    mt4rr(i, 4) = mt4rr(j, 4)
                    mt4rr(j, 4) = temp
                    
                    temp = mt4rr(i, 3)
                    mt4rr(i, 3) = mt4rr(j, 3)
                    mt4rr(j, 3) = temp
                    
                    temp = mt4rr(i, 2)
                    mt4rr(i, 2) = mt4rr(j, 2)
                    mt4rr(j, 2) = temp
                    
                    temp = mt4rr(i, 1)
                    mt4rr(i, 1) = mt4rr(j, 1)
                    mt4rr(j, 1) = temp
                End If
            Next j
        Next i
        ReDim mt4arr(1 To UBound(mt4rr), 1 To 4)
        For i = 1 To UBound(mt4rr)
            Select Case mt4rr(i, 4)
            Case Is >= 3
             fnu1 = fnu1 + 1
             mt4arr(fnu1, 1) = mt4rr(i, 1)
             If mt4rr(i, 1) = "" Then mt4arr(fnu1, 1) = 0
             mt4arr(fnu1, 2) = mt4rr(i, 2)
             mt4arr(fnu1, 3) = mt4rr(i, 3)
             mt4arr(fnu1, 4) = mt4rr(i, 4)
            End Select
        Next i
        Sheets("注資不成功").[G1] = "日期"
        Sheets("注資不成功").[H1] = D
        Sheets("注資不成功").[G2] = "交易帳?"
        Sheets("注資不成功").[H2] = "姓名"
        Sheets("注資不成功").[I2] = "會員帳?"
        Sheets("注資不成功").[J2] = "失敗次數"
        Sheets("注資不成功").Range("G3:J3").Resize(UBound(mt4arr)) = mt4arr '注資不成功
    End If
End Sub

Public  Sub failClear()
    Sheets("OA").Cells.ClearContents
    Sheets("注資不成功").Range("A:J").ClearContents
    Sheets("注資不成功").Range("A:J").ClearContents
End Sub