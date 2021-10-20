Attribute VB_Name = "DW"

Public  Sub dwRecord()
    Sheets("贈金紀錄").Select
    Sheets("贈金紀錄").Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    On Error Resume Next
    Cells.Select
    Selection.NumberFormatLocal = "G/通用格式"
    On Error GoTo 0
    Dim Drecord, DW, Ndw()
    Drecord = Sheets("贈金紀錄").[A1].CurrentRegion
    'vitonu = Sheets("贈金紀錄").Cells(Rows.Count, 1).End(xlUp).Row
    'Drecord = Sheets("贈金紀錄").Range("A1:V" & vitonu)
    Set DWDic = CreateObject("SCRIPTING.DICTIONARY")
    For i = 1 To UBound(Drecord)
        If Drecord(i, 3) = "加额完成" And Drecord(i, 4) = "审核完成" Then DWDic(Drecord(i, 14)) = Drecord(i, 18)
    Next i
    Sheets("DW").Select
    Dnu1 = Sheets("DW").Cells(1, 2).End(xlDown).Row + 1
    Dnu2 = Sheets("DW").Cells(Rows.Count, 2).End(xlUp).Row
    DW = Sheets("DW").Range("A" & Dnu1 & ":G" & Dnu2)
    ReDim Ndw(1 To UBound(DW), 1 To 1)
    For i = 1 To UBound(DW)
        Ndw(i, 1) = DWDic(DW(i, 1))
    Next i
    Range("H" & Dnu1).Resize(UBound(Ndw)) = Ndw
    Sheets("DW").Range("A" & Dnu1 - 1 & ":H" & Dnu1 - 1).Select
    Selection.AutoFilter
    Call ComboBox2_Change
End Sub

Public  Sub DW()
    Sheets("DW").Range("A1").Select
    On Error Resume Next
    ActiveSheet.PasteSpecial Format:="Unicode 文本", Link:=False, DisplayAsIcon _
        :=False, NoHTMLFormatting:=True
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    On Error GoTo 0
    Columns("D").Select
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByRows
    Columns("E").Select
    Selection.Replace What:="#*", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
    Columns("F").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
End Sub

Private  Sub ComboBox2_Change()
    Dim index, nu1, nu2 As Long, i, nu3 As Integer
    Selection.AutoFilter
    index = ComboBox2.ListIndex
    nu1 = Sheets("DW").Cells(1, 2).End(xlDown).Row
    nu2 = Sheets("DW").Cells(Rows.Count, 2).End(xlUp).Row
    nu3 = Sheets("DW").Cells(Rows.Count, 18).End(xlUp).Row
    Sheets("DW").Range("A" & nu1 & ":H" & nu1).Select
    Selection.AutoFilter
    ComboBox2.List = Sheets("DW").Range("R2:R" & nu3).Value
    i = 1
    Select Case index
    Case 0
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=5, Criteria1:=Array( _
            "Bonus UKOIL", "Bonus USOIL", "Cancel Withdrawal", "Deposit", "Withdrawal", "Cancel Deposit", "Withdrawal Transfer", "Cancel Withdrawal Transfer", "Deposit Transfer", "Cancel Deposit Transfer"), _
            Operator:=xlFilterValues
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:= _
            "=" & Cells(i + 1, 18), Operator:=xlOr, Criteria2:="="
    Case 1
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 2, 18)
    Case 2
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 3, 18)
    Case 3
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 4, 18)
    Case 4
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 5, 18)
    Case 5
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 6, 18)
    Case 6
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 7, 18)
    Case 7
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 8, 18)
    Case 8
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 9, 18)
    Case 9
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 10, 18)
    Case 10
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 11, 18)
    Case 11
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 12, 18)
    Case 12
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 13, 18)
    Case 13
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 14, 18)
    Case 14
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 15, 18)
    Case 15
    ActiveSheet.Range("A" & nu1 & ":H" & nu2).AutoFilter Field:=8, Criteria1:="=" & Cells(i + 16, 18)
    End Select
End Function

Public  Sub dwDelete()
    Sheets("DW").Range("A:H").ClearContents
    Sheets("DW").Range("A:H").ClearContents
    Sheets("贈金紀錄").Cells.ClearContents
End Sub