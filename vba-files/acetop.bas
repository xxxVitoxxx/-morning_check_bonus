Attribute VB_Name = "acetop"

Public  Sub acetopgo()
    Dim Balance, Bal()
    vvv = Cells(Rows.Count, 8).End(xlUp).Row
    syarr = Range("H2:H" & vvv - 1)
    Balance = Sheets("GMAIL").[A1].CurrentRegion
    ReDim TheNew(1 To UBound(syarr), 1 To 3)
    Set dic = CreateObject("SCRIPTING.DICTIONARY")
    Set gmail = CreateObject("SCRIPTING.DICTIONARY")
    ReDim Bal(1 To UBound(Balance) * 20, 1 To 5)
    For i = 3 To UBound(Balance)
        For j = 36 To 55
            gmail(Balance(1, j)) = gmail(Balance(1, j)) + Balance(i, j)
            Total1 = Total1 + Balance(i, j)
            If Balance(i, j) > 0 Then
                nu = nu + 1
                Bal(nu, 1) = Balance(i, 1)
                Bal(nu, 2) = "LIVE01"
                Bal(nu, 3) = Round(Balance(i, j), 0)
                Bal(nu, 4) = "活动赠金"
                Bal(nu, 5) = Balance(1, j)
                dic(Bal(nu, 5)) = dic(Bal(nu, 5)) + Bal(nu, 3)
                Total2 = Total2 + Bal(nu, 3)
            End If
            
        Next j
    Next i
    Sheets("餘額模板").Range("A2:E2").Resize(UBound(Bal)) = Bal
    Sheets("Data").Range("A2:E2").Resize(UBound(Bal)) = Bal
    For i = 1 To UBound(syarr)
        TheNew(i, 1) = Round(dic(syarr(i, 1)), 2)
        TheNew(i, 2) = Round(gmail(syarr(i, 1)), 2)
        TheNew(i, 3) = TheNew(i, 1) = TheNew(i, 2)
    Next i
    Range("I2:K2").Resize(UBound(TheNew)) = TheNew
    Range("I" & vvv) = Round(Total1, 2)
    Range("J" & vvv) = Round(Total2, 2)
    Range("K" & vvv) = Round(Total1, 2) = Round(Total2, 2)
End Sub

Public  Sub acetop()
    Sheets("GMAIL").Select
    Sheets("GMAIL").[A1].Select
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    Sheets("餘額模板").Select
End Sub

Public  Sub delete()
    Sheets("餘額模板").Range("A2:E" & Range("A2").End(xlDown).Row).ClearContents
    Sheets("餘額模板").Range("I2:K" & Range("I2").End(xlDown).Row).ClearContents
    Sheets("Data").Range("A2:E" & Range("A2").End(xlDown).Row).ClearContents
    Sheets("GMAIL").Cells.ClearContents
End Sub

Public  Sub excel()
    UserForm1.Show
    If UserForm1.TextBox1.Value = "" Then
        Exit Sub
    Else
        Sheets("Data").Copy
        ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\ACETOP" & UserForm1.TextBox1.Value & "月餘額模板.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWindow.Close
        UserForm1.TextBox1.Value = ""
    End If
    'UserForm1.TextBox1.Value
End Sub