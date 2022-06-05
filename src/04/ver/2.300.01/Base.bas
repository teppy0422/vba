Attribute VB_Name = "Base"
 Public myVer As String
 Public myIP As String

Public Function verCheck(ByVal book As Workbook) As String
    Dim aa As Integer, bb As Integer
    aa = InStr(book.Name, "Sjp")
    If aa = 0 Then MsgBox "名前はSjpから始まる必要があります。修正してください。": Exit Function
    bb = InStr(book.Name, "_")
    If bb = 0 Then MsgBox "Sjp*.***.**_←Verの数字の後ろにはアンダーバーがある必要があります。修正してください。": Exit Function
    verCheck = Mid(book.Name, 4, bb - 4)
    If Not IsNumeric(Mid(book.Name, 4, 1)) Then
        MsgBox "Sjp*.***.**_←ファイル名は必ずこの名前からはじまる必要があります。"
    End If
End Function

Function GetIPAddress() As String
    Dim NetAdapters, objNic, strIPAddress
    Set NetAdapters = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") _
                           .ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & _
                           "Where (IPEnabled = TRUE)")

    For Each objNic In NetAdapters 'ネットワークアダプターは、複数ある場合がある
        For Each strIPAddress In objNic.IPAddress 'IPは、複数割り当てられている場合がある
            GetIPAddress = strIPAddress
            Exit For        ' １回のみ
        Next
        Exit For        ' １回のみ
    Next
End Function

Public Function checkSpace(ByVal addRess As String)
    If InStr(addRess, "\") = 1 Then '\\10.7.120.44とか
        addRess = Left(addRess, InStr(Mid(addRess, 3), "\") + 1)
    Else
        addRess = Left(addRess, InStr(addRess, "\") - 2)
    End If
    Dim FSO As Object, DrvLetter As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    DrvLetter = addRess
    If DrvLetter = "" Then
        Set FSO = Nothing
        Exit Function
    End If
    If FSO.DriveExists(DrvLetter) Then
        Dim maxSize As Long, nowSize As Long
        maxSize = Format(FSO.GetDrive(DrvLetter).TotalSize / 1024 / 1024 / 1024, "0")
        nowSize = Format(FSO.GetDrive(DrvLetter).AvailableSpace / 1024 / 1024 / 1024, "0")
        checkSpace = "容量:" & nowSize & "/" & maxSize & "GB (" & Format(nowSize / maxSize * 100, "0") & "%)"
    Else
        checkSpace = ""
    End If
    Set FSO = Nothing
End Function
'msgfFg = true で見つからないメッセージ有り
'endFlg = true で見つからない時に停止
Function checkSheet(ByVal sheetName As String, ByVal wb As Workbook, msgFlg As Boolean, endFlg As Boolean) As Boolean
    Dim S As Worksheet, myMsg As String
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    checkSheet = True
    sp = Split(sheetName, ";")
    For i = LBound(sp) To UBound(sp)
        On Error Resume Next
        Set S = wb.Sheets(sp(i))
        On Error GoTo 0
        If Not S Is Nothing = False Then
            myMsg = myMsg & sp(i) & vbLf
            checkSheet = False
        End If
    Next i
    If checkSheet = False Then
        If msgFlg Then MsgBox myMsg & "上記のシートが見つかりません": Stop
        If endFlg = True Then End
    End If
End Function
