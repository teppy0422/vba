Attribute VB_Name = "Module000"
'クリップボードクリア
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long

Public Function connect_Server()
    
    If Left(Mid(myAddress(0, 1), 3), InStr(Mid(myAddress(0, 1), 3), "\") - 1) <> "10.7.1.35" Then Exit Function
    
    Dim key() As Byte
    Dim iv() As Byte
    Dim data() As Byte
    Dim objCipher As Cipher

    key = StringUtility.stringToByte("12345678abcdefgh")
    iv = StringUtility.stringToByte("hgfedcba87654321")
    data = "1234566"
    
    Set objCipher = New Cipher

    Call objCipher.encrypt(key, iv, data)

    Dim jjjStr As String
    Dim jjjStrSP
    jjjStr = "233.74.120.122.48.182.112.122.237.168.200.43.31.67.86.34"
    jjjStrSP = Split(jjjStr, ".")
    
    For i = LBound(jjjStrSP) To UBound(jjjStrSP)
        data(i) = jjjStrSP(i)
    Next i
    
    Call objCipher.decrypt(key, iv, data)
        
    Dim myAddressString As String, myPass As String, myAcount As String
    myAddressString = "\\10.7.1.35\plus"
    myAcount = "nim.jp.yazaki.com\plus"
    Dim ans As Boolean
    ans = Ck_NetWork(myAddressString) '接続できるか確認
    If ans = True Then Exit Function
    
    Dim oNetwork As IWshRuntimeLibrary.WshNetwork
    Set oNetwork = New IWshRuntimeLibrary.WshNetwork
    
    Call oNetwork.MapNetworkDrive("", myAddressString, False, myAcount, StringUtility.byteToString(data))

End Function

Public Function remove_server()

    Dim oNetwork As IWshRuntimeLibrary.WshNetwork
    Set oNetwork = New IWshRuntimeLibrary.WshNetwork
    
    Dim myAddress As String
    myAddress = "\\10.7.1.35"
    Dim ans As Boolean
    ans = Ck_NetWork(myAddress) '接続できるか確認
    If ans = True Then
        Call oNetwork.RemoveNetworkDrive(myAddress, True, True)
    End If

End Function

Public Function Ck_NetWork(myAddress)
   Dim WshShell As Object
   Dim Def_Dir As String
  
    Set WshShell = CreateObject("WScript.Shell")
   Def_Dir = WshShell.CurrentDirectory
   On Error Resume Next
   WshShell.CurrentDirectory = myAddress
   If err.number <> 0 Then
      Ck_NetWork = False
      err.Clear
   Else
      Ck_NetWork = True
      WshShell.CurrentDirectory = Def_Dir
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork2()
    Dim WshShell As Object
    Dim Ping_Str As String
    PING_CMD = "ping -n 1 -w 100 10.7.120.117" ' & vbLf & "ping -n 1 -w 100 10.7.120.117"
    PING_CMD = "dir"
    PING_CMD = "C:\18_部材一覧+\ps\actTakePicture.bat"
    
    Set WshShell = CreateObject("WScript.Shell")
    Ping_Str = WshShell.Exec("%ComSpec% /c " & PING_CMD).StdOut.ReadAll
     
     Dim picPath(2) As String
     picPath(0) = InStr(Ping_Str, "http")
     picPath(1) = InStr(Ping_Str, ".JPG")
     picPath(2) = Mid(Ping_Str, picPath(0), picPath(1) - picPath(0) + 4)
     picPath(2) = Replace(picPath(2), " ", "")
     picPath(2) = Replace(picPath(2), Chr(13), "")
     picPath(2) = Replace(picPath(2), Chr(10), "")
     
     Shell "EXPLORER.EXE " & picPath(2)
    'Debug.Print (Ping_St)
    保存先 = "C:\18_部材一覧+\ps\temp.jpg"
    Result = URLDownloadToFile(0, picPath(2), 保存先, 0, 0)
    
    Set WshShell = Nothing
End Function

 Public Function Ck_NetWork3()
   Dim WshShell As Object
   Dim Ping_St As String
   PING_CMD = "Invoke-WebRequest -Uri http://192.168.122.1:8080/sony/camera -Method POST -Body '{" & _
   Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "actTakePicture" & Chr(34) & ", " & _
   Chr(34) & "params" & Chr(34) & ": [], " & _
   Chr(34) & "id" & Chr(34) & ": 1, " & _
   Chr(34) & "version" & Chr(34) & ": " & Chr(34) & "1.0" & Chr(34) & "}'"
  
    Set WshShell = CreateObject("WScript.Shell")
   Ping_St = WshShell.Exec(PING_CMD)
   
   If InStr(1, Ping_St, "Lost = 0") > 0 Then
      MsgBox "ネットワークに接続しています", 64
   Else
      MsgBox "現在ネットワークに接続されていません", 48
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork4()
    Dim duf As String
    Debug.Print buf
    Close #1

   Dim oExec As Object
   Dim Ping_Str As String
   Dim cmdStr As String
  
   cmdStr = "Invoke-WebRequest -Uri http://192.168.122.1:8080/sony/camera -Method POST -Body '{" & _
   Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "actTakePicture" & Chr(34) & ", " & _
   Chr(34) & "params" & Chr(34) & ": [], " & _
   Chr(34) & "id" & Chr(34) & ": 1, " & _
   Chr(34) & "version" & Chr(34) & ": " & Chr(34) & "1.0" & Chr(34) & "}'"
   
   Set oExec = CreateObject("Wscript.shell").Exec("powershell -NoLogo -ExecutionPolicy Bypass -Scope CurrentUser  -Command " & cmdStr)
   Do While oExec.Status = 0
        Sleep 100
   Loop
   
   Debug.Print (oExec.StdOut.ReadAll)
   
   If InStr(1, Ping_St, "Lost = 0") > 0 Then
      MsgBox "ネットワークに接続しています", 64
   Else
      MsgBox "現在ネットワークに接続されていません", 48
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork5()

   Dim oExec As Object
   Dim Ping_Str As String
   Dim cmdStr As String
  
   cmdStr = "Invoke-WebRequest -Uri http://192.168.122.1:8080/sony/camera -Method POST -Body '{" & _
   Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "actTakePicture" & Chr(34) & ", " & _
   Chr(34) & "params" & Chr(34) & ": [], " & _
   Chr(34) & "id" & Chr(34) & ": 1, " & _
   Chr(34) & "version" & Chr(34) & ": " & Chr(34) & "1.0" & Chr(34) & "}'"
    
    Debug.Print (cmdStr)
   Stop

   Set oExec = CreateObject("Wscript.shell").Exec("powershell -ExecutionPolicy RemoteSigned -Command " & cmdStr)
   Do While oExec.Status = 0
        Sleep 100
   Loop
   
   Debug.Print (oExec.StdOut.ReadAll)
   
   If InStr(1, Ping_St, "Lost = 0") > 0 Then
      MsgBox "ネットワークに接続しています", 64
   Else
      MsgBox "現在ネットワークに接続されていません", 48
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork6()
    Call addressSet(ThisWorkbook)
    Dim temp As String
    Dim duf
    Open "C:\18_部材一覧+\ps\actTakePicture.bat" For Input As #1
    Do Until EOF(1)
        Line Input #1, temp
        buf = buf & " & " & temp
    Loop
    Close #1
    buf = Mid(buf, 4)
    
   Dim oExec As Object
   Dim Ping_Str As String
   Dim cmdStr As String
    
   Dim sh As New IWshRuntimeLibrary.WshShell
   Dim ex As WshExec
   
   Stop
   
   Set ex = sh.Exec("cmd.exe /c " & buf)
   
   Stop
   
   Set WshShell = CreateObject("WScript.Shell")
   Ping_St = WshShell.Exec("cmd.exe /c " & buf).StdOut.ReadAll
   
   Dim obj As IWshRuntimeLibrary.WshShell
   Call obj.Run("cmd.exe ""C:\18_部材一覧+\ps\actTakePicture.bat""", 1, WaitOnreturn:=True)
     
   Set oExec = CreateObject("Wscript.shell").Exec(duf)
   Do While oExec.Status = 0
        Sleep 100
   Loop
   
   Debug.Print (oExec.StdOut.ReadAll)
   
   If InStr(1, Ping_St, "Lost = 0") > 0 Then
      MsgBox "ネットワークに接続しています", 64
   Else
      MsgBox "現在ネットワークに接続されていません", 48
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork7()
    Dim WshShell As Object
    Dim Ping_Str As String
    PING_CMD = "ping -n 1 -w 100 10.7.120.117" ' & vbLf & "ping -n 1 -w 100 10.7.120.117"
    PING_CMD = "dir"
    PING_CMD = "C:\18_部材一覧+\ps\actTakePicture.ps1"
    
    Set oExec = CreateObject("Wscript.shell").Exec("powershell -ExecutionPolicy RemoteSigned -Scope Bypass -Command " & PING_CMD)
    Do While oExec.Status = 0
        Sleep 100
    Loop
   
    Debug.Print (oExec.StdOut.ReadAll)
    
    Set WshShell = CreateObject("WScript.Shell")
    Ping_Str = WshShell.Exec("%ComSpec% /c " & PING_CMD).StdOut.ReadAll
     
     Dim picPath(2) As String
     picPath(0) = InStr(Ping_Str, "http")
     picPath(1) = InStr(Ping_Str, ".JPG")
     picPath(2) = Mid(Ping_Str, picPath(0), picPath(1) - picPath(0) + 4)
     picPath(2) = Replace(picPath(2), " ", "")
     picPath(2) = Replace(picPath(2), Chr(13), "")
     picPath(2) = Replace(picPath(2), Chr(10), "")
     
     Shell "EXPLORER.EXE " & picPath(2)
    'Debug.Print (Ping_St)
    
    Set WshShell = Nothing
End Function
Public Function OneDimensionalsort(ByRef myRan As Variant)
    Dim i As Long
    
    For i = LBound(myRan) To UBound(myRan)
        
    Next i
End Function
Public Function BubbleSort2(ByRef argAry() As Variant, ByVal keyPos As Long)
    Dim vSwap
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(argAry, 1) To UBound(argAry, 1)
        For j = UBound(argAry, 1) To i Step -1
            If val(argAry(i, keyPos)) > val(argAry(j, keyPos)) Then
                For k = LBound(argAry, 2) To UBound(argAry, 2)
                    vSwap = argAry(i, k)
                    argAry(i, k) = argAry(j, k)
                    argAry(j, k) = vSwap
                Next
            End If
        Next j
    Next i
 End Function
 
Public Function BubbleSort3(ByRef argAry As Variant, ByVal keyPos As Long, ByVal keyPos2 As Long)
    Dim vSwap
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(argAry, 1) + 1 To UBound(argAry, 1)
        For j = i + 1 To UBound(argAry, 1)
            If val(argAry(i, keyPos)) > val(argAry(j, keyPos)) Then
                For k = LBound(argAry, 2) To UBound(argAry, 2)
                    vSwap = argAry(i, k)
                    argAry(i, k) = argAry(j, k)
                    argAry(j, k) = vSwap
                Next
            ElseIf val(argAry(i, keyPos)) = val(argAry(j, keyPos)) Then
                If val(argAry(i, keyPos2)) < val(argAry(j, keyPos2)) Then
                    For k = LBound(argAry, 2) To UBound(argAry, 2)
                        vSwap = argAry(i, k)
                        argAry(i, k) = argAry(j, k)
                        argAry(j, k) = vSwap
                    Next
                End If
            End If
        Next j
    Next i
 End Function
 
Public Function BubbleSort4(ByRef argAry() As Variant, ByVal keyPos As Long, ByVal keyPos2 As Long, ByVal keyPos3 As Long)
    Dim vSwap
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(argAry, 1) + 1 To UBound(argAry, 1)
        For j = i + 1 To UBound(argAry, 1)
            If val(argAry(i, keyPos)) > val(argAry(j, keyPos)) Then
                For k = LBound(argAry, 2) To UBound(argAry, 2)
                    vSwap = argAry(i, k)
                    argAry(i, k) = argAry(j, k)
                    argAry(j, k) = vSwap
                Next
            ElseIf val(argAry(i, keyPos)) = val(argAry(j, keyPos)) Then
                If val(argAry(i, keyPos2)) > val(argAry(j, keyPos2)) Then
                    For k = LBound(argAry, 2) To UBound(argAry, 2)
                        vSwap = argAry(i, k)
                        argAry(i, k) = argAry(j, k)
                        argAry(j, k) = vSwap
                    Next
                ElseIf val(argAry(i, keyPos2)) = val(argAry(j, keyPos2)) Then
                    If val(argAry(i, keyPos3)) > val(argAry(j, keyPos3)) Then
                        For k = LBound(argAry, 2) To UBound(argAry, 2)
                            vSwap = argAry(i, k)
                            argAry(i, k) = argAry(j, k)
                            argAry(j, k) = vSwap
                        Next k
                    End If
                End If
            End If
        Next j
    Next i
 End Function
 Public Function BubbleSort5(ByRef argAry As Variant, ByVal keyPos As Long, ByVal keyPos2 As Long)
    Dim vSwap
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(argAry, 2) + 1 To UBound(argAry, 2)
        For j = i + 1 To UBound(argAry, 2)
            If val(argAry(keyPos, i)) > val(argAry(keyPos, j)) Then
                For k = LBound(argAry, 1) To UBound(argAry, 1)
                    vSwap = argAry(k, i)
                    argAry(k, i) = argAry(k, j)
                    argAry(k, j) = vSwap
                Next
            ElseIf val(argAry(keyPos, i)) = val(argAry(keyPos, j)) Then
                If val(argAry(keyPos2, i)) < val(argAry(keyPos2, j)) Then
                    For k = LBound(argAry, 1) To UBound(argAry, 1)
                        vSwap = argAry(k, i)
                        argAry(k, i) = argAry(k, j)
                        argAry(k, j) = vSwap
                    Next
                End If
            End If
        Next j
    Next i
 End Function
Public Function compare_Text(ByVal textA As String, ByVal textB As String) As Boolean
    Dim i As Long
    '値が数字型なら数として比較
    If IsNumeric(textA) And IsNumeric(textB) Then
        If Int(textA) > Int(textB) Then
            compare_Text = True
        Else
            compare_Text = False
        End If
    Else
        For i = 1 To Len(textA)
            If Mid(textA, i, 1) > Mid(textB, i, 1) Then
                compare_Text = True
                Exit Function
            ElseIf Mid(textA, i, 1) < Mid(textB, i, 1) Then
                compare_Text = False
                Exit Function
            End If
        Next i
    End If
End Function
Sub QuickSort(ByRef argAry() As Variant, ByVal lngMin As Long, ByVal lngMax As Long, ByVal keyPos As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim vBase As Variant
    Dim vSwap As Variant
    vBase = argAry(Int((lngMin + lngMax) / 2), keyPos)
    i = lngMin
    j = lngMax
    Do
        Do While argAry(i, keyPos) < vBase
            i = i + 1
        Loop
        Do While argAry(j, keyPos) > vBase
            j = j - 1
        Loop
        If i >= j Then Exit Do
        For k = LBound(argAry, 2) To UBound(argAry, 2)
            vSwap = argAry(i, k)
            argAry(i, k) = argAry(j, k)
            argAry(j, k) = vSwap
        Next
        i = i + 1
        j = j - 1
    Loop
    If (lngMin < i - 1) Then
        Call QuickSort(argAry, lngMin, i - 1, keyPos)
    End If
    If (lngMax > j + 1) Then
        Call QuickSort(argAry, j + 1, lngMax, keyPos)
    End If
 End Sub
 
 Function changeRowCol(ByVal myRan As Variant)
    Dim changedRan As Variant
    a = UBound(myRan, 2)
    b = UBound(myRan, 1)
    ReDim changedRan(a, b)
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        For x = LBound(myRan) To UBound(myRan)
            changedRan(i, x) = myRan(x, i)
        Next x
    Next i
    changeRowCol = changedRan
 End Function

Function ReplaceLR(ByRef myRan As Variant)
    Dim myStr(3) As String
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
         myStr(0) = myRan(1, i) 'ひだり
         myStr(1) = myRan(2, i) 'みぎ
         myStr(2) = myRan(3, i)
         myStr(3) = myRan(4, i)
         'ひだりが数字じゃなければ、みぎをひだりにもってくる
         If myStr(0) = "" Then
            myRan(1, i) = myStr(1)
            myRan(2, i) = myStr(0)
            myRan(3, i) = myStr(3)
            myRan(4, i) = myStr(2)
         End If
         'みぎが小さい場合は左右交換
         If IsNumeric(myStr(0)) = True And IsNumeric(myStr(1)) = True Then
            If val(myStr(1)) < val(myStr(0)) Then
                myRan(1, i) = myStr(1)
                myRan(2, i) = myStr(0)
                myRan(3, i) = myStr(3)
                myRan(4, i) = myStr(2)
            End If
         End If
    Next i
End Function

Function ReplaceLR_先ハメ誘導(ByRef myRan As Variant)
    Dim myStr(3) As String
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
         myStr(0) = myRan(2, i) 'ひだり
         myStr(1) = myRan(3, i) 'みぎ
         myStr(2) = myRan(4, i)
         myStr(3) = myRan(5, i)
         'If myStr(0) = "803" And myStr(1) = "35" Then Stop
         'ひだりが空欄なら左右交換
         If myStr(0) = "" Then
'            myRan(2, i) = myStr(1)
'            myRan(3, i) = myStr(0)
'            myRan(4, i) = myStr(3)
'            myRan(5, i) = myStr(2)
         '右がサブナンバーと同じなら左右交換
         ElseIf myRan(UBound(myRan) - 4, i) = myStr(1) Then
            myRan(2, i) = myStr(1)
            myRan(3, i) = myStr(0)
            myRan(4, i) = myStr(3)
            myRan(5, i) = myStr(2)
         'みぎが小さい場合は左右交換
         ElseIf IsNumeric(myStr(0)) = True And IsNumeric(myStr(1)) = True Then
'            If val(myStr(1)) < val(myStr(0)) Then
'                myRAN(2, i) = myStr(1)
'                myRAN(3, i) = myStr(0)
'                myRAN(4, i) = myStr(3)
'                myRAN(5, i) = myStr(2)
'            End If
         End If
    Next i
End Function
Function SumRan(ByRef myRan As Variant)
    Dim count As Long, i As Long, ii As Long, x As Long
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        count = 1
        myRan(0, i) = count
        For ii = LBound(myRan, 2) + 1 To UBound(myRan, 2)
            If i <> ii Then
                If myRan(1, i) & "_" & myRan(2, i) & "_" & myRan(6, i) = myRan(1, ii) & "_" & myRan(2, ii) & "_" & myRan(6, ii) Then
                    count = count + 1
                    myRan(0, i) = count
                    '構成ナンバー
                    myRan(7, i) = myRan(7, i) & ";" & myRan(7, ii)
                    '仕上寸法
                    If myRan(5, i) = "0" Then myRan(5, i) = myRan(5, ii)
                    For x = LBound(myRan) To UBound(myRan)
                        myRan(x, ii) = ""
                    Next x
                End If
            End If
        Next ii
    Next i
End Function

Function evaluationRan(ByVal myRan As Variant)
    Dim hyokaRan As Variant
    ReDim hyokaRan(7, 0)
    Dim 端末str As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If myRan(0, i) <> "" Then
            If i = LBound(myRan, 2) Then
                hyokaRan(0, UBound(hyokaRan, 2)) = "端末No"
                hyokaRan(1, UBound(hyokaRan, 2)) = "端末矢崎品番"
                hyokaRan(2, UBound(hyokaRan, 2)) = "評価ポイント"
                hyokaRan(3, UBound(hyokaRan, 2)) = "評価分類"
                hyokaRan(4, UBound(hyokaRan, 2)) = "部品分類"
                hyokaRan(5, UBound(hyokaRan, 2)) = "親端末No"
                hyokaRan(6, UBound(hyokaRan, 2)) = "接続数"
                hyokaRan(7, UBound(hyokaRan, 2)) = "subNumber"
            Else
                For x = 1 To 2
                    If myRan(x, i) <> "" Then
                        ReDim Preserve hyokaRan(7, UBound(hyokaRan, 2) + 1)
                        端末str = myRan(x, i)
                        'If 端末str = "22" Then Stop
                        hyokaRan(0, UBound(hyokaRan, 2)) = myRan(x, i)
                        hyokaRan(1, UBound(hyokaRan, 2)) = myRan(x + 2, i)
                        hyokaRan(2, UBound(hyokaRan, 2)) = myRan(0, i) ^ 2
                        部品分類str = 部材詳細の読み込み(端末矢崎品番変換(myRan(x + 2, i)), "部品分類_")
                        hyokaRan(4, UBound(hyokaRan, 2)) = 部品分類str
                        Select Case Left(部品分類str, 3)
                            Case "003", "008", "004", "006"
                                優先 = "1"
                            Case "001"
                                優先 = "2"
                            Case "052", "020", "056"
                                優先 = "9"
                            Case Else
                                Debug.Print 部品分類str, myRan(x, i), myRan(x + 2, i)
                                Stop '登録されていない部品分類
                        End Select
                        hyokaRan(3, UBound(hyokaRan, 2)) = 優先
                        hyokaRan(6, UBound(hyokaRan, 2)) = 1
                        If 優先 = "9" Then hyokaRan(5, UBound(hyokaRan, 2)) = "999"
                    Else
                        GoTo line20
                    End If
                    For ii = LBound(myRan, 2) + 1 To UBound(myRan, 2)
                        If myRan(0, ii) <> "" Then
                            For xx = 1 To 2
                                If i <> ii Or x <> xx Then
                                    If 端末str = myRan(xx, ii) Then
                                        hyokaRan(2, UBound(hyokaRan, 2)) = hyokaRan(2, UBound(hyokaRan, 2)) + myRan(0, ii) ^ 2
                                        hyokaRan(6, UBound(hyokaRan, 2)) = hyokaRan(6, UBound(hyokaRan, 2)) + 1
                                        myRan(xx, ii) = ""
                                    End If
                                End If
                            Next xx
                        End If
                    Next ii
                    '優先が1でも行き先が1箇所の場合は2に下げる
                    If 優先 = "1" And hyokaRan(6, UBound(hyokaRan, 2)) = 1 Then
                        hyokaRan(3, UBound(hyokaRan, 2)) = "2"
                    End If
line20:
                Next x
            End If
        End If
    Next i
    evaluationRan = hyokaRan
End Function

Function search相手端末評価(ByVal myRan As Variant, ByVal 相手端末str As String)
    If 相手端末str = "" Then
        search相手端末評価 = 1
        Exit Function
    End If
    Dim 相手端末評価lng As Long: 相手端末評価lng = 0
    For i = LBound(myRan) + 1 To UBound(myRan)
        For j = LBound(myRan) + 1 To UBound(myRan)
            For x = 1 To 2
                If 相手端末str = myRan(j, x) Then
                    If myRan(j, 6) = "" Then
                        If myRan(j, 0) > 相手端末評価lng Then
                            相手端末評価lng = myRan(j, 0)
                        End If
                    End If
                End If
            Next x
        Next j
    Next i
    search相手端末評価 = 相手端末評価lng
End Function

Function search端末評価RAN(ByVal myRan As Variant, ByVal 端末str As String, myPos As Integer)
    For i = LBound(myRan) To UBound(myRan)
        If 端末str = myRan(i, 0) Then
            search端末評価RAN = myRan(i, myPos)
            Exit Function
        End If
    Next i
End Function

Function search端末評価RAN_2pos(ByVal myRan As Variant, ByVal 端末str1 As String, ByVal 端末str2 As String, myPos As Integer)
    For i = LBound(myRan) To UBound(myRan)
        If 端末str1 = myRan(i, 0) And 端末str2 = myRan(i, 1) Then
            search端末評価RAN_2pos = myRan(i, myPos)
            Exit Function
        End If
    Next i
End Function

Function search端末電線数RAN(ByVal myRan As Variant, ByVal 端末str1 As String, ByVal 端末str2 As String, ByVal 接続Gstr As String, myPos As Integer)
    For i = LBound(myRan) To UBound(myRan)
        If 端末str1 = myRan(i, 1) And 端末str2 = myRan(i, 2) And 接続Gstr = myRan(i, 6) Then
            search端末電線数RAN = myRan(i, myPos)
            Exit Function
        End If
    Next i
End Function

Public Function checkClipboard() As Integer
    Dim cbData As New DataObject
    Dim Result As Variant
    Sleep 5
line05:
    On Error Resume Next
    cbData.GetFromClipboard
    If err.number <> 0 Then
        GoTo line05
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Result = Application.ClipboardFormats
    If err.number <> 0 Then Stop
    On Error GoTo 0
    
    'If result(1) = -1 Then Stop '2.200.96
    Do While Result(1) = -1 '0=テキスト,2=画像,3=ビットマップ
        Sleep 5
        On Error Resume Next
        Result = Application.ClipboardFormats
        If err.number <> 0 Then Stop
        On Error GoTo 0
    Loop
    checkClipboard = Result(1)
    Sleep 5
    DoEvents
    Sleep 5
End Function

Public Sub clearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub

Public Function readAccessDB(ByVal path As String) As Variant
        
    Dim adoCON      As New ADODB.Connection
    Dim adoRS       As New ADODB.Recordset
    Dim strSQL      As String
    Dim odbdDB      As Variant
    Dim wSheetName  As Variant
    Dim i           As Integer
 
    'カレントディレクトリのデータベースパスを取得
    odbdDB = path
 
    'データベースに接続する
    adoCON.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" _
                        & "Data Source=" & odbdDB & ""
    adoCON.Open
 
    'DB接続用SQL
    strSQL = "SELECT サブ図棚番.* FROM サブ図棚番 ORDER BY サブ図棚番.棚番;"
 
    'レコードセットを開く
    adoRS.Open strSQL, adoCON, adOpenDynamic
    Dim myRan() As Variant
    Dim fieldCount As Integer
    fieldCount = adoRS.fields.count
    ReDim myRan(fieldCount - 1, 0)
    
    'テーブルの読み込み
    Do Until adoRS.EOF
        ReDim Preserve myRan(fieldCount - 1, UBound(myRan, 2) + 1)
        For i = 0 To fieldCount - 1
            myRan(i, UBound(myRan, 2)) = adoRS(i)
        Next i
        adoRS.MoveNext
    Loop

    'クローズ処理
    adoRS.Close
    Set adoRS = Nothing
    adoCON.Close
    Set adoCON = Nothing
    
    readAccessDB = myRan
    
End Function
Public Function checkConect(ByVal path As String, ByVal myType As Integer) As Boolean '0=ファイル,1=フォルダ
    Dim FSO As New Scripting.FileSystemObject
    
    If myType = 0 Then checkConect = FSO.FileExists(path)
    If myType = 1 Then checkConect = FSO.FolderExists(path)
    
    Set FSO = Nothing
End Function
'objType:0電線,1付属
Sub processingBlink(ByVal ws As Worksheet, ByVal obj As Object, ByVal objType As Long, ByVal groupName As String)

    If objType = 0 Then
        
        Call clearClipboard
        Sleep 10
        obj.Copy
        Call checkClipboard
        ws.Paste
        Selection.Left = obj.Left
        Selection.Top = obj.Top
        '点滅用にオートシェイプを変更
'        Selection.ShapeRange.Fill.Visible = msoFalse
        Selection.ShapeRange.Fill.Transparency = 0.8
'        Selection.ShapeRange.Fill.Solid
'        Selection.ShapeRange.Fill.ForeColor.RGB = tempcolor
        'Selection.ShapeRange.Line.Visible = False
        tempcolor = Selection.ShapeRange.Fill.ForeColor
        
        Selection.ShapeRange.Line.ForeColor.RGB = tempcolor
        Selection.ShapeRange.Line.Weight = 3
        
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
        
        
'        Selection.ShapeRange.Line.Glow.Radius = 3
'        Selection.ShapeRange.Glow.color.RGB = tempcolor
'        Selection.ShapeRange.Glow.Transparency = 0.5
    
    Else
        
        Call clearClipboard
        Sleep 10
        obj.Copy
        Call checkClipboard
        ws.Paste
        
        Selection.ShapeRange.Left = obj.Left
        Selection.ShapeRange.Top = obj.Top
        tempcolor = Selection.ShapeRange.Fill.ForeColor
        
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 0, 0)
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
'        Set newobj = obj.Duplicate
'        newobj.Left = obj.Left
'        newobj.Top = obj.Top
'        '点滅用にオートシェイプを変更
'        'Selection.ShapeRange.Line.Visible = False
'        tempcolor = newobj.Fill.ForeColor
'        'Selection.ShapeRange.Fill.ForeColor.RGB = tempcolor
'        newobj.Line.ForeColor.RGB = RGB(255, 0, 0)
'        newobj.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
'        newobj.Select False
        画像URL = ""
'        If obj.Name Like "*CLAMP*" Then 画像URL = myAddress(0, 1) & "\付属部品イラスト\CLAMP_001.png"
'        If obj.Name Like "*HOLDER*" Then 画像URL = myAddress(0, 1) & "\付属部品イラスト\HOLDER_001.png"
'        If obj.Name Like "*TUBE*" Then 画像URL = myAddress(0, 1) & "\付属部品イラスト\TUBE_001.png"
'        If obj.Name Like "*COVER*" Then 画像URL = myAddress(0, 1) & "\付属部品イラスト\COVER_001.png"
'        If obj.Name Like "*GROMMET*" Then 画像URL = myAddress(0, 1) & "\付属部品イラスト\GROMMET_001.png"
'        If obj.Name Like "*OTHER*" Then 画像URL = myAddress(0, 1) & "\付属部品イラスト\OTHER_001.png"
        
'        ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, obj.Left, obj.Top, obj.Width, obj.Height).Select
'        Selection.ShapeRange.Adjustments.Item(1) = 0.15
        If 画像URL = "" Then
            'Selection.ShapeRange.Fill.ForeColor.RGB = RGB(230, 230, 0)
'            Selection.ShapeRange.Fill.Visible = msoFalse
'            Selection.ShapeRange.Fill.Transparency = 0.5
        Else
            Selection.ShapeRange.Fill.UserPicture 画像URL
            Selection.ShapeRange.Line.Visible = msoFalse
            Selection.ShapeRange.Left = 0
            Selection.ShapeRange.Fill.Transparency = 0
            'Set ob = ActiveSheet.Shapes.AddPicture(画像URL, False, True, obj.Left, obj.Top, obj.Width, obj.Height)
        End If
    End If
    
    On Error Resume Next
    aa = Empty
    aa = ws.Shapes(groupName).GroupItems.count
    bb = Empty
    bb = ws.Shapes("temp").Width
    On Error GoTo 0
'    If Not IsEmpty(aa) Then
'        If aa = 1 Then
'            ws.Shapes(groupName).Name = "temp"
'            ws.Shapes("temp").Select False
'        Else
'        End If
'    End If
    
    If Not IsEmpty(aa) Then
        ws.Shapes(groupName).Select False
    ElseIf Not IsEmpty(bb) Then
        ws.Shapes("temp").Select False
    End If
        
    aa = Empty
    On Error Resume Next
    aa = Selection.count
    On Error GoTo 0
    If aa = Empty Then
        Selection.Name = "temp"
    Else
        Selection.Group.Name = groupName
    End If
End Sub

Function sortKnumber(ByVal Knumber As String)
    Dim sp As Variant, i As Long, ii As Long, Str_i As String, Str_ii As String, Swap As String
    sp = Split(Knumber, ",")
    Dim lng As Long
    For i = LBound(sp) To UBound(sp)
        For ii = UBound(sp) To i Step -1
            If i <> ii Then
                If val(sp(i)) > val(sp(ii)) Then
                    Swap = sp(i)
                    sp(i) = sp(ii)
                    sp(ii) = Swap
                End If
            End If
        Next ii
    Next i
    sortKnumber = Join(sp, ",")
End Function

Public Function exChangeHTMLcolor(色呼, clocode1, clocode2, clofont)
    Dim 色呼a As String, 色呼b As String
    Dim 変換前 As String
    With wb(0).Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        色呼 = Replace(色呼, " ", "")
        If InStr(色呼, "/") = 0 Then
            色呼a = 色呼
            色呼b = ""
        Else
            色呼a = Left(色呼, InStr(色呼, "/") - 1)
            色呼b = Mid(色呼, InStr(色呼, "/") + 1)
        End If
        
        If 色呼 = "" Then
            clocode1 = "FFF"
            clocode2 = "FFF"
            clofont = "000"
            mysel.Select
            Exit Function
        End If
        '色の登録確認
        検索色 = 色呼a
        Set 検索x = .Columns(key.Column).Find(検索色, , , 1)
        If 検索x Is Nothing Then GoTo errFlg
        
        変換前 = 検索x.Offset(0, 2)
        clocode1s = Split(変換前, ",")
        clocode1 = Format(Hex(clocode1s(0)), "00") & Format(Hex(clocode1s(1)), "00") & Format(Hex(clocode1s(2)), "00")
        変換前 = 検索x.Offset(0, 3)
        clofonts = Split(変換前, ",")
        clofont = Format(Hex(clofonts(0)), "00") & Format(Hex(clofonts(1)), "00") & Format(Hex(clofonts(2)), "00")
        
        clocode2 = clocode1
        If 色呼b <> "" Then
            '色の登録確認
            検索色 = 色呼b
            Set 検索x = .Columns(key.Column).Find(検索色, , , 1)
            If 検索x Is Nothing Then GoTo errFlg
            
            変換前 = 検索x.Offset(0, 2)
            clocode2s = Split(変換前, ",")
            clocode2 = Format(Hex(clocode2s(0)), "00") & Format(Hex(clocode2s(1)), "00") & Format(Hex(clocode2s(2)), "00")
        End If
    End With

Exit Function
errFlg:
    MsgBox "登録されていない色 " & 色呼a & " を含んでいます。登録してください。"
    Call 最適化もどす
    With wb(0).Sheets("color")
        .Select
        .Cells(.Cells(.Rows.count, key.Column).End(xlUp).Row + 1, key.Column) = 検索色
    End With
    
    End
Return
End Function

Function setWorkRan(Optional ByRef 端末サブRAN As Variant) As Variant
    
    Call checkSheet("PVSW_RLTF;端末一覧", wb(0), True, True)
    
    '端末一覧から使用するサブナンバーをゲット
    With wb(0).Sheets("端末一覧")
        Dim myKey As Variant, i As Long, 端末 As String, サブran() As Variant, foundFlag As Boolean, サブ As String
        ReDim サブran(0, 0)
        Set myKey = .Cells.Find(製品品番Ran(1, 1), , , 1)
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            サブ = .Cells(i, myKey.Column)
            foundFlag = False
            If サブ <> "" Then
                For x = LBound(サブran, 2) To UBound(サブran, 2)
                    If サブran(0, x) = サブ Then
                        foundFlag = True
                        Exit For
                    End If
                Next x
                If foundFlag = False Then
                    ReDim Preserve サブran(0, UBound(サブran, 2) + 1)
                    サブran(0, UBound(サブran, 2)) = サブ
                End If
            End If
        Next i
        If UBound(サブran, 2) = 0 Then
            MsgBox "[端末一覧]にサブナンバーがありません。"
            Stop
        End If
        サブran = WorksheetFunction.transpose(サブran) 'bubbleSort2の為に入れ替える
        Call BubbleSort2(サブran, 1)
        サブran = WorksheetFunction.transpose(サブran) 'bubbleSort2の為に入れ替える
    End With
    
    '端末一覧から端末№毎のサブナンバーをゲット
    With wb(0).Sheets("端末一覧")
        ReDim 端末サブRAN(2, 0)
        Dim 矢崎 As String
        Dim 端末Col As Long: 端末Col = .Cells.Find("端末№", , , 1).Column
        Dim 矢崎Col As Long: 矢崎Col = .Cells.Find("端末矢崎品番", , , 1).Column
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            サブ = .Cells(i, myKey.Column)
            端末 = .Cells(i, 端末Col)
            矢崎 = .Cells(i, 矢崎Col)
            If サブ <> "" Then
                ReDim Preserve 端末サブRAN(2, UBound(端末サブRAN, 2) + 1)
                端末サブRAN(0, UBound(端末サブRAN, 2)) = サブ
                端末サブRAN(1, UBound(端末サブRAN, 2)) = 端末
                端末サブRAN(2, UBound(端末サブRAN, 2)) = 矢崎
            End If
        Next i
    End With
    
    'PVSW_RLTFから条件をゲット
    Set myKey = ws(0).Cells.Find(製品品番Ran(1, 1), , , 1)
    '使用するフィールド名のセット
    Dim fieldname As String: fieldname = myKey.Value & ",RLTFtoPVSW_,始点側端末識別子,終点側端末識別子,始点側キャビティ,終点側キャビティ,接続G_,両端ハメ,構成_,色呼_"
    ff = Split(fieldname, ",")
    ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    
    Dim addFieldCount As Long: addFieldName = "maxCount,subSubNumber,作業順,cav"
    addFieldCount = 4
    a = UBound(ff) + addFieldCount
    
    Dim myRan() As Variant, sp As Variant
    ReDim myRan(a, 0)
    sp = Split(fieldname & "," & addFieldName, ",")
    For x = LBound(myRan) To UBound(myRan)
        myRan(x, 0) = sp(x)
    Next x
    
    Dim lastRow As Long
    lastRow = ws(0).Cells(ws(0).Rows.count, myKey.Column).End(xlUp).Row
    
    'サブナンバー順に電線条件をセットしていく
    Dim y As Long, サブstr As String, r As Long
    For y = LBound(サブran) + 1 To UBound(サブran)
        For x = 0 To 1
            For i = myKey.Row + 1 To lastRow
                サブstr = ws(0).Cells(i, myKey.Column).Value
                両端ハメ = ws(0).Cells(i, f(7)).Value
                If サブran(y) = サブstr Then
                    If 両端ハメ = CStr(x) Then
                        ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
                        For r = LBound(myRan) To UBound(myRan) - addFieldCount
                            myRan(r, UBound(myRan, 2)) = ws(0).Cells(i, f(r)).Value
                        Next r
                    End If
                End If
            Next i
        Next x
    Next y
    
    Call ReplaceLR_先ハメ誘導(myRan) '端末№の数を小さい方が左に移動
    
    '作業端末を決める
    Dim cav As String
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        サブstr = myRan(0, i)
        '片端ハメ
        If myRan(7, i) = "0" Then
            For x = 2 To 3
                端末 = myRan(x, i)
                cav = myRan(x + 2, i)
                foundFlag = False
                '端末が親なら最後ハメ
                If サブstr = 端末 Then
                    myRan(UBound(myRan) - 1, i) = "1000"
                    myRan(UBound(myRan) - 0, i) = cav
                    foundFlag = True
                Else
                    For r = LBound(端末サブRAN, 2) + 1 To UBound(端末サブRAN, 2)
                        If 端末 = 端末サブRAN(1, r) Then
                            If サブstr = 端末サブRAN(0, r) Then
                                myRan(UBound(myRan) - 1, i) = 端末
                                myRan(UBound(myRan) - 0, i) = cav
                                foundFlag = True
                                Exit For
                            End If
                        End If
                    Next r
                End If
                If foundFlag = True Then Exit For
            Next x
        Else ' 両端ハメ
            Dim flag As Boolean: flag = False
            For x = 2 To 3
                端末 = myRan(x, i)
                cav = myRan(x + 2, i)
                If サブstr <> 端末 Then
                    myRan(UBound(myRan) - 1, i) = 端末
                    myRan(UBound(myRan) - 0, i) = cav
                Else
                    flag = True
                End If
            Next x
            '親に繋がらない両端ハメの時作業順を最後にする
            If flag = flase Then
                myRan(UBound(myRan) - 1, i) = "2000"
            End If
        End If
    Next i
    
    'delete
    myRan = WorksheetFunction.transpose(myRan) 'bubbleSort2の為に入れ替える
    Call BubbleSort4(myRan, 1, UBound(myRan, 2) - 1, 8)
    myRan = WorksheetFunction.transpose(myRan) 'bubbleSort2の為に入れ替える
    
    '最大ステップ数,サブステップカウント数を確認
    Dim 両ハメ As String, outPath As String, myCount As Long, 端末1bak As String, 端末2bak As String, 部品1 As String, 部品2 As String
    Dim subStr As String, subStrbak As String, subCount As Long
    '作業ステップの最大値を確認
    Dim maxCount As Long
    myCount = 0: 画像名bak = "": subCount = 1
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        skipFlag = False
        subStr = myRan(1, i)
        両ハメ = myRan(8, i)
        If subStr <> subStrbak And subStrbak <> "" Then subCount = 0
        If subStrbak = "999" Or subStrbak = "999" Then
            skipFlag = True
            GoTo line15
        End If
        If 両ハメ = "0" Then
            If myRan(UBound(myRan) - 1, i) = "1000" Then
                画像名 = myRan(1, i) & "_" & 両ハメ
            Else
                画像名 = myRan(UBound(myRan) - 1, i) & "_" & 両ハメ
            End If
        Else
            If myRan(UBound(myRan) - 1, i) = "2000" Then
                画像名 = myRan(3, i) & "_" & myRan(4, i) & "_" & 両ハメ
            Else
                画像名 = myRan(UBound(myRan) - 1, i) & "_" & 両ハメ
            End If
        End If
line15:
        If 画像名 <> 画像名bak And 画像名bak <> "" And skipFlag = False Then
            myCount = myCount + 1
            subCount = subCount + 1
        End If
        myRan(UBound(myRan) - 2, i) = subCount
        subStrbak = subStr
        画像名bak = 画像名
    Next i
    
    maxCount = myCount
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        myRan(UBound(myRan) - 3, i) = maxCount
    Next i
    
    setWorkRan = myRan
    
End Function

Function setWorkRanV2(ByVal pNumber As String) As Variant
    pNumber = Replace(pNumber, " ", "")
    Call checkSheet("PVSW_RLTF;端末一覧", wb(0), True, True)
    
    'テキストファイル読み込み
    Dim myPath As String
    myPath = wb(0).path & "\09_AutoSub\" & Replace(pNumber, " ", "") & "_term.txt"
    termRan = readTextToArray(myPath)
    
    Dim words As String: words = "端末No,端末矢崎品番,親端末No"
    Dim ff As Variant: ff = Split(words, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    
    Dim fields As Variant, x As Long, i As Long
    For x = 0 To UBound(termRan)
        For i = 0 To UBound(f)
            If ff(i) = termRan(x, 1) Then
                f(i) = x
                Exit For
            End If
        Next i
    Next x
    
    '親端末をセット
    Dim foundFlag As Boolean, 親端末Ran() As Variant, 親端末 As String
    ReDim 親端末Ran(0, 0)
    For i = LBound(termRan, 2) + 2 To UBound(termRan, 2)
        親端末 = termRan(f(2), i)
        foundFlag = False
        If 親端末 <> "" Then
            For x = LBound(親端末Ran, 2) To UBound(親端末Ran, 2)
                If 親端末Ran(0, x) = 親端末 Then
                    foundFlag = True
                    Exit For
                End If
            Next x
            If foundFlag = False Then
                ReDim Preserve 親端末Ran(0, UBound(親端末Ran, 2) + 1)
                親端末Ran(0, UBound(親端末Ran, 2)) = 親端末
            End If
        End If
    Next i
    '昇順にする
    親端末Ran = WorksheetFunction.transpose(親端末Ran)
    BubbleSort2 親端末Ran, 1
    親端末Ran = WorksheetFunction.transpose(親端末Ran)
    
    '端末,端末矢崎品番,親端末をセット
    Dim 端末Ran() As Variant
    ReDim 端末Ran(2, 0)
    For i = LBound(termRan, 2) + 2 To UBound(termRan, 2)
        If termRan(f(0), i) <> "" Then
            ReDim Preserve 端末Ran(2, UBound(端末Ran, 2) + 1)
            For x = 0 To UBound(f)
                端末Ran(x, UBound(端末Ran, 2)) = termRan(f(x), i)
            Next x
        End If
    Next i
    
    'テキストファイル読み込み
    myPath = wb(0).path & "\09_AutoSub\" & Replace(pNumber, " ", "") & "_wireSum.txt"
    wireRan = readTextToArray(myPath)
    
    words = "始点側端末識別子,終点側端末識別子,接続G_,親端末No"
    Dim dd As Variant: dd = Split(words, ",")
    Dim d() As Variant: ReDim d(UBound(dd))
    
    For x = 0 To UBound(wireRan)
        For i = 0 To UBound(d)
            If dd(i) = wireRan(x, 1) Then
                d(i) = x
                Exit For
            End If
        Next i
    Next x
    
    'wordsの条件をセット
    Dim 電線RAN() As Variant
    ReDim 電線RAN(UBound(dd), 0)
    For i = LBound(wireRan, 2) + 2 To UBound(wireRan, 2)
            ReDim Preserve 電線RAN(UBound(dd), UBound(電線RAN, 2) + 1)
        For x = 0 To UBound(d)
            電線RAN(x, UBound(電線RAN, 2)) = wireRan(d(x), i)
        Next x
    Next i
    
    'PVSW_RLTFの条件をmyRanにセット
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    Set myKey = ws(0).Cells.Find(製品品番Ran(1, 1), , , 1)
    '使用するフィールド名のセット
    Dim fieldname As String: fieldname = myKey.Value & ",RLTFtoPVSW_,始点側端末識別子,終点側端末識別子,始点側キャビティ,終点側キャビティ,接続G_,両端ハメ,構成_,色呼_"
    ff = Split(fieldname, ",")
    ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    
    Dim addFieldCount As Long: addFieldName = "subNumber,親端末No,maxCount,subSubNumber,作業順,cav"
    addFieldCount = 6
    a = UBound(ff) + addFieldCount
    
    Dim myRan() As Variant, sp As Variant
    ReDim myRan(a, 0)
    sp = Split(fieldname & "," & addFieldName, ",")
    For x = LBound(myRan) To UBound(myRan)
        myRan(x, 0) = sp(x)
    Next x

    '親端末順に電線条件をセット
    Dim y As Long, 親端末str As String, r As Long, lastRow As Long, 端末1 As String, 端末2 As String, 接続G As String, 構成 As String
    lastRow = ws(0).Cells(ws(0).Rows.count, myKey.Column).End(xlUp).Row
    For y = LBound(親端末Ran) + 1 To UBound(親端末Ran)
        For x = 0 To 1
            For i = myKey.Row + 1 To lastRow
                If ws(0).Cells(i, myKey.Column).Value <> "" Then
                    構成 = ws(0).Cells(i, f(2)).Value
                    
                    端末1 = ws(0).Cells(i, f(2)).Value
                    端末2 = ws(0).Cells(i, f(3)).Value
                    接続G = ws(0).Cells(i, f(6)).Value
                    親端末str = search親端末(端末1, 端末2, 接続G, 電線RAN, d, dd)
                    両端ハメ = ws(0).Cells(i, f(7)).Value
                    If 親端末Ran(y) = 親端末str Then
                        If 両端ハメ = CStr(x) Then
                            ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
                            For r = LBound(myRan) To UBound(myRan) - addFieldCount
                                myRan(r, UBound(myRan, 2)) = ws(0).Cells(i, f(r)).Value
                            Next r
                            myRan(UBound(myRan) - 4, UBound(myRan, 2)) = 親端末str
                        End If
                    End If
                End If
            Next i
        Next x
    Next y
    
'    Call export_ArrayToSheet(myRan, "myRAN", False)


    Call ReplaceLR_先ハメ誘導(myRan) '端末№の数を小さい方が左に移動
    
    '作業端末を決める
    Dim cav As String
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        親端末str = myRan(UBound(myRan) - 4, i)
        '片端ハメ
        If myRan(7, i) = "0" Then
            For x = 2 To 3
                端末 = myRan(x, i)
                cav = myRan(x + 2, i)
                foundFlag = False
                '端末が親なら最後ハメ
                If 親端末str = 端末 Then
                    myRan(UBound(myRan) - 1, i) = "1000"
                    myRan(UBound(myRan) - 0, i) = cav
                    foundFlag = True
                Else
                    For r = LBound(端末Ran, 2) + 1 To UBound(端末Ran, 2)
                        If 端末 = 端末Ran(0, r) Then
                            If 親端末str = 端末Ran(2, r) Then
                                myRan(UBound(myRan) - 1, i) = 端末
                                myRan(UBound(myRan) - 0, i) = cav
                                foundFlag = True
                                Exit For
                            End If
                        End If
                    Next r
                End If
                If foundFlag = True Then Exit For
            Next x
        Else ' 両端ハメ
            Dim flag As Boolean: flag = False
            For x = 2 To 3
                端末 = myRan(x, i)
                cav = myRan(x + 2, i)
                If 親端末str <> 端末 Then
                    myRan(UBound(myRan) - 1, i) = 端末
                    myRan(UBound(myRan) - 0, i) = cav
                Else
                    flag = True
                End If
            Next x
            '親に繋がらない両端ハメの時作業順を最後にする
            If flag = flase Then
                myRan(UBound(myRan) - 1, i) = "2000"
            End If
        End If
    Next i
    
    'delete
    myRan = WorksheetFunction.transpose(myRan) 'bubbleSort2の為に入れ替える
    Call BubbleSort4(myRan, UBound(myRan, 2) - 4, UBound(myRan, 2) - 1, 8)
    myRan = WorksheetFunction.transpose(myRan) 'bubbleSort2の為に入れ替える
    
    '最大ステップ数,サブステップカウント数を確認
    Dim 両ハメ As String, outPath As String, myCount As Long, 端末1bak As String, 端末2bak As String, 部品1 As String, 部品2 As String
    Dim pStr As String, pStrbak As String, pCount As Long
    '作業ステップの最大値を確認
    Dim maxCount As Long
    myCount = 0: 画像名bak = "": pCount = 1
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        skipFlag = False
        pStr = myRan(UBound(myRan) - 4, i)
        両ハメ = myRan(8, i)
        If pStr <> pStrbak And pStrbak <> "" Then pCount = 0
        If pStrbak = "999" Or pStrbak = "999" Then
            skipFlag = True
            GoTo line15
        End If
        If 両ハメ = "0" Then
            If myRan(UBound(myRan) - 1, i) = "1000" Then
                画像名 = myRan(UBound(myRan) - 4, i) & "_" & 両ハメ
            Else
                画像名 = myRan(UBound(myRan) - 1, i) & "_" & 両ハメ
            End If
        Else
            If myRan(UBound(myRan) - 1, i) = "2000" Then
                画像名 = myRan(3, i) & "_" & myRan(4, i) & "_" & 両ハメ
            Else
                画像名 = myRan(3, i) & "_" & myRan(4, i) & "_" & 両ハメ
            End If
        End If
line15:
        If 画像名 <> 画像名bak And 画像名bak <> "" And skipFlag = False Then
            myCount = myCount + 1
            pCount = pCount + 1
        End If
        myRan(UBound(myRan) - 2, i) = pCount
        pStrbak = pStr
        画像名bak = 画像名
    Next i
    
    'サブナンバーの決定
    Dim subCount As Long
    subCount = 0: pStrbak = ""
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        If pStrbak <> myRan(UBound(myRan) - 4, i) Then
            subCount = subCount + 1
        End If
        If myRan(UBound(myRan) - 4, i) = "999" Then subCount = 999
        myRan(UBound(myRan) - 5, i) = subCount
        pStrbak = myRan(UBound(myRan) - 4, i)
    Next i
    
    maxCount = myCount
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        myRan(UBound(myRan) - 3, i) = maxCount
    Next i
    
    setWorkRanV2 = myRan
    
'    myRan = WorksheetFunction.transpose(myRan)
'    Sheets("Sheet3").Range("a1:z" & UBound(myRan)) = myRan
'    myRan = WorksheetFunction.transpose(myRan)
End Function

Sub makeDir(ByVal myPath As String)
    Dim sp As Variant, i As Long, tempDir As String
    sp = Split(myPath, "\")
    For i = LBound(sp) To UBound(sp)
        tempDir = tempDir & sp(i) & "\"
        If i <> UBound(sp) Then
            On Error Resume Next
            If Dir(tempDir, vbDirectory) = "" Then MkDir (tempDir)
            On Error GoTo 0
        End If
    Next i
End Sub


Function import_Sheet(ByVal myPath As String, ByVal sheetName As String) As Worksheet
    Dim tempWB As Workbook
    Workbooks.Open myPath, ReadOnly:=True
    Set tempWB = ActiveWorkbook
    tempWB.Sheets(sheetName).Copy after:=wb(0).Sheets("製品品番")
    tempWB.Close
    Set import_Sheet = ActiveSheet
End Function

Function checkFieldName(ByVal keyWord As String, ByVal ws As Worksheet, ByVal FieldNames As String) As String
    Dim sp As Variant, x As Long, key As Variant, found As Variant, msg As String
    sp = Split(FieldNames, ",")
    With ws
        Set key = .Cells.Find(keyWord, , , 1)
        For x = LBound(sp) To UBound(sp)
            Set found = .Rows(key.Row).Cells.Find(sp(x), , , 1, , , 1)
            If found Is Nothing Then
                msg = msg & "," & sp(x)
            End If
        Next x
        msg = Mid(msg, 2)
    End With
    checkFieldName = msg
End Function

Function revision_Compare(ByVal now As String, ByVal this As String) As String
    If now = "" Then revision_Compare = this: Exit Function
    
    Dim nowChar As Long, thisChar As Long, i As Long
    For i = 1 To Len(now)
        nowChar = Asc(Mid(now, i, 1))
        thisChar = Asc(Mid(this, i, 1))
        If nowChar > thisChar Then revision_Compare = now: Exit Function
        If nowChar < thisChar Then revision_Compare = this: Exit Function
    Next i
    
End Function

Function subTypeCheck(ByVal ran As Variant) As String
    Dim subStr As String, notCAE As Boolean, notPLUS As Boolean, subType As String
    For i = LBound(ran, 2) To UBound(ran, 2)
        subStr = ran(0, 1)
        '4桁じゃない場合はCAEサブじゃない
        If Len(subStr) <> 4 Then notCAE = True
        '数字型で2桁を超える場合且つ888,999じゃない場合はPLUSサブじゃない
        If IsNumeric(subStr) And Len(subStr) > 2 Then
            If subStr <> "999" And subStr <> "888" Then
                notPLUS = True
            End If
        End If
    Next i
    
    If notCAE = True And notPLUS = True Then
        subType = "その他"
    ElseIf notCAE = True And notPLUS = False Then
        subType = "PLUS"
    ElseIf notCAE = False And notPLUS = True Then
        subType = "CAE"
    Else
        subType = "不明"
    End If
    
    subTypeCheck = subType
End Function

Function searchStatus(ByVal ary As Variant, ByVal koseiNumber As String, ByVal terminalNumber As String) As String
    For i = LBound(ary) To UBound(ary)
        If koseiNumber = ary(0, i) Then
            For x = 1 To 2
                If terminalNumber = ary(x, i) Then
                    searchStatus = ary(x + 2, i)
                    End Function
                End If
            Next x
        End If
    Next i
    searchStatus = "notFound"
End Function
