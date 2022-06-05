Attribute VB_Name = "Module000"
Function GetIPAddress() As String

    Dim NetAdapters, objNic, strIPAddress
    Set NetAdapters = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") _
                           .ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & _
                           "Where (IPEnabled = TRUE)")

    For Each objNic In NetAdapters '�l�b�g���[�N�A�_�v�^�[�́A��������ꍇ������
        For Each strIPAddress In objNic.IPAddress 'IP�́A�������蓖�Ă��Ă���ꍇ������
            GetIPAddress = strIPAddress
            Exit For        ' �P��̂�
        Next
        Exit For        ' �P��̂�
    Next

End Function

Sub �摜�Ƃ��ďo��(myPicName)

    Selection.Copy
    
    ActiveSheet.Pictures.Paste.Select
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Set obj = Selection
    Dim myWidth As Single: myWidth = Selection.Width
    Dim myHeight As Single: myHeight = Selection.Height
     '�摜�\��t���p�̖��ߍ��݃O���t���쐬
    Set cht = ActiveSheet.ChartObjects.add(0, 0, myWidth, myHeight).Chart
     '���ߍ��݃O���t�ɓ\��t����
    cht.Paste
    cht.PlotArea.Fill.Visible = mesofalse
    cht.ChartArea.Fill.Visible = msoFalse
    cht.ChartArea.Border.LineStyle = 0
    
    '�T�C�Y����
    ActiveWindow.Zoom = 100
    '��l = 1000
    �{�� = 1
    �{�� = 192 / Selection.Width
    ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
    ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
    
    cht.Export fileName:=ActiveWorkbook.Path & "\" & myPicName & ".bmp", filtername:="bmp"
    
     '���ߍ��݃O���t���폜
    ActiveSheet.Activate
    cht.Parent.Delete
    obj.Delete
    
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True

End Sub

Public Function connect_Server()

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
        
    Dim myAddress As String, myPass As String, myAcount As String
    myAddress = "\\10.7.1.35\plus"
    myAcount = "nim.jp.yazaki.com\plus"
    Dim ans As Boolean
    ans = Ck_NetWork(myAddress) '�ڑ��ł��邩�m�F
    If ans = True Then Exit Function
    
    Dim oNetwork As IWshRuntimeLibrary.WshNetwork
    Set oNetwork = New IWshRuntimeLibrary.WshNetwork
    
    Call oNetwork.MapNetworkDrive("", myAddress, False, myAcount, StringUtility.byteToString(data))

End Function
Public Function remove_server()
    
    Dim oNetwork As IWshRuntimeLibrary.WshNetwork
    Set oNetwork = New IWshRuntimeLibrary.WshNetwork
    
    Dim myAddress As String
    myAddress = "\\10.7.1.35\plus"
    Call oNetwork.RemoveNetworkDrive(myAddress, True, True)

End Function

Public Function Ck_NetWork(myAddress)
   Dim WshShell As Object
   Dim Def_Dir As String
  
    Set WshShell = CreateObject("WScript.Shell")
   Def_Dir = WshShell.CurrentDirectory
   On Error Resume Next
   WshShell.CurrentDirectory = myAddress
   If Err.Number <> 0 Then
      Ck_NetWork = False
      Err.Clear
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
    PING_CMD = "C:\18_���ވꗗ+\ps\actTakePicture.bat"
    
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
    �ۑ��� = "C:\18_���ވꗗ+\ps\temp.jpg"
    result = URLDownloadToFile(0, picPath(2), �ۑ���, 0, 0)
    
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
      MsgBox "�l�b�g���[�N�ɐڑ����Ă��܂�", 64
   Else
      MsgBox "���݃l�b�g���[�N�ɐڑ�����Ă��܂���", 48
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
      MsgBox "�l�b�g���[�N�ɐڑ����Ă��܂�", 64
   Else
      MsgBox "���݃l�b�g���[�N�ɐڑ�����Ă��܂���", 48
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
      MsgBox "�l�b�g���[�N�ɐڑ����Ă��܂�", 64
   Else
      MsgBox "���݃l�b�g���[�N�ɐڑ�����Ă��܂���", 48
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork6()
    Call �A�h���X�Z�b�g(ThisWorkbook)
    Dim temp As String
    Dim duf
    Open "C:\18_���ވꗗ+\ps\actTakePicture.bat" For Input As #1
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
   Call obj.Run("cmd.exe ""C:\18_���ވꗗ+\ps\actTakePicture.bat""", 1, WaitOnreturn:=True)
     
   Set oExec = CreateObject("Wscript.shell").Exec(duf)
   Do While oExec.Status = 0
        Sleep 100
   Loop
   
   Debug.Print (oExec.StdOut.ReadAll)
   
   If InStr(1, Ping_St, "Lost = 0") > 0 Then
      MsgBox "�l�b�g���[�N�ɐڑ����Ă��܂�", 64
   Else
      MsgBox "���݃l�b�g���[�N�ɐڑ�����Ă��܂���", 48
   End If
   Set WshShell = Nothing
End Function

Public Function Ck_NetWork7()
    Dim WshShell As Object
    Dim Ping_Str As String
    PING_CMD = "ping -n 1 -w 100 10.7.120.117" ' & vbLf & "ping -n 1 -w 100 10.7.120.117"
    PING_CMD = "dir"
    PING_CMD = "C:\18_���ވꗗ+\ps\actTakePicture.ps1"
    
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

Public Function BubbleSort2(ByRef argAry() As Variant, ByVal keyPos As Long)
    Dim vSwap
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(argAry, 1) To UBound(argAry, 1)
        For j = UBound(argAry, 1) To i Step -1
            If Val(argAry(i, keyPos)) > Val(argAry(j, keyPos)) Then
                For k = LBound(argAry, 2) To UBound(argAry, 2)
                    vSwap = argAry(i, k)
                    argAry(i, k) = argAry(j, k)
                    argAry(j, k) = vSwap
                Next
            End If
        Next j
    Next i
 End Function
 
Public Function BubbleSort3(ByRef argAry() As Variant, ByVal keyPos As Long, ByVal keyPos2 As Long)
    Dim vSwap
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(argAry, 1) + 1 To UBound(argAry, 1)
        For j = i + 1 To UBound(argAry, 1)
            If Val(argAry(i, keyPos)) > Val(argAry(j, keyPos)) Then
                For k = LBound(argAry, 2) To UBound(argAry, 2)
                    vSwap = argAry(i, k)
                    argAry(i, k) = argAry(j, k)
                    argAry(j, k) = vSwap
                Next
            ElseIf Val(argAry(i, keyPos)) = Val(argAry(j, keyPos)) Then
                If Val(argAry(i, keyPos2)) < Val(argAry(j, keyPos2)) Then
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
        For X = LBound(myRan) To UBound(myRan)
            changedRan(i, X) = myRan(X, i)
        Next X
    Next i
    changeRowCol = changedRan
 End Function

Function ReplaceLR(ByRef myRan As Variant)
    Dim myStr(3) As String
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
         myStr(0) = myRan(1, i) '�Ђ���
         myStr(1) = myRan(2, i) '�݂�
         myStr(2) = myRan(3, i)
         myStr(3) = myRan(4, i)
         '�Ђ��肪��������Ȃ���΁A�݂����Ђ���ɂ����Ă���
         If myStr(0) = "" Then
            myRan(1, i) = myStr(1)
            myRan(2, i) = myStr(0)
            myRan(3, i) = myStr(3)
            myRan(4, i) = myStr(2)
         End If
         '�݂����������ꍇ�͍��E����
         If IsNumeric(myStr(0)) = True And IsNumeric(myStr(1)) = True Then
            If Val(myStr(1)) < Val(myStr(0)) Then
                myRan(1, i) = myStr(1)
                myRan(2, i) = myStr(0)
                myRan(3, i) = myStr(3)
                myRan(4, i) = myStr(2)
            End If
         End If
    Next i
End Function

Function SumRan(ByRef myRan As Variant)
    Dim count As Long
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        count = 1
        myRan(0, i) = count
        For ii = LBound(myRan, 2) + 1 To UBound(myRan, 2)
            If i <> ii Then
                If myRan(1, i) & "_" & myRan(2, i) = myRan(1, ii) & "_" & myRan(2, ii) Then
                    count = count + 1
                    myRan(0, i) = count
                    myRan(0, ii) = ""
                    myRan(1, ii) = ""
                    myRan(2, ii) = ""
                    myRan(3, ii) = ""
                    myRan(4, ii) = ""
                    '�d�㐡�@
                    If myRan(5, i) = "0" Then myRan(5, i) = myRan(5, ii)
                    myRan(5, ii) = ""
                End If
            End If
        Next ii
    Next i
End Function

Function evaluationRan(ByVal myRan As Variant)
    Dim hyokaRan As Variant
    ReDim hyokaRan(6, 0)
    Dim �[��str As String
    For i = LBound(myRan, 2) + 1 To UBound(myRan, 2)
        If myRan(0, i) <> "" Then
            For X = 1 To 2
                If myRan(X, i) <> "" Then
                    ReDim Preserve hyokaRan(6, UBound(hyokaRan, 2) + 1)
                    �[��str = myRan(X, i)
                    'If �[��str = "22" Then Stop
                    hyokaRan(0, UBound(hyokaRan, 2)) = myRan(X, i)
                    hyokaRan(1, UBound(hyokaRan, 2)) = myRan(X + 2, i)
                    hyokaRan(2, UBound(hyokaRan, 2)) = myRan(0, i) ^ 2
                    ���i����str = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(myRan(X + 2, i)), "���i����_")
                    hyokaRan(4, UBound(hyokaRan, 2)) = ���i����str
                    Select Case Left(���i����str, 3)
                        Case "003", "008"
                            �D�� = "1"
                        Case "001"
                            �D�� = "2"
                        Case "052", "020"
                            �D�� = "9"
                        Case Else
                            Stop '�o�^����Ă��Ȃ����i����
                    End Select
                    hyokaRan(3, UBound(hyokaRan, 2)) = �D��
                    hyokaRan(6, UBound(hyokaRan, 2)) = 1
                    If �D�� = "9" Then hyokaRan(5, UBound(hyokaRan, 2)) = "999"
                Else
                    GoTo line20
                End If
                For ii = LBound(myRan, 2) + 1 To UBound(myRan, 2)
                    If myRan(0, ii) <> "" Then
                        For xx = 1 To 2
                            If i <> ii Or X <> xx Then
                                If �[��str = myRan(xx, ii) Then
                                    hyokaRan(2, UBound(hyokaRan, 2)) = hyokaRan(2, UBound(hyokaRan, 2)) + myRan(0, ii) ^ 2
                                    hyokaRan(6, UBound(hyokaRan, 2)) = hyokaRan(6, UBound(hyokaRan, 2)) + 1
                                    myRan(xx, ii) = ""
                                End If
                            End If
                        Next xx
                    End If
                Next ii
                '�D�悪1�ł��s���悪1�ӏ��̏ꍇ��2�ɉ�����
                If �D�� = "1" And hyokaRan(6, UBound(hyokaRan, 2)) = 1 Then
                    hyokaRan(3, UBound(hyokaRan, 2)) = "2"
                End If
line20:
            Next X
        End If
    Next i
    evaluationRan = hyokaRan
End Function

Function search����[���]��(ByVal myRan As Variant, ByVal ����[��str As String)
    If ����[��str = "" Then
        search����[���]�� = 1
        Exit Function
    End If
    Dim ����[���]��lng As Long: ����[���]��lng = 0
    For i = LBound(myRan) + 1 To UBound(myRan)
        For j = LBound(myRan) + 1 To UBound(myRan)
            For X = 1 To 2
                If ����[��str = myRan(j, X) Then
                    If myRan(j, 6) = "" Then
                        If myRan(j, 0) > ����[���]��lng Then
                            ����[���]��lng = myRan(j, 0)
                        End If
                    End If
                End If
            Next X
        Next j
    Next i
    search����[���]�� = ����[���]��lng
End Function

Function search�[���]��RAN(ByVal myRan As Variant, ByVal �[��str As String, myPos As Integer)
    For i = LBound(myRan) To UBound(myRan)
        If �[��str = myRan(i, 0) Then
            search�[���]��RAN = myRan(i, myPos)
            Exit Function
        End If
    Next i
End Function

Function search�[���]��RAN_2pos(ByVal myRan As Variant, ByVal �[��str1 As String, ByVal �[��str2 As String, myPos As Integer)
    For i = LBound(myRan) To UBound(myRan)
        If �[��str1 = myRan(i, 0) And �[��str2 = myRan(i, 1) Then
            search�[���]��RAN_2pos = myRan(i, myPos)
            Exit Function
        End If
    Next i
End Function

Function search�[���d����RAN(ByVal myRan As Variant, ByVal �[��str1 As String, ByVal �[��str2 As String, myPos As Integer)
    For i = LBound(myRan) To UBound(myRan)
        If �[��str1 = myRan(i, 1) And �[��str2 = myRan(i, 2) Then
            search�[���d����RAN = myRan(i, myPos)
            Exit Function
        End If
    Next i
End Function

Function readTextToArray(ByVal myPath As String)
        Dim myRan As Variant
        Dim Target As New FileSystemObject
        If Dir(myPath) = "" Then readTextToArray = False: Exit Function
        Dim intFino As Variant
        intFino = FreeFile
        Open myPath For Input As #intFino
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            a = UBound(temp)
            If IsEmpty(myRan) Then
                ReDim myRan(a, 0)
            End If
            ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
            For X = LBound(temp) To UBound(temp)
                myRan(X, UBound(myRan, 2)) = temp(X)
            Next X
        Loop
        Close #intFino
        readTextToArray = myRan
End Function
