Attribute VB_Name = "M23_IE"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'API_�摜���_�E�����[�h
Public Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    
Dim ���ޏڍ�_�^�C�g��Ran As Range

Sub fajdlajf()

    Set area = Range("aa21:jk1142")
    
    For Each a In area
        If a.Value <> "" Or a.Value <> Empty Then
            If a.Borders(xlEdgeBottom).LineStyle <> xlContinuous Then
                a.Select
                Stop
            End If
        End If
    Next a

End Sub

Sub ie_�ʒm�����擾(�Ԏ�str)

    Dim �����\��(1 To 20) As String
    Dim iD As String
    Dim myRyakuDir As String
    Dim mailURL(2) As String
    Dim ���(2) As String

    With Sheets("�ݒ�")
        mailURL(0) = .Cells.Find("�ʒm���A�h���X_", , , 1).Offset(0, 1).Value
        mailURL(1) = .Cells.Find("�ʒm���A�h���X_", , , 1).Offset(1, 1).Value
        mailURL(2) = .Cells.Find("�ʒm���A�h���X_", , , 1).Offset(2, 1).Value
    End With
    ���(0) = "��"
    ���(1) = "��"
    ���(2) = "��"
    '�}�����ۊǗp�̃t�H���_
'    myRyakuDir = ActiveWorkbook.PAth & "\�}����"
'    If Dir(myRyakuDir, vbDirectory) = "" Then MkDir myRyakuDir
    'IE�̋N��
    Dim objIE As Object '�ϐ����`���܂�
    Dim ieVerCheck As Variant
    Set objIE = CreateObject("InternetExplorer.Application")
    Set objSFO = CreateObject("Scripting.FileSystemObject")
'    Select Case Application.OperatingSystem
'    Case "Windows (32-bit) NT 6.01"
'        Set objIE = CreateObject("InternetExplorer.Application") '�I�u�W�F�N�g���쐬���܂��B
'    Case Else
'        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
'    End Select
'    objIE.Visible = True
    
    ieVerCheck = Val(objSFO.GetFileVersion(objIE.FullName))
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10�ȍ~(���Ԃ�)
    Else
        Set objIE = CreateObject("InternetExplorer.Application") '�m��񂯂�
    End If
    
    objIE.Visible = True
    '��L��64-bit NT 6.01�Ȃ̂�32bit�Ɣ��f�����s��̎b��΍�
    On Error Resume Next
    objIE.Navigate mailURL(p)
    a = objIE.readystate
    b = objIE.busy
    Debug.Print Err.Number
    If Err.Number = -2147417848 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
        objIE.Navigate mailURL(p)
    End If
    
    On Error GoTo 0
    
    '���ATrue�Ō�����悤�ɂ��܂�
    '�����������y�[�W��\�����܂�
    '���1 ���O�C�����
'   objIE.document.all.Item(�A�J�E���gID).Value = �A�J�E���g
'   objIE.document.all.Item(�p�XID).Value = �p�X
'   objIE.document.all.Item("btnLogin").Click '���O�C���N���b�N
'   Call �y�[�W�\����҂�(objIE)
'   '���2 �g�p���ӏ��
'   objIE.document.all.Item("btnOK").Click 'OK�N���b�N
'   Call �y�[�W�\����҂�(objIE)
'   '���3 ���C���y�[�W
'   objIE.document.all.Item("btnYzk").Click '���i�Ԃ���̌���
'   Call �y�[�W�\����҂�(objIE)
'loop
   With ActiveSheet
        Dim key As Range: Set key = .Cells.Find("key_", , , 1)
        Dim key2 As Range: Set key2 = .Cells.Find("�ʒm����_", , , 1)
        Dim lastRow As Long: lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .UsedRange.Columns.count
        Dim �ʒm��Row As Long: �ʒm��Row = key2.Row
        Dim �ʒm��Col As Long: �ʒm��Col = key2.Column
        Dim ���tCol As Long: ���tCol = .Rows(key2.Row).Find("���t_", , , 1).Column
        Dim ���Rcol As Long: ���Rcol = .Rows(key2.Row).Find("���R_", , , 1).Column
        Dim �ύX�v�_col As Long: �ύX�v�_col = .Cells.Find("�ύX�v�__", , , 1).Column
        Dim �ŏI�擾��Row As Long: �ŏI�擾��Row = .Cells.Find("�ŏI�擾��", , , 1).Row
        Dim �Ԏ�Row As Long: �Ԏ�Row = .Cells.Find("�Ԏ�_", , , 1).Row
        Dim �ʒm�� As String

        '�ʒm�����̓o�^
        Dim �ʒm��RAN() As Variant, j As Long
        GoSub �ʒm���̓o�^

        '�Ώۂ̐��i�i�Ԃ̓_�����v�Z
        For X = key.Offset(0, 1).Column To lastCol
            �Ԏ� = .Cells(�Ԏ�Row, X)
            If �Ԏ�str = "" Or InStr(�Ԏ�str, �Ԏ�) > 0 Then
                ���i�i�� = .Cells(key.Row, X)
                If ���i�i�� <> "" Then
                    Total = Total + 1
                End If
            End If
        Next X
        
        Dim myText As String, mytext2 As String, myTextA As String, myTextTR As String
        Dim aa(6) As Long
        
        For X = key.Offset(0, 1).Column To lastCol
            For p = LBound(mailURL) To UBound(mailURL)
                �Ԏ� = .Cells(�Ԏ�Row, X)
                If �Ԏ�str = "" Or InStr(�Ԏ�str, �Ԏ�) > 0 Then
                    ���i�i�� = Replace(.Cells(key.Row, X), " ", "")
                    If ���i�i�� <> "" Then
                        .Cells(key.Row, X).Select
                        '�Ώۃy�[�W�̕\��
                        objIE.Navigate mailURL(p)
                        Call �y�[�W�\����҂�(objIE)
                        '�i�ԓ���
                        Select Case p
                            Case 0, 1
                            objIE.document.all.Item("hinban").Value = ���i�i��
                            hensu = 2
                            Case 2
                            objIE.document.all.Item("s_hinban").Value = ���i�i��
                            hensu = 2
                        End Select
                        Call �y�[�W�\����҂�(objIE)
                        
                        '�����N���b�N
                        Call �{�^���N���b�N(objIE, "����")
                        Call �y�[�W�\����҂�(objIE)
                        
                        '��ʏ��̎擾
                        For i = 0 To objIE.document.getElementsByTagName("tr").Length - hensu
                            Select Case p
                                Case 0
                                    myText = objIE.document.getElementsByTagName("tr")(i + 1).outerhtml
                                    If InStr(StrConv(myText, vbUpperCase), "HREF") > 0 Then
                                        URL = objIE.document.getElementsByTagName("a")(i - URL����count).href
                                        c = 0
                                    Else
                                        URL = ""
                                        c = 1
                                        URL����count = URL����count + 1
                                    End If
                                    a = ����(myText, ">", 3 - c)
                                    b = ����(myText, "<", 4 - c)
                                    �ʒm�� = Mid(myText, a + 1, b - a - 1)
                                    a = ����(myText, ">", 6 - c - c)
                                    b = ����(myText, "<", 7 - c - c)
                                    ���t = CDate(Mid(myText, a + 1, b - a - 1))
                                    a = ����(myText, ">", 10 - c - c)
                                    b = ����(myText, "<", 11 - c - c)
                                    ���R = Mid(myText, a + 1, b - a - 1)
                                    a = ����(myText, ">", 14 - c - c)
                                    b = ����(myText, "<", 15 - c - c)
                                    �ݕ� = Mid(myText, a + 1, b - a - 1)
                                    ���i = ""
                                Case 1
                                    myText = objIE.document.getElementsByTagName("tr")(i + 1).outerhtml
                                    'PDF�̃����N��L���̊m�F_�����������遨http://10.7.1.35/nim_intra/70_busyobetsu/30_sekkei/program/hentu/hinban_result.asp
                                    If InStr(StrConv(myText, vbUpperCase), "HREF") > 0 Then
                                        URL = objIE.document.getElementsByTagName("a")(i - URL����count).href
                                        c = 0
                                    Else
                                        URL = ""
                                        c = 1
                                        URL����count = URL����count + 1
                                    End If
                                    a = ����(myText, ">", 3 - c)
                                    b = ����(myText, "<", 4 - c)
                                    �ʒm�� = Mid(myText, a + 1, b - a - 1)
                                    a = ����(myText, ">", 6 - c - c)
                                    b = ����(myText, "<", 7 - c - c)
                                    ���t = CDate(Mid(myText, a + 1, b - a - 1))
                                    ���R = "�݌v�ύX"
                                    a = ����(myText, ">", 12 - c - c)
                                    b = ����(myText, "<", 13 - c - c)
                                    �ݕ� = Mid(myText, a + 1, b - a - 1)
                                    ���i = ""
                                Case 2
                                    myText = objIE.document.getElementsByTagName("tr")(i + 1).outerhtml
                                    If InStr(StrConv(myText, vbUpperCase), "HREF") > 0 Then
                                        
                                        URL = objIE.document.getElementsByTagName("a")(i - URL����count).href
                                        c = 0
                                    Else
                                        URL = ""
                                        c = 1
                                        URL����count = URL����count + 1
                                    End If
                                    a = ����(myText, ">", 3 - c)
                                    b = ����(myText, "<", 4 - c)
                                    �ʒm�� = Mid(myText, a + 1, b - a - 1)
                                    a = ����(myText, ">", 6 - c - c)
                                    b = ����(myText, "<", 7 - c - c)
                                    ���t = CDate(Mid(myText, a + 1, b - a - 1))
                                    ���R = "���i�ύX"
                                    a = ����(myText, ">", 14 - c - c)
                                    b = ����(myText, "<", 15 - c - c)
                                    �ݕ� = Mid(myText, a + 1, b - a - 1)
                                    a = ����(myText, ">", 10 - c - c)
                                    b = ����(myText, "<", 11 - c - c)
                                    ���i = Mid(myText, a + 1, b - a - 1)
                            End Select
                            
                            addRow = 0
                            '�o�^���Ă邩�m�F
                            flg = False
                            For r = LBound(�ʒm��RAN, 2) To UBound(�ʒm��RAN, 2)
                                If �ʒm�� = �ʒm��RAN(0, r) And ���(p) = �ʒm��RAN(2, r) Then
                                    addRow = �ʒm��RAN(1, r)
                                    Exit For
                                End If
                            Next r
                            
                            '�����ꍇ�o�^
                            If addRow = 0 Then
                                flg = True
                                For r = LBound(�ʒm��RAN, 2) To UBound(�ʒm��RAN, 2)
                                    If ���t < �ʒm��RAN(3, r) Then
                                        addRow = �ʒm��RAN(1, r)
                                        .Rows(addRow).Insert
                                        .Range(.Cells(key2.Row + 1, 1), .Cells(key2.Row + 1, key.Column)).Copy .Range(.Cells(addRow, 1), .Cells(addRow, key.Column))
                                        .Range(.Cells(addRow, 1), .Cells(addRow, key.Column)).ClearContents
                                        .Range(.Cells(addRow, key.Column + 1), .Cells(addRow, .Columns.count)).ClearFormats
                                        Exit For
                                        Stop
                                    End If
                                Next r
                            End If
                            
'                            '�����ꍇ�o�^
'                            If addRow = 0 Then
'                                addRow = .Cells(.Rows.Count, key2.Column).End(xlUp).Row + 1
'                                ReDim Preserve �ʒm��RAN(3, UBound(�ʒm��RAN, 2) + 1)
'                                �ʒm��RAN(0, UBound(�ʒm��RAN, 2)) = �ʒm��
'                                �ʒm��RAN(1, UBound(�ʒm��RAN, 2)) = addRow
'                                �ʒm��RAN(2, UBound(�ʒm��RAN, 2)) = ���(p)
'                                .Cells(addRow, x).Select
'                            End If
                            '�o��
                            If addRow = 0 Then
                                addRow = .Cells(.Rows.count, key2.Column).End(xlUp).Row + 1
                            End If
                            .Cells(addRow, key2.Column + 0) = �ʒm��
                            .Cells(addRow, key2.Column - 1) = ���(p)
                            .Cells(addRow, key2.Column).NumberFormat = "@"
                            If URL <> "" Then
                                .Hyperlinks.add anchor:=.Cells(addRow, key2.Column), address:=URL, ScreenTip:="", TextToDisplay:=CStr(�ʒm��)
                            Else
                                .Cells(addRow, key2.Column).Font.Underline = False
                            End If
                            Select Case p
                                Case 0
                                .Cells(addRow, key2.Column).Font.color = RGB(0, 0, 255)
                                .Cells(addRow, ���Rcol).Font.color = RGB(0, 0, 255)
                                .Cells(addRow, �ύX�v�_col).Font.color = RGB(0, 0, 0)
                                �ݕ� = Left(�ݕ�, 1) & Mid(�ݕ�, 3, 1) & Mid(�ݕ�, 5, 1)
                                Case 1
                                .Cells(addRow, key2.Column).Font.color = RGB(255, 0, 255)
                                .Cells(addRow, ���Rcol).Font.color = RGB(255, 0, 255)
                                .Cells(addRow, �ύX�v�_col).Font.color = RGB(0, 0, 0)
                                �ݕ� = Left(�ݕ�, 1) & Mid(�ݕ�, 3, 1) & Mid(�ݕ�, 5, 1)
                                Case 2
                                .Cells(addRow, key2.Column).Font.color = RGB(0, 100, 0)
                                .Cells(addRow, ���Rcol).Font.color = RGB(0, 100, 0)
                                .Cells(addRow, �ύX�v�_col).Font.color = RGB(0, 100, 0)
                                .Cells(addRow, �ύX�v�_col) = CStr(���i)
                                �ݕ� = Left(�ݕ�, 1) & Mid(�ݕ�, 3, 1) & Mid(�ݕ�, 5, 1)
                            End Select
                            
                            .Cells(addRow, ���Rcol) = ���R
                            .Cells(addRow, ���tCol).NumberFormat = "yy/mm/dd"
                            .Cells(addRow, ���tCol) = ���t
                            .Cells(addRow, X).NumberFormat = "@"
                            .Cells(addRow, X).HorizontalAlignment = xlCenter
                            .Cells(addRow, X).VerticalAlignment = xlCenter
                            .Cells(addRow, X).Font.Bold = True
                            .Cells(addRow, X) = �ݕ�
                            .Cells(addRow, X).Borders.Weight = xlThin
                            .Cells(addRow, X).Select
                            .Rows(addRow).RowHeight = 27
                            If flg = True Then GoSub �ʒm���̓o�^
                            
                        Next i
                        .Cells(�ŏI�擾��Row, X) = Date
                        If p = 0 Then
                            onetime = DateDiff("s", mytime, Time)
                            totaltime = totaltime + onetime
                            count = count + 1
                            counttime = totaltime / count
                            Application.StatusBar = "  " & count & "/" & Total & "  �c��: " & Int(((Total - count) * counttime) / 60)
                            mytime = Time
                        End If
                     End If
                 End If
                 URL����count = 0
            Next p
        Next X
        '���ёւ������珑���������̂ŕ��ёւ������Ȃ�
'        Stop
'        addRow = .Cells(Rows.Count, key2.Column).End(xlUp).Row
'        With .Sort.SortFields
'            .Clear
'            .Add key:=Range(Cells(key2.Row, ���tCol).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
''            .Add key:=Range(Cells(1, �D��2).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'        End With
'        .Sort.SetRange .Range(.Rows(key2.Row), .Rows(addRow))
'        .Sort.Header = xlYes
'        .Sort.MatchCase = False
'        .Sort.Orientation = xlTopToBottom
'        .Sort.Apply
'        '.Rows(key2.Row + 1 & ":" & addRow).Sort key1:=.Cells(key2.Row, ���tCol), order1:=xlAscending
        Application.StatusBar = False
        
        objIE.Quit
        Set objIE = Nothing
    End With
    
    MsgBox "�X�V���������܂����B"
    
Exit Sub

�ʒm���̓o�^:

        ReDim �ʒm��RAN(3, 0): j = 0
        With ActiveSheet
            lastRow = .UsedRange.Rows.count
            For ii = key2.Row + 1 To lastRow
                If .Cells(ii, �ʒm��Col) <> "" Then
                    ReDim Preserve �ʒm��RAN(3, j)
                    �ʒm��RAN(0, j) = .Cells(ii, �ʒm��Col)
                    �ʒm��RAN(1, j) = ii
                    �ʒm��RAN(2, j) = .Cells(ii, �ʒm��Col - 1)
                    �ʒm��RAN(3, j) = .Cells(ii, ���tCol)
                    j = j + 1
                End If
            Next ii
        End With
Return

End Sub

Public Function a�擾_�R�l�N�^��_�R�l�N�^�ɐ�(ByVal objIE As Object, iD, �R�l�N�^�ɐ�)
  �R�l�N�^�ɐ� = ""
    �������� = "�R�l�N�^�ɐ�"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    �R�l�N�^�ɐ� = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a�擾_���}�_�E�����[�h(ByVal objIE As Object, myRyakuDir, ���i�i��)
    �������� = "<img src="
    For a = 0 To 1
        '���}�̃{�^��id��������Ώ������Ȃ�
        If InStr(objIE.document.all(0).outerhtml, "ctl01_dispRyaku_btnDraw") = 0 Then Exit Function
        objIE.document.all.Item("ctl01_dispRyaku_edtText").Value = ���i�i��
        objIE.document.all.Item("ctl01_dispRyaku_rgpReverse_" & a).Click      '0=���ʎ� 1=���ʎ�
        objIE.document.all.Item("ctl01_dispRyaku_cmbText")(3).Selected = True '�e�L�X�g����
        objIE.document.all.Item("ctl01_dispRyaku_chkOriginalSize").Checked = True     '�`��
        objIE.document.all.Item("ctl01_dispRyaku_btnDraw").Click              '�`��
        
        Call �y�[�W�\����҂�(objIE)
        For X = 0 To objIE.document.all.tags("img").Length - 1  '�v�f�̐�
            �f�[�^ = objIE.document.all.tags("img")(X).outerhtml
            aaa = InStr(�f�[�^, ��������)
            If aaa = 0 Then GoTo line0
            ���}URL = "http://10.1.33.95/DesignSource" & Mid(�f�[�^, Len(��������) + 3)
            ���}URL = Left(���}URL, Len(���}URL) - 2)
            ���}�ۑ�PASS = myRyakuDir & "\" & ���i�i�� & "_" & a & "_" & Format(X, "000") & ".emf"
            '�_�E�����[�h�̎��s
            Ret = URLDownloadToFile(0, ���}URL, ���}�ۑ�PASS, 0, 0)
line0:
        Next X
    Next a
End Function
Public Function a�擾_���Ӑ�i��(ByVal objIE As Object, iD, ByVal i As Long)
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML
    On Error GoTo 0
    Dim �f�[�^s As Variant
    Dim �^�C�g��AddCol As Long
    �f�[�^s = Split(�f�[�^, vbLf)
    For Each �f�[�^o In �f�[�^s
        a = InStr(�f�[�^o, "<th"): If a <> 0 Then GoTo line10
        aa = InStr(�f�[�^o, Chr(34) & ">"): If aa = 0 Then GoTo line10
        aaa = Mid(�f�[�^o, aa + 2)
        bb = InStr(aaa, "<"): If bb = 0 Then GoTo line10
        bbb = Left(aaa, bb - 1)
        cc = InStr(aaa, Chr(34) & ">"): If cc = 0 Then GoTo line10
        ccc = Mid(aaa, cc + 2)
        dd = InStr(ccc, "<"): If dd = 0 Then GoTo line10
        ddd = Left(ccc, dd - 1)
        ���Ӑ於 = Replace(bbb, "&nbsp;", "")
        ���Ӑ�i�� = Replace(ddd, "&nbsp;", "")
    '���ޏڍׂ���T���č��ڂ�������Βǉ�
    With Sheets("���ޏڍ�")
        Set ���Ӑ於find = ���ޏڍ�_�^�C�g��Ran.Find(���Ӑ於 & "_", LookAt:=xlWhole)
        If ���Ӑ於find Is Nothing Then
            Dim �^�C�g��Row As Long: �^�C�g��Row = ���ޏڍ�_�^�C�g��Ran.Row
             �^�C�g��AddCol = .Cells(�^�C�g��Row, .Columns.count).End(xlToLeft).Column + 1
            .Cells(�^�C�g��Row - 1, �^�C�g��AddCol) = "���Ӑ於"
            .Cells(�^�C�g��Row, �^�C�g��AddCol) = ���Ӑ於 & "_"
        Else
            �^�C�g��AddCol = ���Ӑ於find.Column
        End If
            .Cells(i, �^�C�g��AddCol).NumberFormat = "@"
            .Cells(i, �^�C�g��AddCol) = ���Ӑ�i��
    End With
line10:
    Next
End Function

Public Function a�擾_�`���[�u�O�a(ByVal objIE As Object, iD, �`���[�u�O�a)
  �`���[�u�O�a = ""
    �������� = "�`���[�u�O�a"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    �`���[�u�O�a = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a�擾_�`���[�u���a(ByVal objIE As Object, iD, �`���[�u���a)
  �`���[�u���a = ""
    �������� = "�`���[�u���a"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    �`���[�u���a = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a�擾_�`���[�u����(ByVal objIE As Object, iD, �`���[�u����)
  �`���[�u���� = ""
    �������� = "�`���[�u����"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    �`���[�u���� = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function

Public Function a�擾_�`���[�u�i��(ByVal objIE As Object, iD, �`���[�u�i��)
  �`���[�u�i�� = ""
    �������� = "�`���[�u�i��"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    �`���[�u�i�� = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a�擾_�N�����v�^�C�v(ByVal objIE As Object, �N�����v�^�C�v)
  �N�����v�^�C�v = ""
  
    �������� = "�N�����v�^�C�v"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmIndivs").outertext 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = Mid(ccc, Len(��������) + 1)
    �N�����v�^�C�v = Replace(ddd, vbLf, "")
    
End Function

Public Function a�擾_�d�オ��O�a(ByVal objIE As Object, �d�オ��O�a)
  �d�オ��O�a = ""
  
    �������� = "�d�オ��O�a"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmIndivs").innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ";")
    ccc = InStr(bbb + 1, �f�[�^, ";")
    ddd = InStr(ccc + 1, �f�[�^, ";")
    eee = InStr(ddd + 1, �f�[�^, ">")
    zzz = InStr(eee + 1, �f�[�^, "<")
    �d�オ��O�a = Mid(�f�[�^, eee + 1, zzz - eee - 1)
      
End Function

Public Function �y�[�W�\����҂�(ByRef objIE As Object)

    While objIE.readystate <> 4 Or objIE.busy = True '.ReadyState <> 4�̊Ԃ܂��B
        DoEvents  '�d���̂Ō����Ȑl���邯�ǁB
        Sleep 1
        'Call ���z�L�[����(�V�t�g)
    Wend
    
End Function

Public Function a�擾_���}(ByVal objIE As Object, ���}URL, ���}��)
  ���}URL = "": ���}�� = 0
  
    ���}�� = objIE.document.Images.Length - 1
  
    For r = 1 To objIE.document.Images.Length - 1
  
        ���}URL = objIE.document.Images(1).src
    Next r
  
      
End Function

Public Function a�擾_���i���(ByVal objIE As Object, ���i���)
  ���i��� = ""
  
    �������� = "���i���"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(ccc + 1, �f�[�^, "<")
    
    If aaa <> 0 Then ���i��� = Mid(�f�[�^, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a�擾_���i����(ByVal objIE As Object, ���i����)
  ���i���� = ""
  
    �������� = "���i����"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(ccc + 1, �f�[�^, "<")
    
    If aaa <> 0 Then ���i���� = Mid(�f�[�^, ccc + 1, zzz - ccc - 1)
      
End Function
Public Function a�擾_���i����(ByVal objIE As Object, ���i����)
  ���i���� = ""
  
    �������� = "���i����"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(ccc + 1, �f�[�^, "<")
    
    If aaa <> 0 Then ���i���� = Mid(�f�[�^, ccc + 1, zzz - ccc - 1)
      
End Function
Public Function a�擾_�o�^�H��(ByVal objIE As Object, �o�^�H��)
  �o�^�H�� = ""
  
    �������� = "�o�^�H��"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(ccc + 1, �f�[�^, "<")
        
    If aaa <> 0 Then �o�^�H�� = Mid(�f�[�^, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a�擾_����(ByVal objIE As Object, ���̕i��)
  ���̕i�� = "": �f�[�^ = ""
  
    �������� = "����"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdEmtrCommon").innerHTML 'JAIRS���
    On Error GoTo 0
    
    If �f�[�^ = "" Then
        �������� = "�i��"
        On Error Resume Next
        �f�[�^ = objIE.document.getElementById("ctl01_grdJairsCommon").innerHTML 'JAIRS���
        On Error GoTo 0
    End If
    
    If �f�[�^ = "" Then Stop '��L�̂ǂ����������Ȃ�
        
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(ccc + 1, �f�[�^, "<")
        
    If aaa <> 0 Then ���̕i�� = Mid(�f�[�^, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a�擾_���i�F(ByVal objIE As Object, ���i�F)
  ���i�F = "": �f�[�^ = ""
  
    �������� = "�F"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdJairsSpecs").innerHTML 'JAIRS�̎d�l
    On Error GoTo 0
        
    'If �f�[�^ = "" Then Stop '��L�̂ǂ����������Ȃ�
        
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(ccc + 1, �f�[�^, "<")
        
    If aaa <> 0 Then ���i�F = Mid(�f�[�^, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a�擾_�d��(ByVal objIE As Object, �d��)
  �d�� = ""
  
    �������� = "�d��"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdJairsSize").innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    �d�� = Mid(eee, fff + 1, Len(eee) - fff + 1)
      
End Function

Public Function a�擾_�����\��(ByVal objIE As Object, �����\��)
  
    �������� = "�ǔ�"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdEmtrComp").innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, Len(��������) + aaa + 1, Len(�f�[�^))
    
    For i = 1 To 20
        ccc = InStr(bbb, "target")
        If ccc = 0 Then �����\��(i) = "": GoTo line10
        ddd = Mid(bbb, ccc, Len(bbb))
        eee = InStr(ddd, ">")
        fff = InStr(ddd, "<")
        ggg = Mid(ddd, eee + 1, fff - eee - 1)
        �����\��(i) = ggg
        
        bbb = Mid(bbb, ccc + fff, Len(bbb))
line10:
    Next i
          
End Function

Public Function a�擾_��������(ByVal objIE As Object, ��������, ByVal ���i�i��)
    �������� = ""
    
    Dim �����N�ԍ� As Long
    'NotFound�m�F
    �f�[�^ = objIE.document.getElementById("ctl00_lblErrMsg").innerHTML
    �������� = �f�[�^
    If �������� = "Not Found." Then Exit Function
    
    '���������_�����m�F
    �f�[�^ = objIE.document.getElementById("ctl00_grdList").innerHTML
    aaa = InStrRev(�f�[�^, "grdList")
    bbb = Mid(�f�[�^, aaa + 8, 100)
    zzz = InStr(bbb, "'")
    �_�� = Mid(�f�[�^, aaa + 8, zzz - 1)
    
    '�_������������ꍇ�A�����N���N���b�N
    If �_�� > 0 Then
    '�����Naaa = InStrRev(�f�[�^, ">" & Replace(���i�i��, "-", "") & "<")
    '�����Nbbb = Left(�f�[�^, �����Naaa)
    '�����Nccc = InStrRev(�����Nbbb, "grdList")
    '�����N�A�h���X = Mid(�����Nbbb, �����Nccc, 9 + Len(�_��))
    'objIe.document.all.Item("javascript:__doPostBack('ctl00$grdList','grdList$0')").Click
    
    '�����N�ԍ��ŊJ��(�_��+4�Ō�������ׁA�m���ł͂Ȃ�����)
    �����Naaa = InStrRev(�f�[�^, ">" & Replace(���i�i��, "-", "") & "<")
    If �����Naaa <> 0 Then
        �����Nbbb = Left(�f�[�^, �����Naaa)
        �����Nccc = InStrRev(�����Nbbb, "$")
        �����Nzzz = InStrRev(�����Nbbb, "'")
        �����N�ԍ� = Mid(�����Nbbb, �����Nccc + 1, �����Nzzz - (�����Nccc + 1))
    Else
        �������� = "NotMatch"
    End If
    
    objIE.document.Links(4).Click
    
    End If
    
    Call �y�[�W�\����҂�(objIE)
        '�\�����ꂽ�i�Ԃƌ����������i�Ԃ��}�b�`���邩�m�F
        �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML
        aaa = InStr(�f�[�^, "�x�a�l�R�[�h")
        aaa�ȉ� = Mid(�f�[�^, aaa + 1, Len(�f�[�^) - aaa)
        bbb = InStr(aaa�ȉ�, ">")
        bbb�ȉ� = Mid(aaa�ȉ�, bbb + 1, Len(aaa�ȉ�) - bbb)
        ccc = InStr(bbb�ȉ�, ">")
        ccc�ȉ� = Mid(bbb�ȉ�, ccc + 1, Len(bbb�ȉ�) - ccc)
        zzz = InStr(ccc�ȉ�, "<")
        �\�����ꂽ���i�i�� = Left(ccc�ȉ�, zzz - 1)
        '�\�����ꂽ���i�i�� = ObjIE.Document.all.Item("ctl00_txtYbm").Value
        '�\�����ꂽ���i�i�� = Replace(�\�����ꂽ���i�i��, "%", "")
        '�\�����ꂽ���i�i�� = Replace(�\�����ꂽ���i�i��, "-", "")
        
        If Replace(�\�����ꂽ���i�i��, "-", "") <> Replace(���i�i��, "-", "") Then
            '���������i�Ԃƕ\�����ꂽ�i�Ԃ̏ƍ�
            If Replace(�\�����ꂽ���i�i��, "-", "") Like "*" & Replace(���i�i��, "-", "") Then
                �������� = "Found"
            Else
                Stop '���������i�Ԃƕ\�����ꂽ�i�Ԃ̌㔼���قȂ�
            End If
        Else
                �������� = "Found"
        End If
    
End Function

Public Function �{�^���N���b�N(ByRef objIE As Object, buttonValue As String) '�s�v����
    Dim objInput As Object
    
    For Each objInput In objIE.document.getElementsByTagName("input")
        If objInput.Value = buttonValue Then
            objInput.Click
            Exit For
        End If
    Next
End Function

Public Function ��ʏ��擾a(ByVal objIE As Object) '�s�v����

Dim ���s�� As Long

    'ObjIE.document.getElementsByName("q")(0).Value = "������"
  For Each obj In objIE.document.all  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    With Sheets("���O")
        nextGyo = .Range("a" & .Rows.count).End(xlUp).Row + 1
        �l = obj.innertext
        'Call ���s�̉񐔂𒲂ׂ�(�l, ���s��)
        'For a = 1 To ���s��
        .Range("a" & nextGyo) = �l
        .Range("b" & nextGyo) = "ID=" & obj.iD
        'Next a
    End With
  Next
  
End Function

Public Function ��ʏ��擾(ByVal objIE As Object) '�s�v����

    'ObjIE.document.getElementsByName("q")(0).Value = "������"
  For Each obj In objIE.document.getElementsByTagName("a")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    Sheets("���O").Range("a" & Sheets("���O").Range("a" & Rows.count).End(xlUp).Row + 1) = "a_innertext=" & obj.innertext & "  " & "ID=" & obj.iD           '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
  Next
  
  For Each obj In objIE.document.getElementsByTagName("input")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    Sheets("���O").Range("a" & Sheets("���O").Range("a" & Rows.count).End(xlUp).Row + 1) = "input_innertext=" & obj.innertext & "  " & "ID=" & obj.iD           '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
  Next
  
  For Each obj In objIE.document.getElementsByTagName("btn")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    Sheets("���O").Range("a" & Sheets("���O").Range("a" & Rows.count).End(xlUp).Row + 1) = "btn_innertext=" & obj.innertext & "  " & "ID=" & obj.iD & " " & obj.Name         '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
  Next

End Function

Sub IE_open_sample() '�Q�l
  
  j = 0
  
  Set objIE = CreateObject("InternetExplorer.Application")  'IE���J���ۂ̂���
  objIE.Visible = True                                      'IE���J���ۂ̂���
  objIE.Navigate "http://www.yahoo.co.jp/"                  '�J�������T�C�g��URL���w��
  
  Do While objIE.readystate <> 4                            '�T�C�g���J�����܂ő҂i���񑩁j
    Do While objIE.busy = True                              '�T�C�g���J�����܂ő҂i���񑩁j
    Loop
  Loop
  
  For Each obj In objIE.document.getElementsByTagName("a")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    If obj.innertext = "�t�@�C�i���X" Then                  '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
      obj.Click                                             '��L�ɊY������^�O���N���b�N
      Exit For                                              '��L������AFor Each�@�`�@Next�𔲂���
    End If
  Next                                                      '���̃^�O������

  Sleep (1000)                                              '1�b�҂�
  
  Do While objIE.readystate <> 4                            '�T�C�g���J�����܂ő҂i���񑩁j
    Do While objIE.busy = True                              '�T�C�g���J�����܂ő҂i���񑩁j
    
    Loop
  Loop
  
  For Each obj In objIE.document.getElementsByTagName("input")  '�\������Ă���T�C�g��input�^�O�����ϐ�obj�ɃZ�b�g
                                                                '�einput�^�O�P�ʂɈȉ��̏��������{
    If obj.iD = "searchText" Then                           '�^�O��id�����usearchText�v�̏ꍇ�A�ȉ��̏��������{
      obj.Value = "�C�V��"                                  '�e�L�X�g�{�b�N�X�Ɂu�C�V���v��}��
    Else
      If obj.iD = "searchButton" Then                       '�^�O��id�����usearchButton�v�̏ꍇ�A�ȉ��̏��������{
        obj.Click                                           '�Y����input�^�O���N���b�N
        Exit For                                            '��L������AFor Each�@�`�@Next�𔲂���
      End If
    End If
  Next                                                      '���̃^�O������

End Sub

Public Function ����a(ByVal objIE As Object, ��������, �G�������g)

    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(�G�������g).innerHTML 'PTM���
    On Error GoTo 0
    aa = ����(�f�[�^, ��������, 1)
    If aa = 0 Then Exit Function
    �f�[�^a = Mid(�f�[�^, aa)
    bb = ����(�f�[�^a, "<", 3)
    �f�[�^b = Left(�f�[�^a, bb - 1)
    cc = InStrRev(�f�[�^b, ">")
    ����a = Mid(�f�[�^b, cc + 1)
    ����a = Replace(����a, "&nbsp;", "")
      
End Function
Public Function ����(�\�[�X, ��������, �q�b�g��)
    Dim myCount As Long
    For i = 1 To Len(�\�[�X)
        If �������� = Mid(�\�[�X, i, Len(��������)) Then
            myCount = myCount + 1
            If �q�b�g�� = myCount Then
                ���� = i
                Exit Function
            End If
        End If
    Next i
    
End Function
