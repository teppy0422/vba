Attribute VB_Name = "M10_WEB_���ތ���"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'API_�摜���_�E�����[�h
Public Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    
Dim ���ޏڍ�_�^�C�g��Ran As Range

Function ieVerCheck() As Integer

  Set objIEA = CreateObject("InternetExplorer.Application")
  Set objSFO = CreateObject("Scripting.FileSystemObject")

  ieVerCheck = val(objSFO.GetFileVersion(objIEA.FullName))

  Set objIEA = Nothing
  Set objSFO = Nothing

End Function

Public Sub open_dsw()
Attribute open_dsw.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim �����\��(1 To 20) As String
    Dim iD As String
    Dim myRyakuDir As String

    addressSet ThisWorkbook

    Dim gyo As Long: gyo = 10
    
    With wb(0).Sheets("WEB")
        �A�J�E���g = .Range("c" & gyo)
        �A�J�E���gID = .Range("d" & gyo)
        �p�X = .Range("e" & gyo)
        �p�XID = .Range("f" & gyo)
        ���O�C��btn = .Range("g" & gyo)
        �A�h���Xstr = .Range("h" & gyo)
        �E�B���h�E�� = .Range("i" & gyo)
        �u���E�U = .Range("j" & gyo)
    End With
'
'    With Sheets("A0_���ޏڍ�")
'        Dim ���ޏڍ�_�^�C�g��Row As Long: ���ޏڍ�_�^�C�g��Row = .Cells.Find("���i�i��_").Row
'        Set ���ޏڍ�_�^�C�g��Ran = .Range(.Cells(���ޏڍ�_�^�C�g��Row, 1), .Cells(���ޏڍ�_�^�C�g��Row, .Columns.count))
'
'        �^�C�g������ = "��������_,���i���_,���i����_,���́E�i��_,�F_,�o�^�H��_,�d��_,�d�オ��O�a_,���}��,���}URL,�N�����v�^�C�v_,�`���[�u�i��_,�`���[�u���a_,�R�l�N�^�ɐ�_,���i����_,�����\��01,�敪_,���i�i��_,���l_,��������_,�R�l�N�^�F_,�h���敪_,���b�N�ʒu���@_,���b�N�����敪_,�[�q��̌^�敪_,���b�L�敪_,�t�@�~���[_,�I�X���X_,�`���[�u�O�a_,�`���[�u����_"
'        �^�C�g������s = Split(�^�C�g������, ",")
'        '�����^�C�g��������������Βǉ�
'        Dim addCol As Long, checkTitle As Variant, x As Long
'        For x = LBound(�^�C�g������s) To UBound(�^�C�g������s)
'            Set checkTitle = .Cells.Find(�^�C�g������s(x), , , 1)
'            If checkTitle Is Nothing Then
'                addCol = .Cells(���ޏڍ�_�^�C�g��Row, .Columns.count).End(xlToLeft).Column + 1
'                .Cells(���ޏڍ�_�^�C�g��Row, addCol).Value = �^�C�g������s(x)
'            End If
'        Next x
'
'        Dim myCol() As Long
'        ReDim myCol(UBound(�^�C�g������s))
'        For i = LBound(�^�C�g������s) To UBound(�^�C�g������s)
'            myCol(i) = ���ޏڍ�_�^�C�g��Ran.Find(�^�C�g������s(i), , , 1).Column
'        Next i
'    End With
    
    '���}�̃_�E�����[�h�p�̃t�H���_
    If Dir(myRyakuDir, vbDirectory) = "" Then MkDir myRyakuDir
    'IE�̋N��
    Dim objIE As Object '�ϐ����`���܂��B
    Dim ieVerCheck As Variant

    Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
    Set objSFO = CreateObject("Scripting.FileSystemObject")

    ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
    
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10�ȍ~(���Ԃ�)
    End If
    
    objIE.Visible = True      '���ATrue�Ō�����悤�ɂ��܂��B
    
    '�����������y�[�W��\�����܂��B
   objIE.Navigate �A�h���Xstr
   Call �y�[�W�\����҂�(objIE)
  
   '���1 ���O�C�����
   objIE.document.all.Item(�A�J�E���gID).Value = �A�J�E���g
   objIE.document.all.Item(�p�XID).Value = �p�X
   objIE.document.all.Item("btnLogin").Click '���O�C���N���b�N
   Call �y�[�W�\����҂�(objIE)
   '���2 �g�p���ӏ��
   objIE.document.all.Item("btnOK").Click 'OK�N���b�N
   Call �y�[�W�\����҂�(objIE)
   '���3 ���C���y�[�W
   objIE.document.all.Item("btnYzk").Click '���i�Ԃ���̌���
   Call �y�[�W�\����҂�(objIE)
'loop
    
    Set ws(0) = wb(0).ActiveSheet
   With ws(0)
         ���i�i�� = .Cells(ActiveCell.Row, .Cells.Find("���i�i��", , , 1).Column).Value
'        lastgyo = .Cells(.Rows.count, myCol(17)).End(xlUp).Row
'        For i = 6 To lastgyo
'            If .Cells(i, myCol(19)) = "" Then
'                �敪 = .Cells(i, myCol(16))
'                If Len(�敪) = 1 Then
'                    ���i�i�� = .Cells(i, myCol(17))
                    '�i�ԓ���
                    objIE.document.all.Item("ctl00_txtYbm").Value = "%" & ���i�i�� & "%"
                    Call �y�[�W�\����҂�(objIE)
                    '�����N���b�N
                    objIE.document.all.Item("ctl00_btnSearch").Click
                    Call �y�[�W�\����҂�(objIE)
                    '�i�ԏ��̎擾
                    Call a�擾_��������(objIE, ��������, ���i�i��)
                    
                    If �������� = "Not Found." Then
'                        .Cells(i, myCol(19)) = "NotFound"
                    ElseIf �������� = "NotMatch" Then
'                        .Cells(i, myCol(19)) = "NotMatch"
                    Else
                        'PTM
                        ���i��� = ����a(objIE, "���i���", "ctl01_grdPtmCommn")
                        ���i���� = ����a(objIE, "���i����", "ctl01_grdPtmCommn")
                        ���i���� = ����a(objIE, "���i����", "ctl01_grdPtmCommn")
                        �o�^�H�� = ����a(objIE, "�o�^�H��", "ctl01_grdPtmCommn")
                        'JAIRS
                        ���̕i�� = ����a(objIE, "����", "ctl01_grdEmtrCommon")
                        If ���̕i�� = "" Then ���̕i�� = ����a(objIE, "�i��", "ctl01_grdJairsCommon")
                        
                        ���i�F = ����a(objIE, "�F", "ctl01_grdJairsSpecs")
                        �t�@�~���[ = ����a(objIE, "�t�@�~���[", "ctl01_grdJairsSpecs")
                        �I�X���X = ����a(objIE, "�I�X/���X", "ctl01_grdJairsSpecs")
                        'JAIRS�d�l
                        �d�� = ����a(objIE, "�d��", "ctl01_grdJairsSize")
                        '�^�C�v = ����a(objIE, "�^�C�v", "ctl01_grdJairsSpecs")
                        Call a�擾_�����\��(objIE, �����\��)
                        '���}
                        Call a�擾_���}(objIE, ���}URL, ���}��)
                        '�P���d��
                        �d�オ��O�a = ����a(objIE, "�d�オ��O�a", "ctl01_grdPtmIndivs")
                        '�N�����v�^�C�v
                        �N�����v�^�C�v = ����a(objIE, "�N�����v�^�C�v", "ctl01_grdPtmIndivs")
                        '�`���[�u
                        �`���[�u�i�� = ����a(objIE, "�`���[�u�i��", "ctl01_grdPtmIndivs")
                        �`���[�u���� = ����a(objIE, "�`���[�u����", "ctl01_grdPtmIndivs")
                        �`���[�u���a = ����a(objIE, "�`���[�u���a", "ctl01_grdPtmIndivs")
                        �`���[�u�O�a = ����a(objIE, "�`���[�u�O�a", "ctl01_grdPtmIndivs")
                        '�R�l�N�^
                        �R�l�N�^�ɐ� = ����a(objIE, "�R�l�N�^�ɐ�", "ctl01_grdPtmIndivs")
                        �R�l�N�^�F = ����a(objIE, "�R�l�N�^�F", "ctl01_grdPtmIndivs")
                        �R�l�N�^�h���敪 = ����a(objIE, "�h���敪", "ctl01_grdPtmIndivs")
                        ���b�L�敪 = ����a(objIE, "���b�L�敪", "ctl01_grdPtmIndivs")
                        ���b�N�ʒu���@ = ����a(objIE, "���b�N�ʒu���@", "ctl01_grdPtmIndivs")
                        ���b�N�����敪 = ����a(objIE, "���b�N�����敪", "ctl01_grdPtmIndivs")
                        �[�q��̌^�敪 = ����a(objIE, "�[�q��̌^�敪", "ctl01_grdPtmIndivs")
                        
                        '���Ӑ�i��
                        iD = "ctl01_grdJairsCustomers"
                        'Call a�擾_���Ӑ�i��(objIE, iD, i)
                        '���}
                        iD = "ctl01_dispRyaku_btnDraw"
                        
'                        Call a�擾_���}�_�E�����[�h(objIE, �A�h���X(0) & "\202_���}", ���i�i��, �A�h���Xstr) '���ɍ��W�𒲂ׂ��}���ύX���ꂽ��ēx���W�𒲂ׂ�K�v������̂ňꎞ�I�ɃR�����g�s
'                        Call a�擾_���}�_�E�����[�h(objIE, �A�h���X(1) & "\202_���}", ���i�i��, �A�h���Xstr) '���ɍ��W�𒲂ׂ��}���ύX���ꂽ��ēx���W�𒲂ׂ�K�v������̂ňꎞ�I�ɃR�����g�s
                        
'                        .Cells(i, myCol(0)).Value = ��������
'                        .Cells(i, myCol(1)).Value = Replace(���i���, "&nbsp;", " ")
'                        .Cells(i, myCol(2)).Value = Replace(���i����, "&nbsp;", " ")
'                        .Cells(i, myCol(3)).Value = Replace(���̕i��, "&nbsp;", " ")
'                        .Cells(i, myCol(4)).Value = Replace(���i�F, "&nbsp;", " ")
'                        .Cells(i, myCol(5)).Value = Replace(�o�^�H��, "&nbsp;", " ")
'                        .Cells(i, myCol(6)).Value = Replace(�d��, "&nbsp;", " ")
'
'                        .Cells(i, myCol(7)).Value = �d�オ��O�a
'                        .Cells(i, myCol(8)).Value = ���}��
'                        .Cells(i, myCol(9)).Value = ���}URL
'
'                        .Cells(i, myCol(10)).Value = �N�����v�^�C�v
'                        '�`���[�u
'                        .Cells(i, myCol(11)).Value = �`���[�u�i��
'                        .Cells(i, myCol(12)).Value = �`���[�u���a
'                        .Cells(i, myCol(13)).Value = �R�l�N�^�ɐ�
'                        .Cells(i, myCol(14)).Value = ���i����
'                        For x = 1 To 20
'                            .Cells(i, 54 + myCol(15)).Value = �����\��(x)
'                        Next x
'                        .Cells(i, myCol(20)).Value = �R�l�N�^�F
'                        .Cells(i, myCol(21)).Value = �R�l�N�^�h���敪
'                        .Cells(i, myCol(22)).Value = ���b�N�ʒu���@
'                        .Cells(i, myCol(23)).Value = ���b�N�����敪
'                        .Cells(i, myCol(24)).Value = �[�q��̌^�敪
'                        .Cells(i, myCol(25)).Value = ���b�L�敪
'                        .Cells(i, myCol(26)).Value = �t�@�~���[
'                        .Cells(i, myCol(27)).Value = �I�X���X
'
'                        .Cells(i, myCol(28)).Value = �`���[�u�O�a
'                        .Cells(i, myCol(29)).Value = �`���[�u����
                        
                    End If
'                    ActiveWindow.ScrollRow = i
'                    Sleep 1000
'                End If
'            End If
'        Next i
   End With

   Set objIE = Nothing
   Set objSFO = Nothing
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    �R�l�N�^�ɐ� = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a�擾_���}�_�E�����[�h(ByVal objIE As Object, myRyakuDir, ���i�i��, �A�h���X)
    �������� = "<IMG SRC="
    For a = 0 To 1
        '���}�̃{�^��id��������Ώ������Ȃ�
        'Stop '���Ƃ���Win7�Ȃ�0�AWin10�Ȃ�2�BWin7�ł�2�œ���m�F�ς�
        If InStr(objIE.document.all(2).outerHTML, "ctl01_dispRyaku_btnDraw") = 0 Then Exit Function
        objIE.document.all.Item("ctl01_dispRyaku_edtText").Value = ���i�i��
        objIE.document.all.Item("ctl01_dispRyaku_rgpReverse_" & a).Click      '0=���ʎ� 1=���ʎ�
        objIE.document.all.Item("ctl01_dispRyaku_cmbText")(3).Selected = True '�e�L�X�g����
        objIE.document.all.Item("ctl01_dispRyaku_chkOriginalSize").Checked = True     '�`��
        objIE.document.all.Item("ctl01_dispRyaku_btnDraw").Click              '�`��
        
        Call �y�[�W�\����҂�(objIE)
        For x = 0 To objIE.document.all.tags("img").Length - 1  '�v�f�̐�
            �f�[�^ = objIE.document.all.tags("img")(x).outerHTML
            aaa = InStr(StrConv(�f�[�^, vbUpperCase), ��������)
            If aaa = 0 Then GoTo line0
            ���}URL = Left(�A�h���X, InStrRev(�A�h���X, "/") - 1) & Mid(�f�[�^, Len(��������) + 3)
            
            ���}URL = Left(���}URL, Len(���}URL) - 2)
            ���}�ۑ�PASS = myRyakuDir & "\" & ���i�i�� & "_" & a & "_" & Format(x, "000") & ".emf"
            '�_�E�����[�h�̎��s
            Ret = URLDownloadToFile(0, ���}URL, ���}�ۑ�PASS, 0, 0)
line0:
        Next x
    Next a
End Function
Public Function a�擾_���Ӑ�i��(ByVal objIE As Object, iD, ByVal i As Long)
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(iD).innerHTML
    On Error GoTo 0
    Dim ii As Long
    Dim �^�C�g��AddCol As Long
    For i = 1 To Len(�f�[�^)
        If Mid(�f�[�^, i, 3) = "<TD" Then
            ���Ӑ於 = "": ���Ӑ�i�� = ""
            flg = False: flg1 = False: flg2 = False: flg3 = False: flg4 = False
            For ii = i + 1 To Len(�f�[�^)
                If Mid(�f�[�^, ii, 1) = "<" Then
                    flg = False
                    flg1 = True
                End If
                If flg1 = False Then
                    If flg = True Then
                        ���Ӑ於 = ���Ӑ於 & Mid(�f�[�^, ii, 1)
                    End If
                    
                    If Mid(�f�[�^, ii, 1) = ">" Then flg = True
                End If
                If flg1 = True Then
                   
                    If flg2 = True Then
                        If Mid(�f�[�^, ii, 1) = "<" Then
                            i = ii
                            flg4 = True
                            Exit For
                        End If
                        
                        If flg3 = True Then
                            ���Ӑ�i�� = ���Ӑ�i�� & Mid(�f�[�^, ii, 1)
                        End If
                        If Mid(�f�[�^, ii, 1) = ">" Then flg3 = True
                    End If
                    If Mid(�f�[�^, ii, 3) = "<TD" Then flg2 = True
                End If
            Next ii
        End If
        
        If flg4 = True Then
            ���Ӑ於 = Replace(���Ӑ於, "&nbsp;", "")
            ���Ӑ�i�� = Replace(���Ӑ�i��, "&nbsp;", "")
                '���ޏڍׂ���T���č��ڂ�������Βǉ�
            With Sheets("A0_���ޏڍ�")
                Set ���Ӑ於find = ���ޏڍ�_�^�C�g��Ran.Find(���Ӑ於 & "_", lookat:=xlWhole)
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
        End If
    Next i
    
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    �`���[�u�i�� = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a�擾_�N�����v�^�C�v(ByVal objIE As Object, �N�����v�^�C�v)
  �N�����v�^�C�v = ""
  
    �������� = "�N�����v�^�C�v"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmIndivs").outerText 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = InStr(1, �f�[�^, ��������)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = Mid(Ccc, Len(��������) + 1)
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
    Ccc = InStr(bbb + 1, �f�[�^, ";")
    ddd = InStr(Ccc + 1, �f�[�^, ";")
    eee = InStr(ddd + 1, �f�[�^, ">")
    zzz = InStr(eee + 1, �f�[�^, "<")
    �d�オ��O�a = Mid(�f�[�^, eee + 1, zzz - eee - 1)
      
End Function

Public Function �y�[�W�\����҂�(ByRef objIE As Object)

    While objIE.ReadyState <> 4 Or objIE.Busy = True '.ReadyState <> 4�̊Ԃ܂��B
        DoEvents  '�d���̂Ō����Ȑl���邯�ǁB
        Call ���z�L�[����(�V�t�g)
    Wend
    
End Function

Public Function a�擾_���}(ByVal objIE As Object, ���}URL, ���}��)
  ���}URL = "": ���}�� = 0
  
    ���}�� = objIE.document.Images.Length - 1
  
    For r = 1 To objIE.document.Images.Length - 1
        ���}URL = objIE.document.Images(1).src
    Next r
      
End Function

Public Function ����a(ByVal objIE As Object, ��������, �G�������g)
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById(�G�������g).innerHTML 'PTM���
    On Error GoTo 0
    aa = ����(�f�[�^, ��������, 1)
    If aa = 0 Then Exit Function
    �f�[�^a = Mid(�f�[�^, aa)
    bb = ����(�f�[�^a, "<", 3)
    �f�[�^b = Left(�f�[�^a, bb - 1)
    Cc = InStrRev(�f�[�^b, ">")
    ����a = Mid(�f�[�^b, Cc + 1)
    ����a = Replace(����a, "&nbsp;", "")
End Function

Public Function a�擾_���i����(ByVal objIE As Object, ���i����)
  ���i���� = ""
  
    �������� = "���i����"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    Ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(Ccc + 1, �f�[�^, "<")
    
    If aaa <> 0 Then ���i���� = Mid(�f�[�^, Ccc + 1, zzz - Ccc - 1)
      
End Function
Public Function a�擾_���i����(ByVal objIE As Object, ���i����)
  ���i���� = ""
  
    �������� = "���i����"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    Ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(Ccc + 1, �f�[�^, "<")
    
    If aaa <> 0 Then ���i���� = Mid(�f�[�^, Ccc + 1, zzz - Ccc - 1)
      
End Function
Public Function a�擾_�o�^�H��(ByVal objIE As Object, �o�^�H��)
  �o�^�H�� = ""
  
    �������� = "�o�^�H��"
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM���
    aaa = InStr(1, �f�[�^, ��������)
    bbb = InStr(aaa + Len(��������) + 1, �f�[�^, ">")
    Ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(Ccc + 1, �f�[�^, "<")
        
    If aaa <> 0 Then �o�^�H�� = Mid(�f�[�^, Ccc + 1, zzz - Ccc - 1)
      
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
    Ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(Ccc + 1, �f�[�^, "<")
        
    If aaa <> 0 Then ���̕i�� = Mid(�f�[�^, Ccc + 1, zzz - Ccc - 1)
      
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
    Ccc = InStr(bbb + 1, �f�[�^, ">")
    zzz = InStr(Ccc + 1, �f�[�^, "<")
        
    If aaa <> 0 Then ���i�F = Mid(�f�[�^, Ccc + 1, zzz - Ccc - 1)
      
End Function

Public Function a�擾_�d��(ByVal objIE As Object, �d��)
  �d�� = ""
  
    �������� = "�d��"
    On Error Resume Next
    �f�[�^ = objIE.document.getElementById("ctl01_grdJairsSize").innerHTML 'JAIRS�K�p�T�C�Y
    On Error GoTo 0
    If �f�[�^ = "" Then Exit Function
    aaa = ����(�f�[�^, ��������, 1)
    If aaa = 0 Then Exit Function
    bbb = Mid(�f�[�^, aaa)
    Ccc = ����(bbb, "<", 3)
    ddd = Left(bbb, Ccc - 1)
    eee = InStrRev(ddd, ">")
    �d�� = Mid(ddd, eee + 1)
      
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
        Ccc = InStr(bbb, "target")
        If Ccc = 0 Then �����\��(i) = "": GoTo line10
        ddd = Mid(bbb, Ccc, Len(bbb))
        eee = InStr(ddd, ">")
        fff = InStr(ddd, "<")
        ggg = Mid(ddd, eee + 1, fff - eee - 1)
        �����\��(i) = ggg
        
        bbb = Mid(bbb, Ccc + fff, Len(bbb))
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
    
    '���X�g�̌������ʐ����m�F
    Dim myCount As Long: myCount = 0
    �f�[�^ = objIE.document.getElementById("ctl00_grdList").innerHTML '�N���w�肵��ID����ForEach�ŎQ�Ƃ�����@�����Ă�������
    �f�[�^sp = Split(�f�[�^, vbCrLf)
    For i = LBound(�f�[�^sp) To UBound(�f�[�^sp)
        Debug.Print �f�[�^sp(i)
        If InStr(�f�[�^sp(i), "javascript") > 0 Then
            myCount = myCount + 1
        End If
    Next i
    
    '�������ʂ���������ꍇ
    If myCount > 1 Then
        For Each objLink In objIE.document.getElementsByTagName("A")
            Debug.Print objLink.innerText
            If objLink.innerText = Replace(���i�i��, "-", "") Then
                Debug.Print ���i�i��, objLink.innerText, objLink.href
                'Debug.Print objLink.href
                objIE.Navigate objLink.href
                Exit For
    '        If objLink.innerText = anchorText Then
    '            objIE.navigate objLink.href
    '            Exit For
    '        End If
            ElseIf objLink.innerText = "450" & Replace(���i�i��, "-", "") Then 'VS
                Debug.Print ���i�i��, objLink.innerText, objLink.href
                'Debug.Print objLink.href
                objIE.Navigate objLink.href
                Exit For
            End If
        Next
    End If

'    '�_������������ꍇ�A�����N���N���b�N
'    If �_�� > 0 Then
'        '�����Naaa = InStrRev(�f�[�^, ">" & Replace(���i�i��, "-", "") & "<")
'        '�����Nbbb = Left(�f�[�^, �����Naaa)
'        '�����Nccc = InStrRev(�����Nbbb, "grdList")
'        '�����N�A�h���X = Mid(�����Nbbb, �����Nccc, 9 + Len(�_��))
'        'objIe.document.all.Item("javascript:__doPostBack('ctl00$grdList','grdList$0')").Click
'
'        '�����N�ԍ��ŊJ��(�_��+4�Ō�������ׁA�m���ł͂Ȃ�����)
'        �����Naaa = InStrRev(�f�[�^, ">" & Replace(���i�i��, "-", "") & "<")
'        If �����Naaa <> 0 Then
'            �����Nbbb = left(�f�[�^, �����Naaa)
'            �����Nccc = InStrRev(�����Nbbb, "$")
'            �����Nzzz = InStrRev(�����Nbbb, "'")
'            �����N�ԍ� = Mid(�����Nbbb, �����Nccc + 1, �����Nzzz - (�����Nccc + 1))
'        Else
'            �������� = "NotMatch"
'        End If
'
'        objIE.document.Links(4).Click
'
'    End If
    
    Call �y�[�W�\����҂�(objIE)

    '�\�����ꂽ�i�Ԃƌ����������i�Ԃ��}�b�`���邩�m�F
    �f�[�^ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML
        
    aa = ����(�f�[�^, "�x�a�l�R�[�h", 1)
    If aa = 0 Then Exit Function
    �f�[�^a = Mid(�f�[�^, aa)
    bb = ����(�f�[�^a, "<", 3)
    �f�[�^b = Left(�f�[�^a, bb - 1)
    Cc = InStrRev(�f�[�^b, ">")
    v = Mid(�f�[�^b, Cc + 1)
    �\�����ꂽ���i�i�� = Replace(v, "&nbsp;", "")
        
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
        Debug.Print objInput.Value
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
        �l = obj.innerText
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
    Sheets("���O").Range("a" & Sheets("���O").Range("a" & Rows.count).End(xlUp).Row + 1) = "a_innertext=" & obj.innerText & "  " & "ID=" & obj.iD           '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
  Next
  
  For Each obj In objIE.document.getElementsByTagName("input")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    Sheets("���O").Range("a" & Sheets("���O").Range("a" & Rows.count).End(xlUp).Row + 1) = "input_innertext=" & obj.innerText & "  " & "ID=" & obj.iD           '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
  Next
  
  For Each obj In objIE.document.getElementsByTagName("btn")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    Sheets("���O").Range("a" & Sheets("���O").Range("a" & Rows.count).End(xlUp).Row + 1) = "btn_innertext=" & obj.innerText & "  " & "ID=" & obj.iD & " " & obj.Name         '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
  Next

End Function

Sub IE_open_sample() '�Q�l
  
  j = 0
  
  Set objIE = CreateObject("InternetExplorer.Application")  'IE���J���ۂ̂���
  objIE.Visible = True                                      'IE���J���ۂ̂���
  objIE.Navigate "http://www.yahoo.co.jp/"                  '�J�������T�C�g��URL���w��
  
  Do While objIE.ReadyState <> 4                            '�T�C�g���J�����܂ő҂i���񑩁j
    Do While objIE.Busy = True                              '�T�C�g���J�����܂ő҂i���񑩁j
    Loop
  Loop
  
  For Each obj In objIE.document.getElementsByTagName("a")  '�\������Ă���T�C�g�̃A���J�[�^�O�����ϐ�obj�ɃZ�b�g
                                                            '�e�A���J�[�^�O�P�ʂɈȉ��̏��������{
    If obj.innerText = "�t�@�C�i���X" Then                  '�A���J�[�^�O�̕\�����e���u�t�@�C�i���X�v�̏ꍇ�Ɉȉ��̏��������{
      obj.Click                                             '��L�ɊY������^�O���N���b�N
      Exit For                                              '��L������AFor Each�@�`�@Next�𔲂���
    End If
  Next                                                      '���̃^�O������

  Sleep (1000)                                              '1�b�҂�
  
  Do While objIE.ReadyState <> 4                            '�T�C�g���J�����܂ő҂i���񑩁j
    Do While objIE.Busy = True                              '�T�C�g���J�����܂ő҂i���񑩁j
    
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

Sub �G�N�X�|�[�g_���ޏڍ�1273()
    Dim Timer2 As Single
    Timer2 = Timer
'    Dim �ďo��flg As Boolean
'    �ďo��flg = True
    
    Set wb(0) = ActiveWorkbook
    Call addressSet(wb(0))

    With wb(0).Sheets("A0_���ޏڍ�")
        'myRan�Ƀ��[�h�Ɨ�ԍ��Əo�͍s�̒l����āA�s���ɒl��ς��Ȃ���e�L�X�g�o�͂��Ă݂�myRan(,2)�ɍs���̒l����� ��VBA��join�͓񎟔z��Ŏ��s�ł��Ȃ������߂�
        Dim myWords As String, myWords2 As String, myWords2Col2 As Long, myWordsSP
        myWords = "���i�i��_,��������_,�R�l�N�^�ɐ�_,���i���_,���i����_,�N�����v�^�C�v_,���l_,�h���敪_,�F_,���b�L�敪_,�t�@�~���[_,�I�X���X_"
        myWordsSP = Split(myWords, ",")
        Dim myRan(), r As Long
        ReDim myRan(UBound(myWordsSP), 1)
        For r = LBound(myRan) To UBound(myRan)
            myRan(r, 0) = myWordsSP(r)                                             'myWordsSP������
            myRan(r, 1) = .Cells.Find(myWordsSP(r), , , 1).Column    '���ވꗗ�ł̗�ԍ�
            If myWordsSP(r) = "�F_" Then myWords2Col2 = r
        Next r
        '�R�l�N�^�F���u�����N�̏ꍇ�͐F���g�p����ׂ̒ǋL
        Dim myword2 As String, myWords2Col As Long, �R�l�N�^�F As String
        myWords2 = "�R�l�N�^�F_"
        myWords2Col = .Cells.Find(myWords2, , , 1).Column
        
        Dim lastRow As Long, key, i As Long, ��������str As String, �o��flg As Boolean
        Set key = .Cells.Find(myRan(0, 0), , , 1)
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For i = key.Row + 1 To lastRow
            '�o��flg = .Cells(i, myRan(0, 1)).Value: If �o��flg = True And �ďo��flg = False Then GoTo line20
            ��������str = .Cells(i, myRan(1, 1)).Value: If ��������str <> "Found" Then GoTo line20
            �R�l�N�^�F = .Cells(i, myWords2Col).Value
            'text�o�͂���
            Dim ���i�i��str As String: ���i�i��str = .Cells(i, myRan(0, 1)).Value
            Dim FSO As New FileSystemObject ' FileSystemObject
            Dim TS As TextStream            ' TextStream
            Dim strREC As String            ' �����o�����R�[�h���e
            Set TS = FSO.CreateTextFile(fileName:=myAddress(0) & "\300_���ޏڍ�\" & ���i�i��str & ".txt", overwrite:=True)
            Dim text1 As String, myValue As String
            Dim IntFlNo As Integer: IntFlNo = FreeFile
            TS.WriteLine Join(myWordsSP, ",")
            For r = LBound(myRan) To UBound(myRan)  'myWords��0��1�̗v�f�͖�������
                If myWords2Col2 = r Then
                    'myValue = PTMorJCMP(�R�l�N�^�F, .Cells(i, myRan(r, 1)))
                Else
                    myValue = .Cells(i, myRan(r, 1))
                End If
                text1 = text1 & "," & myValue
            Next r
            text1 = Mid(text1, 2)
            TS.WriteLine text1
            TS.Close
            text1 = ""
            Set TS = Nothing
            Set FSO = Nothing
line20:
        Next i
        Debug.Print Round(Timer - Timer2, 1) & "s"
    End With
End Sub



Sub ie_test()

    Dim �����\��(1 To 20) As String
    Dim iD As String
    Dim myRyakuDir As String, gyo As Long

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Call addressSet(myBook)

    gyo = 10
    With Sheets("WEB")
        �A�J�E���g = .Range("c" & gyo)
        �A�J�E���gID = .Range("d" & gyo)
        �p�X = .Range("e" & gyo)
        �p�XID = .Range("f" & gyo)
        ���O�C��btn = .Range("g" & gyo)
        �A�h���Xstr = .Range("h" & gyo)
        �E�B���h�E�� = .Range("i" & gyo)
        �u���E�U = .Range("j" & gyo)
    End With
        
    With Sheets("A0_���ޏڍ�")
        Dim ���ޏڍ�_�^�C�g��Row As Long: ���ޏڍ�_�^�C�g��Row = .Cells.Find("���i�i��_").Row
        Set ���ޏڍ�_�^�C�g��Ran = .Range(.Cells(���ޏڍ�_�^�C�g��Row, 1), .Cells(���ޏڍ�_�^�C�g��Row, .Columns.count))
        
        �^�C�g������ = "��������_,���i���_,���i����_,���́E�i��_,�F_,�o�^�H��_,�d��_,�d�オ��O�a_,���}��,���}URL,�N�����v�^�C�v_,�`���[�u�i��_,�`���[�u���a�~�O�a-����_,�R�l�N�^�ɐ�_,���i����_,�����\��01,�敪_,���i�i��_,���l_,��������_,�R�l�N�^�F_,�h���敪_,���b�N�ʒu���@_,���b�N�����敪_,�[�q��̌^�敪_,���b�L�敪_,�t�@�~���[_,�I�X���X_"
        
        �^�C�g������s = Split(�^�C�g������, ",")
        Dim myCol() As Long
        ReDim myCol(UBound(�^�C�g������s))
        For i = LBound(�^�C�g������s) To UBound(�^�C�g������s)
            myCol(i) = ���ޏڍ�_�^�C�g��Ran.Find(�^�C�g������s(i), , , 1).Column
        Next i
    End With
    
    '���}�̃_�E�����[�h�p�̃t�H���_
    If Dir(myRyakuDir, vbDirectory) = "" Then MkDir myRyakuDir
    'IE�̋N��
    Dim objIE As Object '�ϐ����`���܂��B
    Dim ieVerCheck As Variant

    Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
    Set objSFO = CreateObject("Scripting.FileSystemObject")

    ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
    
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10�ȍ~(���Ԃ�)
    End If
    
    objIE.Visible = True      '���ATrue�Ō�����悤�ɂ��܂��B
    
    '�����������y�[�W��\�����܂��B
   objIE.Navigate �A�h���Xstr
   Call �y�[�W�\����҂�(objIE)
  
   '���1 ���O�C�����
   objIE.document.all.Item(�A�J�E���gID).Value = �A�J�E���g
   objIE.document.all.Item(�p�XID).Value = �p�X
   objIE.document.all.Item("btnLogin").Click '���O�C���N���b�N
   Call �y�[�W�\����҂�(objIE)
   '���2 �g�p���ӏ��
   objIE.document.all.Item("btnOK").Click 'OK�N���b�N
   Call �y�[�W�\����҂�(objIE)
   '���3 ���C���y�[�W
   objIE.document.all.Item("btnYzk").Click '���i�Ԃ���̌���
   Call �y�[�W�\����҂�(objIE)
'loop
   With Sheets("A0_���ޏڍ�")
        lastgyo = .Cells(.Rows.count, myCol(17)).End(xlUp).Row
        For i = 6 To lastgyo
            If .Cells(i, myCol(19)) = "" Then
                �敪 = .Cells(i, myCol(16))
                If Len(�敪) = 1 Then
                    ���i�i�� = .Cells(i, myCol(17))
                    '�i�ԓ���
                    objIE.document.all.Item("ctl00_txtYbm").Value = "%" & ���i�i�� & "%"
                    Call �y�[�W�\����҂�(objIE)
                    '�����N���b�N
                    objIE.document.all.Item("ctl00_btnSearch").Click
                    Call �y�[�W�\����҂�(objIE)
                    '�i�ԏ��̎擾
                    Call a�擾_��������(objIE, ��������, ���i�i��)
                    
                    If �������� = "Not Found." Then
                        .Cells(i, myCol(19)) = "NotFound"
                    ElseIf �������� = "NotMatch" Then
                        .Cells(i, myCol(19)) = "NotMatch"
                    Else
                        'PTM
                        ���i��� = ����a(objIE, "���i���", "ctl01_grdPtmCommn")
                        ���i���� = ����a(objIE, "���i����", "ctl01_grdPtmCommn")
                        ���i���� = ����a(objIE, "���i����", "ctl01_grdPtmCommn")
                        �o�^�H�� = ����a(objIE, "�o�^�H��", "ctl01_grdPtmCommn")
                        'JAIRS
                        ���̕i�� = ����a(objIE, "����", "ctl01_grdEmtrCommon")
                        If ���̕i�� = "" Then ���̕i�� = ����a(objIE, "�i��", "ctl01_grdJairsCommon")
                        
                        ���i�F = ����a(objIE, "�F", "ctl01_grdJairsSpecs")
                        �t�@�~���[ = ����a(objIE, "�t�@�~���[", "ctl01_grdJairsSpecs")
                        �I�X���X = ����a(objIE, "�I�X/���X", "ctl01_grdJairsSpecs")
                        'JAIRS�d�l
                        �d�� = ����a(objIE, "�d��", "ctl01_grdJairsSize")
                        Call a�擾_�����\��(objIE, �����\��)
                        '���}
                        Call a�擾_���}(objIE, ���}URL, ���}��)
                        '�P���d��
                        �d�オ��O�a = ����a(objIE, "�d�オ��O�a", "ctl01_grdPtmIndivs")
                        '�N�����v�^�C�v
                        �N�����v�^�C�v = ����a(objIE, "�N�����v�^�C�v", "ctl01_grdPtmIndivs")
                        '�`���[�u
                        �`���[�u�i�� = ����a(objIE, "�`���[�u�i��", "ctl01_grdPtmIndivs")
                        �`���[�u���� = ����a(objIE, "�`���[�u����", "ctl01_grdPtmIndivs")
                        �`���[�u���a = ����a(objIE, "�`���[�u���a", "ctl01_grdPtmIndivs")
                        �`���[�u�O�a = ����a(objIE, "�`���[�u�O�a", "ctl01_grdPtmIndivs")
                        
                        '�R�l�N�^
                        �R�l�N�^�ɐ� = ����a(objIE, "�R�l�N�^�ɐ�", "ctl01_grdPtmIndivs")
                        �R�l�N�^�F = ����a(objIE, "�R�l�N�^�F", "ctl01_grdPtmIndivs")
                        �R�l�N�^�h���敪 = ����a(objIE, "�h���敪", "ctl01_grdPtmIndivs")
                        ���b�L�敪 = ����a(objIE, "���b�L�敪", "ctl01_grdPtmIndivs")
                        ���b�N�ʒu���@ = ����a(objIE, "���b�N�ʒu���@", "ctl01_grdPtmIndivs")
                        ���b�N�����敪 = ����a(objIE, "���b�N�����敪", "ctl01_grdPtmIndivs")
                        �[�q��̌^�敪 = ����a(objIE, "�[�q��̌^�敪", "ctl01_grdPtmIndivs")
                        
                        '���Ӑ�i��
                        iD = "ctl01_grdJairsCustomers"
                        Call a�擾_���Ӑ�i��(objIE, iD, i)
                        '���}
                        iD = "ctl01_dispRyaku_btnDraw"
                        
                        Call a�擾_���}�_�E�����[�h(objIE, myAddress(0) & "\202_���}", ���i�i��, �A�h���Xstr) '���ɍ��W�𒲂ׂ��}���ύX���ꂽ��ēx���W�𒲂ׂ�K�v������̂ňꎞ�I�ɃR�����g�s
                        Call a�擾_���}�_�E�����[�h(objIE, myAddress(1) & "\202_���}", ���i�i��, �A�h���Xstr) '���ɍ��W�𒲂ׂ��}���ύX���ꂽ��ēx���W�𒲂ׂ�K�v������̂ňꎞ�I�ɃR�����g�s
                        
                        
                        .Cells(i, myCol(0)).Value = ��������
                        .Cells(i, myCol(1)).Value = Replace(���i���, "&nbsp;", " ")
                        .Cells(i, myCol(2)).Value = Replace(���i����, "&nbsp;", " ")
                        .Cells(i, myCol(3)).Value = Replace(���̕i��, "&nbsp;", " ")
                        .Cells(i, myCol(4)).Value = Replace(���i�F, "&nbsp;", " ")
                        .Cells(i, myCol(5)).Value = Replace(�o�^�H��, "&nbsp;", " ")
                        .Cells(i, myCol(6)).Value = Replace(�d��, "&nbsp;", " ")
                        
                        .Cells(i, myCol(7)).Value = �d�オ��O�a
                        .Cells(i, myCol(8)).Value = ���}��
                        .Cells(i, myCol(9)).Value = ���}URL
                        
                        .Cells(i, myCol(10)).Value = �N�����v�^�C�v
                        '�`���[�u
                        .Cells(i, myCol(11)).Value = �`���[�u�i��
                        .Cells(i, myCol(12)).Value = �`���[�u���a & "�~" & �`���[�u�O�a & "-" & �`���[�u����
                        .Cells(i, myCol(13)).Value = �R�l�N�^�ɐ�
                        .Cells(i, myCol(14)).Value = ���i����
                        For x = 1 To 20
                            .Cells(i, 54 + myCol(15)).Value = �����\��(x)
                        Next x
                        .Cells(i, myCol(20)).Value = �R�l�N�^�F
                        .Cells(i, myCol(21)).Value = �R�l�N�^�h���敪
                        .Cells(i, myCol(22)).Value = ���b�N�ʒu���@
                        .Cells(i, myCol(23)).Value = ���b�N�����敪
                        .Cells(i, myCol(24)).Value = �[�q��̌^�敪
                        .Cells(i, myCol(25)).Value = ���b�L�敪
                        .Cells(i, myCol(26)).Value = �t�@�~���[
                        .Cells(i, myCol(27)).Value = �I�X���X
                    End If
                    ActiveWindow.ScrollRow = i
                End If
            End If
        Next i
   End With

   Set objIE = Nothing
   Set objSFO = Nothing
End Sub

Public Function dsw_open() As Variant
    
    addressSet ThisWorkbook
    
    Dim getStrings As String, getSplit As Variant, g() As Variant, i As Long
    getStrings = "����,�T�C�g��,�A�J�E���g,�A�J�E���gID,�p�X,�p�XID,���O�C��bt,�A�h���X"
    getSplit = Split(getStrings, ",")
    ReDim g(UBound(getSplit))
    For i = LBound(getSplit) To UBound(getSplit)
        g(i) = wb(0).Sheets("WEB").Cells.Find(getSplit(i), , , 1, , , 1).Offset(1, 0)
    Next i
    
    With Sheets("WEB")
        �A�J�E���g = g(2)
        �A�J�E���gID = g(3)
        �p�X = g(4)
        �p�XID = g(5)
        ���O�C��btn = g(6)
        �A�h���Xstr = g(7)
    End With
    
    'IE�̋N��
    Dim objIE As Object
    Dim ieVerCheck As Variant
    
    Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
    Set objSFO = CreateObject("Scripting.FileSystemObject")
    
    ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
    
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10�ȍ~(���Ԃ�)
    End If
    
    objIE.Visible = True      'True�Ō�����悤�ɂ���
    
    objIE.Navigate �A�h���Xstr
    Call �y�[�W�\����҂�(objIE)
    
    '���1 ���O�C�����
    objIE.document.all.Item(�A�J�E���gID).Value = �A�J�E���g
    objIE.document.all.Item(�p�XID).Value = �p�X
    objIE.document.all.Item("btnLogin").Click '���O�C���N���b�N
    Call �y�[�W�\����҂�(objIE)
    '���2 �g�p���ӏ��
    objIE.document.all.Item("btnOK").Click 'OK�N���b�N
    Call �y�[�W�\����҂�(objIE)
    '���3 ���C���y�[�W
    objIE.document.all.Item("btnYzk").Click '���i�Ԃ���̌���
    Call �y�[�W�\����҂�(objIE)
    
    Set dsw_open = objIE
    
    Set objIE = Nothing
    Set objSFO = Nothing
   
End Function

Public Function dsw_search(ByVal objIE As Object, ByVal searchWord As String) As Variant
    
    '�i�ԓ���
    objIE.document.all.Item("ctl00_txtYbm").Value = "%" & searchWord & "%"
    Call �y�[�W�\����҂�(objIE)
    '�����N���b�N
    objIE.document.all.Item("ctl00_btnSearch").Click
    Call �y�[�W�\����҂�(objIE)
    '�i�ԏ��̎擾
    Call a�擾_��������(objIE, ��������, searchWord)
                    
    If �������� = "Not Found." Then
        DSW = "False"
    Else
        Dim FieldStrings As String, i As Long, fieldStringSplit As Variant, a As Long
        FieldStrings = "���i�i��_,��������_,�R�l�N�^�ɐ�_,���i���_,���i����_,�N�����v�^�C�v_,���l_," & _
            "�h���敪_,�F_,���b�L�敪_,�t�@�~���[_,�I�X���X_,�`���[�u���a_,�`���[�u�O�a_,�`���[�u����_,���́E�i��_"
        fieldStringSplit = Split(FieldStrings, ",")
        a = UBound(fieldStringSplit)
        Dim myArray() As Variant
        ReDim myArray(a, 1)
        For i = LBound(myArray) To UBound(myArray)
            myArray(i, 0) = fieldStringSplit(i)
        Next i
        
        myArray(0, 1) = searchWord
        myArray(1, 1) = ��������
        myArray(6, 1) = "" '���l
        'PTM
        myArray(3, 1) = ����a(objIE, "���i���", "ctl01_grdPtmCommn")
        myArray(4, 1) = ����a(objIE, "���i����", "ctl01_grdPtmCommn")
        ���i���� = ����a(objIE, "���i����", "ctl01_grdPtmCommn")
        �o�^�H�� = ����a(objIE, "�o�^�H��", "ctl01_grdPtmCommn")
        'JAIRS
        ���̕i�� = ����a(objIE, "����", "ctl01_grdEmtrCommon")
        If ���̕i�� = "" Then ���̕i�� = ����a(objIE, "�i��", "ctl01_grdJairsCommon")
        myArray(15, 1) = ���̕i��
        
        ���i�F = ����a(objIE, "�F", "ctl01_grdJairsSpecs")
        myArray(10, 1) = ����a(objIE, "�t�@�~���[", "ctl01_grdJairsSpecs")
        myArray(11, 1) = ����a(objIE, "�I�X/���X", "ctl01_grdJairsSpecs")
        'JAIRS�d�l
        �d�� = ����a(objIE, "�d��", "ctl01_grdJairsSize")
        Call a�擾_�����\��(objIE, �����\��)
        '���}
        Call a�擾_���}(objIE, ���}URL, ���}��)
        '�P���d��
        �d�オ��O�a = ����a(objIE, "�d�オ��O�a", "ctl01_grdPtmIndivs")
        '�N�����v�^�C�v
        myArray(5, 1) = ����a(objIE, "�N�����v�^�C�v", "ctl01_grdPtmIndivs")
        '�`���[�u
        �`���[�u�i�� = ����a(objIE, "�`���[�u�i��", "ctl01_grdPtmIndivs")
        myArray(12, 1) = ����a(objIE, "�`���[�u���a", "ctl01_grdPtmIndivs")
        myArray(13, 1) = ����a(objIE, "�`���[�u�O�a", "ctl01_grdPtmIndivs")
        myArray(14, 1) = ����a(objIE, "�`���[�u����", "ctl01_grdPtmIndivs")
        
        '�R�l�N�^
        myArray(2, 1) = ����a(objIE, "�R�l�N�^�ɐ�", "ctl01_grdPtmIndivs")
        �R�l�N�^�F = ����a(objIE, "�R�l�N�^�F", "ctl01_grdPtmIndivs")
        �R�l�N�^�F = Mid(�R�l�N�^�F, 4)
        myArray(7, 1) = ����a(objIE, "�h���敪", "ctl01_grdPtmIndivs")
        If �R�l�N�^�F <> "" Then ���i�F = �R�l�N�^�F
        myArray(8, 1) = ���i�F
        myArray(9, 1) = ����a(objIE, "���b�L�敪", "ctl01_grdPtmIndivs")
        ���b�N�ʒu���@ = ����a(objIE, "���b�N�ʒu���@", "ctl01_grdPtmIndivs")
        ���b�N�����敪 = ����a(objIE, "���b�N�����敪", "ctl01_grdPtmIndivs")
        �[�q��̌^�敪 = ����a(objIE, "�[�q��̌^�敪", "ctl01_grdPtmIndivs")
        
        '���Ӑ�i��
'        iD = "ctl01_grdJairsCustomers"
'        Call a�擾_���Ӑ�i��(objIE, iD, i)
        '���}
        Dim �A�h���Xstr As String
        �A�h���Xstr = Left(objIE.locationurl, InStrRev(objIE.locationurl, "/"))
        iD = "ctl01_dispRyaku_btnDraw"
        Call a�擾_���}�_�E�����[�h(objIE, myAddress(1, 1) & "\202_���}", searchWord, �A�h���Xstr) '���ɍ��W�𒲂ׂ��}���ύX���ꂽ��ēx���W�𒲂ׂ�K�v������̂ňꎞ�I�ɃR�����g�s
        dsw_search = myArray
        
    End If

   Set objIE = Nothing

End Function





