Attribute VB_Name = "M10_�O���A�v������"
'�E�C���h�E�n���h���ɂ�鑼�A�v������
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hParent As Long, ByVal hChildAfter As Long, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As Long
        '�������ASC1�������n������x��
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
        '����������̂܂ܓn�����瑬�����ǁAClass�ɂ���Ă͎g�p�s��
Declare Function SendMessageStr Lib "user32.dll" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal msg As Long, _
        ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDlgItem Lib "user32" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long

'�萔_���
Public Const WM_SETFOCUS = &H7
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLEN = &HE
Public Const WM_ALT = &H12
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104 'ALT�Ƃ�����
Public Const WM_COMMAND = &H111&
Public Const WM_SYSCOMMAND = &H112
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203 '�_�u���N���b�N
Public Const WM_IME_CHAR = &H286     '�����R�[�h���M
Public Const WM_CLEAR = &H303
'WM_KEY*�Ƃ̈Ⴂ�͂悭������񂯂ǁA�������ł�����
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291
'�萔_���X�g�{�b�N�X
Public Const LB_GETTEXT = &H189     '������
Public Const LB_GETTEXTLEN = &H18A  '������
Public Const LB_GETCOUNT = &H18B    '�v�f��
Public Const LB_SETCURSEL = &H186   '�w�肵�����ڂ�I��
Public Const LB_SETTOPINDEX = &H197 '�w�肵�����ڂ����X�g�㕔�ɕ\��
'�萔_�R���{�{�b�N�X
Public Const CB_GETTEXT = &H148
Public Const CB_GETTEXTLEN = &H149
Public Const CB_GETCOUNT = &H146
Public Const CB_SETCURSEL = &H14E
Public Const CB_SETTOPINDEX = &H15C
Public Const CB_SELECTSTRING = &H14D '�g���邩�킩���
Public Const CB_SHOWDROPDOWN = &H14F '�h���b�v�_�E�����X�g�̕\��_0����_1�J��
'�萔_�{�^��
Public Const BN_CLICKED = 0&
Public Const BM_SETCHECK = &HF1      '0�O��_1�����
Public Const BM_CLICK = &HF5
Public Const BM_GETCHECK = &HF0      '���W�Ior�`�F�b�N�{�b�N�X�̏�Ԃ�m��
'�萔_���̑�
Public Const SC_CLOSE = &HF060
Public Const EM_SETSEL = &HB1
'�萔_���z�L�[�R�[�h
Public Const VK_MENU = &H12 'ALT
Public Const VK_RETURNE = &HD 'ENTER
'�ϐ�
Public �������x���V�X�e��Path As String
Public myHND(10) As String
Public myHNDtemp As String
Dim i As Long
Dim Index, Ret, Rep As Integer

Sub Control_YcEditor()
    ���i�i��str = "8216136D40     "
    �ݕ�str = "test"
    Set myBook = ActiveWorkbook
    
    'Symbol�f�[�^�̍쐬
    Call SQL_YcEditor_Symbol(RAN, myBook, ���i�i��str)
    Dim i As Long
    '�o�͐�Dir��������΍쐬
    Dim outPath(1) As String
    outPath(0) = myBook.Path & "\81_���ʌ���date_�Ȉ�"
    If Dir(outPath(0), vbDirectory) = "" Then MkDir outPath(0)
    outPath(1) = outPath(0) & "\" & Replace(���i�i��str, " ", "")
    If Dir(outPath(1), vbDirectory) = "" Then MkDir outPath(1)
    '�o�͐�book���쐬
    Set wb(3) = Workbooks.add
    Application.DisplayAlerts = False
    wb(3).SaveAs fileName:=outPath(1) & "\" & Replace(���i�i��str, " ", "") & "_" & Replace(�ݕ�str, " ", "")
    Application.DisplayAlerts = True
    '�o��sheet���쐬
    wb(3).Worksheets.add
    wb(3).Sheets(1).Name = "Symbol"
    wb(3).Sheets(2).Name = "WH"
    
    With wb(3).Sheets("Symbol")
        .Activate
        .Cells.NumberFormat = "@"
        .Columns(1).NumberFormat = 0
        .Cells.Font.Name = "�l�r �S�V�b�N"
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            For X = LBound(RAN, 1) To UBound(RAN, 1)
                .Cells(Y, X + 1) = RAN(X, Y)
            Next X
        Next Y
        '���ёւ�
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range(Rows(1), Rows(lastRow)).Sort key1:=Range("a1"), Order1:=xlAscending, Header:=xlNo
        Dim endPoint As Long
        endPoint = .Cells(lastRow, 1) + 200
        If endPoint > 1900 Then endPoint = 1900
    End With
    'WH�f�[�^�̍쐬
    Call SQL_YcEditor_WH(RAN, myBook, ���i�i��str)
    With wb(3).Sheets("WH")
        .Activate
        .Cells.NumberFormat = "@"
        .Cells.Font.Name = "�l�r �S�V�b�N"
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            For X = LBound(RAN, 1) To UBound(RAN, 1)
                .Cells(Y, X + 1) = RAN(X, Y)
            Next X
        Next Y
        '���ёւ�
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range(Rows(1), Rows(lastRow)).Sort key1:=Range("a1"), Order1:=xlAscending, Header:=xlNo
        Dim endKosei As Long
        endKosei = CLng(.Cells(lastRow, 1))
    End With
    '�{���_�[�A�W���C���g�A�V�[���h�h�������̉�H�����C��
    Stop
    
    'YcEditor�ɏo��
    With myBook.Sheets("�ݒ�")
        For i = 0 To 10
            AppPath = .Cells.Find("YcEditor_exe", , , 1).Offset(i, 1)
            If Dir(AppPath) <> "" Then Exit For
        Next i
    End With
    '�����J�n
    Call Control_�A�v���N��(AppPath)
    '�n���h���擾
line00:
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    '�t�@�C�����J��
    SetForegroundWindow myHND(0) '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
    PostMessage myHND(0), WM_SYSKEYDOWN, VK_MENU, &H20380001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("F"), &H20210001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("O"), &H20180001
    '�V�K�쐬���N���b�N
    If myHND(0) = 0 Then GoTo line00
line01:
    myHND(1) = FindWindow("TfrmFile", "�t�@�C���I��")
    If myHND(1) = 0 Then GoTo line01
    myHND(2) = FindWindowEx(myHND(1), 0&, "TButton", "�V�K")
    If myHND(2) = 0 Then GoTo line01
    PostMessage myHND(2), BM_CLICK, 0, 0
    '���i�i�Ԃ����
line02:
    myHND(3) = FindWindow("TForm", "�V�K�t�@�C���̍쐬")
    If myHND(3) = 0 Then GoTo line02
    myHND(4) = FindWindowEx(myHND(3), 0&, "TEdit", vbNullString)
    If myHND(4) = 0 Then GoTo line02
    Call input_Message(myHND(4), Replace(���i�i��str, " ", "") & "_" & Replace(�ݕ�str, " ", ""))
    'ENTER
line03:
    myHND(5) = FindWindowEx(myHND(3), 0&, "TButton", "OK")
    If myHND(5) = 0 Then GoTo line03
    Call Control_Click(myHND(5), "SEND", &H1&)
    '�w�b�_�[�ҏW
    '���i�i��
line04:
    myHND(6) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(7) = FindWindowEx(myHND(6), 0&, "TfrmHeader", vbNullString)
    myHND(8) = FindWindowEx(myHND(7), 0&, "TEdit", "00000000000000000000")
    If myHND(8) = 0 Then GoTo line04:
    SendMessage myHND(8), EM_SETSEL, 0, 20 '������I��
    SendMessage myHND(8), WM_CLEAR, 0, 0 '�������N���A
    SendMessage myHND(8), WM_LBUTTONDOWN, 0&, 0& '�I��
    SendMessage myHND(8), WM_LBUTTONUP, 0&, 0&
    a = SendMessageStr(myHND(8), WM_SETTEXT, 0&, Replace(���i�i��str, " ", "") & "_" & Replace(�ݕ�str, " ", ""))
    'WH���i�i��
    myHND(8) = FindWindowEx(myHND(7), 0&, "TEdit", "00000000")
    SendMessage myHND(8), EM_SETSEL, 0, 20 '������I��
    SendMessage myHND(8), WM_CLEAR, 0, 0 '�������N���A
    SendMessage myHND(8), WM_LBUTTONDOWN, 0&, 0& '�I��
    SendMessage myHND(8), WM_LBUTTONUP, 0&, 0&
    a = SendMessageStr(myHND(8), WM_SETTEXT, 0&, ByVal Right(Replace(���i�i��str, " ", ""), 8))
    PostMessage myHND(8), WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
    '�����߲��
    myHND(8) = FindWindowEx(myHND(7), 0&, "TEdit", "100")
    SendMessage myHND(8), EM_SETSEL, 0, 20 '������I��
    SendMessage myHND(8), WM_CLEAR, 0, 0 '�������N���A
    SendMessage myHND(8), WM_LBUTTONDOWN, 0&, 0& '�I��
    SendMessage myHND(8), WM_LBUTTONUP, 0&, 0&
    a = SendMessageStr(myHND(8), WM_SETTEXT, 0&, "1900")
    PostMessage myHND(8), WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
    'YC�@��
    myHND(8) = FindWindowEx(myHND(7), 0&, "TComboBox", vbNullString)
    Index = SendMessage(myHND(8), CB_GETCOUNT, 0, 0)
    For i = 0 To Index - 1
        myLen = SendMessage(myHND(8), CB_GETTEXTLEN, i, 0)
        Dim myStr As String: myStr = String(myLen, vbNullChar)
        Ret = SendMessageStr(myHND(8), CB_GETTEXT, i, myStr)
        a = SendMessage(myHND(8), CB_SETTOPINDEX, i, 0)
        b = SendMessage(myHND(8), CB_SETCURSEL, i, 0)
        c = SendMessage(myHND(8), BM_CLICK, i, 0)
        If i = 18 Then Exit For
    Next i
    PostMessage myHND(8), WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
    '����Ӱ��
    myHNDtemp = GetWindow(myHND(8), 2) '���̃n���h��
    Index = SendMessage(myHNDtemp, CB_GETCOUNT, 0, 0)
    For i = 0 To Index - 1
        a = SendMessage(myHNDtemp, CB_SETTOPINDEX, i, 0)
        b = SendMessage(myHNDtemp, CB_SETCURSEL, i, 0)
        c = SendMessage(myHNDtemp, BM_CLICK, i, 0)
        If i = 2 Then Exit For
    Next i
    PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
    'Symbol�f�[�^�ҏW
    Stop '�V���{���f�[�^�ҏW���A�N�e�B�u�ɂ���
    myHND(2) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(3) = FindWindowEx(myHND(2), 0&, "TfrmSymbol", "�V���{���f�[�^�ҏW")
    SetForegroundWindow myHND(3) '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
    'PostMessage myHND(3), &H222, &H408E4, &H4081A
    myHND(4) = GetWindow(myHND(3), 5)
    myHND(5) = GetWindow(myHND(4), 1)
    myHNDtemp = myHND(5)
    '�V���{���f�[�^�ҏW�ɉ�H������^����
    Dim pageMax As Long: pageMax = 100
    With wb(3).Sheets("Symbol")
        For s = 1 To endPoint
            If s > pageMax Then
                myHNDtemp = myHND(5)
                pageMax = pageMax + 100
            End If
            Set myPoint = .Cells.Columns(1).Find(s, , , 1)
            If Not (myPoint Is Nothing) Then
                ��H����str = myPoint.Offset(0, 1)
                SetForegroundWindow myHNDtemp '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '�I��
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                'SendMessageStr myHNDtemp, WM_SETTEXT, 0&, ��H����str
                Call input_Message(myHNDtemp, CStr(��H����str))
            End If
            'Debug.Print Hex(myHNDtemp), s Mod 100
            'Call input_Message(myHNDtemp, CStr(��H����str))
            'SendMessage myHNDtemp, &H281, &H1, &HC000000F
            'SendMessage myHNDtemp, &H281, &H0, &HC000000F
            'SendMessage myHNDtemp, &H1, 1, 0&
            'SendMessage myHNDtemp, BM_CLICK, 0, 0
            'bbb = PostMessage(myHNDtemp, WM_KEYUP, &HD, 0)
            PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
            myHNDtemp = GetWindow(myHNDtemp, 3) '��̃n���h��
            Sleep 50
        Next s
    End With
    Stop '�����܂�
    'W/H�f�[�^�ҏW
    myHND(3) = FindWindowEx(myHND(2), 0&, "TfrmWH", "�v�^�g�f�[�^�ҏW")
    myHND(4) = GetWindow(myHND(3), 5)
    myHND(5) = GetWindow(myHND(4), 1)
    myHNDtemp = myHND(5)
    'WH�ҏW�ɉ�H������^����
    pageMax = 60
    With wb(3).Sheets("WH")
        For s = 1 To endKosei
            Set mykosei = .Cells.Columns(1).Find(Format(s, "0000"), , , 1)
            If Not (mykosei Is Nothing) Then
                ��H����Astr = mykosei.Offset(0, 1)
                ��H����Bstr = mykosei.Offset(0, 2)
                SetForegroundWindow myHNDtemp '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '�I��
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                Call input_Message(myHNDtemp, CStr(��H����Astr))
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
                Sleep 50
                myHNDtemp = GetWindow(myHNDtemp, 3) '���̃n���h��
                SetForegroundWindow myHNDtemp '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '�I��
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                Call input_Message(myHNDtemp, CStr(��H����Bstr))
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
                Sleep 50
            Else
                SetForegroundWindow myHNDtemp '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '�I��
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
                Sleep 50
                myHNDtemp = GetWindow(myHNDtemp, 3) '���̃n���h��
                SetForegroundWindow myHNDtemp '�X�e�b�v�C�����ƍőO�ʂɂȂ�Ȃ����璍��
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '�I��
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTER�Ŋm��
                Sleep 50
            End If
            If s = pageMax Then
                myHNDtemp = myHND(5) '�擪�̃n���h��
                pageMax = pageMax + 60
            Else
                myHNDtemp = GetWindow(myHNDtemp, 3) '���̃n���h��
            End If
        Next s
    End With
    
    Stop
    Stop
    
    a = SendMessage(myHND(0), WM_KEYDOWN, WM_ALT, 0)
    Call Control_Click(&H5&, "SEND", "ThunderRT6FormDC")
    Call Control_Click(&H2&, "SEND", "ThunderRT6FormDC")  '�V�`���f�[�^

    Call �������x���V�X�e��_��荞��(myTextDir, �t�@�C����) '�捞�f�[�^�t�@�C���I��
    '�ҋ@
    '�㏑���m�F�A�捞�������������܂����B�ǂ��炩�m�F
    Do
        �㏑���m�F = �������x���V�X�e��_�����m�F("#32770", "�����m�F", "�͂�(&Y)", &H6&)  '�㏑���m�F
        
        �捞�������� = �������x���V�X�e��_�����m�F("#32770", "�����m�F", "OK", &H2&)  '�捞�������������܂���
        If �捞�������� <> 0 Then Exit Do
    Loop
    
    Call �������x���V�X�e��_���O��t���ĕۑ�(&H1&, myTextDir & "\" & �Ǘ��i���o�[ & "_KairoMat_3.txt")
    '�G���[�ɂȂ�̂ōċN��
    Call �������x���V�X�e��_����(&H1&)             '����
    Sleep 1000
    Call Control_�A�v���N��(�������x���V�X�e��Path)
    '���ޏ��v��
    Call �������x���V�X�e��_�Ǘ��i���o�[�I��(�Ǘ��i���o�[)
    Call �������x���V�X�e��_���ޏ��v�ʕ\��
    Call �������x���V�X�e��_���O��t���ĕۑ�(&H1&, myTextDir & "\" & �Ǘ��i���o�[ & "_MRP.txt")
    Call �������x���V�X�e��_����2(&H1&)
    
    Set JJF = Nothing
    Set FSO = Nothing
    If myCount = 0 Then
        a = MsgBox("�Ώۂ̃t�@�C����������Ȃ��ׁA���������s�ł��܂���ł����B" & vbCrLf & _
                    "���݂̏ꏊ�ɏ����������t�@�C��(RFLT??-B?.txt)�����鎖���m�F���Ă�����s���Ă��������B" & vbCrLf & _
                    "" & vbCrLf & _
                    "���݂̏ꏊ: " & ActiveWorkbook.Path, vbOKOnly, "PLUS+")
    Else
        a = MsgBox("�������I�����܂���", vbOKOnly, "PLUS+")
    End If
End Sub

Public Sub �������x���V�X�e��_���ޏ��v�ʕ\��()
line10:
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "ThunderRT6Frame", "�����Ώۋ敪�I��")
    If myHND(1) = 0& Then GoSub ��荞�ݑҋ@
    '�`�F�b�N�����
    For i = 2 To 5
        myHND(2) = GetDlgItem(myHND(1), i)
        Rep = SendMessage(myHND(2), BM_SETCHECK, 1, 0)
    Next i
    '�\��
    myHND(3) = GetDlgItem(myHND(0), &H8&)
    SendMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + &H8&, myHND(3)
    '�t�@�C���o��
    myHND(4) = GetDlgItem(myHND(0), &H9&)
    PostMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + &H9&, myHND(4)
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub
Public Sub �������x���V�X�e��_�Ǘ��i���o�[�I��(NMB�i���o�[)
line10:
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = FindWindowEx(myHND(0), 0&, "ThunderRT6Frame", "�����ΏۊǗ��m���I��")
    myHND(1) = FindWindowEx(myHND(0), 0&, "ThunderRT6TextBox", vbNullString)
    myHND(2) = FindWindowEx(myHND(0), 0&, "ThunderRT6ComboBox", vbNullString)
    myHND(3) = FindWindowEx(myHND(2), 0&, "Edit", vbNullString)
    If myHND(3) = 0& Then GoSub ��荞�ݑҋ@

    Index = SendMessage(myHND(2), CB_GETCOUNT, 0, 0)
    For i = 0 To Index - 1
        myLen = SendMessage(myHND(2), CB_GETTEXTLEN, i, 0)
        Dim myStr As String: myStr = String(myLen, vbNullChar)
        Ret = SendMessageStr(myHND(2), CB_GETTEXT, i, myStr)
        Debug.Print myStr
        If NMB�i���o�[ = myStr Then
            a = SendMessage(myHND(2), CB_SHOWDROPDOWN, 1, 0)
            b = SendMessage(myHND(2), CB_SETCURSEL, i, 0)
            c = SendMessage(myHND(2), WM_LBUTTONDOWN, i, 0)
            Exit For
        End If
    Next i
    
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Public Sub Control_�A�v���N��(myPath)
line10:
        myHND(0) = FindWindow("TfrmMain", vbNullString)
        If myHND(0) = 0& Then GoSub �A�v���̋N��
    Sleep 100
    Exit Sub
�A�v���̋N��:
    
    On Error GoTo myErr
        ChDrive Left(myPath, 2)
        ChDir Left(myPath, InStrRev(myPath, "\") - 1)
        Shell myPath
    On Error GoTo 0
    
    myCount = myCount + 1: If myCount > 10 Then Stop
    GoTo line10
    
myErr:
    If Err.Number = 76 Or Err.Number = 53 Then
        MsgBox "�V�[�g[�ݒ�]��YcEditor�̃A�h���X������������܂���B" & vbCrLf & vbCrLf _
             & "YCEditor.exe�̕ۑ��A�h���X���m�F���ďC�����Ă��������B"
    End If
    Sheets("�ݒ�").Activate
    Sheets("�ݒ�").Cells.Find("YcEditor_exe", , , 1).Offset(0, 1).Activate
    End
End Sub

Public Sub �������x���V�X�e��_��荞��(��荞�݃t�@�C��Path, �t�@�C����)
    'Drive�ɕ�����n��
    myDrive = Left(��荞�݃t�@�C��Path, 1)
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = GetDlgItem(myHND(0), &H5&)
    myASC = Asc(myDrive)
    b = SendMessage(myHND(0), WM_IME_CHAR, myASC, 0)
    '�f�B���N�g���̑I��
    Dim myStr As String:
    temp = Split(��荞�݃t�@�C��Path, "\")
    For i = LBound(temp) To UBound(temp)
        Dim myFolder As String
        If i = 0 Then myFolder = StrConv(temp(i), vbLowerCase) & "\" Else myFolder = temp(i)
        'Dir�̎Q��
        myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
        myHND(0) = GetDlgItem(myHND(0), &H7&)
        Index = SendMessage(myHND(0), LB_GETCOUNT, 0, 0)
        For i2 = 0 To Index - 1
            myLen = SendMessage(myHND(0), LB_GETTEXTLEN, i2, 0)
            myStr = String(myLen, vbNullChar)
            Ret = SendMessageStr(myHND(0), LB_GETTEXT, i2, myStr)
            If StrConv(myStr, vbUpperCase) = StrConv(myFolder, vbUpperCase) Then
                c = SendMessage(myHND(0), LB_SETCURSEL, i2, 0)
                D = SendMessage(myHND(0), WM_LBUTTONDBLCLK, i2, 0)
                Exit For
            End If
        Next i2
line10:
    Next i
    '�t�@�C���̑I��
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = GetDlgItem(myHND(0), &H6&)
    Index = SendMessage(myHND(0), LB_GETCOUNT, 0, 0)
        For i3 = 0 To Index - 1
            myLen = SendMessage(myHND(0), LB_GETTEXTLEN, i3, 0)
            myStr = String(myLen, vbNullChar)
            Ret = SendMessageStr(myHND(0), LB_GETTEXT, i3, myStr)
            If myStr = �t�@�C���� Then
                c = SendMessage(myHND(0), LB_SETCURSEL, i3, 0)
                D = SendMessage(myHND(0), WM_LBUTTONDBLCLK, i3, 0)
            End If
        Next i3
    '��荞�݃{�^��������
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(1) = GetDlgItem(myHND(0), &H9&)
    PostMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + &H9&, myHND(1)
Sleep 100
End Sub
Public Function �������x���V�X�e��_�����m�F(myClass, myCaption, myCaption2, myID)
    '�㏑�����Ă������ł���?

    myHND(0) = FindWindow(myClass, myCaption)
    myHND(1) = FindWindowEx(myHND(0), 0&, vbNullString, myCaption2)
    �������x���V�X�e��_�����m�F = myHND(1)
    If myHND(1) = 0 Then Exit Function
    SendMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + myID, myHND(1)

End Function
Public Sub Control_Click(myHWND, ���, CtrlID)
    Dim myDlg, myBtn, myStat As Long
    Dim myCount As Long
line10:
    SendMessage myHWND, BM_CLICK, 0, 0
    
    Exit Sub
    
    
    Stop
    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then Exit Sub
    Select Case ���
        Case "SEND": SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
        Case "POST": PostMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
    End Select
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub

Public Sub �������x���V�X�e��_�N���b�N2(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    Dim myStr As String: myStr = String(15, vbNullChar)
line10:
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "ThunderRT6Frame", vbNullString)
    If myWHND = 0& Then GoSub ��荞�ݑҋ@

    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then Exit Sub

    '�N���b�N
    SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myStat = GetDlgItem(myWHND, &H6&)
    'ret = SendMessageStr(myStat, WM_GETTEXT, 15, myStr)
    Ret = SendMessage(myStat, LB_SETSEL, 1, -1)
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub


Public Sub �������x���V�X�e��_�N���b�N3(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    Dim myStr As String: myStr = String(15, vbNullChar)
line10:
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = FindWindowEx(myWHND, 0&, "ThunderRT6ListBox", vbNullString)
    If myWHND = 0& Then GoSub ��荞�ݑҋ@

    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then Exit Sub

    '�N���b�N
    SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myStat = GetDlgItem(myWHND, &H6&)
    'ret = SendMessageStr(myStat, WM_GETTEXT, 15, myStr)
    Ret = SendMessage(myStat, LB_SETSEL, 1, -1)
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub
Public Sub �������x���V�X�e��_����(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    
line10:
    myWHND = FindWindow("ThunderRT6FormDC", "�������x���V�X�e��")
    Rep = PostMessage(myWHND, WM_SYSCOMMAND, SC_CLOSE, 0)  '����(�����Ċm�F�E�B���h�E���J���̂ŏ�����҂��Ȃ�POST���g��)

    myWHND = FindWindow("#32770", "�����m�F")
    If myWHND = 0& Then GoSub ��荞�ݑҋ@
    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then GoSub ��荞�ݑҋ@
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 1000 Then Stop 'post�Ȃ̂ŕԂ�l������ ���̐����傫�����Ȃ����񂩂�
GoTo line10
End Sub
Public Sub �������x���V�X�e��_����2(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    
line10:
    
    myWHND = FindWindow("#32770", "���ޏ��v�ʕ\��")
    If myWHND = 0& Then GoSub ��荞�ݑҋ@
    myBtn = GetDlgItem(myWHND, &H2&)
    SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + &H2&, myBtn
    
    myWHND = FindWindow("ThunderRT6FormDC", "�������x���V�X�e��")
    Rep = PostMessage(myWHND, WM_SYSCOMMAND, SC_CLOSE, 0)  '����(�����Ċm�F�E�B���h�E���J���̂ŏ�����҂��Ȃ�POST���g��)
Sleep 100
    myWHND = FindWindow("#32770", "���ޏ��v�ʕ\��")
    If myWHND = 0& Then GoSub ��荞�ݑҋ@
    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then GoSub ��荞�ݑҋ@
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Public Sub �������x���V�X�e��_���O��t���ĕۑ�(CtrlID, �ۑ��t���p�X)
    Dim myWHND, myWHND2, myWHND3, lngRC, myDlg, myStat, myCount As Long
    Dim myText, myASC As String
line10:
    myWHND = FindWindow("#32770", "���O��t���ĕۑ�")
    myWHND = FindWindowEx(myWHND, 0&, "DUIViewWndClassName", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "DirectUIHWND", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "FloatNotifySink", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "ComboBox", vbNullString)
    myStat = FindWindowEx(myWHND, 0&, "Edit", vbNullString)
    If myWHND = 0& Then GoSub ��荞�ݑҋ@
    
    For i = 1 To Len(�ۑ��t���p�X)
        myText = Mid(�ۑ��t���p�X, i, 1)
        myASC = Asc(myText)
        lngRC = SendMessage(myStat, WM_IME_CHAR, myASC, 0)
        Sleep 1
    Next i
    
    myWHND = FindWindow("#32770", "���O��t���ĕۑ�")
    myBtn = GetDlgItem(myWHND, CtrlID)
    '�N���b�N
line20:
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
    If Rep <> 0 Then GoTo line20
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Sub Sample_�R���{�{�b�N�X�̑I��() 'Drive�I��_FileList���؂�ւ��Ȃ��̂Ŏg�p���Ȃ�
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = GetDlgItem(myHND(0), &H5&)
    
    Index = SendMessage(myHND(0), CB_GETCOUNT, 0, 0)
    
    For i = 0 To Index - 1
        myLen = SendMessage(myHND(0), CB_GETTEXTLEN, i, 0)
        Dim myStr As String: myStr = String(myLen, vbNullChar)
        Ret = SendMessageStr(myHND(0), CB_GETTEXT, i, myStr)
        a = SendMessage(myHND(0), CB_SETTOPINDEX, i, 0)
        b = SendMessage(myHND(0), CB_SETCURSEL, i, 0)
        c = SendMessage(myHND(0), BM_CLICK, i, 0)
    Next i
End Sub

Public Function input_Message(myStat, myMessage)
    
    For i = 1 To Len(myMessage)
        myText = Mid(myMessage, i, 1)
        myASC = Asc(myText)
        Sleep 10
        lngRC = SendMessage(myStat, WM_IME_CHAR, myASC, 0)
        If lngRC <> 0 Then Stop
    Next i
End Function

Public Sub Control_�V�K�쐬()
    Dim myWHND, myWHND2, myWHND3, lngRC, myDlg, myStat, myCount As Long
    Dim myText, myASC As String
line10:
    myWHND = FindWindow("TfrmFile", "�t�@�C���I��")
    myWHND = FindWindowEx(myWHND, 0&, "TButton", "�V�K")
    myWHND = FindWindowEx(myWHND, 0&, "DirectUIHWND", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "FloatNotifySink", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "ComboBox", vbNullString)
    myStat = FindWindowEx(myWHND, 0&, "Edit", vbNullString)
    If myWHND = 0& Then GoSub ��荞�ݑҋ@
    
    For i = 1 To Len(�ۑ��t���p�X)
        myText = Mid(�ۑ��t���p�X, i, 1)
        myASC = Asc(myText)
        lngRC = SendMessage(myStat, WM_IME_CHAR, myASC, 0)
        Sleep 1
    Next i
    
    myWHND = FindWindow("#32770", "���O��t���ĕۑ�")
    myBtn = GetDlgItem(myWHND, CtrlID)
    '�N���b�N
line20:
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
    If Rep <> 0 Then GoTo line20
Sleep 100
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Public Function inputString(myStr)
    For e = 1 To Len(myStr)
        myText = Mid(myStr, e, 1)
        myASC = Asc(myText)
        lngRC = SendMessage(myHND(0), WM_IME_CHAR, myASC, 0)
    Next e
    Sleep 300
End Function

Public Sub clickButton(mySelect)
    Dim myCount As Long
    
line10:
    If myHND(0) = 0& Then GoSub ��荞�ݑҋ@
    myHND(1) = FindWindow("ExToolBoxClass", vbNullString)
    If myHND(1) = 0& Then GoSub ��荞�ݑҋ@
    Select Case mySelect
        Case "Enter"
            myHND(2) = FindWindowEx(myHND(1), 0&, "Button", "ENTER")
            myBtn = GetDlgItem(myHND(1), &H218)
            a = SendMessage(myBtn, BM_CLICK, 0, 0)
            Sleep 500
        Case "Home"
            myHND(2) = FindWindowEx(myHND(1), 0&, "Button", "HOME")
            myBtn = GetDlgItem(myHND(1), &H213)
            a = SendMessage(myBtn, BM_CLICK, 0, 0)
            Sleep 100
        Case "Pause"
            myHND(2) = FindWindowEx(myHND(1), 0&, "Button", "CLEAR")
            myBtn = GetDlgItem(myHND(1), &H217)
            a = SendMessage(myBtn, BM_CLICK, 0, 0)
        Sleep 100
    Case Else
        Stop
    End Select
    
Exit Sub
��荞�ݑҋ@:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub

Sub Macro1()
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    Const WM_APP = &H8000

    Ret = PostMessage(myHND(0), WM_APP + 15620, 12, 20380001)
    
    SendMessage myHND(0), &H92, 0, &H19EE84
    SendMessage myHND(0), &H11F, &HFFFF0000, 0
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(2) = FindWindowEx(myHND(1), 0&, "TfrmSymbol", "�V���{���f�[�^�ҏW")
    myHND(3) = GetWindow(myHND(2), 5)
    myHND(4) = GetWindow(myHND(3), 1)
    myHNDtemp = myHND(4)
    PostMessage myHND(0), WM_IME_KEYDOWN, 164, 1
    PostMessage myHND(0), WM_IME_KEYDOWN, 70, 1
    PostMessage myHND(0), WM_IME_KEYUP, 70, 0
    PostMessage myHND(0), WM_IME_KEYUP, 164, 0
    AppActivate myHND(0)
    PostMessage myHND(0), &H106, VK_RMENU, 0
    Stop
    Stop
    For i = 1 To 13
        PostMessage myHND(0), WM_KEYDOWN, i, 0
        PostMessage myHND(0), WM_KEYUP, i, 0
    Next i
    Call input_Message(myHND(0), CStr(��H����str))
End Sub

Sub test_notepad()
    myHND(0) = FindWindow("notepad", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "Edit", vbNullString)
    
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(2) = FindWindowEx(myHND(1), 0&, "TfrmSymbol", "�V���{���f�[�^�ҏW")
    myHND(3) = GetWindow(myHND(2), 5)
    myHND(4) = GetWindow(myHND(3), 1)
    SetForegroundWindow myHND(0) '�f�o�b�O�ŃX�e�b�v������őO�ʂɂȂ�Ȃ����璍��
    PostMessage myHND(0), WM_SYSKEYDOWN, VK_MENU, &H20380001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("F"), &H20210001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("O"), &H20180001
End Sub
