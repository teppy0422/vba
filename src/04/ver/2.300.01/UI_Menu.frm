VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Menu 
   Caption         =   "menu"
   ClientHeight    =   6120
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7840
   OleObjectBlob   =   "UI_Menu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "UI_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














































Private clsForm As New clsUserForm
Private THEME As Long, THEMEgray1 As Long, THEMEgray2 As Long
Private THEMEwhite As Long

Private Sub initFormSetting()

    Me.BorderColor = THEME
    
    Me.Labeltitle.Top = 1
    Me.Labeltitle.Left = 1
    Me.Labeltitle.Width = Me.Width - 3
    Me.Labeltitle.BackColor = THEME
    
    Me.btnClose.Top = 1
    Me.btnClose.Left = Me.Labeltitle.Width - Me.btnClose.Width + 1
    
    Me.btnHelp.Top = 1
    Me.btnHelp.Left = Me.btnClose.Left - Me.btnHelp.Width - 3
    
    Me.myVerup.Top = 1
    
    Me.current.Top = 1
    Me.myVerup.BackColor = THEME
    Me.myVerup.ForeColor = white
    
    Me.Label0.ForeColor = black
    
End Sub

Private Sub NormalizeSet()
    Me.btnClose.BackColor = THEME
    Me.btnClose.ForeColor = clsForm.GetColor(white)
    Me.btnHelp.BackColor = THEME
    Me.btnHelp.ForeColor = clsForm.GetColor(white)
    Me.myVerup.BackColor = THEME
    Me.myVerup.ForeColor = clsForm.GetColor(white)
    Me.current.BackColor = THEME
    Me.current.ForeColor = clsForm.GetColor(white)
End Sub
Private Sub Normalizeset_tag()
        Me.tag1.ForeColor = clsForm.GetColor(gray02)
        Me.tag2.ForeColor = clsForm.GetColor(gray02)
End Sub
Private Sub current_Click()
    Shell "C:\Windows\explorer.exe " & ThisWorkbook.path, vbNormalFocus
End Sub

Private Sub current_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.current.BackColor = clsForm.GetColor(white)
    Me.current.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub Image5_Click()
    Me.MultiPage1.Value = 1
End Sub

Private Sub Image5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.ChangeCursor Hand
End Sub

Private Sub Image6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.ChangeCursor Hand
End Sub

Private Sub in01_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_00.Show
End Sub

Private Sub in01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in01.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub in02_Click()
    'If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_02.Show
End Sub

Private Sub in02_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in02.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub in03_Click()
    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_07.Show
End Sub

Private Sub in03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in03.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub in04_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_08.Show
End Sub

Private Sub in04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in04.ForeColor = vbRed
    clsForm.ChangeCursor Hand
End Sub

Private Sub Label11_Click()
    
End Sub

Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Me.in01.ForeColor = THEMEgray2
        Me.in02.ForeColor = THEMEgray2
        Me.in03.ForeColor = THEMEgray2
        Me.in04.ForeColor = THEMEgray2
        Call NormalizeSet
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Me.out01.ForeColor = THEMEgray2
        Me.out03.ForeColor = THEMEgray2
        Me.out04.ForeColor = THEMEgray2
        Me.out05.ForeColor = THEMEgray2
        Me.out06.ForeColor = THEMEgray2
        Me.out07.ForeColor = THEMEgray2
        Me.out08.ForeColor = THEMEgray2
        Me.out09.ForeColor = THEMEgray2
        
        Me.in01.ForeColor = THEMEgray2
        Me.in02.ForeColor = THEMEgray2
        Me.in03.ForeColor = THEMEgray2
        Me.in04.ForeColor = THEMEgray2
        Call NormalizeSet
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call Normalizeset_tag
End Sub

Private Sub myVerup_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_04.Show
End Sub

Private Sub out01_Click()
    aa = MsgBox("����͌������ł��B" & vbLf & "���s���܂���?", vbYesNo, "��H�}�g���N�X")
    If aa <> 6 Then Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    Call ��H�}�g���N�X�쐬_������
End Sub

Private Sub out01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out01.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out03_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_06.Show
End Sub

Private Sub out03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out03.ForeColor = vbRed
    clsForm.ChangeCursor Hand
End Sub

Private Sub out04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out04.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out04_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    If Shift = 1 Then �T���v���쐬���[�h = True
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_01.Show
End Sub

Private Sub out05_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_03.Show
End Sub

Private Sub out05_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out05.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out06_Click()
    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_05.Show
End Sub

Private Sub out06_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out06.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out07_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_09.Show
End Sub

Private Sub out07_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out07.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out08_Click()
    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_10.Show
End Sub

Private Sub out09_Click()
'    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_11.Show
End Sub

Private Sub out08_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out08.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out09_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out09.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub tag1_Click()
    Me.MultiPage1.Value = 0
End Sub

Private Sub tag1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.tag1.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub tag2_Click()
    Me.MultiPage1.Value = 1
End Sub

Private Sub tag2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.tag2.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.FormDrag Me.Name, Button
End Sub
Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.btnClose.BackColor = clsForm.GetColor(red)
    Me.btnClose.ForeColor = clsForm.GetColor(white)
    clsForm.ChangeCursor Hand
End Sub

Private Sub UserForm_Initialize()
    Set wb(0) = ThisWorkbook
    myIP = GetIPAddress
    addressSet wb(0)
    Me.Label_IP.Caption = myIP
    Call �Q�ƕs������΂��̃t�H���_���쐬����
    Call �Q�Ɛݒ�̕ύX
    
    Call connect_Server
    
    myVer = Base.verCheck(ThisWorkbook)
    Call �f�B���N�g���쐬
    Call �K�v�t�@�C���̎擾
    Call �œK��
    HDsize.Caption = checkSpace(myAddress(0, 1))
    '�t�H�[���f�U�C��
    On Error GoTo ErrHandler
    Static initCompleted As Boolean
    If initCompleted = False Then
        initCompleted = True
        THEME = clsForm.GetColor(TBLUE)         ' Choose theme colors
        THEMEgray1 = RGB(100, 100, 100)
        THEMEgray2 = RGB(220, 220, 220)
        THEMEwhite = RGB(255, 255, 255)
        THEMEred = RGB(255, 0, 0)
        
        clsForm.NonTitleBar Me.Name                      ' Set Flat style
        Call initFormSetting
    End If
    GoTo Finally
ErrHandler:
    Call MsgBox(err.Description, , "����+:��O���������܂����B")
Finally:
    Me.startupposition = 2
    On Error GoTo 0
    Me.btnClose.BackColor = THEME
    Me.btnClose.ForeColor = clsForm.GetColor(white)
    Me.btnHelp.BackColor = THEME
    Me.btnHelp.ForeColor = clsForm.GetColor(white)
    Me.myVerup.BackColor = THEME
    Me.myVerup.ForeColor = clsForm.GetColor(white)
    Me.current.BackColor = THEME
    Me.current.ForeColor = clsForm.GetColor(white)
    
    Me.Labeltitle.Caption = "���Y����+" & myVer
    Dim FSO As New FileSystemObject
    '�A�h���X�ɃA�N�Z�X�ł��邩���ׂ�
    With ActiveWorkbook.Sheets("�ݒ�")
        Dim �A�h���Xb As Variant, myMsg As String
        For i = 0 To 2
            �A�h���Xb = myAddress(i, 1)
            Select Case i
            Case 0, 1
                If FSO.FolderExists(�A�h���Xb) = False Then
                    myMsg = myMsg & myAddress(i, 0) & " �̃t�H���_��������܂���" & vbCrLf
                Else
                    myMsg = myMsg & myAddress(i, 0) & " �̃t�H���_��������܂���" & vbCrLf
                End If
            Case 2
                If �A�h���Xb = "" Then
                    myMsg = myMsg & "����IP�ł� " & myAddress(i, 0) & " �̓o�^������Ă��܂���"
                Else
                    If FSO.FileExists(�A�h���Xb) = False Then
                        myMsg = myMsg & myAddress(i, 0) & " �̃t�@�C����������܂���" & vbCrLf
                    Else
                        myMsg = myMsg & myAddress(i, 0) & " �̃t�@�C����������܂���" & vbCrLf
                    End If
                End If
            End Select
        Next i
    End With
    
    '�A�h���X�m�F�̌���
    With Label0
        .Caption = myMsg
        If InStr(myMsg, "������܂���") > 0 Then
            .ForeColor = RGB(255, 0, 0)
        Else
            .ForeColor = RGB(255, 255, 255)
        End If
    End With
    Debug.Print Label0.ForeColor
    Set FSO = Nothing
    '�t�B�[���h���̃`�F�b�N
    Call fieldAdd("���i�i��", "�t�B�[���h��_���i�i��", 2)
    Call fieldAdd("PVSW_RLTF", "�t�B�[���h��_�ʏ�", 1)
    Call fieldAdd("PVSW_RLTF", "�t�B�[���h��_�ǉ�", 2)
    Call fieldAdd("PVSW_RLTF", "�t�B�[���h��_�ǉ�2", 2)
    
    Call �œK�����ǂ�
End Sub
'**********************************
'top label
'**********************************
Private Sub labelTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call NormalizeSet
    clsForm.ChangeCursor Cross
End Sub

Private Sub labelTitle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.FormDrag Me.Name, Button
End Sub
'**********************************
'close button
'**********************************
Private Sub btnClose_Click()
    Unload Me
End Sub
'**********************************
'help button
'**********************************
Private Sub btnHelp_Click()
    buf = "http://10.7.1.35/nim_intra/40_program/plus/41_web/myweb/index.html "
    'If Dir(buf, vbDirectory) <> "" Then
        'buf = buf & "\myWeb\index.html"
        'IE�̋N��
        Dim objIE As Object '�ϐ����`���܂��B
        Dim ieVerCheck As Variant
    
        Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
        Set objSFO = CreateObject("Scripting.FileSystemObject")
    
        ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
        
        'Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
        
        If ieVerCheck >= 11 Then
            Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10�ȍ~(���Ԃ�)
        End If
        
        objIE.Visible = True      '���ATrue�Ō�����悤�ɂ��܂��B
        
        '�����������y�[�W��\�����܂��B
       objIE.Navigate buf
       
       Set objIE = Nothing
    'End If
End Sub

Private Sub btnHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.btnHelp.BackColor = clsForm.GetColor(white)
    Me.btnHelp.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub myVerup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.myVerup.BackColor = clsForm.GetColor(red)
    Me.myVerup.ForeColor = clsForm.GetColor(white)
    clsForm.ChangeCursor Hand
End Sub
'**********************************
'excute button
'**********************************
Private Sub btnExcute_Click()
'    Me.btnExcute.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub btnExcute_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.ChangeCursor Hand
End Sub

'**********************************
'bottom label
'**********************************
Private Sub labelBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call NormalizeSet
End Sub

Private Sub CommandButton5_Click()
    aa = MsgBox("����͌������ł��B" & vbLf & "���s���܂���?", vbYesNo, "��H�}�g���N�X")
    If aa <> 6 Then Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    Call ��H�}�g���N�X�쐬_������
End Sub

Private Sub CommandButton8_Click()
    If Label0.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    Unload UI_Menu
    UI_70.Show
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call NormalizeSet
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub

Private Sub UserForm_Terminate()
    Application.WindowState = xlMaximized
End Sub

Private Sub Version_Click()

End Sub
