VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_04 
   Caption         =   "VerUp"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   OleObjectBlob   =   "UI_04.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






































































































Public thisVer As String
Public newVer As String

 Sub fjkajfdaljdfka()
    PlaySound "�����Ă�"
    
    Call �A�h���X�Z�b�g(ActiveWorkbook)
    Path = �A�h���X(0) & "\ver"
    If Dir(Path, vbDirectory) = "" Then MkDir (Path)

    Path = Path & "\" & Mid(ThisWorkbook.Name, 6, InStr(ThisWorkbook.Name, "_") - 6)
    If Dir(Path, vbDirectory) = "" Then MkDir (Path)
    
    myCount = VBC_Export(Path)
    Call Sheet_Ver_Export(Path)
    
    DoEvents
    
    If myCount = 0 Then
        MsgBox "�G�N�X�|�[�g�o����R�[�h������܂���ł����B"
    Else
        MsgBox myCount & " �_�̃R�[�h���G�N�X�|�[�g���܂����B"
    End If
    
    Unload UI_04
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Shift = 1 Then
        PlaySound "�����Ă�"
        
        Call �A�h���X�Z�b�g(ActiveWorkbook)
        Path = �A�h���X(0) & "\ver"
        If Dir(Path, vbDirectory) = "" Then MkDir (Path)
    
        Path = Path & "\" & Mid(ThisWorkbook.Name, 6, InStr(ThisWorkbook.Name, "_") - 6)
        If Dir(Path, vbDirectory) = "" Then MkDir (Path)
        Dim newVer As String: newVer = Mid(Path, InStrRev(Path, "\") + 1)
        
        myCount = VBC_Export(Path)
        Call Sheet_Ver_Export(Path)
        Call MakeShortcut(Path)
        DoEvents
        
        If myCount = 0 Then
            MsgBox "�G�N�X�|�[�g�o����R�[�h������܂���ł����B"
        Else
            MsgBox myCount & " �_�̃R�[�h���G�N�X�|�[�g���܂����B"
            Call ���O�o��("test", "test", "VerExport = " & newVer)
        End If
        
        Unload UI_04
    End If
End Sub

Private Sub CommandButton2_Click()
    If CB0.Value = "" Then MsgBox "�o�[�W������I�����Ď��s���Ă��������B": End
    If Left(ThisWorkbook.Name, Len(mySystemName)) <> mySystemName Then MsgBox "�t�@�C������" & mySystemName & "����n�܂��Ă���K�v������܂��B" & vbCrLf & "���O���C�����ĉ������B": End
    newVer = CB0.Value
    aa = MsgBox("VerUp�����s���܂��B" & vbCrLf & thisVer & " �� " & newVer & vbCrLf & "�����̓s����A�ʃu�b�N����̎��s�ɂȂ�܂��B�o�[�W�����A�b�v�̎��s�{�^���������Ă��������B", vbYesNo): If aa = vbNo Then End
    PlaySound "�����Ă�"
    
    Call DeleteDefinedNames '���O�̒�`���d��������x���o�邩��폜����
    mywb = ActiveWorkbook.FullName
    Workbooks.Open �A�h���X(0) & "\VerUp.xlsm"
    Set wb(0) = ActiveWorkbook
    
    With wb(0).Sheets("Sheet1")
        .Cells(1, 1) = �A�h���X(0) & "\ver\" & newVer
        .Cells(2, 1) = mywb
    End With
    
    Call ���O�o��("test", "test", "VerUP" & thisVer & "��" & newVer)
    
    Unload UI_04
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("���ǂ�")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()
    
    Dim buf As String, msg As String
    Dim ����(1) As String
    Dim myDateTime
    Dim nowVer As String
    
    nowVer = ActiveWorkbook.Name
    nowVer = Mid(nowVer, InStr(nowVer, "+") + 1, InStr(nowVer, "_") - InStr(nowVer, "+") - 1)
    
    Me.Caption = nowVer
    Call �A�h���X�Z�b�g(ActiveWorkbook)
    buf = Dir(�A�h���X(0) & "\ver\", vbDirectory)
    Do While buf <> ""
        If Replace(buf, ".", "") <> "" Then
            ����(0) = ����(0) & "," & buf
            ����(1) = ����(1) & "," & FileDateTime(�A�h���X(0) & "\ver\" & buf)
        End If
        buf = Dir()
    Loop
    ����(0) = Mid(����(0), 2)
    ����(1) = Mid(����(1), 2)
    Debug.Print msg
    
    ����0s = Split(����(0), ",")
    ����1s = Split(����(1), ",")
    With CB0
        .RowSource = ""
        For i = LBound(����0s) To UBound(����0s)
            .AddItem ����0s(i)
            If ����1s(i) > myDateTime Then myindex = i
            myDateTime = ����1s(i)
        Next i
        .ListIndex = UBound(����0s)
    End With
    
    newVer = CB0.Value
    thisVer = Mid(ThisWorkbook.Name, Len(mySystemName) + 1, InStr(ThisWorkbook.Name, "_") - Len(mySystemName) - 1)
    
    If thisVer = newVer Then
        messe.Caption = "�o�[�W�����͍ŐV�ł�"
    ElseIf thisVer < newVer Then
        messe.Caption = "�V�����o�[�W����������܂�"
        messe.ForeColor = RGB(255, 0, 0)
    Else
        messe.Caption = "���̃o�[�W���������V�����ł��B" & vbCrLf & "�G�N�X�|�[�g�����s���Ă��������B"
        messe.ForeColor = RGB(255, 0, 0)
        CommandButton2.Visible = False
    End If
    
End Sub
