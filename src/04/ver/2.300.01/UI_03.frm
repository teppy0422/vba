VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_03 
   Caption         =   "�z���}"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_03.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


















































































Private Sub CB5_Change()
    
End Sub

Private Sub CB0_Change()
    Dim ����(1) As String
    Dim ����2(1) As String
    'CB0.Text
    With ActiveWorkbook.Sheets("���i�i��")
        Set myKey = .Cells.Find("�^��", , , 1)
        Set myKey = .Rows(myKey.Row).Find(CB0.Text, , , 1)
        Set mykey2 = .Rows(myKey.Row).Find("����", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For y = myKey.Row + 1 To lastRow
            If InStr(����(0), "," & .Cells(y, myKey.Column)) & "," = 0 Then
                ����(0) = ����(0) & "," & .Cells(y, myKey.Column) & ","
                ����2(0) = ����2(0) & "," & .Cells(y, mykey2.Column) & ","
            End If
        Next y
        If Len(����(0)) <= 2 Then
            ����(0) = ""
            ����s = Empty
        Else
            ����(0) = Mid(����(0), 2)
            ����(0) = Left(����(0), Len(����(0)) - 1)
            ����s = Split(����(0), ",,")
            ����2(0) = Mid(����2(0), 2)
            ����2(0) = Left(����2(0), Len(����2(0)) - 1)
            ����2s = Split(����2(0), ",,")
        End If
    End With
    
    With CB1
        .RowSource = ""
        .Clear
        If Not IsEmpty(����s) Then
            For i = LBound(����s) To UBound(����s)
                .AddItem
                .List(i, 0) = ����s(i)
                .List(i, 1) = ����2s(i)
            Next i
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub CB1_Change()
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "")
    If ���i�i��RANc <> 1 Then
        myLabel.Caption = "���i�i�ԓ_�����ُ�ł��B"
        myLabel.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        myLabel.Caption = ""
    End If
End Sub

Private Sub CommandButton4_Click()
    PlaySound "���ǂ�"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    mytime = time
    PlaySound "��������"
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "")
    ���str = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "����"), 1)
    Dim aa As Long
    On Error Resume Next
    a = wb(0).Sheets("���_" & ���str).Cells(1, 1)
    On Error GoTo 0
    If a = Empty Then
        aa = MsgBox("����_" & ���str & " �̃V�[�g���쐬����Ă��܂���B" & vbLf & _
               "��n���}�����̍쐬�ɂȂ�܂��B", vbYesNo)
        If aa = 6 Then
            �z���}�쐬temp = "1"
            GoTo line10
        End If
        Exit Sub
    Else
        �z���}�쐬temp = "0"
    End If
    
    With wb(0).Sheets("���_" & ���str)
        .Activate
        Call �z���}�쐬
    End With
line10:
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "") '�z���}�쐬�̎��ɓ������̐��i�i�Ԃ��Z�b�g�����̂Ń��Z�b�g
    If ���i�i��RANc <> 1 Then
        myLabel.Caption = "���i�i�ԓ_�����ُ�ł��B"
        myLabel.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    Unload Me
    
    Set wb(0) = ActiveWorkbook
    
    cb�I�� = "1,4,1,1,0,-1,-1"
    ���^�p�x����flag = False
    �}���}�`�� = 21
    �������� = cbx1
    Call �n���}�쐬_Ver220098(cb�I��, CB0.Value, CB1.Value)
    If cbx2 = True Then ��n���_�� = True
    If cbx1 = True Then
        '�U���i�r�`�F�b�N�L��
        Call �z���}�쐬one3(���i�i��Ran, "�n���}_" & CB0.Value & "_" & Replace(CB1.Value, " ", ""))
        Call ���O�o��("test", "test", "�z���U��" & CB1.Value)
    Else
        Call �z���}�쐬one(���i�i��Ran, "�n���}_" & CB0.Value & "_" & Replace(CB1.Value, " ", ""))
        Call ���O�o��("test", "test", "�z���}��" & CB1.Value)
    End If
    Call �œK�����ǂ�
    PlaySound "���񂹂�"
    
    Dim myMsg As String: myMsg = "�쐬���܂���" & vbCrLf & DateDiff("s", mytime, time) & "s"
    If �z���}�쐬temp = "1" Then myMsg = myMsg & vbCrLf & vbCrLf & "������W�f�[�^��������Ȃ������̂Ō�n���}�f�[�^�̂ݍ쐬���܂����B"
    aa = MsgBox(myMsg, vbOKOnly, "���Y����+�z���U��")
End Sub

Private Sub CommandButton6_Click()
    
End Sub

Private Sub UserForm_Initialize()
    Dim ����(1) As String
    With ActiveWorkbook.Sheets("���i�i��")
        Set myKey = .Cells.Find("�^��", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For x = myKey.Column To lastCol
            ����(0) = ����(0) & "," & .Cells(myKey.Row, x)
        Next x
        ����(0) = Mid(����(0), 2)
    End With
    ����s = Split(����(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
            If ����s(i) = "���C���i��" Then myindex = i
        Next i
        .ListIndex = myindex
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub
