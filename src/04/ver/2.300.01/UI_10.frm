VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_10 
   Caption         =   "�T�u�i���o�[�̏o��"
   ClientHeight    =   4485
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4320
   OleObjectBlob   =   "UI_10.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
































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
        Label_alert.Caption = "���i�i�ԓ_�����ُ�ł��B"
        Label_alert.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        Label_alert.Caption = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    Call addressSet(wb(0))
    
    If myAddress(2, 1) = "" Then
        Call MsgBox("����IP�ł̓t�@�C�����o�^����Ă��܂���B", vbOKOnly, "Sjp+")
        Exit Sub
    End If
    
    If Label_alert.ForeColor = 255 Then MsgBox "�ݒ���m�F���Ă�������", , "���s�ł��܂���": Exit Sub
    PlaySound ("�����Ă�")
    
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "")
    
    Dim �ݕ�str As String
    
    ���i�i��str = UI_10.CB1.Value
    �ݕ�str = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "��z"), 1)
    Dim myMessage As String
    If CheckBox_wireEfu Then
        myMessage = myMessage & _
                               "�d���̃T�u�����X�V���܂��B" & vbCrLf & vbCrLf & _
                               "    �f�[�^��: ���̃u�b�N�̃V�[�g[PVSW_RLTF]" & vbCrLf & _
                               "    �o�͐�F" & myAddress(2, 1) & vbCrLf & vbCrLf & _
                               "    �����w��������V�X�e���ŕt�^����T�u���ł��B" & vbCrLf & vbCrLf
    End If
    If CheckBox_tubeEfu Then
        myMessage = myMessage & _
                               "�`���[�u�̃T�u���ƒ[�������X�V���܂��B" & vbCrLf & vbCrLf & _
                               "�f�[�^��: ���̃u�b�N�̃V�[�g[���i���X�g,�[���ꗗ]" & vbCrLf & _
                               "�o�͐�F" & myAddress(3, 1) & vbCrLf & vbCrLf & _
                               "�`���[�u�G�t���SYS�ŕt�^����T�u���ł��B" & vbCrLf & vbCrLf
    End If
    If CheckBox_partsEfu Then
        myMessage = myMessage & _
                               "�p�[�c�̃T�u���ƒ[�������X�V���܂��B" & vbCrLf & vbCrLf & _
                               "�f�[�^��: [�[���ꗗ]=�T�u��,[���i���X�g]=����ȊO" & vbCrLf & _
                               "�o�͐�F���̃u�b�N��[���i�G�t](�b��)" & vbCrLf & vbCrLf & _
                               "�⋋�i���i�Ǘ��ŕt�^����T�u���ł��B" & vbCrLf & vbCrLf
    End If
    
    If myMessage = "" Then Exit Sub
    Dim a As Long
    a = MsgBox(myMessage, vbYesNo, "�T�u�i���o�[�X�V")
    If a = 6 Then
        Unload Me
        If CheckBox_wireEfu Then Call PVSWcsv����G�t����p�T�u�i���o�[txt�o��_Ver2012(myIP, CheckBox_stepNumberAdd)
        If CheckBox_tubeEfu Then Call export_tubeEfu(myIP)
        If CheckBox_partsEfu Then Call export_partEfu(���i�i��str, �ݕ�str)
        MsgBox "�o�͂��܂���"
    End If
End Sub

Private Sub CommandButton4_Click()
    PlaySound "���ǂ�"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    '�폜����
    Set wb(0) = ThisWorkbook
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "��������"
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "")
    
    Unload Me
    
    Call checkSheet("PVSW_RLTF;�[���ꗗ", wb(0), True, True)
    
    '�[���ꗗ����g�p����T�u�i���o�[���Q�b�g
    With wb(0).Sheets("�[���ꗗ")
        Dim myKey As Variant, i As Long, �[�� As String, �T�uran() As Variant, foundFlag As Boolean, �T�u As String
        ReDim �T�uran(0, 0)
        Set myKey = .Cells.Find(���i�i��Ran(1, 1), , , 1)
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            �T�u = .Cells(i, myKey.Column)
            foundFlag = False
            If �T�u <> "" Then
                For x = LBound(�T�uran, 2) To UBound(�T�uran, 2)
                    If �T�uran(0, x) = �T�u Then
                        foundFlag = True
                        Exit For
                    End If
                Next x
                If foundFlag = False Then
                    ReDim Preserve �T�uran(0, UBound(�T�uran, 2) + 1)
                    �T�uran(0, UBound(�T�uran, 2)) = �T�u
                End If
            End If
        Next i
        If UBound(�T�uran, 2) = 0 Then
            MsgBox "[�[���ꗗ]�ɃT�u�i���o�[������܂���B"
            Stop
        End If
        �T�uran = WorksheetFunction.transpose(�T�uran) 'bubbleSort2�ׂ̈ɓ���ւ���
        Call BubbleSort2(�T�uran, 1)
        �T�uran = WorksheetFunction.transpose(�T�uran) 'bubbleSort2�ׂ̈ɓ���ւ���
    End With
    
    '�[���ꗗ����[�������̃T�u�i���o�[���Q�b�g
    With wb(0).Sheets("�[���ꗗ")
        Dim �[���T�uRAN()
        ReDim �[���T�uRAN(1, 0)
        Dim �[��Col As Long: �[��Col = .Cells.Find("�[����", , , 1).Column
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            �T�u = .Cells(i, myKey.Column)
            �[�� = .Cells(i, �[��Col)
            If �T�u <> "" Then
                ReDim Preserve �[���T�uRAN(1, UBound(�[���T�uRAN, 2) + 1)
                �[���T�uRAN(0, UBound(�[���T�uRAN, 2)) = �T�u
                �[���T�uRAN(1, UBound(�[���T�uRAN, 2)) = �[��
            End If
        Next i
    End With
    
    'PVSW_RLTF����������Q�b�g
    Set myKey = ws(0).Cells.Find(���i�i��Ran(1, 1), , , 1)
    '�g�p����t�B�[���h���̃Z�b�g
    Dim fieldname As String: fieldname = myKey.Value & ",RLTFtoPVSW_,�n�_���[�����ʎq,�I�_���[�����ʎq,�n�_���L���r�e�B,�I�_���L���r�e�B,�ڑ�G_,���[�n��,�\��_"
    ff = Split(fieldname, ",")
    ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    a = UBound(ff) + 2
    
    Dim lastRow As Long
    lastRow = ws(0).Cells(ws(0).Rows.count, myKey.Column).End(xlUp).Row
    
    '�T�u�i���o�[���ɓd���������Z�b�g���Ă���
    Dim myRan() As Variant, y As Long, �T�ustr As String, r As Long
    ReDim myRan(a, 0)
    For y = LBound(�T�uran) + 1 To UBound(�T�uran)
        For x = 0 To 1
            For i = myKey.Row + 1 To lastRow
                �T�ustr = ws(0).Cells(i, myKey.Column).Value
                ���[�n�� = ws(0).Cells(i, f(7)).Value
                If �T�uran(y) = �T�ustr Then
                    If ���[�n�� = CStr(x) Then
                        ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
                        For r = LBound(myRan) To UBound(myRan) - 2
                            myRan(r, UBound(myRan, 2)) = ws(0).Cells(i, f(r)).Value
                        Next r
                    End If
                End If
            Next i
        Next x
    Next y
    
    Call �œK�����ǂ�
    PlaySound "���񂹂�"
    
    Dim myMsg As String: myMsg = "�쐬���܂���" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "��n���U��_SSC����")
End Sub

Private Sub myLabel_Click()
    
End Sub

Private Sub Label8_Click()

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
