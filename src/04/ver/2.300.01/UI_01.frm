VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_01 
   Caption         =   "�n���}�쐬"
   ClientHeight    =   8430
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9580
   OleObjectBlob   =   "UI_01.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





















































Private Sub B0_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 4
    CB1.ListIndex = 1
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
    'PIC00.Picture = LoadPicture(myaddress(0,1) & "\�n���}sample_" & "4511000000" & ".jpg")
End Sub

Private Sub B1_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 1
    CB1.ListIndex = 4
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub B2_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 2
    CB1.ListIndex = 0
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 1
    CB5.ListIndex = 5
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = True
    cbxQR.Value = False
End Sub

Private Sub B3_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 0
    CB1.ListIndex = 0
    CB2.ListIndex = 0
    CB3.ListIndex = 0
    CB4.ListIndex = 0
    CB5.ListIndex = 0
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 0
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub B4_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 1
    CB1.ListIndex = 2
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 3
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 2
    cbx0.Value = False
    cbx1.Value = True
    cbx2.Value = False
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub CB0_Change()
    Call CB�I��ύX
End Sub

Private Sub CB1_Change()
    Call CB�I��ύX
End Sub

Private Sub CB2_Change()
    Call CB�I��ύX
End Sub

Private Sub CB3_Change()
    Call CB�I��ύX
End Sub

Private Sub CB4_Change()
    Call CB�I��ύX
End Sub

Private Sub CB5_Change()
    If CB5.Value = "" Then Exit Sub
    With ActiveWorkbook.Sheets("���i�i��")
        Set key = .Cells.Find("�^��", , , 1)
        myCol = .Rows(key.Row).Find(CB5.Value, , , 1).Column
        lastRow = .Cells(.Rows.count, .Cells.Find("���C���i��", , , 1).Column).End(xlUp).Row
        Dim ���� As String: ���� = ""
        For i = key.Row + 1 To lastRow
            If InStr(����, "," & .Cells(i, myCol) & ",") = 0 Then
                ���� = ���� & "," & .Cells(i, myCol) & ","
            End If
        Next i
    
    End With
    ���� = Mid(����, 2)
    ���� = Left(����, Len(����) - 1)
    ����s = Split(����, ",,")
    With CB6
        .RowSource = ""
        .Clear
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = -1
    End With
    Call CB�I��ύX
End Sub

Public Function CB�I��ύX()
    
    If CB5.Value = "���C���i��" Then ��� = 0 Else ��� = 1
    
    ��� = ��� & CB0.ListIndex & CB1.ListIndex & CB2.ListIndex & CB3.ListIndex & CB4.ListIndex & CB9.ListIndex & CB8.ListIndex
    
    If CB0.ListIndex = 0 Then
        PIC00.Visible = False
    Else
        PIC00.Visible = True
        On Error Resume Next
        PIC00.Picture = LoadPicture(myAddress(0, 1) & "\menu\" & ��� & ".bmp")
        On Error GoTo 0
    End If
    If �T���v���쐬���[�h = True Then
        sample.Visible = True
        sample.Caption = ���
    End If
End Function

Public Function �}�W�b�N�I��ύX()
    Call addressSet(wb(0))
    
    ��� = CB8.Value
    
    ��� = Replace(���, "-1", "0")
    
    If Left(���, 1) = "0" Then ��� = "0000000000"
    
    PIC01.Picture = LoadPicture(myAddress(0, 1) & "\" & ��� & ".jpg")
End Function

Private Sub CheckBox1_Click()

End Sub

Private Sub CBa_Change()

End Sub

Private Sub CB6_Change()
    Call CB�I��ύX
End Sub

Private Sub CB8_Change()
    Call CB�I��ύX
End Sub

Private Sub cbx0_Click()
    If cbx0.Value = True Then
        cbxQR.Visible = True
    Else
        cbxQR.Visible = False
        cbxQR.Value = False
    End If
End Sub

Private Sub cbx1_Change()
    If cbx1.Value = True Then
        CB7.Visible = True
        CB7.ListIndex = 0
    Else
        CB7.Visible = False
        CB7.ListIndex = -1
    End If
End Sub

Private Sub cbx2_Click()
    If cbx2.Value = True Then
        CB9.Visible = True
        Label9.Visible = True
    Else
        CB9.Visible = False
        Label9.Visible = False
    End If
End Sub

Private Sub cbx4_Click()
    If cbx4 = True Then
        If Dir(myAddress(3, 1)) = "" Then
            �R�����g.Visible = True
            �R�����g.Caption = "��n����ƎҎ擾�悪������܂���B�ݒ���m�F���Ă��������B"
            �R�����g.ForeColor = RGB(255, 0, 0)
        End If
    Else
        �R�����g.Caption = "��n����ƎҎ擾�悪������܂����B"
        �R�����g.ForeColor = 0
        �R�����g.Visible = False
    End If
End Sub

Private Sub cbx7_Click()

End Sub

Private Sub CommandButton1_Click()
    �t�H�[������̌Ăяo�� = True
    mytime = time
    
    If CB6.ListIndex = -1 Then
        �R�����g.Visible = True
        �R�����g.Caption = "�쐬�Ώۂ��w�肳��Ă��܂���B"
        Beep
        Exit Sub
    End If
    
    If �R�����g.ForeColor = RGB(255, 0, 0) Then
        MsgBox "Sheet[�ݒ�]�̌�n����Ǝ҈ꗗ�擾_�̃A�h���X��������܂���B"
        Exit Sub
    End If
    �[���i���o�[�\�� = True
    
    PlaySound ("��������")
    cb�I�� = CB0.ListIndex
    cb�I�� = cb�I�� & "," & CB1.ListIndex
    cb�I�� = cb�I�� & "," & CB2.ListIndex
    cb�I�� = cb�I�� & "," & CB3.ListIndex
    cb�I�� = cb�I�� & "," & CB4.ListIndex
    cb�I�� = cb�I�� & "," & CB9.ListIndex
    cb�I�� = cb�I�� & "," & CB10.ListIndex
    
    �}���}�`�� = CB8.List(CB8.ListIndex, 1)
    
    Call ���i�i��RAN_set2(���i�i��Ran, CB5.Value, CB6.Value, "")
    
    �F�Ŕ��f = cbx2
    ��d�W�~flg = cbx5
    ��n����Ǝ� = cbx4
    QR��� = cbxQR
    is_setProcessColor = cbx6
    ��n���_�� = cbx7
    
    ���^�p�x����flag = True
    
    If ���i�i��RANc = 0 Then
        �R�����g.Visible = True
        �R�����g.Caption = "�Y�����鐻�i�i�Ԃ�����܂���B" & vbCrLf _
                         & "�Ⴆ�ΑI�������������A" & vbCrLf & "[PVSW_RLTF]�ɍ݂�܂���B"
        Beep
        Exit Sub
    End If
    
    Set myBook = ActiveWorkbook
    
    If ���i�i��RANc <> 1 And cbx0.Value = True Then
        �R�����g.Visible = True
        �R�����g.Caption = "�i�Ԃ���������ׁA�T�u�}�쐬�s�B"
        Beep
        Exit Sub
    End If
    
    Unload UI_01
    
    Call PVSWcsv���[�̃V�[�g�쐬_Ver2001
    
    If ��n����Ǝ� = True Then
        If ���i�i��RANc = 1 Then
            myAddress(3, 1) = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "��n����ƎҎ擾"), 1)
            If Dir(myAddress(3, 1)) <> "" Then
                Set wb(3) = Workbooks.Open(fileName:=myAddress(3, 1), UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
                Call SQL���ǂ�_��n����Ǝ�(��n����Ǝ�ran, CB6.Value)
                Application.DisplayAlerts = False
                wb(3).Close
                Application.DisplayAlerts = True
                '��n���_�łׂ̈Ɍ�n�����RAN��CAV�̏��Ƃ�����Ƃ�
                Dim tempArray As Variant
                tempArray = readSheetToRan3(wb(0).Sheets("PVSW_RLTF���["), "�d�����ʖ�", CB6.Value & ",RLTFtoPVSW_,�[�����ʎq,�L���r�e�B,�\��_,�n��", "", 1)
                
                For i = LBound(��n����Ǝ�ran, 2) + 1 To UBound(��n����Ǝ�ran, 2)
                    For ii = LBound(tempArray, 2) To UBound(tempArray, 2)
                        If tempArray(5, ii) = "��" Then
                            If ��n����Ǝ�ran(0, i) = tempArray(4, ii) Then
                                ��n����Ǝ�ran(3, i) = tempArray(2, ii)
                                ��n����Ǝ�ran(4, i) = tempArray(3, ii)
                                Exit For
                            End If
                        End If
                    Next ii
                Next i
            End If
        Else
            MsgBox "���i�i�Ԃ��P�_�𒴂���ꍇ�A��n����Ǝ҂̕\���͖����Ή����Ă��܂���B"
            Exit Sub
        End If
    End If
    
    
    '�T���v���쐬���[�h
    Dim cb5str(1) As String: cb5str(0) = "���C���i��": cb5str(1) = "����"
    Dim cb6str(1) As String: cb6str(0) = "8211136Y82     ": cb6str(1) = "G"
    If �T���v���쐬���[�h = True Then
        For i0 = 1 To CB0.ListCount - 1
            For i1 = 0 To CB1.ListCount - 1
                For i2 = 0 To CB2.ListCount - 1
                    For i3 = 0 To CB3.ListCount - 1
                        For i4 = 0 To CB4.ListCount - 1
                            For i8 = 0 To CB8.ListCount - 1
                                i9 = -1
                                �}���}�`�� = CB8.List(i8, 1)
                                cb�I�� = i0 & "," & i1 & "," & i2 & "," & i3 & "," & i4 & "," & i9
                                For ii = 0 To 1
                                    Call ���i�i��RAN_set2(���i�i��Ran, cb5str(ii), cb6str(ii), "")
                                    Call �n���}�쐬_Ver220098(cb�I��, cb5str(ii), cb6str(ii))
                                  
                                    On Error Resume Next
                                    ActiveSheet.Shapes.Range("324_1").Select
                                    If err.number = 1004 Then
                                        ActiveSheet.Shapes.Range("324_7").Select
                                    End If
                                    On Error GoTo 0
                                    Call �摜�Ƃ��ďo��(ii & Replace(cb�I��, ",", "") & i8)
                                Next ii
                            Next i8
                        Next i4
                    Next i3
                Next i2
            Next i1
        Next i0
    Else
        Call �n���}�쐬_Ver220098(cb�I��, CB5.Value, CB6.Value)
    End If
    
    
    
    
    
    If cbx7.Value = True Then Call ��n����Ǝҕ�_�_�ŉ摜�쐬(CB6.Value, CB12.Value)
    If cbx3.Value = True Then Call ���������V�X�e���p�f�[�^�쐬v2182(CB6.Value)
    If cbx1.Value = True Then Call �n���}�̈���p�f�[�^�쐬(CB7.Value, CB5.Value & Replace(CB6.Value, " ", ""))
    'MsgBox "�쐬���������܂����B"
    
    If cbx0.Value = True Then
        msg = �T�u�}�쐬_Ver220116(CB6.Value, ���i�i��Ran)
        If msg <> "" Then
            DoEvents
            a = MsgBox("���̃T�u��[�[���ꗗ]�ɂ�����[PVSW_RLTF]�ɂ���܂���B" & vbCrLf & "�}�������ŕύX�ɂȂ��Ă��܂��񂩁H" & vbCrLf & vbCrLf & msg, , "�A���}�b�`�G���[")
        End If
    End If
    
    If �}���}�s�� <> "" Then
        �}���}�s��sp = Split(�}���}�s��, "_")
        msg = "���̒[���Ń}���}�̐����s���B�}���}�������Ă�����ɏI�����Ă��܂���B"
        msg = msg & Join(�}���}�s��sp, vbCrLf)
    End If
    
    Call ���O�o��("test", "test", "�n���}" & cb�I�� & CB5.Value & CB6.Value)
    
    myBook.Activate
    If �T���v���쐬���[�h = True Then
    
    Else
        MsgBox "�쐬����= " & DateDiff("s", mytime, time) & " s"
    End If
    
    'Call �������܂���(myBook)
    'mybook.VBProject.VBComponents(Sheets("�n���}_" & CB5.Value & Replace(CB6.Value, " ", "")).CodeName).CodeModule.AddFromFile myaddress(0,1) & "\002_��A���쐬_�}���}.txt"
End Sub

Private Sub OptionButton1_Click()
    
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("���ǂ�")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 4
    CB1.ListIndex = 1
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub CommandButton6_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 2
    CB1.ListIndex = 0
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 0
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = True
    cbx4.Value = False
    cbx5.Value = True
    cbxQR.Value = False
End Sub

Private Sub ST01_Change()

End Sub

Private Sub PIC_Click()

End Sub

Private Sub CommandButton7_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 6
    CB1.ListIndex = 2
    CB2.ListIndex = 0
    CB3.ListIndex = 0
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = 13

    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 0
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = True
    cbx5.Value = False
    cbx7.Value = True
    cbxQR.Value = False
    
    
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub PIC00_Click()

End Sub

Private Sub UserForm_Initialize()
    Set wb(0) = ThisWorkbook
    Dim ����(12) As String
    ����(0) = "�}���쐬���Ȃ�,�d���T�C�Y�̂�,�|�C���g,��H����,�\��,����[��,��n����ƃi���o�["
    ����(1) = "�������Ȃ�,��n���͐Ԑ�,��n���͏���������,��n���͓h��Ԃ�,��n���̂ݕ\��"
    ����(2) = "�\�����Ȃ�,��n�����i(�H��40)"
    ����(3) = "�ϊ����Ȃ�,�ϊ�����"
    ����(4) = "�g�p���Ȃ�,�g�p����"
    ����(6) = "A4-�^�e,A4-��,A3-�^�e,A3-��"
    ����(8) = "Tear,Oval,Heart" '�}���}�̌`��
    ����(10) = "160,9,21" '�}���}�̔ԍ�
    ����(11) = "�\�����Ȃ�,��n����=0�Ȃ�\��,��n���� <> 0�Ȃ�\��"
    
    ����(12) = "2�E"
        
    With ActiveWorkbook.Sheets("�ݒ�")
        Set myKey = .Cells.Find("�n���F_", , , 1)
        lastRow = myKey.Offset(0, 1).End(xlDown).Row - myKey.Row
        For i = 0 To lastRow
            ����(9) = ����(9) & "," & myKey.Offset(i, 2)
        Next i
        ����(9) = Mid(����(9), 2)
    End With
    
    With ActiveWorkbook.Sheets("���i�i��")
        Set myKey = .Cells.Find("�^��", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For x = myKey.Column To lastCol
            If .Cells(myKey.Row, x).Offset(-1, 0) = 1 Then
                ����(5) = ����(5) & "," & .Cells(myKey.Row, x)
            End If
        Next x
        ����(5) = Mid(����(5), 2)
        Set myKey = Nothing
    End With
    
    ����s = Split(����(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With

    ����s = Split(����(1), ",")
    With CB1
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(2), ",")
    With CB2
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(3), ",")
    With CB3
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(4), ",")
    With CB4
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(5), ",")
    With CB5
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(6), ",")
    With CB7
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(8), ",")
    ����s2 = Split(����(10), ",")
    With CB8
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem
             .List(i, 0) = ����s(i)
             .List(i, 1) = ����s2(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(9), ",")
    With CB9
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = -1
    End With
    ����s = Split(����(11), ",")
    With CB10
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
    
    ����s = Split(����(12), ",")
    With CB12
        .RowSource = ""
        For i = LBound(����s) To UBound(����s)
            .AddItem ����s(i)
        Next i
        .ListIndex = 0
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub

Private Sub �R�����g_Click()

End Sub


