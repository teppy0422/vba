VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_08 
   Caption         =   "�T�u����"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   OleObjectBlob   =   "UI_08.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_08"
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
        For Y = myKey.Row + 1 To lastRow
            If InStr(����(0), "," & .Cells(Y, myKey.Column)) & "," = 0 Then
                ����(0) = ����(0) & "," & .Cells(Y, myKey.Column) & ","
                ����2(0) = ����2(0) & "," & .Cells(Y, mykey2.Column) & ","
            End If
        Next Y
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
    Call ���i�i��RAN_set2(���i�i��RAN, CB0.Value, CB1.Value, "")
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
    mytime = Time
    PlaySound "��������"
    Call ���i�i��RAN_set2(���i�i��RAN, CB0.Value, CB1.Value, "")

    Unload Me
    Set wb(0) = ActiveWorkbook
    
    '�g�p����t�B�[���h���̃Z�b�g
    Dim fieldName As String: fieldName = "RLTFtoPVSW_,�n�_���[�����ʎq,�I�_���[�����ʎq,�n�_���[�����i��,�I�_���[�����i��,�d�㐡�@_"
    ff = Split(fieldName, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    For X = LBound(ff) To UBound(ff)
        f(X) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(X), , , 1).Column
    Next X
    a = UBound(ff) + 1
    '�d�������Z�b�g����z��
    Dim �[���d����RAN As Variant
    ReDim �[���d����RAN(a, 0)
    '�t�B�[���h����z��ɓ����
    For X = LBound(ff) To UBound(ff)
        �[���d����RAN(X, 0) = ff(X)
    Next X
    '�Ώۂ̃O���[�v���ɏ���
    Dim ���C���i��i As Integer
    ���C���i��i = ���i�i��RAN_read(���i�i��RAN, "���C���i��")
    For i = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
        ���i�i��str = ���i�i��RAN(���C���i��i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '���i�i�Ԃ̃t�B�[���h���L�[�Ƃ��ăZ�b�g
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(���i�i��str, , , 1)
            Dim lastRow As Long
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    If .Cells(Y, f(0)) = "Found" Then
                        ReDim Preserve �[���d����RAN(a, UBound(�[���d����RAN, 2) + 1)
                        For X = LBound(f) + 1 To UBound(f)
                            �[���d����RAN(X, UBound(�[���d����RAN, 2)) = .Cells(Y, f(X))
                        Next X
                        '�[���d����RAN(0, UBound(�[���d����RAN, 2)) = 1
                    End If
                End If
            Next Y
        End With
       Call ReplaceLR(�[���d����RAN)
       Call SumRan(�[���d����RAN)
       '��H�̑�������e�����߂�
       Dim �[���]��RAN()
       �[���]��RAN = evaluationRan(�[���d����RAN) '�D�悪9�̎��T�u�i���o�[999
       �[���]��RAN = changeRowCol(�[���]��RAN)
       Call BubbleSort3(�[���]��RAN, 3, 2)
       �[���d����RAN = changeRowCol(�[���d����RAN)
    Next i
    
    '�[���T�u�i���o�[999�̓d���T�u�i���o�[��999�ɂ���
    For i = LBound(�[���]��RAN) To UBound(�[���]��RAN)
        �[��str = �[���]��RAN(i, 0)
        �T�ustr = �[���]��RAN(i, 5)
        If �T�ustr <> "" Then
            For ii = LBound(�[���d����RAN) To UBound(�[���d����RAN)
                For X = 1 To 2
                    If �[��str = �[���d����RAN(ii, X) Then
                        �[���d����RAN(ii, 6) = �T�ustr
                        Exit For
                    End If
                Next X
            Next ii
        End If
    Next i
    
    '�]���̍����e����ɃT�u�i���o�[��z�z���Ă���
    For i = LBound(�[���]��RAN) + 1 To UBound(�[���]��RAN)
        �[��str = �[���]��RAN(i, 0)
        ����[����str = �[���]��RAN(i, 6)
        'If �[��str = "250" Then Stop
        If �[���]��RAN(i, 5) <> "" Then GoTo line20
        �[���]��RAN(i, 5) = �[��str
        For j = LBound(�[���d����RAN) + 1 To UBound(�[���d����RAN)
            If �[���d����RAN(j, 6) = "" Then '�܂��T�u�i���o�[�����܂��Ė������
                For X = 1 To 2
                    If �[��str = �[���d����RAN(j, X) Then
                        �[���]��lng = �[���d����RAN(j, 0)
                        If X = 1 Then ����[��str = �[���d����RAN(j, 2)
                        If X = 2 Then ����[��str = �[���d����RAN(j, 1)
                        '��������[������1�̏ꍇ�ɑ���[���̃T�u�i���o�[�ɕύX
                        If ����[����str = "1" Then
                            ����[���T�ustr = search�[���]��RAN(�[���]��RAN, ����[��str, 5)
                            If ����[���T�ustr <> "" Then
                                �[���]��RAN(i, 5) = ����[���T�ustr
                                GoTo line20
                            End If
                        End If
                        ����[���D�� = search�[���]��RAN(�[���]��RAN, ����[��str, 3)
                        If ����[���D�� = "1" Then GoTo line15
                        ����[���]��lng = search����[���]��(�[���d����RAN, ����[��str)
                        If �[���]��lng >= ����[���]��lng Then
                            �[���d����RAN(j, 6) = �[��str
                            For ii = LBound(�[���]��RAN) + 1 To UBound(�[���]��RAN)
                                If �[���]��RAN(ii, 0) = ����[��str Then
                                    If �[���]��RAN(ii, 5) = "" Then
                                        �[���]��RAN(ii, 5) = �[��str
                                    End If
                                    Exit For
                                End If
                            Next ii
                        End If
                    End If
line15:
                Next X
            End If
        Next j
line20:
    Next i
    
    '�q����Ȃ������d����]���̍����[���ɑ}���悤�ɂ���
    For i = LBound(�[���d����RAN) + 1 To UBound(�[���d����RAN)
        If �[���d����RAN(i, 0) <> "" Then
            If �[���d����RAN(i, 6) = "" Then
                �[��1str = �[���d����RAN(i, 1)
                �[��2str = �[���d����RAN(i, 2)
                �[���]��1str = search�[���]��RAN(�[���]��RAN, �[��1str, 2)
                �[���]��2str = search�[���]��RAN(�[���]��RAN, �[��2str, 2)
                If �[���]��1str > �[���]��2str Then
                    �����T�ustr = search�[���]��RAN(�[���]��RAN, �[��1str, 5)
                Else
                    �����T�ustr = search�[���]��RAN(�[���]��RAN, �[��2str, 5)
                End If
                �[���d����RAN(i, 6) = �����T�ustr
            End If
        End If
    Next i
    
    addRow = 1
    For Y = LBound(�[���d����RAN) To UBound(�[���d����RAN)
        If �[���d����RAN(Y, 0) <> "" Then
            For X = LBound(�[���d����RAN, 2) To UBound(�[���d����RAN, 2)
                With Sheets("temp")
                    .Cells(addRow, X + 1) = �[���d����RAN(Y, X)
                End With
            Next X
            addRow = addRow + 1
        End If
    Next Y
    
    addRow = 1
    For Y = LBound(�[���]��RAN) To UBound(�[���]��RAN)
            For X = LBound(�[���]��RAN, 2) To UBound(�[���]��RAN, 2)
                With Sheets("temp2")
                    .Cells(addRow, X + 1) = �[���]��RAN(Y, X)
                End With
            Next X
            addRow = addRow + 1
    Next Y
    Stop
    
    'PVSW_RLTF�ɃT�u�i���o�[��z�z
    For i = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
        ���i�i��str = ���i�i��RAN(���C���i��i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '���i�i�Ԃ̃t�B�[���h���L�[�Ƃ��ăZ�b�g
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(���i�i��str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    If .Cells(Y, f(0)) = "Found" Then
                        �[��str1 = .Cells(Y, f(1))
                        �[��str2 = .Cells(Y, f(2))
                        '�[��str1�������������ɑ�����
                        swapflg = False
                        If �[��str1 = "" Then swapflg = True
                        If IsNumeric(�[��str1) = True And IsNumeric(�[��str2) = True Then
                            If Val(�[��str1) > Val(�[��str2) Then
                                swapflg = True
                            End If
                        End If
                        If swapflg = True Then
                            vSwap = �[��str2
                            �[��str2 = �[��str1
                            �[��str1 = vSwap
                        End If
                        If �[��str1 & �[��str2 <> "" Then
                            �T�ustr = search�[���d����RAN(�[���d����RAN, �[��str1, �[��str2, 6)
                            If �T�ustr = "" Then Stop '�S���҂ɘA��
                            .Cells(Y, myKey.Column) = �T�ustr
                            .Cells(Y, myKey.Column).Interior.color = RGB(129, 216, 208)
                        End If
                    End If
                End If
            Next Y
        End With
    Next i
    
    '�[���ꗗ�ɃT�u�i���o�[��z�z
    For i = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
        ���i�i��str = ���i�i��RAN(���C���i��i, i)
        Set ws(3) = wb(0).Sheets("�[���ꗗ")
        With ws(3)
            Dim myCol(1) As Integer
            myCol(0) = .Cells.Find("�[�����i��", , , 1).Column
            myCol(1) = .Cells.Find("�[����", , , 1).Column
            '���i�i�Ԃ̃t�B�[���h���L�[�Ƃ��ăZ�b�g
            Set myKey = ws(3).Cells.Find(���i�i��str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    �[�����i��str = .Cells(Y, myCol(0))
                    �[��str = .Cells(Y, myCol(1))
                    �T�ustr = search�[���]��RAN_2pos(�[���]��RAN, �[��str, �[�����i��str, 5)
                    If �T�ustr = "" Then Stop
                    .Cells(Y, myKey.Column) = �T�ustr
                    .Cells(Y, myKey.Column).Interior.color = RGB(129, 216, 208)
                End If
            Next Y
        End With
    Next i
    
    Stop
    
    Call �œK�����ǂ�
    PlaySound "���񂹂�"
    
    Dim myMsg As String: myMsg = "�쐬���܂���" & vbCrLf & DateDiff("s", mytime, Time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "���Y����+�z���U��")
End Sub

Private Sub UserForm_Initialize()
    Dim ����(1) As String
    With ActiveWorkbook.Sheets("���i�i��")
        Set myKey = .Cells.Find("�^��", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For X = myKey.Column To lastCol
            ����(0) = ����(0) & "," & .Cells(myKey.Row, X)
        Next X
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
