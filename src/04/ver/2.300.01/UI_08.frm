VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_08 
   Caption         =   "�T�u����"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
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

    Unload Me
    Call checkSheet("PVSW_RLTF;�[���ꗗ", wb(0), True, True)
    
    Call PVSWcsv���[�̃V�[�g�쐬_Ver2001
    Call PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬_2017
    
    '�g�p����t�B�[���h���̃Z�b�g
    Dim fieldname As String: fieldname = "RLTFtoPVSW_,�n�_���[�����ʎq,�I�_���[�����ʎq,�n�_���[�����i��,�I�_���[�����i��,�d�㐡�@_,�ڑ�G_,�\��_,����_"
    ff = Split(fieldname, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    a = UBound(ff) + 1
    '�d�������Z�b�g����z��
    Dim �[���d����RAN As Variant
    ReDim �[���d����RAN(a, 0)
    '�t�B�[���h����z��ɓ����
    For x = LBound(ff) To UBound(ff)
        �[���d����RAN(x, 0) = ff(x)
    Next x
    �[���d����RAN(UBound(�[���d����RAN), 0) = "�e�[��No"
    
    '�Ώۂ̃O���[�v���ɏ���
    Dim ���C���i��i As Integer
    ���C���i��i = ���i�i��RAN_read(���i�i��Ran, "���C���i��")
    For i = LBound(���i�i��Ran, 2) + 1 To UBound(���i�i��Ran, 2)
        ���i�i��str = ���i�i��Ran(���C���i��i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '���i�i�Ԃ̃t�B�[���h���L�[�Ƃ��ăZ�b�g
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(���i�i��str, , , 1)
            Dim lastRow As Long
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    If .Cells(y, f(0)) = "Found" Then
                        ReDim Preserve �[���d����RAN(a, UBound(�[���d����RAN, 2) + 1)
                        For x = LBound(f) + 1 To UBound(f)
                            �[���d����RAN(x, UBound(�[���d����RAN, 2)) = .Cells(y, f(x))
                        Next x
                        '�[���d����RAN(0, UBound(�[���d����RAN, 2)) = 1
                    End If
                End If
            Next y
        End With
       Call ReplaceLR(�[���d����RAN)
    
        Call SumRan(�[���d����RAN) '���[�̍s���悪�����A�ڑ�G�������ꍇ�܂Ƃ߂�
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
                For x = 1 To 2
                    If �[��str = �[���d����RAN(ii, x) Then
                        �[���d����RAN(ii, UBound(�[���d����RAN, 2)) = �T�ustr
                        Exit For
                    End If
                Next x
            Next ii
        End If
    Next i
    
    'Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    'todo
    '�T�u�i���o�[��z�z
    For ii = LBound(�[���d����RAN) To UBound(�[���d����RAN)
        '�ڑ�G�ɂ�锻�f
        Select Case Left(�[���d����RAN(ii, 6), 1)
            Case "T"
                '�������Ȃ�
            Case "E", "J", "B"
                �[���d����RAN(ii, UBound(�[���d����RAN, 2)) = "999"
            Case "W"
                �[���d����RAN(ii, UBound(�[���d����RAN, 2)) = "999"
        End Select
        '����_�ɂ�锻�f
        Select Case Left(�[���d����RAN(ii, 8), 1)
            Case "E"
                �[���d����RAN(ii, UBound(�[���d����RAN, 2)) = "999"
        End Select
    Next ii
    'Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    
    '�]���̍����e����ɃT�u�i���o�[��z�z���Ă���
    For i = LBound(�[���]��RAN) + 1 To UBound(�[���]��RAN)
        �[��str = �[���]��RAN(i, 0)
        ����[����str = �[���]��RAN(i, 6)
        'If �[��str = "250" Then Stop
        If �[���]��RAN(i, 5) <> "" Then GoTo line20
        �[���]��RAN(i, 5) = �[��str
        For j = LBound(�[���d����RAN) + 1 To UBound(�[���d����RAN)
'            If j = 227 Then Stop
'            If i = 3 And j = 207 Then Stop
            If �[���d����RAN(j, UBound(�[���d����RAN, 2)) = "" Then '�܂��T�u�i���o�[�����܂��Ė������
                For x = 1 To 2
                    If �[��str = �[���d����RAN(j, x) Then
                        �[���]��lng = �[���d����RAN(j, 0)
                        If x = 1 Then ����[��str = �[���d����RAN(j, 2)
                        If x = 2 Then ����[��str = �[���d����RAN(j, 1)
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
                            If �[���d����RAN(j, UBound(�[���d����RAN, 2)) = "" Then
                                �[���d����RAN(j, UBound(�[���d����RAN, 2)) = �[��str
                            End If
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
                Next x
            End If
        Next j
line20:
    Next i
    
    'Call export_ArrayToSheet(�[���]��RAN, "�[���]��RAN", False)
    
    'Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    '�q����Ȃ������d����]���̍����[���ɑ}���悤�ɂ���
    For i = LBound(�[���d����RAN) + 1 To UBound(�[���d����RAN)
        If �[���d����RAN(i, 0) <> "" Then
            If �[���d����RAN(i, UBound(�[���d����RAN, 2)) = "" Then
                �[��1str = �[���d����RAN(i, 1)
                �[��2str = �[���d����RAN(i, 2)
                �[���]��1str = search�[���]��RAN(�[���]��RAN, �[��1str, 2)
                �[���]��2str = search�[���]��RAN(�[���]��RAN, �[��2str, 2)
                If �[���]��1str > �[���]��2str Then
                    �����T�ustr = search�[���]��RAN(�[���]��RAN, �[��1str, 5)
                Else
                    �����T�ustr = search�[���]��RAN(�[���]��RAN, �[��2str, 5)
                End If
                �[���d����RAN(i, UBound(�[���d����RAN, 2)) = �����T�ustr
            End If
        End If
    Next i
    
    'Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    
    '�[��888�̔z�z
    Dim �[��str1 As String, �[��str2 As String, �ڑ�Gstr As String
    For i = LBound(�[���d����RAN) + 1 To UBound(�[���d����RAN)
        �[��str1 = �[���d����RAN(i, 1)
        �[��str2 = �[���d����RAN(i, 2)
        If �[��str1 & �[��str2 = "" Then �[���d����RAN(i, UBound(�[���d����RAN, 2)) = "999"
    Next i
    
'    Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    
    'RLFTtoPVSW_���󗓂̏ꍇ���O����
    For i = LBound(�[���d����RAN) To UBound(�[���d����RAN)
        If i > UBound(�[���d����RAN) Then Exit For
        If �[���d����RAN(i, 0) = "" Then
            �[���d����RAN = removeArrayIndex(�[���d����RAN, i)
        End If
    Next i
    
   'Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    
    '�[���̃T�u�i���o�[���d���̃T�u�i���o�[�ɖ����ꍇ�Ac�ɂ���
    Dim foundFlg As Boolean
    For i = LBound(�[���]��RAN) + 1 To UBound(�[���]��RAN)
        foundFlg = False
        �T�ustr = �[���]��RAN(i, 5)
        For ii = LBound(�[���d����RAN) + 1 To UBound(�[���d����RAN)
            If �T�ustr = �[���d����RAN(ii, UBound(�[���d����RAN, 2)) Then
                foundFlg = True
                Exit For
            End If
        Next ii
        If foundFlg = False Then
            �[���]��RAN(i, 5) = "c"
        End If
    Next i
    
    '�[���]��RAN���e�L�X�g�o��
    Dim myTextPath As String
    myTextPath = wb(0).path & dirString_09
    makeDir myTextPath
    myTextPath = myTextPath & Replace(���i�i��str, " ", "") & "_term.txt"
    export_Array_ShiftJis �[���]��RAN, myTextPath, ","
    
    '�[���d����RAN���e�L�X�g�o��
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(���i�i��str, " ", "") & "_wiresum.txt"
    export_Array_ShiftJis �[���d����RAN, myTextPath, ","
    
    'Call export_ArrayToSheet(�[���d����RAN, "�[���d����RAN", False)
    
    '�d�����ɃT�u�i���o�[(��Ə�)�ƃX�e�b�v�i���o�[�����߂�
    Dim myRan As Variant
    myRan = setWorkRanV2(���i�i��str)
    
'    Call export_ArrayToSheet(�[���]��RAN, "�[���]��RAN", False)
    
    '�d�����̃T�u�i���o�[��[�����̃T�u�i���o�[�ɓn��
    Dim subNumber As String, �e�[��str As String
    For y = LBound(�[���]��RAN) To UBound(�[���]��RAN)
        �e�[��str = �[���]��RAN(y, 5)
        If �e�[��str = "c" Then
            subNumber = "c"
        Else
            subNumber = searchRan_ver2(myRan, �e�[��str, "�e�[��No", "subNumber")
        End If
        �[���]��RAN(y, UBound(�[���]��RAN, 2)) = subNumber
    Next y
    �[���]��RAN = WorksheetFunction.transpose(�[���]��RAN)
    
    'Call export_ArrayToSheet(myRan, "myRan", True)
    
    'PVSW_RLTF�ɃT�u�i���o�[��z�z
    For i = LBound(���i�i��Ran, 2) + 1 To UBound(���i�i��Ran, 2)
        ���i�i��str = ���i�i��Ran(���C���i��i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '���i�i�Ԃ̃t�B�[���h���L�[�Ƃ��ăZ�b�g
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(���i�i��str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    If .Cells(y, f(0)) = "Found" Then
                        �\��str = Left(.Cells(y, f(7)), 4)
                        subNumber = searchRan_ver2(myRan, �\��str, "�\��_", "subNumber")
                        .Cells(y, myKey.Column) = subNumber
                        .Cells(y, myKey.Column).Interior.color = theme_color1
                    End If
                End If
            Next y
        End With
    Next i
    
    '�[���ꗗ�ɃT�u�i���o�[��z�z
    For i = LBound(���i�i��Ran, 2) + 1 To UBound(���i�i��Ran, 2)
        ���i�i��str = ���i�i��Ran(���C���i��i, i)
        Set ws(3) = wb(0).Sheets("�[���ꗗ")
        With ws(3)
            Dim myCol(1) As Integer
            myCol(0) = .Cells.Find("�[�����i��", , , 1).Column
            myCol(1) = .Cells.Find("�[����", , , 1).Column
            '���i�i�Ԃ̃t�B�[���h���L�[�Ƃ��ăZ�b�g
            Set myKey = ws(3).Cells.Find(���i�i��str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    �[�����i��str = .Cells(y, myCol(0))
                    �[��str = .Cells(y, myCol(1))
                    subNumber = searchRan_ver2(�[���]��RAN, �[��str & "," & �[�����i��str, "�[��No,�[�����i��", "subNumber")
                    If subNumber = "" Then Stop
                    .Cells(y, myKey.Column) = subNumber
                    .Cells(y, myKey.Column).Interior.color = theme_color1
                End If
            Next y
        End With
    Next i
    
    '�z����e�L�X�g�t�@�C���o�͂���
    '�[���]��RAN���e�L�X�g�o��
    �[���]��RAN = WorksheetFunction.transpose(�[���]��RAN)
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(���i�i��str, " ", "") & "_term.txt"
    export_Array_ShiftJis �[���]��RAN, myTextPath, ","
    
    '�[���d����RAN���e�L�X�g�o��
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(���i�i��str, " ", "") & "_wireSum.txt"
    export_Array_ShiftJis �[���d����RAN, myTextPath, ","
    
    'myRAN���e�L�X�g�o��
    myRan = WorksheetFunction.transpose(myRan)
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(���i�i��str, " ", "") & "_wire.txt"
    export_Array_ShiftJis myRan, myTextPath, ","
    
    Call �œK�����ǂ�
    PlaySound "���񂹂�"
'
'    addRow = 1
'    For y = LBound(�[���d����RAN) To UBound(�[���d����RAN)
'            For x = LBound(�[���d����RAN, 2) To UBound(�[���d����RAN, 2)
'                With Sheets("temp")
'                    .Cells(addRow, x + 1) = �[���d����RAN(y, x)
'                End With
'            Next x
'            addRow = addRow + 1
'    Next y
'
'    addRow = 1
'    For y = LBound(�[���]��RAN) To UBound(�[���]��RAN)
'        For x = LBound(�[���]��RAN, 2) To UBound(�[���]��RAN, 2)
'            With Sheets("temp2")
'                .Cells(addRow, x + 1) = �[���]��RAN(y, x)
'            End With
'        Next x
'        addRow = addRow + 1
'    Next y

    Dim myMsg As String: myMsg = "�������܂���" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "���Y����+�T�u��������")
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
