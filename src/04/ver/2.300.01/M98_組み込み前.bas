Attribute VB_Name = "M98_�g�ݍ��ݑO"
Public Function �ȈՃ`�F�b�J�[�p�|�C���g�i���o�[�z�zver2180()
    Set wb(0) = ActiveWorkbook
    Set ws(0) = wb(0).ActiveSheet
    ���i�i��str = "8216136D40     "
    
    Call ���i�i��RAN_set2(���i�i��Ran, "���C���i��", ���i�i��str, "")
    ����str = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "����"), 1)
    Set ws(1) = wb(0).Sheets("���_" & ����str)
    With ws(0)
        Dim prodC As Long, yazaC As Long, termC As Long, cavvC As Long, kaniC As Long, lastRow As Long, jiguC As Long
        Set myKey = .Cells.Find("�[�����i��", , , 1)
        prodC = .Rows(myKey.Row).Find(���i�i��str, , , 1).Column
        yazaC = myKey.Column
        termC = .Rows(myKey.Row).Find("�[����", , , 1).Column
        cavvC = .Rows(myKey.Row).Find("Cav", , , 1).Column
        kaniC = .Rows(myKey.Row).Find("�ȈՃ|�C���g", , , 1).Column
        jiguC = .Rows(myKey.Row).Find("����Row", , , 1).Column
        lastRow = .Cells(.Rows.count, termC).End(xlUp).Row
        .Cells(myKey.Row - 1, kaniC) = ���i�i��str
        .Cells(myKey.Row - 1, jiguC) = ����str
        '����Row�̔z�z
        Dim ����Row As Long
        For i = myKey.Row + 1 To lastRow
            �[�� = .Cells(i, termC)
            With ws(1)
                ����Row = .Cells.Find(�[��, , , 1).Row
            End With
            .Cells(i, jiguC) = ����Row
        Next i
        '����Row���Ƀ\�[�g
        
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, jiguC).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, termC).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, yazaC).addRess), Order:=xlAscending
            .add key:=Range(Cells(1, cavvC).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(myKey.Row + 1), Rows(lastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        '�|�C���g�i���o�[�̔z�z
        Dim setFlag As Boolean, myPoint As Long
        startRow = myKey.Row + 1: myPoint = 1
        For i = myKey.Row + 1 To lastRow
            ���it = .Cells(i, prodC)
            If ���it <> "" Then setFlag = True
            ��� = .Cells(i, yazaC)
            �[�� = .Cells(i, termC)
            cav = .Cells(i, cavvC)
            ���next = .Cells(i + 1, yazaC)
            �[��next = .Cells(i + 1, termC)
            If ��� & "_" & �[�� <> ���next & "_" & �[��next Then
                If setFlag = True Then
                    For ii = startRow To i
                        .Cells(ii, kaniC) = myPoint
                        myPoint = myPoint + 1
                    Next ii
                    '�����R�l�N�^��10��
                    If myPoint Mod 10 <> 0 Then
                        myPoint = (myPoint \ 10) * 10 + 11
                    Else
                        myPoint = (myPoint \ 10) * 10 + 1
                    End If
                End If
                setFlag = False
                startRow = i + 1
            End If
        Next i
    
    End With
    
End Function

Public Function �ގ��R�l�N�^�ꗗb�쐬()

    With Sheets("�[���ꗗ")
        Set myKey = .Cells.Find("�[�����i��", , , 1)
        Set mykey2 = .Cells.Find("�[����", , , 1)
        lastRow = myKey.End(xlDown).Row
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        Dim �ގ��ꗗ() As String
        ReDim �ގ��ꗗ(4, 0)
        Dim add As Long
        For i = myKey.Row + 1 To lastRow
            
            �[����� = .Cells(i, myKey.Column)
            �[�� = .Cells(i, mykey2.Column)
            For x = mykey2.Column + 1 To lastCol
                ���i�i�� = .Cells(myKey.Row, x)
                �T�u = .Cells(i, x)
                'Stop
                For r = 1 To ���i�i��RANc
                    If ���i�i�� = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "���C���i��"), r) Then

                        GoTo line10
                    End If
                Next r
                GoTo line20
line10:
                For y = LBound(�ގ��ꗗ, 2) To UBound(�ގ��ꗗ, 2)
                    If �ގ��ꗗ(0, y) = �[����� Then
                        If �ގ��ꗗ(1, y) = �[�� Then
                            If �T�u = "" Then
                                �ގ��ꗗ(4, y) = �ގ��ꗗ(4, y) & "0"
                            Else
                                �ގ��ꗗ(2, y) = �T�u
                                �ގ��ꗗ(4, y) = �ގ��ꗗ(4, y) & "1"
                            End If
                            GoTo line20
                        End If
                    End If
                Next y
line15:
                '�V�K���i�Ԃ̒ǉ�
                add = add + 1
                ReDim Preserve �ގ��ꗗ(4, add)
                �ގ��ꗗ(0, add) = �[�����
                �ގ��ꗗ(1, add) = �[��
                �ގ��ꗗ(2, add) = �T�u
                �ގ��ꗗ(3, add) = "1"
                If �T�u = "" Then
                    �ގ��ꗗ(4, add) = �ގ��ꗗ(4, add) & "0"
                Else
                    �ގ��ꗗ(4, add) = �ގ��ꗗ(4, add) & "1"
                End If
line20:
            Next x
        Next i
    End With
    
    Stop
    With ActiveWorkbook.Sheets("Sheet28")
        .Select
        .Cells.Clear
        .Cells.NumberFormat = "@"
        .Cells(2, 1) = "�[�����i��"
        .Cells(2, 4) = "�[����"
        .Cells(2, 5) = "�T�u��"
        For x = 1 To ���i�i��RANc
            .Cells(2, 5 + x) = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "���C���i��"), x)
            .Cells(1, 5 + x) = Mid(.Cells(2, 5 + x), 8, 3)
        Next x
        For i = LBound(�ގ��ꗗ, 2) + 1 To UBound(�ގ��ꗗ, 2)
            .Cells(i + 2, 1) = �ގ��ꗗ(0, i)
            .Cells(i + 2, 4) = �ގ��ꗗ(1, i)
            .Cells(i + 2, 5) = �ގ��ꗗ(2, i)
            For x = 1 To Len(�ގ��ꗗ(4, i))
                If Mid(�ގ��ꗗ(4, i), x, 1) <> "1" Then
                    .Cells(i + 2, 6 + x - 1) = Mid(�ގ��ꗗ(4, i), x, 1)
                End If
            Next x
        Next i
        Stop
        '���ёւ�
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 1).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 4).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 5).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(3), Rows(i + 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
End Function

Public Function ��������H�}�g���N�X��������擾()

'���i�i�Ԃ��Q�Ƃ��Ă��Ȃ�
'�ŏ��̏����Ŗ�蕔�i���Q�Ƃ��Ă��Ȃ�

Dim ����book As String: ����book = ActiveWorkbook.Name
Dim ����sheet As String: ����sheet = "PVSW_RLTF"
Dim ������c(5) As Long, ������(1) As Variant

Dim ����book As Workbook, C As Long, x As Long

'����book = "�H��H��د��+82162-6AT80,B40-000 5��.xlsm"

Dim wb As Workbook
For Each wb In Workbooks
    If wb.Name <> ����book Then
        C = C + 1
    End If
Next

If C <> 0 Then MsgBox "��������s���鎞�͑��̃u�b�N����Ă��������B": End

Dim OpenFileName As String
OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")
Workbooks.Open OpenFileName, ReadOnly:=True
Set ����book = ActiveWorkbook

Dim ����sheet As String: ����sheet = "PVSW"
Dim ������c(1) As Long, ������(1) As Variant

With Workbooks(����book).Sheets(����sheet) '2
    Dim myKey As Variant: Set myKey = .Cells.Find("�d�����ʖ�", , , 1)
    ������c(0) = .Rows(myKey.Row).Find("�n�_����H����", , , 1).Column
    ������c(1) = .Rows(myKey.Row).Find("�I�_����H����", , , 1).Column
    ������c(2) = .Rows(myKey.Row).Find("�n�_���[�����ʎq", , , 1).Column
    ������c(3) = .Rows(myKey.Row).Find("�I�_���[�����ʎq", , , 1).Column
    ������c(4) = .Rows(myKey.Row).Find("�n�_���L���r�e�B", , , 1).Column
    ������c(5) = .Rows(myKey.Row).Find("�I�_���L���r�e�B", , , 1).Column
    ���i�i��s = .Rows(myKey.Row - 3).Find("���i�i��s", , , 1).Column
    Set ���i�i��e = .Rows(myKey.Row - 3).Find("���i�i��e", , , 1)
    If ���i�i��e Is Nothing Then
        ���i�i��e = ���i�i��s
    Else
        ���i�i��e = ���i�i��e.Column
    End If
    ����lastrow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
    For x = 0 To 5
        .Range(.Cells(myKey.Row + 1, ������c(x)), .Cells(����lastrow, ������c(x))).Interior.Pattern = xlNone
        .Range(.Cells(myKey.Row + 1, ������c(x)), .Cells(����lastrow, ������c(x))).Font.color = 0
        .Range(.Cells(myKey.Row + 1, ������c(x)), .Cells(����lastrow, ������c(x))).Font.Bold = falase
    Next x
End With

With ����book.Sheets(����sheet) '1
    Set key = .Cells.Find("�\��No.", , , 1)
    ������c(0) = .Rows(key.Row).Find("��A", , , 1).Column
    ������c(1) = .Rows(key.Row).Find("��B", , , 1).Column
    ����lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
    Set key2 = .Cells.Find("key1_", , , 1)
    ����lastcol = .Cells(key2.Row, .Columns.count).End(xlToLeft).Column
End With

Dim �������i As String * 15
With ����book.Sheets(����sheet)
    For x = key2.Column + 1 To ����lastcol
        �������i = .Cells(key2.Row, x)
        For y = key.Row + 2 To ����lastRow
            myCount = 0
            �\�� = .Cells(y, key.Column)
            If �\�� = "" Then GoTo line10
            
            ���i�g���� = .Cells(y, x)
            If ���i�g���� <> "" Then
                For i2 = 0 To 1
                    Set ������(i2) = .Cells(y, ������c(i2))
                Next i2
                
                With Workbooks(����book).Sheets(����sheet)
                    Set ����xx = .Rows(myKey.Row).Find(�������i, , , 1)
                    If ����xx Is Nothing Then GoTo line20
                    For y2 = myKey.Row + 1 To ����lastrow
                        �����g���� = .Cells(y2, ����xx.Column)
                        If �����g���� = "" Then GoTo line05
                        �\��2 = Left(.Cells(y2, myKey.Column), 4)
                        If �\�� = �\��2 Then
                            For i = 0 To 1
                                Set ������(i) = .Cells(y2, ������c(i))
                                For i2 = 0 To 1
                                    'Debug.Print ������(i), ������(i2)
                                    If ������(i).Value = ������(i2).Value Then
                                        
                                        .Cells(y2, ������c(i + 0)).Font.color = ������(i2).Font.color
                                        .Cells(y2, ������c(i + 2)).Font.color = ������(i2).Font.color
                                        .Cells(y2, ������c(i + 4)).Font.color = ������(i2).Font.color
                                        .Cells(y2, ������c(i + 0)).Font.Bold = True
                                        .Cells(y2, ������c(i + 2)).Font.Bold = True
                                        .Cells(y2, ������c(i + 4)).Font.Bold = True
                                        
                                        '�w�i�F
'                                        If ������(i2).Interior.color <> 16777215 Then
'                                            .Cells(y2, ������c(i + 0)).Interior.color = ������(i2).Interior.color
'                                            .Cells(y2, ������c(i + 2)).Interior.color = ������(i2).Interior.color
'                                            .Cells(y2, ������c(i + 4)).Interior.color = ������(i2).Interior.color
'                                        End If
                                        
                                        myCount = myCount + 1
                                    End If
                                Next i2
                            Next i
                            If myCount >= 2 Then
                                 ����book.Sheets(����sheet).Cells(y, x).Interior.color = 16764159
                            End If
                            GoTo line10
                        End If
line05:
                    Next y2
                End With
            End If
line10:
        Next y
line20:
    Next x
End With

'���[���ꗗ�ɐF��t����(��菬�����n���F�ԍ���I��)
'Stop
'PVSW_RLTF����[�������擾
With Workbooks(����book).Sheets("�ݒ�")
    Dim �n���F�ݒ�() As String
    ReDim �n���F�ݒ�(3, 0)
    Set �ݒ�key = .Cells.Find("�n���F_", , , 1)
    i = 0
    Do
        If �ݒ�key.Offset(i, 1) <> "" Then
            add = add + 1
            ReDim Preserve �n���F�ݒ�(3, add)
            �n���F�ݒ�(0, add) = �ݒ�key.Offset(i, 1).Value
            �n���F�ݒ�(1, add) = �ݒ�key.Offset(i, 1).Font.color
            �n���F�ݒ�(2, add) = �ݒ�key.Offset(i, 2).Value
            '�n���F�ݒ�(3, add) = �ݒ�key.Offset(i, 1).Interior.color
        Else
            Exit Do
        End If
        i = i + 1
    Loop
End With

With Workbooks(����book).Sheets(����sheet)
    Dim �[��() As String
    ReDim �[��(4, 0)
    Dim �����[��c(1) As Long
    Dim �������c(1) As Long
    �����[��c(0) = .Rows(myKey.Row).Find("�n�_���[�����ʎq", , , 1).Column
    �����[��c(1) = .Rows(myKey.Row).Find("�I�_���[�����ʎq", , , 1).Column
    �������c(0) = .Rows(myKey.Row).Find("�n�_���[�����i��", , , 1).Column
    �������c(1) = .Rows(myKey.Row).Find("�I�_���[�����i��", , , 1).Column
    add = 0
    For y2 = myKey.Row + 1 To ����lastrow
        For i = 0 To 1
            Set �[��v = .Cells(y2, �����[��c(i))
            Set ���v = .Cells(y2, �������c(i))
            '�n���F�ݒ���Q��
            For i2 = 1 To UBound(�n���F�ݒ�, 2)
                If �[��v.Font.color = �n���F�ݒ�(1, i2) Then 'And �[��v.Interior.color = �n���F�ݒ�(3, i2) Then
                    '�[���ւ̓o�^�L���m�F
                    For i3 = LBound(�[��, 2) To UBound(�[��, 2)
                        If �[��(0, i3) = �[��v.Value And �[��(3, i3) = ���v.Value Then
                            'Stop
                            '�[���ւ̓o�^�ύX
                            If �[��(2, i3) > �n���F�ݒ�(0, i2) Then
                                'Stop
                                �[��(1, i3) = �n���F�ݒ�(1, i2) 'Font.color
                                �[��(2, i3) = �n���F�ݒ�(0, i2) '��Ə��ԍ�
                                '�[��(4, i3) = �n���F�ݒ�(3, i2) 'Interior.color
                            End If
                            GoTo line30
                        End If
                    Next i3
                    'Stop
                    '�[���ւ̐V�K�ǉ�
                    add = add + 1
                    ReDim Preserve �[��(4, add)
                    �[��(0, add) = �[��v.Value
                    �[��(3, add) = ���v.Value
                    �[��(1, add) = �n���F�ݒ�(1, i2)
                    �[��(2, add) = �n���F�ݒ�(0, i2)
                    �[��(4, add) = �n���F�ݒ�(3, i2)
                    GoTo line30
                End If
            Next i2
            Debug.Print �[��v.Font.color
            Stop 'font�F��������Ȃ�����
line30:
        Next i
    Next y2
End With

For i = 1 To UBound(�[��, 2)
    Debug.Print �[��(0, i), �[��(1, i), �[��(2, i), �[��(3, i), �[��(4, i)
Next i

With Workbooks(����book).Sheets("�[���ꗗ")
    Set �[���ꗗkey = .Cells.Find("�[�����i��", , , 1)
    Dim �[���ꗗcol(1) As Long
    �[���ꗗcol(0) = �[���ꗗkey.Column
    �[���ꗗcol(1) = .Cells.Find("�[����", , , 1).Column
    �[���ꗗmaxcol = .Cells(�[���ꗗkey.Row, .Columns.count).End(xlToLeft).Column
    �[���ꗗlastrow = .Cells(.Rows.count, �[���ꗗkey.Column).End(xlUp).Row
    '�z������ɎQ��
    For i = 1 To UBound(�[��, 2)
        '�[���ꗗ���Q��
        For i2 = �[���ꗗkey.Row + 1 To �[���ꗗlastrow
            If �[��(0, i) = .Cells(i2, �[���ꗗcol(1)) Then
                If �[��(3, i) = .Cells(i2, �[���ꗗcol(0)) Then
                    'Stop
                    .Range(.Cells(i2, �[���ꗗcol(1) + 1), .Cells(i2, �[���ꗗmaxcol)).Font.color = �[��(1, i)
                    .Range(.Cells(i2, �[���ꗗcol(1) + 1), .Cells(i2, �[���ꗗmaxcol)).Font.Bold = True
'                    If �[��(4, i) <> 16777215 Then
'                        .Range(.Cells(i2, �[���ꗗcol(1) + 1), .Cells(i2, �[���ꗗmaxcol)).Interior.color = �[��(4, i)
'                    End If
                    GoTo line40
                End If
            End If
        Next i2
        Stop 'PVSW_RLTF�ɂ��邯�ǁA�[���ꗗ�ɂȂ�����
line40:
    Next i
End With
    
    MsgBox "�������������܂����B���̏����͊m���ł͂���܂���B���e���m�F���Ă��������B"
    
End Function

Public Function �{�莮��H���X�g��������擾()

    '���i�i�Ԃ��Q�Ƃ��Ă��Ȃ�
    '�ŏ��̏����Ŗ�蕔�i���Q�Ƃ��Ă��Ȃ�
    
    Dim ����book As String: ����book = ActiveWorkbook.Name
    Dim ����sheet As String: ����sheet = "PVSW_RLTF"
    Dim ������c(5) As Long, ������(1) As Variant
    
    Dim �{��book As String
    
    '�{��book = "�H��H��د��+82162-6AT80,B40-000 5��.xlsm"
    
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name <> ����book Then
            �{��book = wb.Name
            C = C + 1
        End If
    Next
    
    If C > 1 Then MsgBox "�Ώۂ̃u�b�N��1�ȏ゠��܂��B": End
    If C = 0 Then MsgBox "�Ώۂ̃u�b�N���J������ԂŎ��s���Ă��������B": End
    
    Dim �{��sheet As String: �{��sheet = "�d��B (2)"
    
    With Workbooks(����book).Sheets(����sheet) '2
        Dim myKey As Variant: Set myKey = .Cells.Find("�d�����ʖ�", , , 1)
        ������c(0) = .Rows(myKey.Row).Find("�n�_����H����", , , 1).Column
        ������c(1) = .Rows(myKey.Row).Find("�I�_����H����", , , 1).Column
        ������c(2) = .Rows(myKey.Row).Find("�n�_���[�����ʎq", , , 1).Column
        ������c(3) = .Rows(myKey.Row).Find("�I�_���[�����ʎq", , , 1).Column
        ������c(4) = .Rows(myKey.Row).Find("�n�_���L���r�e�B", , , 1).Column
        ������c(5) = .Rows(myKey.Row).Find("�I�_���L���r�e�B", , , 1).Column
        ���i�i��s = .Rows(myKey.Row - 3).Find("���i�i��s", , , 1).Column
        ���i�i��e = .Rows(myKey.Row - 3).Find("���i�i��e", , , 1).Column
        ����lastrow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Dim �����[��(1) As Object
        Dim ����CAV(1) As Object
    End With
    
    Dim �{��[��(1) As Object, �{��CAV(1) As Object, �{��hame(1) As Object
    With Workbooks(�{��book).Sheets(�{��sheet) '1
        Set key = .Cells.Find("�\��", , , 1)
        Set �{��[��(0) = .Rows(key.Row).Find("�[��1", , , 1)
        Set �{��[��(1) = .Rows(key.Row).Find("�[��2", , , 1)
        Set �{��CAV(0) = .Rows(key.Row).Find("����è1", , , 1)
        Set �{��CAV(1) = .Rows(key.Row).Find("����è2", , , 1)
        Set �{��hame(0) = .Rows(key.Row).Find("1�Ƃ�", , , 1)
        Set �{��hame(1) = .Rows(key.Row).Find("2�Ƃ�", , , 1)
        
        �{��lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Set key2 = .Cells.Find("key1_", , , 1)
        �{��lastcol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
    End With
    
    Dim �{�萻�i As String * 15
    With Workbooks(�{��book).Sheets(�{��sheet)
        For x = key2.Column To �{��lastcol
            �{�萻�i = .Cells(key.Row, x)
            For y = key.Row + 1 To �{��lastRow
                myCount = 0
                �\�� = Format(.Cells(y, key.Column), "0000")
                If �\�� = "" Then GoTo line10
                ���i�g���� = .Cells(y, x)
                If ���i�g���� <> "" Then
                
                    With Workbooks(����book).Sheets(����sheet)
                        Set ����xx = .Rows(myKey.Row).Find(�{�萻�i, , , 1)
                        If ����xx Is Nothing Then GoTo line20
                        For y2 = myKey.Row + 1 To ����lastrow
                            �����g���� = .Cells(y2, ����xx.Column)
                            If �����g���� = "" Then GoTo line05
                            �\��2 = Left(.Cells(y2, myKey.Column), 4)
                            If �\�� = �\��2 Then
                                For i = 0 To 1
                                    Set �����[��(i) = .Cells(y2, ������c(i + 2))
                                    Set ����CAV(i) = .Cells(y2, ������c(i + 4))
                                    
                                    For i2 = 0 To 1
                                        'Debug.Print ������(i), �{���(i2)
                                        If �����[��(i) & "_" & ����CAV(i) = �{��[��(i2).Offset(y - key.Row, 0) & "_" & �{��CAV(i2).Offset(y - key.Row, 0) Then
                                            Dim myrgb As Long
                                            
                                            myrgb = �{��hame(i).Offset(y - key.Row, 0).Font.color
                                            .Cells(y2, ������c(i + 0)).Font.color = myrgb
                                            .Cells(y2, ������c(i + 2)).Font.color = myrgb
                                            .Cells(y2, ������c(i + 4)).Font.color = myrgb
                                            .Cells(y2, ������c(i + 0)).Font.Bold = True
                                            .Cells(y2, ������c(i + 2)).Font.Bold = True
                                            .Cells(y2, ������c(i + 4)).Font.Bold = True
                                            
                                            '�w�i�F
    '                                        If �{���(i2).Interior.color <> 16777215 Then
    '                                            .Cells(Y2, ������c(i + 0)).Interior.color = �{���(i2).Interior.color
    '                                            .Cells(Y2, ������c(i + 2)).Interior.color = �{���(i2).Interior.color
    '                                            .Cells(Y2, ������c(i + 4)).Interior.color = �{���(i2).Interior.color
    '                                        End If
                                            
                                            myCount = myCount + 1
                                        End If
                                    Next i2
                                    
                                Next i
                                If myCount >= 2 Then
                                     Workbooks(�{��book).Sheets(�{��sheet).Cells(y, x).Interior.color = 16764159
                                End If
                                GoTo line10
                            End If
line05:
                        Next y2
                    End With
                End If
line10:
            Next y
line20:
        Next x
    End With
    
    '���[���ꗗ�ɐF��t����(��菬�����n���F�ԍ���I��)
    
    'PVSW_RLTF����[�������擾
    With Workbooks(����book).Sheets("�ݒ�")
        Dim �n���F�ݒ�() As String
        ReDim �n���F�ݒ�(3, 0)
        Set �ݒ�key = .Cells.Find("�n���F_", , , 1)
        i = 0
        Do
            If �ݒ�key.Offset(i, 1) <> "" Then
                add = add + 1
                ReDim Preserve �n���F�ݒ�(3, add)
                �n���F�ݒ�(0, add) = �ݒ�key.Offset(i, 1).Value
                �n���F�ݒ�(1, add) = �ݒ�key.Offset(i, 1).Font.color
                �n���F�ݒ�(2, add) = �ݒ�key.Offset(i, 2).Value
                �n���F�ݒ�(3, add) = �ݒ�key.Offset(i, 1).Interior.color
            Else
                Exit Do
            End If
            i = i + 1
        Loop
    End With
    
    With Workbooks(����book).Sheets(����sheet)
        Dim �[��() As String
        ReDim �[��(4, 0)
        Dim �����[��c(1) As Long
        Dim �������c(1) As Long
        �����[��c(0) = .Rows(myKey.Row).Find("�n�_���[�����ʎq", , , 1).Column
        �����[��c(1) = .Rows(myKey.Row).Find("�I�_���[�����ʎq", , , 1).Column
        �������c(0) = .Rows(myKey.Row).Find("�n�_���[�����i��", , , 1).Column
        �������c(1) = .Rows(myKey.Row).Find("�I�_���[�����i��", , , 1).Column
        add = 0
        For y2 = myKey.Row + 1 To ����lastrow
            For i = 0 To 1
                Set �[��v = .Cells(y2, �����[��c(i))
                Set ���v = .Cells(y2, �������c(i))
                '�n���F�ݒ���Q��
                For i2 = 1 To UBound(�n���F�ݒ�, 2)
                    If �[��v.Font.color = �n���F�ݒ�(1, i2) And �[��v.Interior.color = �n���F�ݒ�(3, i2) Then
                        '�[���ւ̓o�^�L���m�F
                        For i3 = LBound(�[��, 2) To UBound(�[��, 2)
                            If �[��(0, i3) = �[��v.Value And �[��(3, i3) = ���v.Value Then
                                'Stop
                                '�[���ւ̓o�^�ύX
                                If �[��(2, i3) > �n���F�ݒ�(0, i2) Then
                                    'Stop
                                    �[��(1, i3) = �n���F�ݒ�(1, i2) 'Font.color
                                    �[��(2, i3) = �n���F�ݒ�(0, i2) '��Ə��ԍ�
                                    �[��(4, i3) = �n���F�ݒ�(3, i2) 'Interior.color
                                End If
                                GoTo line30
                            End If
                        Next i3
                        'Stop
                        '�[���ւ̐V�K�ǉ�
                        add = add + 1
                        ReDim Preserve �[��(4, add)
                        �[��(0, add) = �[��v.Value
                        �[��(3, add) = ���v.Value
                        �[��(1, add) = �n���F�ݒ�(1, i2)
                        �[��(2, add) = �n���F�ݒ�(0, i2)
                        �[��(4, add) = �n���F�ݒ�(3, i2)
                        GoTo line30
                    End If
                Next i2
                Stop 'font�F��������Ȃ�����
line30:
            Next i
        Next y2
    End With
    
    For i = 1 To UBound(�[��, 2)
        Debug.Print �[��(0, i), �[��(1, i), �[��(2, i), �[��(3, i), �[��(4, i)
    Next i
    
    With Workbooks(����book).Sheets("�[���ꗗ")
        Set �[���ꗗkey = .Cells.Find("�[�����i��", , , 1)
        Dim �[���ꗗcol(1) As Long
        �[���ꗗcol(0) = �[���ꗗkey.Column
        �[���ꗗcol(1) = .Cells.Find("�[����", , , 1).Column
        �[���ꗗmaxcol = .Cells(�[���ꗗkey.Row, .Columns.count).End(xlToLeft).Column
        �[���ꗗlastrow = .Cells(.Rows.count, �[���ꗗkey.Column).End(xlUp).Row
        '�z������ɎQ��
        For i = 1 To UBound(�[��, 2)
            '���ވꗗ���Q��
            For i2 = �[���ꗗkey.Row + 1 To �[���ꗗlastrow
                If �[��(0, i) = .Cells(i2, �[���ꗗcol(1)) Then
                    If �[��(3, i) = .Cells(i2, �[���ꗗcol(0)) Then
                        'Stop
                        .Range(.Cells(i2, �[���ꗗcol(1) + 1), .Cells(i2, �[���ꗗmaxcol)).Font.color = �[��(1, i)
                        .Range(.Cells(i2, �[���ꗗcol(1) + 1), .Cells(i2, �[���ꗗmaxcol)).Font.Bold = True
                        If �[��(4, i) <> 16777215 Then
                            .Range(.Cells(i2, �[���ꗗcol(1) + 1), .Cells(i2, �[���ꗗmaxcol)).Interior.color = �[��(4, i)
                        End If
                        GoTo line40
                    End If
                End If
            Next i2
            Stop 'PVSW_RLTF�ɂ��邯�ǁA�[���ꗗ�ɂȂ�����
line40:
        Next i
    End With
    
End Function


Public Function �ƃ��C�A�E�g�}�̍쐬ver2179(CB0, �^��str, ���@str)
    �^��str = Replace(�^��str, " ", "")
    Set wb(0) = ActiveWorkbook
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    
    Call �œK��
    Call addressSet(wb(0))
    
    'Call ���i�i��RAN_set2(���i�i��RAN, CB0.Value, CB1.Value, "")
    Call ���i�i��RAN_set2(���i�i��Ran, CB0, �^��str, "")
    
    myDir = "\10_�ƃ��C�A�E�g\"
    
    '�f�B���N�g���쐬
    If Dir(ActiveWorkbook.path & myDir, vbDirectory) = "" Then
        MkDir ActiveWorkbook.path & myDir
    End If
    If Dir(ActiveWorkbook.path & myDir & �^��str, vbDirectory) = "" Then
        MkDir ActiveWorkbook.path & myDir & �^��str
    End If
    myPath = ActiveWorkbook.path & myDir & �^��str & "\" & ���@str
    If Dir(myPath, vbDirectory) = "" Then
        MkDir myPath
    End If
    Dim myFileStr As String, myNumber As String: myNumber = "000"
    myFileStr = Left(wb(0).Name, InStrRev(wb(0).Name, ".") - 1)
    Do
        If Dir(myPath & "\" & myFileStr & "_" & ���@str & "_" & myNumber & ".xlsm") = "" Then Exit Do
        myNumber = Format(CLng(myNumber) + 1, "000")
    Loop
    Dim myFileName As String
    myFileName = myPath & "\" & myFileStr & "_" & ���@str & "_" & myNumber & ".xlsm"
    
    '�o�͐�book���쐬
    Workbooks.Open myAddress(0, 1) & "\genshi\����_�ƃ��C�A�E�g.xlsm"
    Set wb(1) = ActiveWorkbook
    Application.DisplayAlerts = False
    wb(1).SaveAs fileName:=myFileName, FileFormat:=52
    Application.DisplayAlerts = True
    
    Set ws(1) = wb(1).Sheets("Sheet1")
    
    'PVSW_RLTF�̃f�[�^�擾
    With ws(0)
        '�d����
        Dim myWire As String, myTerm As String, myWireSP, myTermSP
        myWire = "���i�i��s,���i�i��e,RLTFtoPVSW_,�\��_,�i��_,�T�C�Y_,�F_,�F��_,�ؒf��_,����_,RLTFtoPVSW_"
        myWireSP = Split(myWire, ",")
        Dim myWireC(): ReDim myWireC(UBound(myWireSP))
        For x = LBound(myWireSP) To UBound(myWireSP)
            myWireC(x) = .Cells.Find(myWireSP(x), , , 1).Column
        Next x
        '�d���[����
        myTerm = "�n�_����H����,�I�_����H����,�n�_���[�����ʎq,�I�_���[�����ʎq,�n�_���[�����i��,�I�_���[�����i��,�n�_���[�q_,�I�_���[�q_,�n�_���}_,�I�_���}_,�n�_�����i_,�I�_�����i_,�n�_���L���r�e�B,�I�_���L���r�e�B"
        myTermSP = Split(myTerm, ",")
        Dim myTermC(): ReDim myTermC(UBound(myTermSP))
        For x = LBound(myTermSP) To UBound(myTermSP)
            myTermC(x) = .Cells.Find(myTermSP(x), , , 1).Column
        Next x
        '���i�i�Ԗ�
        Dim myProdC: ReDim myProdC(���i�i��RANc)
        For x = LBound(���i�i��Ran, 2) + 1 To UBound(���i�i��Ran, 2)
            myProdC(x) = .Cells.Find(���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "���C���i��"), x), , , 1).Column
        Next x
        '���̑�
        Dim AutoC As Long, SubC As Long
        AutoC = .Cells.Find("�����@", , , 1).Column
        SubC = .Cells.Find("SubNo", , , 1).Column
        Set mykey0 = .Cells.Find("�d�����ʖ�", , , 1)
    End With
    
    With ws(1)
        Set mykey1 = .Cells.Find("�\��", , , 1)
        Dim addCol As Long: addCol = mykey1.Column + 1
        .Cells(2, mykey1.Column) = CB0 & "=" & �^��str
        .Cells(3, mykey1.Column) = ���@str
        For x = LBound(myProdC) + 1 To UBound(myProdC)
            ���i�i��str = ws(0).Cells(mykey0.Row, myProdC(x)).Value
            ���i�i��short = ws(0).Cells(mykey0.Row - 1, myProdC(x)).Value
            �N���� = ws(0).Cells(mykey0.Row - 2, myProdC(x)).Value
            .Cells(24 + x, 1) = ���i�i��str
            .Cells(24 + x, 2) = �N����
            .Cells(24 + x, mykey1.Column) = ���i�i��short
        Next x
        addRow = .Cells(.Rows.count, mykey1.Column).End(xlUp).Row + 1
    End With
    
    With ws(0)
        Dim �����@ As String, SubNo As String, RLTFtoPVSW As String
        lastRow = .Cells(.UsedRange.Rows.count + 1, mykey0.Column).End(xlUp).Row
        sCol = myWireC(0): eCol = myWireC(1)
        For i = mykey0.Row + 1 To lastRow
            �����@ = .Cells(i, AutoC)
            SubNo = .Cells(i, SubC)
            RLTFtoPVSW = .Cells(i, myWireC(10))
            If RLTFtoPVSW <> "Found" Then GoTo nextI
            If �����@ <> ���@str Then GoTo nextI
            If .Cells(i, myWireC(2)) <> "Found" Then GoTo nextI
            '���i�i��RAN�ɂ��邩�m�F
            For x = LBound(myProdC) + 1 To UBound(myProdC)
                If .Cells(i, myProdC(x)) <> "" Then
                    GoTo �o�^
                End If
            Next x
            GoTo nextI '�����̂Ŏ��̍s
�o�^:
            �\�� = .Cells(i, myWireC(3))
            Set �i�� = .Cells(i, myWireC(4))
            �T�C�Y = .Cells(i, myWireC(5))
            �F = .Cells(i, myWireC(6))
            �F�� = .Cells(i, myWireC(7))
            �ؒf�� = .Cells(i, myWireC(8))
            ���� = .Cells(i, myWireC(9))
            Set ��H����0 = .Cells(i, myTermC(0))
            Set ��H����1 = .Cells(i, myTermC(1))
            Set �[��0 = .Cells(i, myTermC(2))
            Set �[��1 = .Cells(i, myTermC(3))
            ���0 = .Cells(i, myTermC(4))
            ���1 = .Cells(i, myTermC(5))
            Set �[�q0 = .Cells(i, myTermC(6))
            Set �[�q1 = .Cells(i, myTermC(7))
            �}���}0 = .Cells(i, myTermC(8))
            �}���}1 = .Cells(i, myTermC(9))
            ���i0 = .Cells(i, myTermC(10))
            ���i1 = .Cells(i, myTermC(11))
            CAV0 = .Cells(i, myTermC(12))
            Cav1 = .Cells(i, myTermC(13))
            ���i�i��str = ""
            For x = LBound(myProdC) + 1 To UBound(myProdC)
                ���i�i��str = ���i�i��str & "," & .Cells(i, myProdC(x))
            Next x
            With ws(1)
                .Cells(7, addCol) = �\��
                .Cells(8, addCol) = �i��
                .Cells(8, addCol).Interior.color = �i��.Interior.color
                .Cells(9, addCol) = �T�C�Y
                .Cells(10, addCol) = �F
                .Activate
                Call �d���F�ŃZ����h��(11, addCol, CStr(�F��))
                .Cells(12, addCol) = �F��
                .Cells(15, addCol) = Left(�[�q0, 4) & vbCrLf & Mid(�[�q0, 5, 4) & vbCrLf & Mid(�[�q0, 9, 2)
                .Cells(15, addCol).Interior.color = �[�q0.Interior.color
                .Cells(16, addCol) = �}���}0
                .Cells(17, addCol) = ���i0
                .Cells(18, addCol) = Left(���0, 4) & vbCrLf & Mid(���0, 5, 4) & vbCrLf & Mid(���0, 9, 2)
                .Cells(19, addCol) = �[��0
                .Cells(20, addCol) = CAV0
                .Cells(22, addCol) = ��H����0
                .Cells(22, addCol).Font.color = ��H����0.Font.color
                .Cells(22, addCol).Font.Bold = True
                .Cells(23, addCol) = ����
                .Cells(24, addCol) = �ؒf��
'                .Cells(addRow, addCol) = �ؒf��
                .Cells(7, addCol + 1) = �\��
                .Cells(8, addCol + 1) = �i��
                .Cells(9, addCol + 1) = �T�C�Y
                .Cells(10, addCol + 1) = �F
                .Activate
                Call �d���F�ŃZ����h��(11, addCol + 1, CStr(�F��))
                .Cells(12, addCol + 1) = �F��
                .Cells(15, addCol + 1) = Left(�[�q1, 4) & vbCrLf & Mid(�[�q1, 5, 4) & vbCrLf & Mid(�[�q1, 9, 2)
                .Cells(15, addCol + 1).Interior.color = �[�q1.Interior.color
                .Cells(16, addCol + 1) = �}���}1
                .Cells(17, addCol + 1) = ���i1
                .Cells(18, addCol + 1) = Left(���1, 4) & vbCrLf & Mid(���1, 5, 4) & vbCrLf & Mid(���1, 9, 2)
                .Cells(19, addCol + 1) = �[��1
                .Cells(20, addCol + 1) = Cav1
                .Cells(22, addCol + 1) = ��H����1
                .Cells(22, addCol + 1).Font.color = ��H����1.Font.color
                .Cells(22, addCol + 1).Font.Bold = True
                .Cells(23, addCol + 1) = ����
                .Cells(24, addCol + 1) = �ؒf��
                '.Cells(addRow, addCol + 1) = �ؒf��
                ���i�i��strSP = Split(���i�i��str, ",")
                For x = LBound(���i�i��strSP) + 1 To UBound(���i�i��strSP)
                    If ���i�i��strSP(x) <> "" Then
                        .Cells(24 + x, addCol) = "1"
                        .Cells(24 + x, addCol + 1) = "1"
                    End If
                Next x
                addCol = addCol + 2
                addRow = addRow + 1
            End With
nextI:
        Next i
        ws(1).PageSetup.PrintArea = .Range(.Cells(1, 3), .Cells(63, addCol - 1)).addRess
'        WS(1).PageSetup.RightHeader = "&L" & "&13 " & Left(WB(0).Name, InStr(WB(0).Name, "_") - 1)
'        Set WS(2) = WB(1).Sheets("Ver")
'        Set verkey = WB(1).Sheets("Ver").Cells.Find("Ver", , , 1)
'        myver = WS(2).Cells(WS(2).Cells(Rows.Count, verkey.Column).End(xlUp).Row, verkey.Column)
'        WS(1).PageSetup.RightHeader = "&L" & "&13 " & Left(WB(0).Name, InStr(WB(0).Name, "_") - 1) & "�ƃ��C�A�E�g+_" & myver
    End With
    
    '���ёւ�
    With ws(1)
        .Sort.SortFields.Clear
        Set myKey = .Cells.Find("�[��" & vbLf & "���" & vbLf & "�i��", , , 1)
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        Set �[��a = .Cells.Find("�[��", , , 1)
        Set cava = .Cells.Find("Cav", , , 1)
        .Sort.SortFields.add key:=Cells(�[��a.Row, myKey.Column), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Sort.SortFields.add key:=Cells(cava.Row, myKey.Column), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        With ws(1).Sort
            .SetRange Range(Columns(myKey.Column + 1), Columns(lastCol))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlLeftToRight
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    '���}�̔z�u
    With ws(1)
        ��H����row = .Cells.Find("��", , , 1).Row
        �[��row = .Cells.Find("�[��", , , 1).Row
        ���row = myKey.Row
        �[��bak = ""
        addpoint = .Rows(lastRow + 1).Top
        For x = myKey.Column + 1 To lastCol
            Set ��H���� = .Cells(��H����row, x)
            �[��str = .Cells(�[��row, x)
            ��� = Replace(.Cells(���row, x), vbCrLf, "")
            If InStr(�z�u�[��, "_" & �[��str & "_") = 0 Then
                If ��H����.Font.color = 5287936 Then
                    Select Case Len(Replace(���, " ", ""))
                        Case 8
                        ���str = Left(���, 4) & "-" & Mid(���, 5, 4)
                        Case 10
                        ���str = Left(���, 4) & "-" & Mid(���, 5, 4) & "-" & Mid(���, 9, 2)
                    End Select
                    �摜URL = myAddress(1, 1) & "\202_���}\" & ���str & "_1_001.emf"
                    On Error Resume Next
                    Set ob = ActiveSheet.Shapes.AddPicture(�摜URL, False, True, .Columns(x).Left, addpoint, 50, 50)
                    ob.LockAspectRatio = msoTrue
                    ob.ScaleHeight 1, msoTrue
                    ob.ScaleWidth 1, msoTrue
                    ob.Name = �[��str
'                    .Pictures.Insert(�摜URL).Name = �[��str
'                    .Shapes.Range(�[��str).Top = addpoint
'                    .Shapes.Range(�[��str).Left = .Columns(x).Left
                    .Shapes.Range(�[��str).Width = .Rows(�[��row).Find(�[��str, , , , , 2, 1).Offset(0, 1).Left - .Columns(x).Left
                    'addpoint = addpoint + .Shapes.Range(�[��str).Height
                    On Error GoTo 0
                    �z�u�[�� = �z�u�[�� & "_" & �[��str & "_"
                End If
            End If
            �[��bak = �[��str
        Next x
        '�F�̈Ӗ����o��
        Set �n���Fkey = wb(0).Sheets("�ݒ�").Cells.Find("�n���F_", , , 1)
        For x = 0 To 14
            If �n���Fkey.Offset(x, 1).Value > 0 Then
                Set �n���Fe = �n���Fkey.Offset(x, 1).End(xlDown)
                Set �n���Fran = wb(0).Sheets("�ݒ�").Range(�n���Fkey.Offset(x, 1).addRess, �n���Fe.Offset(0, 1).addRess)
                Exit For
            End If
        Next x
        Dim �F�̐��� As Shape
        Set �F�̐��� = .Shapes.AddShape(1, 100, 50, 70, 100)
        �F�̐���.Fill.Transparency = 1
        �F�̐���.Line.Visible = msoFalse
        �F�̐���.TextFrame2.TextRange.Font.size = 10
        For p = 1 To �n���Fran.count / �n���Fran.Column
            �F�̐���.TextFrame.Characters.Text = �F�̐���.TextFrame.Characters.Text & "��" & �n���Fran(p, 2) & vbCrLf
            �F�̐���.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        Next p
        �F�̐���.TextFrame.Characters.Text = Mid(�F�̐���.TextFrame.Characters.Text, 1, Len(�F�̐���.TextFrame.Characters.Text) - 1)
        �F�̐���.TextFrame2.WordWrap = msoFalse
        ������ = 1
        For p = 1 To �n���Fran.count / �n���Fran.Column
            �F�̐���.TextFrame2.TextRange.Characters(������, 1).Font.Fill.ForeColor.RGB = �n���Fran(p, 1).Font.color
            ������ = ������ + Len(�n���Fran(p, 2)) + 2
        Next p
        �F�̐���.Name = "�F�̐���"
        �F�̐���.TextFrame2.MarginLeft = 0
        �F�̐���.TextFrame2.MarginRight = 0
        �F�̐���.TextFrame2.MarginTop = 0
        �F�̐���.TextFrame2.MarginBottom = 0
        �F�̐���.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        �F�̐���.Top = 0
        �F�̐���.Left = ws(1).Columns(5).Left
        
        '�[�q�t�@�~���[
        Set �n���Fkey = wb(0).Sheets("�ݒ�").Cells.Find("�[�q�t�@�~���[_", , , 1)
        Dim �n���Frange As Range
        For x = 0 To 9
            If �n���Fkey.Offset(x, 1).Value > 0 Then
'                Set �n���Fe = �n���Fkey.Offset(x, 1).End(xlDown)
'                Set �n���Fran = WB(0).Sheets("�ݒ�").Range(�n���Fkey.Offset(x, 1).Address, �n���Fe.Offset(0, 4).Address)
'                Exit For
                '�g�p�����邩�m�F
                Dim �n���Fcolor As Long: myFlg = False
                �n���Fcolor = �n���Fkey.Offset(x, 1).Interior.color
                For C = myKey.Column + 1 To lastCol
                    If �n���Fcolor = .Cells(15, C).Interior.color Then
                        myFlg = True
                        Exit For
                    End If
                Next C
                If myFlg = True Then
                    If �n���Frange Is Nothing Then
                        Set �n���Frange = wb(0).Sheets("�ݒ�").Range(�n���Fkey.Offset(x, 1), �n���Fkey.Offset(x, 5))
                    Else
                        Set �n���Frange = Union(�n���Frange, wb(0).Sheets("�ݒ�").Range(�n���Fkey.Offset(x, 1), �n���Fkey.Offset(x, 5)))
                    End If
                End If
            End If
        Next x
        If Not �n���Frange Is Nothing Then
            Dim �[�q�F�̐��� As Shape
            Set �[�q�F�̐��� = .Shapes.AddShape(1, 100, 50, 150, 200)
            �[�q�F�̐���.Fill.Transparency = 1
            �[�q�F�̐���.Line.Visible = msoFalse
            �[�q�F�̐���.TextFrame2.TextRange.Font.size = 10
            For p = 1 To �n���Frange.count / 4
                �[�q�F�̐���.TextFrame.Characters.Text = �[�q�F�̐���.TextFrame.Characters.Text & "��" & �n���Frange(p, 4) & vbCrLf
                �[�q�F�̐���.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
            Next p
            �[�q�F�̐���.TextFrame.Characters.Text = Mid(�[�q�F�̐���.TextFrame.Characters.Text, 1, Len(�[�q�F�̐���.TextFrame.Characters.Text) - 1)
            ������ = 1
            For p = 1 To �n���Frange.count / 4
                �[�q�F�̐���.TextFrame2.TextRange.Characters(������, 1).Font.Fill.ForeColor.RGB = �n���Frange(p, 1).Interior.color
                ������ = ������ + Len(�n���Frange(p, 4)) + 2
            Next p
            �[�q�F�̐���.TextFrame2.MarginLeft = 0
            �[�q�F�̐���.TextFrame2.MarginRight = 0
            �[�q�F�̐���.TextFrame2.MarginTop = 0
            �[�q�F�̐���.TextFrame2.MarginBottom = 0
            �[�q�F�̐���.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            �[�q�F�̐���.Top = 0
            �[�q�F�̐���.Left = �F�̐���.Left + �F�̐���.Width + 5
            �[�q�F�̐���.Name = "�[�q�F�̐���"
            �[�q�F�̐���.TextFrame2.WordWrap = msoFalse
            �[�q�F�̐���.Select
            �F�̐���.Select False
            Selection.ShapeRange.Group.Select
            Selection.Name = "�F�̐���"
        End If
    End With
    
    'color�̃V�[�g��n��
    wb(0).Sheets("color").Copy before:=ws(1)
    'onkey��n��
    On Error Resume Next
        'WB(1).VBProject.VBComponents(WS(1).CodeName).CodeModule.AddFromFile myaddress(0,1) & "\OnKey" & "\003_�ƃ��C�A�E�g.txt"
    On Error GoTo 0
    ws(1).Activate
    wb(1).Save
    '�ݒ�t�@�C�����쐬
    Call TEXT�o��_�ݒ�_�ƃ��C�A�E�g�}(myPath & "\�ݒ�_�ƃ��C�A�E�g.txt")
    '���
    Set �n���Fkey = Nothing
    Set �n���Frange = Nothing
    Set �[�q�F�̐��� = Nothing
    Set wb(0) = Nothing
    Set wb(1) = Nothing
    Set ws(0) = Nothing
    Set ws(1) = Nothing
    Set myKey = Nothing
    Set mykey0 = Nothing
    Set mykey1 = Nothing
    Set �i�� = Nothing
    Set ��H���� = Nothing
    Set ��H����0 = Nothing
    Set ��H����1 = Nothing
    Set �[��0 = Nothing
    Set �[��1 = Nothing
    Set �[�q0 = Nothing
    Set �[�q1 = Nothing
    Set �[��a = Nothing
    Set cava = Nothing
    Call �œK�����ǂ�
End Function

Public Function ���i�i�Ԃ̃V�[�g�쐬()
    Set wb(0) = ThisWorkbook

    Dim sTime As Single: sTime = Timer
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "Ver"
    Dim newSheetName As String: newSheetName = "���i�i��"
    Call addressSet(wb(0))
    
    '�������O�̃t�@�C�������邩�m�F
    Dim ws As Worksheet
    myCount = 0
line10:

    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            flg = True
            Exit For
        End If
    Next ws
    
    If flg = True Then
        myCount = myCount + 1
        newSheetName = newSheetName & myCount
        GoTo line10
    End If
        
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Cells.NumberFormat = "@"
    If newSheet.Name = "���i�i��" Then
        newSheet.Tab.color = RGB(255, 192, 0)
    End If
    
    With wb(0).Sheets("�t�B�[���h��")
        Dim �t�B�[���hRAN As Range, lastCol As Long
        Set �t�B�[���hRAN = .Cells.Find("�t�B�[���h��_���i�i��", , , 1).Offset(2, 0)
        lastCol = .Cells(�t�B�[���hRAN.Row, .Columns.count).End(xlToLeft).Column - �t�B�[���hRAN.Column + 1
        Set �t�B�[���hRAN = �t�B�[���hRAN.Offset(-1, 0)
        Set �t�B�[���hRAN = �t�B�[���hRAN.Resize(2, lastCol)
    End With
    
    With newSheet
        x = 2
        y = 5
        For r = 1 To �t�B�[���hRAN.count / 2
            Call �Z���̒��g��S�ēn��(.Cells(y + 0, x), �t�B�[���hRAN(1, r))
            Call �Z���̒��g��S�ēn��(.Cells(y + 1, x), �t�B�[���hRAN(2, r))
            .Columns(x).AutoFit
            x = x + 1
        Next r
        .Columns(1).ColumnWidth = 4
        '�E�B���h�E�̌Œ�
        .Cells(7, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call SetButtonsOnActiveSheet("openMenu")
    
    ���i�i�Ԃ̃V�[�g�쐬 = Round(Timer - sTime, 2)
    
End Function

Public Function TEXT�o��_color_UTF8()
    Call addressSet(ThisWorkbook)
    path = myAddress(1, 1) & "\ps\color.txt"
    Dim i As Integer
    Dim outdats() As String
    With Sheets("color")
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        Dim tempVal As String
        For y = 1 To lastRow
            For x = 1 To lastCol
                tempVal = tempVal & "," & .Cells(y, x)
            Next x
            tempVal = Mid(tempVal, 2)
            ReDim Preserve outdats(y - 1)
            outdats(y - 1) = tempVal
            tempVal = ""
        Next y
    End With
    
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open path For Output As #FileNumber
    '�z��̗v�f���������ďo�͂��܂��B
    Print #FileNumber, Join(outdats, vbCrLf)
    '���̓t�@�C������܂��B
    Close #FileNumber

End Function


