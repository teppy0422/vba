Attribute VB_Name = "Module0"
Public Const mySystemName As String = "���Y����+"
Public NMB���� As String
Public �n���}�^�C�v As String
Public �������i As Long
Public �n���\�� As String
Dim ��n�����i�i�� As String
Public ��n���}�\�� As String
Public ���\�L() As String
Public ���c As Long
Dim �T�u�}���i�i�� As String
Public myFont As String
Public ���i�i��RAN() As Variant 'Sheets(���i�i��)�̃f�[�^�Z�b�g�p
Public ���i�i��RANc As Long
Public ���i�i��R() As String
Public ���i�i��Rc As Long
Public �[���ꗗran() As String '�z���}���}�̂ݍ쐬���̒[���m�F�p
Public �}���}���i�i��() As String
Public newBook As Workbook
Public myBook As Workbook
Public �A�h���X(3) As String
Public myVer As String
Public �}���}�`�� As Long
Public myErrFlg As Boolean
Public �F�Ŕ��f As Boolean
Public �n���F�ݒ�() As String
Public �n����ƕ\�� As String
Public strArray() As String
Public ��d�W�~flg As Boolean
Public wb(9) As Workbook '0=���̃u�b�N�A3=��n����Ǝ҈ꗗ
Public ws(9) As Worksheet
Public sikibetu As Range
Public �z���T�usize() As String
Public �[���i���o�[�\�� As Boolean
Public �}���}�s�� As String
Public �������� As Boolean
Public ��n����Ǝ� As Boolean
Public ��n����Ǝ�RAN() As String
Public ��n����Ǝ҃V�[�g�� As String
Public RLTF�T�u As Boolean
Public MD As Boolean
Public SUB�f�[�^RAN() As String
Public QR��� As Boolean
Public �t�H�[������̌Ăяo�� As Boolean
Public �z���}�쐬temp As String  '�����W�f�[�^���������ǗU���f�[�^��鎞�p
Public �T���v���쐬���[�h As Boolean
Public cavCount As Long
Public ��n���_�� As Boolean '�z���U���Ő�n���ł��_�ł���

Sub PVSWcsv_csv�̃C���|�[�g()
'setup
    Dim thisBookName As String: thisBookName = ActiveWorkbook.Name
    Dim thisBookPath As String: thisBookPath = ActiveWorkbook.Path
    '���͂̐ݒ�(�C���|�[�g�t�@�C��)
    Dim TargetName As String: TargetName = "PVSW_RLTF"
    Dim Target As New FileSystemObject
    Dim targetFolder As Variant: Set targetFolder = Target.GetFolder(thisBookPath & "\" & TargetName).Files
    
    '�o�͂̐ݒ�
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim outY As Long: outY = 1
    Dim outX As Long
    Dim lastgyo As Long: lastgyo = 1
    Dim fileCount As Long: fileCount = 0
    Dim TargetFile As Variant
    Dim aa As String
    
    With Workbooks(thisBookName).Sheets(outSheetName)
        .Cells.NumberFormat = "@"
    End With
'loop
    For Each TargetFile In targetFolder
        Dim csvPath As String: csvPath = TargetFile
        Dim csvName As String: csvName = TargetFile.Name
        Dim LngLoop As Long
        Dim intFino As Integer
        
        intFino = FreeFile
        Open csvPath For Input As #intFino
        Dim inX As Long, addX As Long
        Dim temp
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For inX = LBound(temp) To UBound(temp)
                With Workbooks(thisBookName).Sheets(outSheetName)
                    'Debug.Print (temp(inX))
                    If fileCount <> 0 And Len(temp(inX)) = 15 And outY = 1 Then
                        Dim searchX As Long: searchX = 0
                        Do
                            If Len(.Cells(1, 1).Offset(0, searchX)) <> 15 Then
                                'Stop
                                .Columns(searchX + 1).EntireColumn.Insert
                                .Cells(1, searchX + 1).NumberFormat = "@"
                                .Cells(1, searchX + 1) = temp(inX)
                                If inX = 0 Then addX = searchX
                            Exit Do
                            End If
                        searchX = searchX + 1
                        Loop
                    ElseIf fileCount = 0 Then
                        outX = inX
'                        If lastgyo = 1 Then
'                            .Columns(outX + 1).NumberFormat = "@"
'                            If temp(inX) = "�n�_���[�����ʎq" Then .Columns(outX + 1).NumberFormat = 0
'                            If temp(inX) = "�I�_���[�����ʎq" Then .Columns(outX + 1).NumberFormat = 0
'                            If temp(inX) = "�n�_���L���r�e�BNo." Then .Columns(outX + 1).NumberFormat = 0
'                            If temp(inX) = "�I�_���L���r�e�BNo." Then .Columns(outX + 1).NumberFormat = 0
'                        End If
                        .Cells(lastgyo, outX + 1) = Replace(temp(inX), vbLf, "")
                    ElseIf outY <> 1 Then
                    'Stop
                        outX = inX + addX + 1
                        .Cells(lastgyo, outX).NumberFormat = "@"
                        .Cells(lastgyo, outX) = temp(inX)
                    End If
                End With
            Next inX
        outY = outY + 1
        lastgyo = lastgyo + 1
        Loop
    outY = 1
    fileCount = fileCount + 1
    lastgyo = lastgyo - 1
    Next TargetFile
    
    '���ёւ�
    With Workbooks(thisBookName).Sheets(outSheetName)
        Dim titleRange As Range
        Set titleRange = .Range(.Cells(1, 1), .Cells(1, .Cells(1, 1).End(xlToRight).Column))
        Dim r As Variant
        Dim �D��1 As Long, �D��2 As Long, �D��3 As Long, �D��4 As Long, �D��5 As Long, �D��6 As Long
        For Each r In titleRange
            If r = "�n�_���[�����ʎq" Then �D��1 = r.Column
            If r = "�n�_���L���r�e�BNo." Then �D��2 = r.Column
            If r = "�n�_����H����" Then �D��3 = r.Column
            If r = "�I�_���[�����ʎq" Then �D��4 = r.Column
            If r = "�I�_���L���r�e�BNo." Then �D��5 = r.Column
            If r = "�I�_����H����" Then �D��6 = r.Column
        Next r
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��3).address), Order:=xlAscending
            .add key:=Range(Cells(1, �D��4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��5).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��6).address), Order:=xlAscending
        End With
            .Sort.SetRange Range(Rows(2), Rows(lastgyo))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    
    Set targetFolder = Nothing
    Close #intFino
End Sub

Sub PVSWcsv_csv�̃C���|�[�g_2029()
'setup
    Call �A�h���X�Z�b�g(myBook)
    
    Dim thisBookName As String: thisBookName = ActiveWorkbook.Name
    Dim thisBookPath As String: thisBookPath = ActiveWorkbook.Path
    Dim mySheetName As String: mySheetName = "���i�i��"
    '���͂̐ݒ�(�C���|�[�g�t�H���_)
    Dim TargetName As String: TargetName = "01_PVSW_csv"
    Dim Target As New FileSystemObject
    
    a = Dir(thisBookPath & "\" & TargetName, vbDirectory)
    
    If a = "" Then
        MkDir (thisBookPath & "\" & TargetName)
        MsgBox "PVSW�̃t�@�C����������܂���B�m�F���ĉ������B"
        Shell "C:\Windows\explorer.exe " & thisBookPath & "\" & TargetName, vbNormalFocus
        End
    End If
    
    Dim targetFolder As Variant: Set targetFolder = Target.GetFolder(thisBookPath & "\" & TargetName).Files
    
    '�Ώۂ̃t�@�C�����̊m�F
    Dim TargetFile As Variant: Dim fileCount As Long
    For Each TargetFile In targetFolder
        Dim csvPath As String: csvPath = TargetFile
        Dim csvName As String: csvName = TargetFile.Name
        fileCount = fileCount + 1
    Next TargetFile
    If fileCount = 0 Then
        MsgBox "PVSW�̃t�@�C����������܂���B�m�F���ĉ������B"
        Shell "C:\Windows\explorer.exe " & thisBookPath & "\" & TargetName, vbNormalFocus
        End
    End If
    '�o�͂̐ݒ�
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim outY As Long: outY = 1
    Dim outX As Long
    Dim lastgyo As Long: lastgyo = 1
    fileCount = 0
    Dim aa As String
    
    Dim ws As Worksheet, myCount As Long
    outsheetname2 = outSheetName
line10:
    flg = False
    For Each ws In Worksheets
        If ws.Name = outsheetname2 Then
            myCount = myCount + 1
            outsheetname2 = outSheetName & "_" & myCount
            GoTo line10
            Exit For
        End If
    Next ws
    If myCount <> 0 Then outSheetName = outSheetName & "_" & myCount
    Dim newSheet As Worksheet
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Worksheets.add after:=Sheets("���i�i��")
        'Set newSheet = Worksheets.Add(after:=Worksheets(mySheetName))
        'Set newSheet = Worksheets.Add(after:=Worksheets(mySheetName))
        ActiveSheet.Name = outSheetName
        Sheets(outSheetName).Cells.NumberFormat = "@"
        If outSheetName = "PVSW_RLTF" Then
            ActiveSheet.Tab.color = 14470546
        End If

    End If
    
    With Workbooks(thisBookName).Sheets("�t�B�[���h��")
        Set key = .Cells.Find("�t�B�[���h��_�ʏ�", , , 1).Offset(1, 0)
        Set �t�B�[���hran0 = .Range(.Cells(key.Row, key.Column), .Cells(key.Row + 8, .Cells(key.Row, .Columns.count).End(xlToLeft).Column))
        
        Set key = .Cells.Find("�t�B�[���h��_�ǉ�", , , 1).Offset(1, 0)
        Set �t�B�[���hran1 = .Range(.Cells(key.Row, key.Column), .Cells(key.Row + 1, .Cells(key.Row + 1, .Columns.count).End(xlToLeft).Column))
        
        Set key = .Cells.Find("�t�B�[���h��_�ǉ�2", , , 1).Offset(1, 0)
        Set �t�B�[���hran2 = .Range(.Cells(key.Row, key.Column), .Cells(key.Row + 1, .Cells(key.Row + 1, .Columns.count).End(xlToLeft).Column))
        Set key = Nothing
    End With
'loop
    For Each TargetFile In targetFolder
        csvPath = TargetFile
        csvName = TargetFile.Name
        Dim LngLoop As Long
        Dim intFino As Integer
        
        intFino = FreeFile
        Open csvPath For Input As #intFino
        Dim inX As Long, addX As Long
        Dim temp
        �t�B�[���hflg = False
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For inX = LBound(temp) To UBound(temp)
                With Workbooks(thisBookName).Sheets(outSheetName)
                    'Debug.Print (temp(inX))
                    If fileCount <> 0 And Len(temp(inX)) = 15 And outY = 1 Then
                        Dim searchX As Long: searchX = 0
                        Do
                            If Len(.Cells(1, 1).Offset(0, searchX)) <> 15 Then
                                .Columns(searchX + 1).EntireColumn.Insert
                                .Cells(1, searchX + 1).NumberFormat = "@"
                                .Cells(1, searchX + 1) = temp(inX)
                                If inX = 0 Then addX = searchX
                                GoSub ���i�i�Ԃ̒ǉ�
                            Exit Do
                            End If
                        searchX = searchX + 1
                        Loop
                    ElseIf fileCount = 0 Then
                        outX = inX
                        If lastgyo = 1 Then
                            .Columns(outX + 1).NumberFormat = "@"
                            .Cells(lastgyo, outX + 1) = Replace(temp(inX), vbLf, "")
                          
                            If �t�B�[���hflg = False And Len(temp(inX)) <> 15 Then
                                �t�B�[���hflg = True
                            End If
                            If �t�B�[���hflg = True Then
                                '�t�B�[���h���̒u������
                                Set key = �t�B�[���hran0.Find(temp(inX), , , 1)
                                If key Is Nothing Then
                                    Debug.Print temp(inX)
                                    MsgBox "�F���ł��Ȃ��t�B�[���h�� " & temp(inX) & " ���܂܂�Ă��܂��B" & vbCrLf & _
                                           "�Y������t�B�[���h���̉���" & temp(inX) & " ��ǉ����Ă��������B"
                                    Sheets("�t�B�[���h��").Visible = True
                                    Sheets("�t�B�[���h��").Select
                                    Call �œK�����ǂ�
                                    End
                                Else
                                    .Cells(1, outX + 1) = �t�B�[���hran0(1, key.Column)
                                    .Cells(1, outX + 1).Interior.color = �t�B�[���hran0(1, key.Column).Interior.color
                                    .Cells(1, outX + 1).Borders.LineStyle = �t�B�[���hran0(1, key.Column).Borders.LineStyle
                                    'PVSW�t�B�[���hcol = PVSW�t�B�[���hcol & "," & outX + 1 - �t�B�[���h�擪col
                                End If
                            Else
                                GoSub ���i�i�Ԃ̒ǉ�
                            End If
                        Else
                            .Cells(lastgyo, outX + 1) = Replace(temp(inX), vbLf, "")
                        End If
                        
                    ElseIf outY <> 1 Then
                    'Stop
                        outX = inX + addX + 1
                        .Cells(lastgyo, outX).NumberFormat = "@"
                        .Cells(lastgyo, outX) = temp(inX)
                    End If
                End With
            Next inX
            outY = outY + 1
            lastgyo = lastgyo + 1
        Loop
        outY = 1
        fileCount = fileCount + 1
    'lastgyo = lastgyo - 1
    Next TargetFile


    '�t�B�[���h�����d������ꍇ�A�E�ɂ�������폜
    Dim lastCol As Long
    With Workbooks(thisBookName).Sheets(outSheetName)
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        For X = 1 To lastCol
            For x2 = X + 1 To lastCol
                If .Cells(1, X) = .Cells(1, x2) Then
                    .Columns(x2).Delete
                End If
            Next x2
        Next X
    End With
    
    '�l���𐮂���
    Dim �t�B�[���hRow As Long: �t�B�[���hRow = 6
    With Workbooks(thisBookName).Sheets(outSheetName)
        .Range(.Rows(1), .Rows(�t�B�[���hRow - 1)).Insert
        .Cells(�t�B�[���hRow - 3, 1) = "���i�i��s"
        ���i�i�ԓ_�� = .Rows(�t�B�[���hRow).Find("�d�����ʖ�", , , 1).Column - 1
        
        If ���i�i�ԓ_�� = 1 Then
            
        Else
            .Cells(�t�B�[���hRow - 3, ���i�i�ԓ_��) = "���i�i��e"
        End If
        maxCol = .Cells(�t�B�[���hRow, .Columns.count).End(xlToLeft).Column
        For X = 1 To maxCol
            .Cells(1, X) = "PVSW"
        Next X
        '��H�}�g���N�X�p�̃t�B�[���h��ǉ�
        Dim myField As String: myField = "SubNo,SubNo2,SubNo3,SSC,�����@,�n�I��"
        Dim myFieldSP: myFieldSP = Split(myField, ",")
        For X = LBound(myFieldSP) To UBound(myFieldSP)
            .Columns(���i�i�ԓ_�� + X + 1).Insert
            .Cells(�t�B�[���hRow, ���i�i�ԓ_�� + X + 1).Value = myFieldSP(X)
            .Cells(�t�B�[���hRow, ���i�i�ԓ_�� + X + 1).Interior.color = RGB(5, 5, 5)
            .Cells(�t�B�[���hRow, ���i�i�ԓ_�� + X + 1).Font.color = RGB(250, 250, 250)
        Next X
    End With
    
    '�ǉ��t�B�[���h
    With Workbooks(thisBookName).Sheets(outSheetName)
        lastCol = .Cells(�t�B�[���hRow, .Columns.count).End(xlToLeft).Column
        For Y = 1 To 2
            For X = 1 To �t�B�[���hran1.count / 2
                .Cells(�t�B�[���hRow + Y - 2, lastCol + X) = �t�B�[���hran1(Y, X)
                If �t�B�[���hran1(Y, X).Interior.ColorIndex <> xlNone Then
                    .Cells(�t�B�[���hRow + Y - 2, lastCol + X).Interior.color = �t�B�[���hran1(Y, X).Interior.color
                End If
                .Cells(�t�B�[���hRow + Y - 2, lastCol + X).Borders.LineStyle = �t�B�[���hran1(Y, X).Borders.LineStyle
                .Cells(1, lastCol + X) = "RLTFA"
            Next X
        Next Y
    End With
    
    '�ǉ��t�B�[���h2
    With Workbooks(thisBookName).Sheets(outSheetName)
        lastCol = .Cells(�t�B�[���hRow, .Columns.count).End(xlToLeft).Column
        For Y = 1 To 2
            For X = 1 To �t�B�[���hran2.count / 2
                .Cells(�t�B�[���hRow + Y - 2, lastCol + X) = �t�B�[���hran2(Y, X)
                If �t�B�[���hran2(Y, X).Interior.ColorIndex <> xlNone Then
                    .Cells(�t�B�[���hRow + Y - 2, lastCol + X).Interior.color = �t�B�[���hran2(Y, X).Interior.color
                End If
                .Cells(�t�B�[���hRow + Y - 2, lastCol + X).Borders.LineStyle = �t�B�[���hran2(Y, X).Borders.LineStyle
                .Cells(1, lastCol + X) = "ADD"
            Next X
        Next Y
    End With
    
    With Workbooks(thisBookName).Sheets(outSheetName)
        '���ёւ�
        Dim titleRange As Range
        Set titleRange = .Range(.Cells(�t�B�[���hRow, 1), .Cells(�t�B�[���hRow, .Cells(�t�B�[���hRow, .Columns.count).End(xlToLeft).Column))
        Dim r As Variant
        Dim �D��1 As Long, �D��2 As Long, �D��3 As Long, �D��4 As Long, �D��5 As Long, �D��6 As Long
        For Each r In titleRange
            If r = "�d�����ʖ�" Then �D��1 = r.Column '�u�������Ŏg�p���Ă���_�ɒ���
            If r = "�n�_���L���r�e�BNo." Then �D��2 = r.Column
            If r = "�n�_����H����" Then �D��3 = r.Column
            If r = "�I�_���[�����ʎq" Then �D��4 = r.Column
            If r = "�I�_���L���r�e�BNo." Then �D��5 = r.Column
            If r = "�I�_����H����" Then �D��6 = r.Column
        Next r
        lastgyo = .Cells(.Rows.count, �D��1).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(�t�B�[���hRow, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(�t�B�[���hRow, 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Range(Cells(1, �D��3).Address), Order:=xlAscending
'            .Add key:=Range(Cells(1, �D��4).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Range(Cells(1, �D��5).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Range(Cells(1, �D��6).Address), Order:=xlAscending
        End With
            .Sort.SetRange Range(Rows(�t�B�[���hRow), Rows(lastgyo))
            .Sort.Header = xlYes
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
        '�u������
        .Range(.Cells(�t�B�[���hRow + 1, 1), .Cells(lastgyo, �D��1 - 1)).Replace "1", "0"
        Set myCell = .Rows(�t�B�[���hRow).Find("�d�����ʖ�", , , 1).Offset(-1, 0)
        myCell.Value = "�R�����g"
        myCell.AddComment
        myCell.Comment.Text "Ctrl+R�ŃR�����g�̕\���E��\���̐؂�ւ�"
        myCell.Comment.Shape.TextFrame.AutoSize = True
        '�R�����g�\���؊���b�̃R�s�[
        '.Shapes.Range("�R�����gb").Left = .Cells(2, �D��1 + 1).Left
        '.Shapes.Range("�R�����gb").Top = .Cells(2, �D��1 + 1).Top
        '�E�B���h�E�g�̌Œ�
        .Activate
        .Cells(7, 1).Select
        ActiveWindow.FreezePanes = True
        '�C�x���g�̒ǉ�
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile �A�h���X(0) & "\OnKey" & "\001_PVSW_RLTF_make.txt"
        On Error GoTo 0
    End With
    
    Set targetFolder = Nothing
    Close #intFino
    
    Exit Sub

���i�i�Ԃ̒ǉ�:
    With Sheets(mySheetName)
        Set key2 = .Cells.Find("���C���i��", , , 1)
        .Columns(key2.Column).NumberFormat = "@"
        Set key3 = .Columns(key2.Column).Find(temp(inX), , , 1)
        If key3 Is Nothing Then
            addRow = .Cells(.Rows.count, key2.Column).End(xlUp).Row + 1
            .Cells(addRow, key2.Column) = temp(inX)
            .Cells(addRow, key2.Column).Interior.color = RGB(255, 230, 0)
        End If
    End With
    Return

End Sub

Sub PVSWcsv��RLTFA�����H�����擾_Ver2026()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim i As Long, ii As Long, strArrayS As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Sheets(outSheetName).Activate
    
    With Workbooks(myBookName).Sheets("�ݒ�")
        '�[�q�t�@�~���[
        Dim �[�q�t�@�~���[() As String
        ReDim �[�q�t�@�~���[(5, 0) As String
        ii = 0
        Set key = .Cells.Find("�[�q�t�@�~���[_", , , 1)
        Do
            If key.Offset(ii, 1) = "" Then Exit Do
            ReDim Preserve �[�q�t�@�~���[(5, ii)
            �[�q�t�@�~���[(0, ii) = key.Offset(ii, 1)
            �[�q�t�@�~���[(1, ii) = key.Offset(ii, 2)
            �[�q�t�@�~���[(2, ii) = key.Offset(ii, 3)
            �[�q�t�@�~���[(3, ii) = key.Offset(ii, 1).Interior.color
            �[�q�t�@�~���[(4, ii) = key.Offset(ii, 1).Row
            �[�q�t�@�~���[(5, ii) = key.Offset(ii, 4)
            ii = ii + 1
        Loop
        
        'Call ���ޏڍ�_�[�q�t�@�~���[(�A�h���X(1) & "\���ޏڍ�.txt", �[�q�t�@�~���[)
        
        '�[�q�t�@�~���[�̈ꎞ�ۊ�
        Set key = .Cells.Find("�[�q�t�@�~���[temp_", , , 1)
        .Range(key.Offset(0, 1), key.Offset(10, 1)).Name = "�[�q�t�@�~���[�͈�"
        .Range("�[�q�t�@�~���[�͈�").Clear
        '�d���i��
        Dim �d���i��() As String
        ReDim �d���i��(5, 0) As String
        ii = 0
        Set key = .Cells.Find("�d���i��_", , , 1)
        Do
            If key.Offset(ii, 1) = "" Then Exit Do
            ReDim Preserve �d���i��(5, ii)
            �d���i��(0, ii) = key.Offset(ii, 1)
            �d���i��(1, ii) = key.Offset(ii, 2)
            �d���i��(2, ii) = key.Offset(ii, 3)
            �d���i��(3, ii) = key.Offset(ii, 1).Interior.color
            �d���i��(4, ii) = key.Offset(ii, 1).Row
            �d���i��(5, ii) = key.Offset(ii, 4)
            ii = ii + 1
        Loop
        '�d���i��̈ꎞ�ۊ�
        Set key = .Cells.Find("�d���i��temp_", , , 1)
        .Range(key.Offset(0, 1), key.Offset(10, 1)).Name = "�d���i��͈�"
        .Range("�d���i��͈�").Clear
    End With
    
    Call �œK��
    With Workbooks(myBookName).Sheets(outSheetName)
        Dim PVSW����Row As Long: PVSW����Row = .Cells.Find("�d�����ʖ�", , , 1).Row
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim PVSW���i�i��sCol As Long: PVSW���i�i��sCol = .Cells.Find("���i�i��s", , , 1).Column
        Dim PVSW���i�i��eCol As Long
        On Error Resume Next
        PVSW���i�i��eCol = .Cells.Find("���i�i��e", , , 1).Column
        On Error GoTo 0
        If PVSW���i�i��eCol = 0 Then PVSW���i�i��eCol = PVSW���i�i��sCol
        Dim �^�C�g�� As Range: Set �^�C�g�� = .Rows(PVSW����Row)
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW����Col).End(xlUp).Row
        Dim PVSW�d��sCol As Long: PVSW�d��sCol = .Cells.Find("�d�������擾s", , , 1).Column
        Dim PVSW�d��eCol As Long: PVSW�d��eCol = .Cells.Find("�d�������擾e", , , 1).Column
        Dim PVSW�n�I��Col As Long: PVSW�n�I��Col = .Cells.Find("�n�I��", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW�\��Col As Long: PVSW�\��Col = .Cells.Find("�\��_", , , 1).Column
        Dim �ڑ�Gcol As Long: �ڑ�Gcol = .Cells.Find("�ڑ�G_", , , 1).Column
        Dim PVSW�i��Col As Long: PVSW�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim PVSW�i���Col As Long: PVSW�i���Col = .Cells.Find("�i���_", , , 1).Column
        Dim PVSW�T�C�YCol As Long: PVSW�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim PVSW�T�C�Y�ď�Col As Long: PVSW�T�C�Y�ď�Col = .Cells.Find("�T��_", , , 1).Column
        Dim PVSW�FCol As Long: PVSW�FCol = .Cells.Find("�F_", , , 1).Column
        Dim PVSW�F��Col As Long: PVSW�F��Col = .Cells.Find("�F��_", , , 1).Column
        Dim PVSW��IDcol As Long: PVSW��IDcol = .Cells.Find("��ID_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        'Dim PVSWG��GNoCol As Long: PVSWG��GNoCol = .Cells.Find("G��GNo_", , , 1).Column
        Dim PVSW�T�u0Col As Long: PVSW�T�u0Col = .Cells.Find("���0_", , , 1).Column
        Dim PVSW�d�㐡�@Col As Long: PVSW�d�㐡�@Col = .Cells.Find("�d�㐡�@_", , , 1).Column
        Dim PVSW�ؒf��Col As Long: PVSW�ؒf��Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim PVSW�n��HCol As Long: PVSW�n��HCol = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n�[��Col As Long: PVSW�n�[��Col = .Cells.Find("�n�_���[��_", , , 1).Column
        Dim PVSW�n�[Col As Long: PVSW�n�[Col = .Cells.Find("�n�_���[�q_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n�}Col As Long: PVSW�n�}Col = .Cells.Find("�n�_���}_", , , 1).Column
        Dim PVSW�n�ڑ��\��Col As Long: PVSW�n�ڑ��\��Col = .Cells.Find("�n�_���ڑ��\��_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_�����i_", , , 1).Column
        Dim PVSW�n��2Col As Long: PVSW�n��2Col = .Cells.Find("�n�_�����i2_", , , 1).Column
        Dim PVSW�n��3Col As Long: PVSW�n��3Col = .Cells.Find("�n�_�����i3_", , , 1).Column
        Dim PVSW�n��4Col As Long: PVSW�n��4Col = .Cells.Find("�n�_�����i4_", , , 1).Column
        Dim PVSW�n��5Col As Long: PVSW�n��5Col = .Cells.Find("�n�_�����i5_", , , 1).Column
        Dim PVSW�I��HCol As Long: PVSW�I��HCol = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I�[��Col As Long: PVSW�I�[��Col = .Cells.Find("�I�_���[��_", , , 1).Column
        Dim PVSW�I�[Col As Long: PVSW�I�[Col = .Cells.Find("�I�_���[�q_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I�}Col As Long: PVSW�I�}Col = .Cells.Find("�I�_���}_", , , 1).Column
        Dim PVSW�I�ڑ��\��Col As Long: PVSW�I�ڑ��\��Col = .Cells.Find("�I�_���ڑ��\��_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_�����i_", , , 1).Column
        Dim PVSW�I��2Col As Long: PVSW�I��2Col = .Cells.Find("�I�_�����i2_", , , 1).Column
        Dim PVSW�I��3Col As Long: PVSW�I��3Col = .Cells.Find("�I�_�����i3_", , , 1).Column
        Dim PVSW�I��4Col As Long: PVSW�I��4Col = .Cells.Find("�I�_�����i4_", , , 1).Column
        Dim PVSW�I��5Col As Long: PVSW�I��5Col = .Cells.Find("�I�_�����i5_", , , 1).Column
'       Dim PVSW�n����Col As Long: PVSW�n����Col = .Cells.Find("�n�_������_", , , 1).Column
'       Dim PVSW�I����Col As Long: PVSW�I����Col = .Cells.Find("�I�_������_", , , 1).Column
        'PVSW�̒l
        Dim PVSW�n��HCol2 As Long: PVSW�n��HCol2 = .Cells.Find("�n�_����H����", , , 1).Column
        .Columns(PVSW�n��HCol2).ClearComments
        Dim PVSW�n�[��Col2 As Long: PVSW�n�[��Col2 = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        .Columns(PVSW�n�[��Col2).ClearComments
        Dim PVSW�I��HCol2 As Long: PVSW�I��HCol2 = .Cells.Find("�I�_����H����", , , 1).Column
        .Columns(PVSW�I��HCol2).ClearComments
        Dim PVSW�I�[��Col2 As Long: PVSW�I�[��Col2 = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        .Columns(PVSW�I�[��Col2).ClearComments
        
        Dim PVSW�nCavCol2 As Long: PVSW�nCavCol2 = .Cells.Find("�n�_���L���r�e�B", , , 1).Column
        Dim PVSW�ICavCol2 As Long: PVSW�ICavCol2 = .Cells.Find("�I�_���L���r�e�B", , , 1).Column
        '�A���}�b�`�̃J�E���g�p�z��
        Dim unCount(1, 5) As Long
        
        Dim PVSW�n���Col As Long: PVSW�n���Col = .Cells.Find("�n�_���[�����i��", , , 1).Column
        .Columns(PVSW�n���Col).ClearComments
        Dim PVSW�I���Col As Long: PVSW�I���Col = .Cells.Find("�I�_���[�����i��", , , 1).Column
        .Columns(PVSW�I���Col).ClearComments
        
        .Range(.Cells(PVSW����Row + 1, PVSW�d��sCol), .Cells(PVSWlastRow, PVSW�d��eCol)).Clear
        .Range(.Cells(PVSW����Row + 1, PVSW�d��sCol), .Cells(.Rows.count, PVSW�d��eCol)).NumberFormat = "@"
        .Columns(PVSW�d�㐡�@Col).NumberFormat = 0
        '�}�g���N�X�̐F�𖳂��ɕύX
        .Range(.Cells(PVSW����Row + 1, PVSW���i�i��sCol), .Cells(.Rows.count, PVSW���i�i��eCol)).Interior.Pattern = xlNone
        '��r����
        Dim PVSW��rCol(23) As Long
        PVSW��rCol(0) = PVSW�i��Col
        PVSW��rCol(1) = PVSW�T�C�YCol
        PVSW��rCol(2) = PVSW�T�C�Y�ď�Col
        PVSW��rCol(3) = PVSW�FCol
        PVSW��rCol(4) = PVSW�F��Col
        PVSW��rCol(5) = PVSW����Col
        PVSW��rCol(6) = PVSW����Col
        PVSW��rCol(7) = PVSWJCDFCol
        PVSW��rCol(8) = PVSW�d�㐡�@Col
        PVSW��rCol(9) = PVSW�n��HCol
        PVSW��rCol(10) = PVSW�n�[��Col
        PVSW��rCol(11) = PVSW�n�[Col
        PVSW��rCol(12) = PVSW�n�}Col
        PVSW��rCol(13) = PVSW�n�ڑ��\��Col
        PVSW��rCol(14) = PVSW�n��Col
        PVSW��rCol(15) = PVSW�I��HCol
        PVSW��rCol(16) = PVSW�I�[��Col
        PVSW��rCol(17) = PVSW�I�[Col
        PVSW��rCol(18) = PVSW�I�}Col
        PVSW��rCol(19) = PVSW�I�ڑ��\��Col
        PVSW��rCol(20) = PVSW�I��Col
        PVSW��rCol(21) = PVSW�\��Col
        PVSW��rCol(22) = PVSW�T�u0Col
        PVSW��rCol(23) = �ڑ�Gcol
    End With
    
    Dim in���i�i�� As String, in�\�� As String, in�i�� As String, in�i��� As String, in�T�C�Y As String, in�T�C�Y�ď� As String, in�F As String, in�F�ď� As String, _
        in���� As String, in���� As String, inJCDF As String, in�n�_�[�q As String, in�n�_�}���} As String, in�I�_�[�q As String, in�I�_�}���} As String, _
        in�\���⑫ As String, in�T�u0 As String, in�n�_���i As String, in�I�_���i As String, in�n�_���i2 As String, in�I�_���i2 As String, in��ID As String, _
        in�n�_�ڑ��\�� As String, in�I�_�ڑ��\�� As String, in�n�_���i3 As String, in�I�_���i3 As String, in�n�_���i4 As String, in�I�_���i4 As String, in�n�_���i5 As String, in�I�_���i5 As String
    Dim in�d�㐡�@ As Long, myKey As Variant
    Dim in��H(1) As String, in�[��(1) As String, in�[�q(1) As String, in���i2(1) As String, in�}���}(1) As String, in�ڑ��\��(1) As String, in���i3(1) As String, _
        in���i4(1) As String, in���i5(1) As String, �ڑ�G As String
    
Dim sTime As Single: sTime = Timer
Debug.Print "s"

    For c = 1 To ���i�i��RANc
        Set myKey = �^�C�g��.Find(���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), c), , , 1)
        If myKey Is Nothing Then GoTo nextC

        ���i�i��tc = myKey.Column
        Dim inTXT As String
        Dim RLTF As String
        RLTF = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "RLTF-A"), c)
        inTXT = ActiveWorkbook.Path & "\05_RLTF_A\" & RLTF & ".txt"
        Dim inFNo As Integer
        inFNo = FreeFile
        If Dir(inTXT) = "" Then GoTo nextC
        
        Open inTXT For Input As #inFNo
        Dim in��H�� As Long, in�}���}�� As Long, in����c As Long
        in�N���� = "": in��H�� = 0: in��z���� = ""
        in�}���}�� = 0: in����c = 0
        Do Until EOF(inFNo)
            Line Input #inFNo, aa
            in���i�i�� = Replace(Mid(aa, 1, 15), " ", "")
            If Replace(���i�i��RAN(1, c), " ", "") = in���i�i�� Then
                in�\�� = Mid(aa, 27, 4)
                in�\���⑫ = Mid(aa, 31, 1)
                If Left(in�\��, 1) <> "T" And Left(in�\��, 1) <> "B" And (in�\���⑫ = "0" Or in�\���⑫ = " ") Then
                    in�i�� = Mid(aa, 33, 3)
                    in�i��� = Mid(aa, 41, 10)
                    in�T�C�Y = Mid(aa, 36, 3)
                    If in�T�C�Y <> "   " Then in��H�� = in��H�� + 1
                    in�T�C�Y�ď� = Replace(Mid(aa, 51, 5), " ", "")
                    in�F = Mid(aa, 39, 2)
                    in�F�ď� = Replace(Mid(aa, 56, 7), " ", "")
                    in��ID = Mid(aa, 115, 2)
                    If in��ID = "00" Then in��ID = Empty
                    in���� = Replace(Mid(aa, 87, 1), " ", "")
                    in���� = Replace(Mid(aa, 88, 3), " ", "")
                    If in���� = "A" Or in���� = "G" Or in�敪 = "N" Then in����c = in����c + 1
                    inJCDF = Replace(Mid(aa, 539, 5), " ", "")
                    If inJCDF = "0000" Then inJCDF = Empty
                    'inG��Gno = Mid(aa, 539, 5)
                    
                    in��H(0) = Replace(Mid(aa, 96, 6), " ", "")
                    in�[��(0) = Mid(aa, 69, 3): If in�[��(0) = "000" Then in�[��(0) = "" Else in�[��(0) = CLng(in�[��(0))
                    in�[�q(0) = Replace(Mid(aa, 175, 10), " ", "")
                    
                    in���i2(0) = Replace(Mid(aa, 195, 10), " ", "")
                    in���i3(0) = Replace(Mid(aa, 215, 10), " ", "")
                    in���i4(0) = Replace(Mid(aa, 235, 10), " ", "")
                    in���i5(0) = Replace(Mid(aa, 255, 10), " ", "")
                    in�}���}(0) = Replace(Mid(aa, 167, 2), " ", "")
                    If in�}���}(0) <> "" Then in�}���}�� = in�}���}�� + 1
                    in�ڑ��\��(0) = Replace(Mid(aa, 189, 4), " ", "")
                    If in�ڑ��\��(0) = "0000" Then in�ڑ��\��(0) = Empty
                    
                    in��H(1) = Replace(Mid(aa, 102, 6), " ", "")
                    in�[��(1) = Mid(aa, 72, 3): If in�[��(1) = "000" Then in�[��(1) = "" Else in�[��(1) = CLng(in�[��(1))
                    in�[�q(1) = Replace(Mid(aa, 275, 10), " ", "")
                    in���i2(1) = Replace(Mid(aa, 295, 10), " ", "")
                    in���i3(1) = Replace(Mid(aa, 315, 10), " ", "")
                    in���i4(1) = Replace(Mid(aa, 335, 10), " ", "")
                    in���i5(1) = Replace(Mid(aa, 355, 10), " ", "")
                    in�}���}(1) = Replace(Mid(aa, 171, 2), " ", "")
                    If in�}���}(1) <> "" Then in�}���}�� = in�}���}�� + 1
                    in�ڑ��\��(1) = Replace(Mid(aa, 289, 4), " ", "")
                    If in�ڑ��\��(1) = "0000" Then in�ڑ��\��(1) = Empty
                    in�T�u0 = Mid(aa, 155, 4)
                    in�d�㐡�@ = CLng(Replace(Mid(aa, 64, 5), " ", ""))
                    If in�d�㐡�@ = 0 Then in�d�㐡�@ = CLng(Replace(Mid(aa, 148, 5), " ", ""))
                    
                    in�N���� = "20" & Mid(aa, 482, 2) & "/" & Mid(aa, 484, 2) & "/" & Mid(aa, 486, 2)
                    in��z���� = Mid(aa, 19, 2) & Mid(aa, 23, 1)
                    
                    Debug.Print "in�\��", "in����", "injcdf", "in��id"
                    Debug.Print in�\��, in����, inJCDF, in��ID
                    
                    If in�\�� = "0043" Then Stop
                    
                    '�ڑ���\���O���[�v_2.191.13
                    If in���� = "E" And inJCDF <> Empty Then
                        '�V�[���h
                        �ڑ�G = "E" & Mid(inJCDF, 2)
                    ElseIf in���� = "E" Then
                        '�V�[���h��JCDF�������ꍇ
                        �ڑ�G = "E" & in��ID
                    ElseIf Mid(inJCDF, 1, 1) = "W" Then
                        '�{���_�[
                        �ڑ�G = inJCDF
                    ElseIf in���� = "#" Or in���� = "*" Or in���� = "=" Or in���� = "<" Then
                        'Tw
                        �ڑ�G = "Tw" & in��ID
                    ElseIf inJCDF <> Empty Then
                        'J
                        �ڑ�G = inJCDF
                    ElseIf in���� = "BBB" Or in���� = "RRR" Then
                        �ڑ�G = "BAT"
                    Else
                        �ڑ�G = Empty
                    End If
                    
                    
                    in�n�_�� = "": in�I�_�� = ""
'                    If in�\���⑫ = "0" Or in�\���⑫ = " " Then
'                        �n�_ExitFlg = 0: �I�_ExitFlg = 0
'                        Do
'                            in�\��2 = Mid(aa, 27, 4)
'                            If in�\�� <> in�\��2 Then Stop '���̃f�[�^�ǂݍ���ł��܂����B�s���f�[�^�ɂȂ�
'                            For xx = 0 To 4
'                                in�n�_���itemp = Replace(Mid(aa, 175 + (xx * 20), 10), " ", "")
'                                in�n�_��temp = Mid(aa, 189 + (xx * 20), 4)
'                                If �n�_ExitFlg = 0 Then
'                                    If in�n�_�[�q <> in�n�_���itemp Then in�n�_���i = in�n�_���itemp
'                                    in�n�_�� = in�n�_�� & in�n�_��temp & "/"
'                                    If in�n�_��temp = "0000" Or in�n�_��temp = "    " Then �n�_ExitFlg = 1
'                                End If
'
'                                in�I�_���itemp = Replace(Mid(aa, 275 + (xx * 20), 10), " ", "")
'                                in�I�_��temp = Mid(aa, 289 + (xx * 20), 4)
'                                If �I�_ExitFlg = 0 Then
'                                    If in�I�_�[�q <> in�I�_���itemp Then in�I�_���i = in�I�_���itemp
'                                    in�I�_�� = in�I�_�� & in�I�_��temp & "/"
'                                    If in�I�_��temp = "0000" Or in�I�_��temp = "    " Then �I�_ExitFlg = 1
'                                End If
'
'                                If �n�_ExitFlg = 1 And �I�_ExitFlg = 1 Then Exit Do
'                            Next xx
'                            Line Input #inFNo, aa
'                        Loop
'                        in�n�_�� = Replace(Replace(in�n�_��, "0000/", ""), "    /", "")
'                        in�I�_�� = Replace(Replace(in�I�_��, "0000/", ""), "    /", "")
'                        If Len(in�n�_��) > 1 Then in�n�_�� = Left(in�n�_��, Len(in�n�_��) - 1)
'                        If Len(in�I�_��) > 1 Then in�I�_�� = Left(in�I�_��, Len(in�I�_��) - 1)
'
'                        '�����ꂹ����C������A���������͔�ꂽ���炱��ȏ�l�����񂯂�
'                        'If Len(in�n�_��) <> 4 Then in�n�_�� = ""
'                        'If Len(in�I�_��) <> 4 Then in�I�_�� = ""
'                    End If
                    flg = 0
                    '�V�[�g�������������
                    With Workbooks(myBookName).Sheets(outSheetName)
                        For Y = PVSW����Row + 1 To PVSWlastRow
                            If Left(.Cells(Y, PVSW����Col), 4) = in�\�� Then
                                If .Cells(Y, ���i�i��tc) <> "" Then
                                    '�n�_�I�_�����ւ����m�F
                                    Dim chgFlgA As Long, chgFlgB As Long
                                    If .Cells(Y, PVSW�n�I��Col) = "1" Then
                                        chgFlgA = 1
                                        chgFlgB = 0
                                    Else
                                        chgFlgA = 0
                                        chgFlgB = 1
                                    End If
                                    in�n�_��H = in��H(chgFlgA)
                                    in�n�_�[�� = in�[��(chgFlgA)
                                    in�n�_�[�q = in�[�q(chgFlgA)
                                    in�n�_���i2 = in���i2(chgFlgA)
                                    in�n�_���i3 = in���i3(chgFlgA)
                                    in�n�_���i4 = in���i4(chgFlgA)
                                    in�n�_���i5 = in���i5(chgFlgA)
                                    in�n�_�}���} = in�}���}(chgFlgA)
                                    in�n�_�ڑ��\�� = in�ڑ��\��(chgFlgA)
                                    in�I�_��H = in��H(chgFlgB)
                                    in�I�_�[�� = in�[��(chgFlgB)
                                    in�I�_�[�q = in�[�q(chgFlgB)
                                    in�I�_���i2 = in���i2(chgFlgB)
                                    in�I�_���i3 = in���i3(chgFlgB)
                                    in�I�_���i4 = in���i4(chgFlgB)
                                    in�I�_���i5 = in���i5(chgFlgB)
                                    in�I�_�}���} = in�}���}(chgFlgB)
                                    in�I�_�ڑ��\�� = in�ڑ��\��(chgFlgB)
                                    '�����̔�r
                                    ��r = ""
                                    For X = LBound(PVSW��rCol) To UBound(PVSW��rCol)
                                        ��r = ��r & .Cells(Y, PVSW��rCol(X)) & "_"
                                    Next X
                                    
                                    in��r = in�i�� & "_" & in�T�C�Y & "_" & in�T�C�Y�ď� & "_" & in�F & "_" & in�F�ď� & "_" & in���� & "_" & _
                                             in���� & "_" & inJCDF & "_" & in�d�㐡�@ & "_" & _
                                             in�n�_��H & "_" & in�n�_�[�� & "_" & in�n�_�[�q & "_" & in�n�_�}���} & "_" & in�n�_�ڑ��\�� & "_" & in�n�_���i & "_" & _
                                             in�I�_��H & "_" & in�I�_�[�� & "_" & in�I�_�[�q & "_" & in�I�_�}���} & "_" & in�I�_�ڑ��\�� & "_" & in�I�_���i & "_" & _
                                             in�\�� & "_" & in�T�u0 & "_"
                                                                                                  
                                    If Replace(��r, "_", "") <> "" And ��r <> in��r Then
                                        Debug.Print ��r & vbCrLf & in��r
                                        GoSub �������A���}�b�`�Ȃ̂ōs��ǉ�
                                    End If
                                    
                                    ���i���� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), c)
                                    
                                    '�R�����g��t�����̗�ԍ��̎擾
                                    If unCount(1, �z��ԍ�) = 0 Then
                                        unCount(1, 0) = .Cells(Y, PVSW�n�[��Col2).Column
                                        unCount(1, 1) = .Cells(Y, PVSW�n��HCol2).Column
                                        unCount(1, 2) = .Cells(Y, PVSW�I�[��Col2).Column
                                        unCount(1, 3) = .Cells(Y, PVSW�I��HCol2).Column
                                        unCount(1, 4) = .Cells(Y, PVSW�n���Col).Column
                                        unCount(1, 5) = .Cells(Y, PVSW�I���Col).Column
                                    End If
                                    
                                    �z��ԍ� = 0: Set �Z�� = .Cells(Y, PVSW�n�[��Col2): ��rx = in�n�_�[��
                                    GoSub �Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g
                                    
                                    �z��ԍ� = 1: Set �Z�� = .Cells(Y, PVSW�n��HCol2): ��rx = in�n�_��H
                                    GoSub �Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g
                                    
                                    �z��ԍ� = 2: Set �Z�� = .Cells(Y, PVSW�I�[��Col2): ��rx = in�I�_�[��
                                    GoSub �Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g
                                    
                                    �z��ԍ� = 3: Set �Z�� = .Cells(Y, PVSW�I��HCol2): ��rx = in�I�_��H
                                    GoSub �Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g
                                    
                                    �z��ԍ� = 4: Set �Z�� = .Cells(Y, PVSW�n���Col): ��rx = in�n�_�[�q
                                    If Left(�Z��.Value, 4) = "7009" And Left(��rx, 4) = "7009" Then GoSub �Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g
                                    
                                    �z��ԍ� = 5: Set �Z�� = .Cells(Y, PVSW�I���Col): ��rx = in�I�_�[�q
                                    If Left(�Z��.Value, 4) = "7009" And Left(��rx, 4) = "7009" Then GoSub �Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g
                                    
                                    �ǉ�����s = Y
                                    GoSub �擾�������������
                                    .Cells(�ǉ�����s, ���i�i��tc).Interior.color = RGB(255, 204, 255)
                                    '���葤
                                    '.Cells(y, PVSW�n����Col) = .Cells(y, PVSW�I�[��Col2) & "_" & .Cells(y, PVSW�ICavCol2) & "_" & .Cells(y, PVSW�I��HCol2)
                                    '.Cells(y, PVSW�I����Col) = .Cells(y, PVSW�n�[��Col2) & "_" & .Cells(y, PVSW�nCavCol2) & "_" & .Cells(y, PVSW�n��HCol2)
                                    flg = 1
                                End If
                            End If
                        Next Y
                        '����RLFT�̏�����������Ȃ�����
                        If flg = 0 Then
                            PVSWlastRow = PVSWlastRow + 1
                            �ǉ�����s = PVSWlastRow
                            GoSub �擾�������������
                            .Cells(�ǉ�����s, ���i�i��tc) = "0"
                            .Cells(�ǉ�����s, ���i�i��tc).Interior.color = RGB(255, 204, 255)
                            .Cells(�ǉ�����s, PVSW����Col) = in�\�� & "AA"
                            .Cells(�ǉ�����s, PVSW����Col).Interior.color = RGB(255, 204, 255)
                        End If
                    End With
                End If
            End If
        Loop
        Close #inFNo
line20:
        '�N�����̎擾
        With Workbooks(myBookName).Sheets("���i�i��")
            Dim ���C���i�� As Variant: Set ���C���i�� = .Cells.Find("���C���i��", , , 1)
            Dim seihinRow As Long: seihinRow = .Columns(���C���i��.Column).Find(myKey, , , 1).Row
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("�N����", , , 1).Column).NumberFormat = "yyyy/mm/dd"
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("�N����", , , 1).Column) = in�N����
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��H��", , , 1).Column) = in��H��
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��z", , , 1).Column) = in��z����
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��ϐ�", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��ϐ�", , , 1).Column) = in�}���}��
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��H��AB", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��H��AB", , , 1).Column) = in����c
        End With
        With Workbooks(myBookName).Sheets(outSheetName)
            ����s = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), c)
            .Cells(PVSW����Row - 1, ���i�i��tc) = ����s
            .Cells(PVSW����Row - 2, ���i�i��tc).NumberFormat = "mm/dd"
            .Cells(PVSW����Row - 2, ���i�i��tc).ShrinkToFit = True
            .Cells(PVSW����Row - 2, ���i�i��tc) = in�N����
            .Cells(PVSW����Row - 2, ���i�i��tc).HorizontalAlignment = xlLeft
            .Columns(���i�i��tc).ColumnWidth = Len(����s) * 1.05
        End With
        'Application.StatusBar = c & " / " & ���i�i��RANc
        DoEvents
        Sleep 10
nextC:
    Next c
    '���ޏڍׂ̏���z�z
    'Call ���ޏڍ�_set(�A�h���X(1) & "\���ޏڍ�.txt", "���b�L�敪_", 3, myX)
    
    '���b�L�敪�̔z�z
    Dim strArraySP As Variant
    With Workbooks(myBookName).Sheets(outSheetName)
        For i = PVSW����Row + 1 To PVSWlastRow
            '�n�_���[�q
            �[�q = Replace(.Cells(i, PVSW�n�[Col), "-", "")
            .Cells(i, PVSW�n��Col) = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(�[�q), "���b�L�敪_")
            '�I�_���[�q
            �[�q = Replace(.Cells(i, PVSW�I�[Col), "-", "")
            .Cells(i, PVSW�I��Col) = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(�[�q), "���b�L�敪_")
        Next i
    End With
    
    '�}�g���N�X���`�F�b�N����RLTFtoPVSW��Found�ɂȂ��Ă���̂ɒ��F�������ꍇ�s�𕪂���
    Dim cCel As Object
    With Workbooks(myBookName).Sheets(outSheetName)
        For i = PVSW����Row + 1 To PVSWlastRow
            found = .Cells(i, PVSWRLTFtoPVSWCol)
            If found = "Found" Then
                For X = PVSW���i�i��sCol To PVSW���i�i��eCol
                    Set cCel = .Cells(i, X)
                    If cCel <> "" Then
                        If cCel.Interior.color <> 16764159 Then
                            .Rows(i + 1).Insert
                            .Rows(i).Copy (Rows(i + 1))
                            .Range(.Cells(i + 1, PVSW���i�i��sCol), .Cells(i + 1, PVSW���i�i��eCol)).Interior.Pattern = xlNone
                            For xx = PVSW���i�i��sCol To PVSW���i�i��eCol
                                If .Cells(i, xx).Interior.color = 16764159 Then
                                    .Cells(i + 1, xx) = ""
                                Else
                                    .Cells(i, xx) = ""
                                End If
                            Next xx
                            .Cells(i + 1, PVSWRLTFtoPVSWCol) = "NotFound"
                            .Range(.Cells(i + 1, PVSW�d��sCol + 1), .Cells(i + 1, PVSW�d��eCol)).ClearContents
                            i = i + 1
                            Exit For
                        End If
                    End If
                Next X
            End If
        Next i
    End With
    
    'RLTF�̃T�u0���}�g���N�X�Ɏg�p
    If RLTF�T�u = True Then
        With Workbooks(myBookName).Sheets(outSheetName)
            PVSWlastRow = .Cells(.Rows.count, PVSW����Col).End(xlUp).Row
            For Y = PVSW����Row + 1 To PVSWlastRow
                For X = PVSW���i�i��sCol To PVSW���i�i��eCol
                    Set �Z�� = .Cells(Y, X)
                    in�T�u0 = .Cells(Y, PVSW�T�u0Col)
                    If �Z��.Value <> "" And in�T�u0 <> "" Then
                        �Z��.Value = in�T�u0
                    End If
                Next X
            Next Y
        End With
    End If
    
    Call �œK�����ǂ�
    Application.StatusBar = ""
Exit Sub
    
�������A���}�b�`�Ȃ̂ōs��ǉ�:
    If yyy = 1 Then
        Debug.Print ��r
        Debug.Print in��r
    End If

    With Workbooks(myBookName).Sheets(outSheetName)
        .Rows(Y).Copy
        .Rows(Y).Insert
        Application.CutCopyMode = xlCopy
        PVSWlastRow = PVSWlastRow + 1

        For xxx = PVSW���i�i��sCol To PVSW���i�i��eCol
            If xxx = ���i�i��tc Then
                '.Cells(y, xxx).Interior.Color = RGB(255, 204, 255)
                .Cells(Y, xxx) = ""
                .Cells(Y, xxx).Interior.color = xlNone
            Else
                .Cells(Y + 1, xxx) = ""
                .Cells(Y + 1, xxx).Interior.color = xlNone
            End If
        Next xxx
        For xxx = LBound(PVSW��rCol) To UBound(PVSW��rCol)
            .Cells(Y + 1, PVSW��rCol(xxx)) = ""
        Next xxx
        Y = Y + 1
    End With
Return

�擾�������������:
    With Workbooks(myBookName).Sheets(outSheetName)
        .Cells(�ǉ�����s, PVSW�\��Col) = in�\��
        .Cells(�ǉ�����s, �ڑ�Gcol) = �ڑ�G
        .Cells(�ǉ�����s, PVSWRLTFtoPVSWCol) = "Found"
        .Cells(�ǉ�����s, PVSW�i��Col) = in�i��
        .Cells(�ǉ�����s, PVSW�i���Col) = in�i���
        Call �d���i�팟��(.Cells(�ǉ�����s, PVSW�i��Col), �d���i��)
        .Cells(�ǉ�����s, PVSW�T�C�YCol) = in�T�C�Y
        .Cells(�ǉ�����s, PVSW�T�C�Y�ď�Col) = in�T�C�Y�ď�
        .Cells(�ǉ�����s, PVSW�FCol) = in�F
        .Cells(�ǉ�����s, PVSW�F��Col) = in�F�ď�
        If in��ID = "00" Then in��ID = ""
        .Cells(�ǉ�����s, PVSW��IDcol) = in��ID
        .Cells(�ǉ�����s, PVSW����Col) = in����
        .Cells(�ǉ�����s, PVSW����Col) = in����
        
        If inJCDF = "0000" Then inJCDF = Empty
        .Cells(�ǉ�����s, PVSWJCDFCol) = inJCDF
        .Cells(�ǉ�����s, PVSW�n��HCol) = in�n�_��H
        .Cells(�ǉ�����s, PVSW�n�[��Col) = in�n�_�[��
        .Cells(�ǉ�����s, PVSW�n�[Col) = in�n�_�[�q
        .Cells(�ǉ�����s, PVSW�n��Col) = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(in�n�_�[�q), "�t�@�~���[_")
        .Cells(�ǉ�����s, PVSW�n�}Col) = in�n�_�}���}
        
        If in�n�_�ڑ��\�� = "0000" Then in�n�_�ڑ��\�� = Empty
        .Cells(�ǉ�����s, PVSW�n�ڑ��\��Col) = in�n�_�ڑ��\��
        .Cells(�ǉ�����s, PVSW�n��Col) = in�n�_��
        .Cells(�ǉ�����s, PVSW�n��Col) = in�n�_���i
        .Cells(�ǉ�����s, PVSW�n��2Col) = in�n�_���i2
        .Cells(�ǉ�����s, PVSW�n��3Col) = in�n�_���i3
        .Cells(�ǉ�����s, PVSW�n��4Col) = in�n�_���i4
        .Cells(�ǉ�����s, PVSW�n��5Col) = in�n�_���i5
        .Cells(�ǉ�����s, PVSW�I��HCol) = in�I�_��H
        .Cells(�ǉ�����s, PVSW�I�[��Col) = in�I�_�[��
        .Cells(�ǉ�����s, PVSW�I�[Col) = in�I�_�[�q
        .Cells(�ǉ�����s, PVSW�I��Col) = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(in�I�_�[�q), "�t�@�~���[_")
        .Cells(�ǉ�����s, PVSW�I�}Col) = in�I�_�}���}
        
        If in�I�_�ڑ��\�� = "0000" Then in�I�_�ڑ��\�� = Empty
        .Cells(�ǉ�����s, PVSW�I�ڑ��\��Col) = in�I�_�ڑ��\��
        .Cells(�ǉ�����s, PVSW�I��Col) = in�I�_��
        .Cells(�ǉ�����s, PVSW�I��Col) = in�I�_���i
        .Cells(�ǉ�����s, PVSW�I��2Col) = in�I�_���i2
        .Cells(�ǉ�����s, PVSW�I��3Col) = in�I�_���i3
        .Cells(�ǉ�����s, PVSW�I��4Col) = in�I�_���i4
        .Cells(�ǉ�����s, PVSW�I��5Col) = in�I�_���i5
        .Cells(�ǉ�����s, PVSW�d�㐡�@Col) = in�d�㐡�@
        .Cells(�ǉ�����s, PVSW�T�u0Col) = in�T�u0
    End With
Return

�Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g:
    If CStr(�Z��) <> CStr(��rx) Then
        If �Z��.Comment Is Nothing Then
            Set �R�����g = �Z��.AddComment
            �R�����g.Text ���i���� & "= " & ��rx
            �R�����g.Visible = True
            �R�����g.Shape.Fill.ForeColor.RGB = RGB(255, 204, 255)
            �R�����g.Shape.TextFrame.AutoSize = True
            �R�����g.Shape.TextFrame.Characters.Font.Size = 11
            �R�����g.Shape.Placement = xlMove
            '�R�����g.Shape.PrintObject = True
            unCount(0, �z��ԍ�) = unCount(0, �z��ԍ�) + 1
        Else
            �Z��.Comment.Text �Z��.Comment.Text & vbCrLf & ���i���� & "= " & ��rx
        End If
    End If
Return

End Sub

Sub PVSWcsv��RLTFB�����H�����擾()    '�����BBBBBBBBBBBBBBB

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim i As Long, ii As Long, strArrayS As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Sheets(outSheetName).Activate
    Call �œK��
    
    With Workbooks(myBookName).Sheets(outSheetName)
        Dim PVSW����Row As Long: PVSW����Row = .Cells.Find("�d�����ʖ�", , , 1).Row
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim PVSW���i�i��sCol As Long: PVSW���i�i��sCol = .Cells.Find("���i�i��s", , , 1).Column
        Dim PVSW���i�i��eCol As Long
        On Error Resume Next
        PVSW���i�i��eCol = .Cells.Find("���i�i��e", , , 1).Column
        On Error GoTo 0
        If PVSW���i�i��eCol = 0 Then PVSW���i�i��eCol = PVSW���i�i��sCol
        Dim �^�C�g�� As Range: Set �^�C�g�� = .Rows(PVSW����Row)
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW����Col).End(xlUp).Row
        Dim PVSW�d��sCol As Long: PVSW�d��sCol = .Cells.Find("�d�������擾s", , , 1).Column
        Dim PVSW�d��eCol As Long: PVSW�d��eCol = .Cells.Find("�d�������擾e", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW�\��Col As Long: PVSW�\��Col = .Cells.Find("�\��_", , , 1).Column
        Dim PVSW�i��Col As Long: PVSW�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim PVSW�T�C�YCol As Long: PVSW�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim PVSW�T�C�Y�ď�Col As Long: PVSW�T�C�Y�ď�Col = .Cells.Find("�T��_", , , 1).Column
        Dim PVSW�FCol As Long: PVSW�FCol = .Cells.Find("�F_", , , 1).Column
        Dim PVSW�F��Col As Long: PVSW�F��Col = .Cells.Find("�F��_", , , 1).Column
        Dim PVSW��IDcol As Long: PVSW��IDcol = .Cells.Find("��ID_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        'Dim PVSWG��GNoCol As Long: PVSWG��GNoCol = .Cells.Find("G��GNo_", , , 1).Column
        Dim PVSW�T�u0Col As Long: PVSW�T�u0Col = .Cells.Find("���0_", , , 1).Column
        Dim PVSW�d�㐡�@Col As Long: PVSW�d�㐡�@Col = .Cells.Find("�d�㐡�@_", , , 1).Column
        Dim PVSW�ؒf��Col As Long: PVSW�ؒf��Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim PVSW�n��HCol As Long: PVSW�n��HCol = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n�[��Col As Long: PVSW�n�[��Col = .Cells.Find("�n�_���[��_", , , 1).Column
        Dim PVSW�n�[Col As Long: PVSW�n�[Col = .Cells.Find("�n�_���[�q_", , , 1).Column
        Dim PVSW�n�}Col As Long: PVSW�n�}Col = .Cells.Find("�n�_���}_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_�����i_", , , 1).Column
        Dim PVSW�n��2Col As Long: PVSW�n��2Col = .Cells.Find("�n�_�����i2_", , , 1).Column
        Dim PVSW�I��HCol As Long: PVSW�I��HCol = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I�[��Col As Long: PVSW�I�[��Col = .Cells.Find("�I�_���[��_", , , 1).Column
        Dim PVSW�I�[Col As Long: PVSW�I�[Col = .Cells.Find("�I�_���[�q_", , , 1).Column
        Dim PVSW�I�}Col As Long: PVSW�I�}Col = .Cells.Find("�I�_���}_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_�����i_", , , 1).Column
        Dim PVSW�I��2Col As Long: PVSW�I��2Col = .Cells.Find("�I�_�����i2_", , , 1).Column
'       Dim PVSW�n����Col As Long: PVSW�n����Col = .Cells.Find("�n�_������_", , , 1).Column
'       Dim PVSW�I����Col As Long: PVSW�I����Col = .Cells.Find("�I�_������_", , , 1).Column
        'PVSW�̒l
        Dim PVSW�n��HCol2 As Long: PVSW�n��HCol2 = .Cells.Find("�n�_����H����", , , 1).Column
        Dim PVSW�n�[��Col2 As Long: PVSW�n�[��Col2 = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        Dim PVSW�I��HCol2 As Long: PVSW�I��HCol2 = .Cells.Find("�I�_����H����", , , 1).Column
        Dim PVSW�I�[��Col2 As Long: PVSW�I�[��Col2 = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        
        Dim PVSW�nCavCol2 As Long: PVSW�nCavCol2 = .Cells.Find("�n�_���L���r�e�B", , , 1).Column
        Dim PVSW�ICavCol2 As Long: PVSW�ICavCol2 = .Cells.Find("�I�_���L���r�e�B", , , 1).Column
        '�A���}�b�`�̃J�E���g�p�z��
        Dim unCount(1, 5) As Long
        
        Dim PVSW�n���Col As Long: PVSW�n���Col = .Cells.Find("�n�_���[�����i��", , , 1).Column
        Dim PVSW�I���Col As Long: PVSW�I���Col = .Cells.Find("�I�_���[�����i��", , , 1).Column
        
        '��r����
        Dim PVSW��rCol(0) As Long
        PVSW��rCol(0) = PVSW�ؒf��Col
    End With
    
    Dim in���i�i�� As String, in�\�� As String, in�i�� As String, in�T�C�Y As String, in�T�C�Y�ď� As String, in�F As String, in�F�ď� As String, _
        in���� As String, in���� As String, inJCDF As String, in�n�_�[�q As String, in�n�_�}���} As String, in�I�_�[�q As String, in�I�_�}���} As String, _
        in�\���⑫ As String, in�T�u0 As String, in�n�_���i As String, in�I�_���i As String, in�n�_���i2 As String, in�I�_���i2 As String, in��ID As String
    Dim in�d�㐡�@ As Long, in�ؒf�� As Long, myKey As Variant
    
Dim sTime As Single: sTime = Timer
Debug.Print "s"

    For c = 1 To ���i�i��RANc
        Set myKey = �^�C�g��.Find(���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), c), , , 1)
        If myKey Is Nothing Then GoTo nextC

        ���i�i��tc = myKey.Column
        Dim inTXT As String
        Dim RLTF As String
        RLTF = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "RLTF-B"), c)
        inTXT = ActiveWorkbook.Path & "\06_RLTF_B\" & RLTF & ".txt"
        Dim inFNo As Integer
        inFNo = FreeFile
        If Dir(inTXT) = "" Then GoTo nextC
        
        Open inTXT For Input As #inFNo
        Dim in��H�� As Long, in�}���}�� As Long, in����c As Long
        in�N���� = "": in��H�� = 0: in��z���� = ""
        in�}���}�� = 0: in����c = 0
        Do Until EOF(inFNo)
            Line Input #inFNo, aa
            in���i�i�� = Replace(Mid(aa, 1, 15), " ", "")
            
            in���� = Mid(aa, 88, 1)
            Select Case in����
            Case "A", "G", "N"
                PVSW���i�i�� = Replace(���i�i��RAN(2, c), " ", "")
            Case Else
                PVSW���i�i�� = Replace(���i�i��RAN(1, c), " ", "")
            End Select
            
            If PVSW���i�i�� = in���i�i�� Then
                in�\�� = Mid(aa, 27, 4)
                in�\���⑫ = Mid(aa, 31, 1)
                If Left(in�\��, 1) <> "T" And Left(in�\��, 1) <> "B" And (in�\���⑫ = "0" Or in�\���⑫ = " ") Then
                    
                    in�ؒf�� = CLng(Replace(Mid(aa, 64, 5), " ", ""))
                    in�N���� = "20" & Mid(aa, 482, 2) & "/" & Mid(aa, 484, 2) & "/" & Mid(aa, 486, 2)
                    in��z���� = Mid(aa, 19, 2) & Mid(aa, 23, 1)
                    
                    in�T�C�Y = Mid(aa, 36, 3)
                    If in�T�C�Y <> "   " Then in��H�� = in��H�� + 1
                    in���� = Replace(Mid(aa, 88, 3), " ", "")
                    If in���� = "A" Or in���� = "G" Or in�敪 = "N" Then in����c = in����c + 1
                    in�n�_�}���} = Replace(Mid(aa, 167, 2), " ", "")
                    If in�n�_�}���} <> "" Then in�}���}�� = in�}���}�� + 1
                    in�I�_�}���} = Replace(Mid(aa, 171, 2), " ", "")
                    If in�I�_�}���} <> "" Then in�}���}�� = in�}���}�� + 1

                    flg = 0
                    '�V�[�g�������������
                    With Workbooks(myBookName).Sheets(outSheetName)
                        For Y = PVSW����Row + 1 To PVSWlastRow
                            If Left(.Cells(Y, PVSW����Col), 4) = in�\�� Then
                                If .Cells(Y, ���i�i��tc) <> "" Then
                                    '�����̔�r
                                    ��r = ""
                                    For X = LBound(PVSW��rCol) To UBound(PVSW��rCol)
                                        ��r = ��r & .Cells(Y, PVSW��rCol(X))
                                    Next X
                                    
                                    Dim in��r As String
                                    in��r = in�ؒf��
                                    If Replace(��r, "_", "") <> "" And ��r <> in��r Then
                                        GoSub �������A���}�b�`�Ȃ̂ōs��ǉ�
                                    End If
                                    
                                    �ǉ�����s = Y
                                    GoSub �擾�������������
                                    '.Cells(�ǉ�����s, ���i�i��tc).Interior.color = RGB(255, 204, 255)
                                    '���葤
                                    '.Cells(y, PVSW�n����Col) = .Cells(y, PVSW�I�[��Col2) & "_" & .Cells(y, PVSW�ICavCol2) & "_" & .Cells(y, PVSW�I��HCol2)
                                    '.Cells(y, PVSW�I����Col) = .Cells(y, PVSW�n�[��Col2) & "_" & .Cells(y, PVSW�nCavCol2) & "_" & .Cells(y, PVSW�n��HCol2)
                                    flg = 1
                                End If
                            End If
                        Next Y
                        '����RLFT�̏�����������Ȃ�����
                        If flg = 0 Then
                            Stop '���m�F
                            PVSWlastRow = PVSWlastRow + 1
                            �ǉ�����s = PVSWlastRow
                            GoSub �擾�������������
                            .Cells(�ǉ�����s, ���i�i��tc) = "0"
                            .Cells(�ǉ�����s, ���i�i��tc).Interior.color = RGB(255, 204, 255)
                            .Cells(�ǉ�����s, PVSW����Col) = in�\�� & "AA"
                            .Cells(�ǉ�����s, PVSW����Col).Interior.color = RGB(255, 204, 255)
                        End If
                    End With
                End If
            End If
        Loop
        Close #inFNo
line20:
        '�N�����̎擾
        With Workbooks(myBookName).Sheets("���i�i��")
            Dim ���C���i�� As Variant: Set ���C���i�� = .Cells.Find("���C���i��", , , 1)
            Dim seihinRow As Long: seihinRow = .Columns(���C���i��.Column).Find(myKey, , , 1).Row
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("�N����_", , , 1).Column).NumberFormat = "yyyy/mm/dd"
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("�N����_", , , 1).Column) = in�N����
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��H��_", , , 1).Column) = in��H��
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��z_", , , 1).Column) = in��z����
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��ϐ�_", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��ϐ�_", , , 1).Column) = in�}���}��
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��H��AB_", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(���C���i��.Row).Find("��H��AB_", , , 1).Column) = in����c
        End With
        With Workbooks(myBookName).Sheets(outSheetName)
            ����s = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), c)
            .Cells(PVSW����Row - 1, ���i�i��tc) = ����s
            .Cells(PVSW����Row - 2, ���i�i��tc).NumberFormat = "mm/dd"
            .Cells(PVSW����Row - 2, ���i�i��tc).ShrinkToFit = True
            .Cells(PVSW����Row - 2, ���i�i��tc) = in�N����
            .Cells(PVSW����Row - 2, ���i�i��tc).HorizontalAlignment = xlLeft
            .Columns(���i�i��tc).ColumnWidth = Len(����s) * 1.05
        End With
        'Application.StatusBar = c & " / " & ���i�i��RANc
        DoEvents
        Sleep 10
nextC:
    Next c
    
'    '���ёւ�
'    With Workbooks(myBookName).Sheets(outSheetName)
'        Dim titleRange As Range
'        Set titleRange = .Range(.Cells(PVSW����Row, 1), .Cells(PVSW����Row, .Cells(PVSW����Row, .Columns.Count).End(xlToLeft).Column))
'        Dim r As Variant
'        Dim �D��1 As Long, �D��2 As Long, �D��3 As Long, �D��4 As Long, �D��5 As Long, �D��6 As Long
'        For Each r In titleRange
'            If r = "�d�����ʖ�" Then �D��1 = r.Column '�u�������Ŏg�p���Ă���_�ɒ���
'        Next r
'        lastgyo = .Cells(.Rows.Count, �D��1).End(xlUp).Row
'        With .Sort.SortFields
'            .Clear
'            .add key:=Range(Cells(PVSW����Row, �D��1).Address), Order:=xlAscending, DataOption:=xlSortNormal
'            .add key:=Range(Cells(PVSW����Row, 1).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'        End With
'        .Sort.SetRange Range(Rows(PVSW����Row), Rows(lastgyo))
'        .Sort.Header = xlYes
'        .Sort.MatchCase = False
'        .Sort.Orientation = xlTopToBottom
'        .Sort.Apply
'    End With
    
    Call �œK�����ǂ�
    Application.StatusBar = ""
Exit Sub
    
�������A���}�b�`�Ȃ̂ōs��ǉ�:
    If yyy = 1 Then
        Debug.Print ��r
        Debug.Print in��r
    End If

    With Workbooks(myBookName).Sheets(outSheetName)
        .Rows(Y).Copy
        .Rows(Y).Insert
        Application.CutCopyMode = xlCopy
        PVSWlastRow = PVSWlastRow + 1

        For xxx = PVSW���i�i��sCol To PVSW���i�i��eCol
            If xxx = ���i�i��tc Then
                '.Cells(y, xxx).Interior.Color = RGB(255, 204, 255)
                .Cells(Y, xxx) = ""
                .Cells(Y, xxx).Interior.color = xlNone
            Else
                .Cells(Y + 1, xxx) = ""
                .Cells(Y + 1, xxx).Interior.color = xlNone
            End If
        Next xxx
        For xxx = LBound(PVSW��rCol) To UBound(PVSW��rCol)
            .Cells(Y + 1, PVSW��rCol(xxx)) = ""
        Next xxx
        Y = Y + 1
    End With
Return

�擾�������������:
    With Workbooks(myBookName).Sheets(outSheetName)
        .Cells(�ǉ�����s, PVSW�ؒf��Col) = in�ؒf��
    End With
Return

�Z����RLTF�̔�r���ĈقȂ�Ȃ�R�����g:
    If CStr(�Z��) <> CStr(��rx) Then
        If �Z��.Comment Is Nothing Then
            Set �R�����g = �Z��.AddComment
            �R�����g.Text ���i���� & "= " & ��rx
            �R�����g.Visible = True
            �R�����g.Shape.Fill.ForeColor.RGB = RGB(255, 204, 255)
            �R�����g.Shape.TextFrame.AutoSize = True
            �R�����g.Shape.TextFrame.Characters.Font.Size = 11
            �R�����g.Shape.Placement = xlMove
            '�R�����g.Shape.PrintObject = True
            unCount(0, �z��ԍ�) = unCount(0, �z��ԍ�) + 1
        Else
            �Z��.Comment.Text �Z��.Comment.Text & vbCrLf & ���i���� & "= " & ��rx
        End If
    End If
Return

End Sub

Sub PVSWcsv�Ƀ}�W�b�N�����擾_FromNMB_Ver1918(���i�o��, ���i�_���v)

Dim sTime As Single: sTime = Timer
'Debug.Print "��" & Round(Timer - sTime, 2): sTime = Timer
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    If NMB���� = "" Then NMB���� = "NMB3328_���i�ʉ�H�}�g���N�X.xls"
    'NMB
    Dim nmbBookName As String: nmbBookName = NMB����
    Dim nmbSheetName As String: nmbSheetName = "Sheet1"
    Call NMBset(nmbBookName, nmbSheetName)
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inRow As Long: inRow = .Cells.Find("����_").Row
        Dim inCol_Max As Long: inCol_Max = .UsedRange.Columns.count
        Dim myTitleRange As Range: Set myTitleRange = .Range(.Cells(inRow, 1), .Cells(inRow, inCol_Max))
        Dim inLastRow As Long: inLastRow = .UsedRange.Rows.count
        'Dim out�i��Col As Long: out�i��Col = myTitleRange.Find("�i��_", , , xlWhole).Column
        'Dim out�T�C�YCol As Long: out�T�C�YCol = myTitleRange.Find("�T�C�Y_", , , xlWhole).Column
        'Dim out�T�C�Y��Col As Long: out�T�C�Y��Col = myTitleRange.Find("�T��_", , , xlWhole).Column
        'Dim out�FCol As Long: out�FCol = myTitleRange.Find("�F_", , , xlWhole).Column
        'Dim out�F��Col As Long: out�F��Col = myTitleRange.Find("�F��_", , , xlWhole).Column
        'Dim out����Col As Long: out����Col = myTitleRange.Find("����_", , , xlWhole).Column
        Dim out����col As Long: out����col = myTitleRange.Find("�d�����ʖ�", , , xlWhole).Column
        'Dim out��Col As Long: out��Col = myTitleRange.Find("��H����", , , xlWhole).Column
        'Dim out�[�qCol As Long: out�[�qCol = myTitleRange.Find("�[�q�i��", , , xlWhole).Column
        'Dim out�[��Col As Long: out�[��Col = myTitleRange.Find("�[�����ʎq", , , xlWhole).Column
        Dim outPVSWtoNMB As Long: outPVSWtoNMB = myTitleRange.Find("PVSWtoNMB_", , , xlWhole).Column
        Dim out��Col(1) As Long
        out��Col(0) = myTitleRange.Find("�n�_����H����", , , 1).Column
        out��Col(1) = myTitleRange.Find("�I�_����H����", , , 1).Column
        Dim out�[��Col(1) As Long
        out�[��Col(0) = myTitleRange.Find("�n�_���[�����ʎq", , , 1).Column
        out�[��Col(1) = myTitleRange.Find("�I�_���[�����ʎq", , , 1).Column
        Dim out�[�qCol(1) As Long
        out�[�qCol(0) = myTitleRange.Find("�n�_�[�q_", , , 1).Column
        out�[�qCol(1) = myTitleRange.Find("�I�_�[�q_", , , 1).Column
        Dim out�}Col(1) As Long
        out�}Col(0) = myTitleRange.Find("�n�_�}_", , , 1).Column
        out�}Col(1) = myTitleRange.Find("�I�_�}_", , , 1).Column
        '�V�[���h�p
        Dim out�����i��Col As Long: out�����i��Col = myTitleRange.Find("�����i��", , , 1).Column
        Dim out�}�VCol(1) As Long
        out�}�VCol(0) = myTitleRange.Find("�n�_���}���}�F�P", , , 1).Column
        out�}�VCol(1) = myTitleRange.Find("�I�_���}���}�F�P", , , 1).Column
        Dim out�[�q�VCol(1) As Long
        out�[�q�VCol(0) = myTitleRange.Find("�n�_���[�q�i��", , , 1).Column
        out�[�q�VCol(1) = myTitleRange.Find("�I�_���[�q�i��", , , 1).Column
        
        Dim outABCol As Long: outABCol = myTitleRange.Find("AB_", , , 1).Column
        Dim ���i�i��Col0 As Long: ���i�i��Col0 = 1
        Dim p As Long
        Do
            p = p + 1
            If Len(.Cells(inRow, p)) <> 15 Then Exit Do
        Loop
        Dim ���i�i��Col1 As Long: ���i�i��Col1 = p - 1
        Dim outNMBfeltCol(1) As Long
        outNMBfeltCol(0) = myTitleRange.Find("NMB_Felt0", , , 1).Column
        outNMBfeltCol(1) = myTitleRange.Find("NMB_Felt1", , , 1).Column
        Set myTitleRange = Nothing
    End With
    
    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�i��RAN As Range
        Set ���i�i��RAN = .Range(.Cells(7, 4), .Cells(.Cells(.Rows.count, 3).End(xlUp).Row, 3))
        Dim ���i�g����() As String: ReDim Preserve ���i�g����(1 To ���i�_���v, 2)
        Dim X As Long, i As Long, �g�p�m�Fstr As String: �g�p�m�Fstr = ""
        Dim addX As Long: addX = 0
    End With
        
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbMaxCol As Long: nmbMaxCol = .UsedRange.Columns.count
        Dim nmbTitleRange As Range: Set nmbTitleRange = .Range(.Cells(1, 1), .Cells(1, nmbMaxCol))
        Dim nmbEndRow As Long: nmbEndRow = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim nmb���i�i��Col As Long: nmb���i�i��Col = nmbTitleRange.Find("���i", , , xlWhole).Column
        Dim nmb�\��Col As Long: nmb�\��Col = nmbTitleRange.Find("�\��", , , xlWhole).Column
        Dim nmb�}�W1Col As Long: nmb�}�W1Col = nmbTitleRange.Find("ό�1", , , xlWhole).Column
        Dim nmb�}�W2Col As Long: nmb�}�W2Col = nmbTitleRange.Find("ό�2", , , xlWhole).Column
        Dim nmb��1Col As Long: nmb��1Col = nmbTitleRange.Find("��1", , , xlWhole).Column
        Dim nmb��2Col As Long: nmb��2Col = nmbTitleRange.Find("��2", , , xlWhole).Column
        Dim nmb���i11Col As Long: nmb���i11Col = nmbTitleRange.Find("���i11", , , xlWhole).Column
        Dim nmb���i21Col As Long: nmb���i21Col = nmbTitleRange.Find("���i21", , , xlWhole).Column
        Dim nmb�[��1Col As Long: nmb�[��1Col = nmbTitleRange.Find("�[��1", , , xlWhole).Column
        Dim nmb�[��2Col As Long: nmb�[��2Col = nmbTitleRange.Find("�[��2", , , xlWhole).Column
        Set nmbTitleRange = Nothing
        Dim nmbFelt1 As String
        Dim nmbFelt2 As String
    End With
    Dim z As Long, found As Variant
    Dim ���i�i��use, �����i��, ���i�i��, �\��, ��, �[�q, �[��, AB, PVSWtoNMB As String
    
    For X = 1 To ���i�_���v
        For z = inRow + 1 To inLastRow
            Dim �� As Long: �� = -1
            Dim getFelt As String
            With Workbooks(myBookName).Sheets(mySheetName)
                found = "0"
                ���i�i��use = .Cells(z, X)
                If ���i�i��use = "" Then GoTo line20
                �\�� = Left(.Cells(z, out����col), 4)
                'If �\�� = "0007" Then Stop
                �����i�� = .Cells(z, out�����i��Col).Interior.color
                If �����i�� = 9868950 Then
                    found = "1"
                Else
                    PVSWtoNMB = .Cells(z, outPVSWtoNMB)
                    If PVSWtoNMB = "notFound" Then GoTo line20
                    AB = .Cells(z, outABCol)
                    ���i�i�� = Replace(���i�i��RAN(X, AB), " ", "")
                    If ���i�i�� = "" Then GoTo line20
                    Call NMBseek_�d���[��(���i�i��, �\��, found)
                End If
                If found = "1" Then
                    For xx = 0 To 1
                        getFelt = "0"
                        If �����i�� = 9868950 Then
                            getFelt = .Cells(z, out�}�VCol(xx))
                            get�[�q = .Cells(z, out�[�q�VCol(xx))
                        Else
                            �� = .Cells(z, out��Col(xx))
                            �[�q = .Cells(z, out�[�qCol(xx))
                            �[�� = Format(.Cells(z, out�[��Col(xx)), "000")
                        End If
                        '�񕄕����ŒT��
                        If getFelt = "0" Then
                            If ��1val <> ��2val Then
                                If �� = Replace(��1val, " ", "") Then
                                    getFelt = getFelt1val
                                    get�[�q = ���i11val
                                ElseIf �� = Replace(��2val, " ", "") Then
                                    getFelt = getFelt2val
                                    get�[�q = ���i21val
                                End If
                            End If
                        End If
                        '�[���ŒT��
                        If getFelt = "0" Then
                            If �[��1val <> �[��2val Then
                                If �[�� = �[��1val Then
                                    getFelt = getFelt1val
                                    get�[�q = ���i11val
                                ElseIf �[�� = �[��2val Then
                                    getFelt = getFelt2val
                                    get�[�q = ���i21val
                                End If
                            End If
                        End If
                        '�[�q�ŒT��
                        If getFelt = "0" Then
                            If ���i11val <> ���i21val Then
                                If �[�q = Replace(���i11val, " ", "") Then
                                    getFelt = getFelt1val
                                ElseIf �[�q = Replace(���i21val, " ", "") Then
                                    getFelt = getFelt2val
                                End If
                            End If
                        End If
                        With Workbooks(myBookName).Sheets(mySheetName)
                            If getFelt <> "0" Then
                                '�}���}�o��
                                If .Cells(z, out�}Col(xx)) = Replace(getFelt, " ", "") Or .Cells(z, out�}Col(xx)) = "" Then
                                    .Cells(z, outNMBfeltCol(xx)) = "Found"
                                    .Cells(z, out�}Col(xx)) = Replace(getFelt, " ", "")
                                Else
                                    Debug.Print ���i�i��, �\��, ��, �[�q, �[�� & "=" & .Cells(z, out�}Col(xx)) & "<>" & getFelt
                                    Stop '�}���}��PVSW�̋��ʂƃA���}�b�`
                                End If
                                '�[�q�o��
                                If .Cells(z, out�[�qCol(xx)) = get�[�q Or .Cells(z, out�[�qCol(xx)) = "" Then
                                    .Cells(z, out�[�qCol(xx)).NumberFormat = "@"
                                    .Cells(z, out�[�qCol(xx)) = get�[�q
                                Else
                                    Debug.Print ���i�i��, �\��, ��, �[�� & "=" & �[�q & "<>" & get�[�q
                                    Stop '�[�q��PVSW�̋��ʂƃA���}�b�`
                                End If
                            Else
                                '.Cells(z, out�}Col(xx)) = "Found"
                                .Cells(z, outNMBfeltCol(xx)) = "NotFound"
                                Stop '���̔��f���o���Ȃ�
                            End If
                        End With
                        'Exit For
                    Next xx
                Else
                    '.Cells(z, outFeltCol) = "NotFound"
                    Stop '���i�i�� & �\���Ŕ����o���Ȃ�
                End If
            End With
line20:
        Next z
    Next X
    
    Call NMBrelease
    
'Debug.Print "e " & Round(Timer - sTime, 2): sTime = Timer
End Sub
Sub PVSWcsv���[�Ƀ}�W�b�N�����擾_FromNMB_Ver177(���i�o��, ���i�_���v)

Dim sTime As Single: sTime = Timer
'Debug.Print "��" & Round(Timer - sTime, 2): sTime = Timer
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF���["
    If NMB���� = "" Then NMB���� = "NMB3319_���i�ʉ�H�}�g���N�X.xls"
    'NMB
    Dim nmbBookName As String: nmbBookName = NMB����
    Dim nmbSheetName As String: nmbSheetName = "Sheet1"
    Call NMBset(nmbBookName, nmbSheetName)
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inRow As Long: inRow = Cells.Find("����_").Row
        Dim inCol_Max As Long: inCol_Max = .UsedRange.Columns.count
        Dim myTitleRange As Range: Set myTitleRange = .Range(.Cells(inRow, 1), .Cells(inRow, inCol_Max))
        Dim inLastRow As Long: inLastRow = .UsedRange.Rows.count
        Dim out�i��Col As Long: out�i��Col = myTitleRange.Find("�i��_", , , xlWhole).Column
        Dim out�T�C�YCol As Long: out�T�C�YCol = myTitleRange.Find("�T�C�Y_", , , xlWhole).Column
        Dim out�T�C�Y��Col As Long: out�T�C�Y��Col = myTitleRange.Find("�T��_", , , xlWhole).Column
        Dim out�FCol As Long: out�FCol = myTitleRange.Find("�F_", , , xlWhole).Column
        Dim out�F��Col As Long: out�F��Col = myTitleRange.Find("�F��_", , , xlWhole).Column
        Dim out����Col As Long: out����Col = myTitleRange.Find("����_", , , xlWhole).Column
        Dim out����col As Long: out����col = myTitleRange.Find("�d�����ʖ�", , , xlWhole).Column
        Dim out��Col As Long: out��Col = myTitleRange.Find("��H����", , , xlWhole).Column
        Dim out�[�qCol As Long: out�[�qCol = myTitleRange.Find("�[�q�i��", , , xlWhole).Column
        Dim out�[��Col As Long: out�[��Col = myTitleRange.Find("�[�����ʎq", , , xlWhole).Column
        Dim outPVSWtoNMB As Long: outPVSWtoNMB = myTitleRange.Find("PVSWtoNMB_", , , xlWhole).Column
        Dim outABCol As Long: outABCol = myTitleRange.Find("AB_", , , 1).Column
        Dim ���i�i��Col0 As Long: ���i�i��Col0 = 1
        Dim p As Long
        Do
            p = p + 1
            If Len(.Cells(inRow, p)) <> 15 Then Exit Do
        Loop
        Dim ���i�i��Col1 As Long: ���i�i��Col1 = p - 1
        Dim outFeltCol As Long: outFeltCol = .Cells(1, .Columns.count).End(xlToLeft).Column + 1
        .Cells(1, outFeltCol) = "NMB_Result"
        .Cells(1, outFeltCol + 1) = "NMB_Felt"
        Set myTitleRange = Nothing
    End With
    
    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�i��RAN As Range
        Dim ���i�g����() As String: ReDim Preserve ���i�g����(1 To ���i�_���v, 2)
        Dim X As Long, i As Long, �g�p�m�Fstr As String: �g�p�m�Fstr = ""
        Set ���i�i��RAN = .Range(.Range("d7"), .Range("c" & .Cells(7, 3).End(xlDown).Row))
        Dim addX As Long: addX = 0
        For X = 1 To ���i�_���v
            If ���i�o��(X) = 1 Then
                addX = addX + 1
                ���i�g����(addX, 1) = ���i�i��RAN(X, 1)
                ���i�g����(addX, 2) = ���i�i��RAN(X, 2)
                '�g�p�m�Fstr = �g�p�m�Fstr & .Cells(i, x)
            End If
        Next X
    End With
    

    
'    For x = 1 To ���i�_���v
'        If ���i�o��(x) = 1 Then
'            ���i�g����(x, 1) = ���i�i��Ran(, 1)
'        End If
'    Next x
    
        
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbMaxCol As Long: nmbMaxCol = .UsedRange.Columns.count
        Dim nmbTitleRange As Range: Set nmbTitleRange = .Range(.Cells(1, 1), .Cells(1, nmbMaxCol))
        Dim nmbEndRow As Long: nmbEndRow = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim nmb���i�i��Col As Long: nmb���i�i��Col = nmbTitleRange.Find("���i", , , xlWhole).Column
        Dim nmb�\��Col As Long: nmb�\��Col = nmbTitleRange.Find("�\��", , , xlWhole).Column
        Dim nmb�}�W1Col As Long: nmb�}�W1Col = nmbTitleRange.Find("ό�1", , , xlWhole).Column
        Dim nmb�}�W2Col As Long: nmb�}�W2Col = nmbTitleRange.Find("ό�2", , , xlWhole).Column
        Dim nmb��1Col As Long: nmb��1Col = nmbTitleRange.Find("��1", , , xlWhole).Column
        Dim nmb��2Col As Long: nmb��2Col = nmbTitleRange.Find("��2", , , xlWhole).Column
        Dim nmb���i11Col As Long: nmb���i11Col = nmbTitleRange.Find("���i11", , , xlWhole).Column
        Dim nmb���i21Col As Long: nmb���i21Col = nmbTitleRange.Find("���i21", , , xlWhole).Column
        Dim nmb�[��1Col As Long: nmb�[��1Col = nmbTitleRange.Find("�[��1", , , xlWhole).Column
        Dim nmb�[��2Col As Long: nmb�[��2Col = nmbTitleRange.Find("�[��2", , , xlWhole).Column
        Set nmbTitleRange = Nothing
        Dim nmbFelt1 As String
        Dim nmbFelt2 As String
    End With
    Dim z As Long, found As Variant
    Dim ���i�i��use, ���i�i��, �\��, ��, �[�q, �[��, AB, PVSWtoNMB As String
    
    For X = 1 To addX
        For z = inRow + 1 To inLastRow
            Dim getFelt As String: getFelt = 0
            With Workbooks(myBookName).Sheets("PVSW_RLTF���[")
                ���i�i��use = .Cells(z, X)
                If ���i�i��use = "" Then GoTo line20
                PVSWtoNMB = .Cells(z, outPVSWtoNMB)
                If PVSWtoNMB = "notFound" Then GoTo line20
                AB = .Cells(z, outABCol)
                ���i�i�� = Replace(���i�g����(X, CLng(AB)), " ", "")
                If ���i�i�� = "" Then GoTo line20
                �\�� = Left(.Cells(z, out����col), 4)
                �� = .Cells(z, out��Col)
                �[�q = .Cells(z, out�[�qCol)
                �[�� = Format(.Cells(z, out�[��Col), "000")
                Call NMBseek_�d���[��(���i�i��, Left(�\��, 4), found)
                If found = 1 Then
                    '�����}�W�b�N����
                    If Replace(getFelt1val, " ", "") & Replace(getFelt2val, " ", "") = "" Then
                        getFelt = " "
                    ElseIf getFelt1val = getFelt2val Then
                        getFelt = getFelt1val
                    End If
                    '�񕄕����ŒT��
                    If getFelt = "0" Then
                        If ��1val <> ��2val Then
                            If �� = Replace(��1val, " ", "") Then
                                getFelt = getFelt1val
                            ElseIf �� = Replace(��2val, " ", "") Then
                                getFelt = getFelt2val
                            End If
                        End If
                    End If
                    '�[�q�ŒT��
                    If getFelt = "0" Then
                        If ���i11val <> ���i21val Then
                            If �[�q = Replace(���i11val, " ", "") Then
                                getFelt = getFelt1val
                            ElseIf �[�q = Replace(���i21val, " ", "") Then
                                getFelt = getFelt2val
                            End If
                        End If
                    End If
                    '�[���ŒT��
                    If getFelt = "0" Then
                        If �[��1val <> �[��2val Then
                            If �[�� = �[��1val Then
                                getFelt = getFelt1val
                            ElseIf �[�� = �[��2val Then
                                getFelt = getFelt2val
                            End If
                        End If
                    End If
                    With Workbooks(myBookName).Sheets(mySheetName)
                        If getFelt <> "0" Then
                            If .Cells(z, outFeltCol + 1) = getFelt Or .Cells(z, outFeltCol + 1) = "" Then
                                .Cells(z, outFeltCol) = "Found"
                                .Cells(z, outFeltCol + 1) = getFelt
                            Else
                                Debug.Print ���i�i��, �\��, ��, �[�q, �[��, getFelt
                                Stop '�}�W�b�N�F��PVSW�̋��ʂƃA���}�b�`
                            End If
                        Else
                            .Cells(z, outFeltCol) = "Found"
                            .Cells(z, outFeltCol + 1) = "NotFound"
                            Stop '���̔��f���o���Ȃ�
                        End If
                    End With
                    'Exit For
                Else
                    .Cells(z, outFeltCol) = "NotFound"
                    'Stop '���i�i�� & �\���Ŕ����o���Ȃ�
                End If
            End With
line20:
        Next z
    Next X
    
    Call NMBrelease
    
'Debug.Print "e " & Round(Timer - sTime, 2): sTime = Timer
End Sub


Sub PVSWcsv���[�Ƀ|�C���g�擾()

Dim sTime As Single: sTime = Timer
Debug.Print "��" & Round(Timer - sTime, 2): sTime = Timer
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF���["
    '�|�C���g
    Dim pointSheetName As String: pointSheetName = "�|�C���g�ꗗ"
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inRow As Long: inRow = .Cells.Find("�ؒf��_").Row
        Dim inCol_Max As Long: inCol_Max = .UsedRange.Columns.count
        Dim myTitleRange As Range: Set myTitleRange = .Range(.Cells(inRow, 1), .Cells(inRow, inCol_Max))
        Dim inLastRow As Long: inLastRow = .UsedRange.Rows.count
        
        Dim in�[�����Col As Long: in�[�����Col = myTitleRange.Find("�[�����i��", , , xlWhole).Column
        Dim in�[��Col As Long: in�[��Col = myTitleRange.Find("�[�����ʎq", , , xlWhole).Column
        Dim inCavCol As Long: inCavCol = myTitleRange.Find("�L���r�e�B", , , xlWhole).Column

        Dim outLEDCol As Long: outLEDCol = myTitleRange.Find("LED_", , , xlWhole).Column
        Dim out�|�C���g1Col As Long: out�|�C���g1Col = myTitleRange.Find("�|�C���g1_", , , xlWhole).Column
        Dim out�|�C���g2Col As Long: out�|�C���g2Col = myTitleRange.Find("�|�C���g2_", , , xlWhole).Column
        Dim outFUSEcol As Long: outFUSEcol = myTitleRange.Find("FUSE_", , , xlWhole).Column
        Dim out��d�W�~col As Long: out��d�W�~col = myTitleRange.Find("��d�W�~_", , , xlWhole).Column
        Dim outResultCol As Long: outResultCol = myTitleRange.Find("PVSWtoPOINT_", , , xlWhole).Column
        
        Set myTitleRange = Nothing
    End With
    
    Call POINTset(myBookName, pointSheetName)
    
    Dim i As Long, found As Variant
    Dim �[����� As String, �[�� As String, cav As String
    
        For i = inRow + 1 To inLastRow
            With Workbooks(myBookName).Sheets("PVSW_RLTF���[")
                �[����� = .Cells(i, in�[�����Col)
                �[�� = .Cells(i, in�[��Col)
                cav = .Cells(i, inCavCol)
                Call POINTseek(�[�����, �[��, cav, found)
                If found = 1 Then
                    .Cells(i, outLEDCol) = LEDval
                    .Cells(i, out�|�C���g1Col) = �|�C���g1val
                    .Cells(i, out�|�C���g2Col) = �|�C���g2val
                    .Cells(i, outFUSEcol) = FUSEval
                    .Cells(i, out��d�W�~col) = ��d�W�~val
                    .Cells(i, outResultCol) = "Found"
                Else
                    .Cells(i, outResultCol) = "NotFound"
                End If
            End With
line20:
        Next i
    
    Call POINTrelease
End Sub


Sub NMB�ɒ[�������o��_FromPVSWcsv()
    
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim my�^�C�g��Row As Long: my�^�C�g��Row = .Cells.Find("�i��_").Row
        Dim my�^�C�g��Col As Long: my�^�C�g��Col = .Cells.Find("�i��_").Column
        Dim my�^�C�g��Ran As Range: Set my�^�C�g��Ran = .Range(.Cells(my�^�C�g��Row, 1), .Cells(my�^�C�g��Row, my�^�C�g��Col))
        Dim my�d�����ʖ�Col As Long: my�d�����ʖ�Col = .Cells.Find("�d�����ʖ�").Column
        Dim my�i��Col As Long: my�i��Col = .Cells.Find("�i��_").Column
        Dim my�T�C�YCol As Long: my�T�C�YCol = .Cells.Find("�T�C�Y_").Column
        Dim my�FCol As Long: my�FCol = .Cells.Find("�F_").Column
        Dim my����Col As Long: my����Col = .Cells.Find("����_").Column
        Dim my��1Col As Long: my��1Col = .Cells.Find("�n�_����H����").Column
        Dim my��2Col As Long: my��2Col = .Cells.Find("�I�_����H����").Column
        Dim my�[��1Col As Long: my�[��1Col = .Cells.Find("�n�_���[�����ʎq").Column
        Dim my�[��2Col As Long: my�[��2Col = .Cells.Find("�I�_���[�����ʎq").Column
        Dim my���i11Col As Long: my���i11Col = .Cells.Find("�n�_���[�q�i��").Column
        Dim my���i21Col As Long: my���i21Col = .Cells.Find("�I�_���[�q�i��").Column
        
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row
        
    End With
    
    'NMB
    Dim nmbBookName As String: nmbBookName = "NMB3319_���i�ʉ�H�}�g���N�X.xls"
    Dim nmbSheetName As String: nmbSheetName = "Sheet1"
    
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmb�^�C�g��Ran As Range: Set nmb�^�C�g��Ran = .Range(.Cells(1, 1), .Cells(1, .Cells(1, 1).End(xlToRight).Column))
        Dim nmb���iCol As Long: nmb���iCol = nmb�^�C�g��Ran.Find("���i").Column
        Dim nmb�\��Col As Long: nmb�\��Col = nmb�^�C�g��Ran.Find("�\��").Column
        Dim nmb�i��Col As Long: nmb�i��Col = nmb�^�C�g��Ran.Find("�i��").Column
        Dim nmb�T�C�YCol As Long: nmb�T�C�YCol = nmb�^�C�g��Ran.Find("����").Column
        Dim nmb�FCol As Long: nmb�FCol = nmb�^�C�g��Ran.Find("�F").Column
        Dim nmb��1Col As Long: nmb��1Col = nmb�^�C�g��Ran.Find("��1").Column
        Dim nmb��2Col As Long: nmb��2Col = nmb�^�C�g��Ran.Find("��2").Column
        Dim nmb�[��1Col As Long: nmb�[��1Col = nmb�^�C�g��Ran.Find("�[��1").Column
        Dim nmb�[��2Col As Long: nmb�[��2Col = nmb�^�C�g��Ran.Find("�[��2").Column
        Dim nmb���i11Col As Long: nmb���i11Col = nmb�^�C�g��Ran.Find("���i11").Column
        Dim nmb���i21Col As Long: nmb���i21Col = nmb�^�C�g��Ran.Find("���i21").Column
        
        Dim nmbLastRow As Long: nmbLastRow = .Cells(1, 1).End(xlDown).Row
        Dim nmbResult1Col As Long: nmbResult1Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 1
        Dim nmbResult2Col As Long: nmbResult2Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 2
        Dim nmbResult3Col As Long: nmbResult3Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 3
        Dim nmbResult4Col As Long: nmbResult4Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 4
        Dim nmbResult5Col As Long: nmbResult5Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 5
        Dim nmbResult6Col As Long: nmbResult6Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 6
        Dim nmbResult7Col As Long: nmbResult7Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 7
        .Cells(1, nmbResult2Col) = "�i��"
        .Cells(1, nmbResult3Col) = "�T�C�Y"
        .Cells(1, nmbResult4Col) = "�F"
        .Cells(1, nmbResult5Col) = "��H"
        .Cells(1, nmbResult6Col) = "���i11"
        .Cells(1, nmbResult7Col) = "���i21"
        Set nmb�^�C�g��Ran = Nothing
    End With
    
    Dim findCol As Long
    Dim i As Long
    For i = 2 To nmbLastRow
        With Workbooks(nmbBookName).Sheets(nmbSheetName)
            nmb���i = .Cells(i, nmb���iCol)
            nmb�\�� = .Cells(i, nmb�\��Col)
            nmb�i�� = .Cells(i, nmb�i��Col)
            nmb�T�C�Y = .Cells(i, nmb�T�C�YCol)
            nmb�F = .Cells(i, nmb�FCol)
            nmb��1 = Replace(.Cells(i, nmb��1Col), " ", "")
            nmb��2 = Replace(.Cells(i, nmb��2Col), " ", "")
            nmb�[��1 = .Cells(i, nmb�[��1Col)
            nmb�[��2 = .Cells(i, nmb�[��2Col)
            nmb���i11 = Replace(.Cells(i, nmb���i11Col), " ", "")
            nmb���i21 = Replace(.Cells(i, nmb���i21Col), " ", "")
        End With
        
        findCol = my�^�C�g��Ran.Find(nmb���i).Column
        res�i�� = "": res�T�C�Y = "": res�F = "": res���� = "": res���i11 = "": res���i21 = "": res��H = ""
        For i2 = my�^�C�g��Row + 1 To myLastRow
            With Workbooks(myBookName).Sheets(mySheetName)
                my�l = .Cells(i2, findCol)
                my�\�� = Left(.Cells(i2, my�d�����ʖ�Col), 4)
                my�i�� = .Cells(i2, my�i��Col)
                my�T�C�Y = .Cells(i2, my�T�C�YCol)
                my�F = .Cells(i2, my�FCol)
                my���� = .Cells(i2, my����Col)
                my��1 = .Cells(i2, my��1Col)
                my��2 = .Cells(i2, my��2Col)
                my�[��1 = .Cells(i2, my�[��1Col)
                my�[��2 = .Cells(i2, my�[��2Col)
                my���i11 = Replace(.Cells(i2, my���i11Col), " ", "")
                my���i21 = Replace(.Cells(i2, my���i21Col), " ", "")
            End With
            If my�l = 1 Then
                If my�\�� = nmb�\�� Then
                    '���ʏ���
                    If my�i�� <> nmb�i�� Then res�i�� = my�i��
                    If my�T�C�Y <> nmb�T�C�Y Then res�T�C�Y = my�T�C�Y
                    If my�F <> nmb�F Then res�F = my�F
                    If my���� <> nmb���� Then res���� = my����
                    '1��=1��
                    If my��1 = nmb��1 And my��2 = nmb��2 Then
                        If my���i11 <> nmb���i11 Then res���i11 = my���i11
                        If my���i21 <> nmb���i21 Then res���i21 = my���i21
                    '1��=2��
                    ElseIf my��1 = nmb��2 And my��2 = nmb��1 Then
                        If my���i11 <> nmb���i21 Then res���i21 = my���i11
                        If my���i21 <> nmb���i11 Then res���i11 = my���i21
                    '�[��and�񕄂̑g������������Ȃ�
                    Else
                        res��H = "notFound"
                    End If
                    GoSub result
                    Exit For
                End If
            End If
        Next i2
    Next i
    
Exit Sub
result:
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        .Cells(i, nmbResult1Col) = "Found"
        .Cells(i, nmbResult2Col) = res�i��
        .Cells(i, nmbResult3Col) = res�T�C�Y
        .Cells(i, nmbResult4Col) = res�F
        .Cells(i, nmbResult5Col) = res��H
        .Cells(i, nmbResult6Col) = res���i11
        .Cells(i, nmbResult7Col) = res���i21
    End With
    Return

End Sub

Sub PVSWcsv���[�̃V�[�g�쐬_Ver1932(���i�o��, ���i�_���v, ��n�����i�i��)
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "PVSW_RLTF���["
    
    With Workbooks(myBookName).Sheets(mySheetName)
        'PVSW_RLTF����̃f�[�^
        Dim my�^�C�g��Row As Long: my�^�C�g��Row = .Cells.Find("�i��_").Row
        Dim my�^�C�g��Col As Long: my�^�C�g��Col = .Cells.Find("�i��_").Column
        Dim my�^�C�g��Ran As Range: Set my�^�C�g��Ran = .Range(.Cells(my�^�C�g��Row, 1), .Cells(my�^�C�g��Row, my�^�C�g��Col))
        Dim my�d�����ʖ�Col As Long: my�d�����ʖ�Col = .Cells.Find("�d�����ʖ�").Column
        Dim my��1Col As Long: my��1Col = .Cells.Find("�n�_����H����").Column
        Dim my�[��1Col As Long: my�[��1Col = .Cells.Find("�n�_���[�����ʎq").Column
        Dim myCav1Col As Long: myCav1Col = .Cells.Find("�n�_���L���r�e�BNo.").Column
        Dim my��2Col As Long: my��2Col = .Cells.Find("�I�_����H����").Column
        Dim my�[��2Col As Long: my�[��2Col = .Cells.Find("�I�_���[�����ʎq").Column
        Dim myCav2Col As Long: myCav2Col = .Cells.Find("�I�_���L���r�e�BNo.").Column
'        Dim my����Col As Long: my����Col = .Cells.Find("����No").Column
'        Dim my�����i��Col As Long: my�����i��Col = .Cells.Find("�����i��").Column
'        Dim myJoint1Col As Long: myJoint1Col = .Cells.Find("�n�_��JOINT���").Column
'        Dim myJoint2Col As Long: myJoint2Col = .Cells.Find("�I�_��JOINT���").Column
        Dim my�_�u����1Col As Long: my�_�u����1Col = .Cells.Find("�n�_���_�u����H����").Column
        Dim my�_�u����2Col As Long: my�_�u����2Col = .Cells.Find("�I�_���_�u����H����").Column
        
'        Dim myPVSW�i��col As Long: myPVSW�i��col = .Cells.Find("�d���i��").Column
'        Dim myPVSW�T�C�Ycol As Long: myPVSW�T�C�Ycol = .Cells.Find("�d���T�C�Y").Column
'        Dim myPVSW�Fcol As Long: myPVSW�Fcol = .Cells.Find("�d���F").Column
'        Dim my�}���}11Col As Long: my�}���}11Col = .Cells.Find("�n�_���}���}�F�P").Column
'        Dim my�}���}12Col As Long: my�}���}12Col = .Cells.Find("�n�_���}���}�F�Q").Column
'        Dim my�}���}21Col As Long: my�}���}21Col = .Cells.Find("�I�_���}���}�F�P").Column
'        Dim my�}���}22Col As Long: my�}���}22Col = .Cells.Find("�I�_���}���}�F�Q").Column
'        Dim my���i11Col As Long: my���i11Col = .Cells.Find("�n�_���[�q�i��").Column
'        Dim my���i21Col As Long: my���i21Col = .Cells.Find("�I�_���[�q�i��").Column
'        Dim my���i12Col As Long: my���i12Col = .Cells.Find("�n�_���S����i��").Column
'        Dim my���i22Col As Long: my���i22Col = .Cells.Find("�I�_���S����i��").Column
        Dim my���1Col As Long: my���1Col = .Cells.Find("�n�_����햼��").Column
        Dim my���2Col As Long: my���2Col = .Cells.Find("�I�_����햼��").Column
        Dim my���Ӑ�1Col As Long: my���Ӑ�1Col = .Cells.Find("�n�_���[�����Ӑ�i��").Column
        Dim my���1Col As Long: my���1Col = .Cells.Find("�n�_���[�����i��").Column
        Dim my���Ӑ�2Col As Long: my���Ӑ�2Col = .Cells.Find("�I�_���[�����Ӑ�i��").Column
        Dim my���2Col As Long: my���2Col = .Cells.Find("�I�_���[�����i��").Column
'        Dim myJointGCol As Long: myJointGCol = .Cells.Find("�W���C���g�O���[�v").Column
'        Dim myAB�敪Col As Long: myAB�敪Col = .Cells.Find("A/B�EB/C�敪").Column
'        Dim my�d��YBMCol As Long: my�d��YBMCol = .Cells.Find("�d���x�a�l").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(my�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Set my�^�C�g��Ran = Nothing
        'RLTF����̃f�[�^
        Dim my�i��Col As Long: my�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim my�T�C�YCol As Long: my�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim my�T�C�Y��Col As Long: my�T�C�Y��Col = .Cells.Find("�T��_", , , 1).Column
        Dim my�FCol As Long: my�FCol = .Cells.Find("�F_", , , 1).Column
        Dim my�F��Col As Long: my�F��Col = .Cells.Find("�F��_", , , 1).Column
        Dim my����Col As Long: my����Col = .Cells.Find("����", , , 1).Column
        Dim my����Col As Long: my����Col = .Cells.Find("����", , , 1).Column
        Dim myJCDFcol As Long: myJCDFcol = .Cells.Find("JCDF_", , , 1).Column
        Dim my�n�[Col As Long: my�n�[Col = .Cells.Find("�n�_�[�q_", , , 1).Column
        Dim my�n�}Col As Long: my�n�}Col = .Cells.Find("�n�_�}_", , , 1).Column
        Dim my�I�[Col As Long: my�I�[Col = .Cells.Find("�I�_�[�q_", , , 1).Column
        Dim my�I�}Col As Long: my�I�}Col = .Cells.Find("�I�_�}", , , 1).Column
        Dim my����Col As Long: my����Col = .Cells.Find("����_", , , 1).Column
        Dim myRLTFtoPVSW As Long: myRLTFtoPVSW = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        '�T�u�}�f�[�^_Ver181�̒ǉ��f�[�^
        Dim my�T�uCol As Long: my�T�uCol = .Cells.Find("�T�u", , , 1).Column
        Dim my�n�_�n��Col As Long: my�n�_�n��Col = .Cells.Find("�n�_�n��", , , 1).Column
        Dim my�I�_�n��Col As Long: my�I�_�n��Col = .Cells.Find("�I�_�n��", , , 1).Column
        Dim myABcol As Long: myABcol = .Cells.Find("AB_", , , 1).Column
                
    End With
    
    '���[�N�V�[�g�̒ǉ�
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    
    'PVSW_RLTF to PVSW_RLTF���[
    Dim i As Long, ���i�i��RAN As Variant
    For i = my�^�C�g��Row To myLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            'Set ���i�i��Ran = .Range(.Cells(i, my���i�i��Ran0), .Cells(i, my���i�i��Ran1))
            'If Application.Sum(.Range(.Cells(i, my���i�i��Ran0), .Cells(i, my���i�i��Ran1))) = 0 Then GoTo line20
            
            
            Dim ���i�g����() As String: ReDim Preserve ���i�g����(1 To ���i�_���v, 2)
            Dim X As Long, �g�p�m�Fstr As String: �g�p�m�Fstr = ""
            For X = 1 To ���i�_���v
                If ���i�o��(X) = 1 Then
                    ���i�g����(X, 1) = .Cells(i, X)
                    �g�p�m�Fstr = �g�p�m�Fstr & .Cells(i, X)
                End If
            Next X
            If �g�p�m�Fstr = "" Then GoTo line20
            If i = my�^�C�g��Row Then GoTo line10
            
            Dim �d�����ʖ� As String: �d�����ʖ� = .Cells(i, my�d�����ʖ�Col)
            Dim ��1 As String: ��1 = .Cells(i, my��1Col)
            Dim �[��1 As String: �[��1 = .Cells(i, my�[��1Col)
            Dim Cav1 As String: Cav1 = .Cells(i, myCav1Col)
            Dim ��2 As String: ��2 = .Cells(i, my��2Col)
            Dim �[��2 As String: �[��2 = .Cells(i, my�[��2Col)
            Dim cav2 As String: cav2 = .Cells(i, myCav2Col)
'            Dim ���� As String: ���� = .Cells(i, my����Col)
'            Dim �����i�� As Range: Set �����i�� = .Cells(i, my�����i��Col)
'            Dim �V�[���h�t���O As String: If �����i��.Interior.Color = 9868950 Then �V�[���h�t���O = "S" Else �V�[���h�t���O = ""
'            Dim Joint1 As String: Joint1 = .Cells(i, myJoint1Col)
'            Dim Joint2 As String: Joint2 = .Cells(i, myJoint2Col)
            Dim �_�u����1 As String: �_�u����1 = .Cells(i, my�_�u����1Col)
            Dim �_�u����2 As String: �_�u����2 = .Cells(i, my�_�u����2Col)
'            Dim ���i11 As String: ���i11 = .Cells(i, my���i11Col)
'            Dim ���i21 As String: ���i21 = .Cells(i, my���i21Col)
'            Dim ���i12 As String: ���i12 = .Cells(i, my���i12Col)
'            Dim ���i22 As String: ���i22 = .Cells(i, my���i22Col)
            Dim ���1 As String: ���1 = .Cells(i, my���1Col)
            Dim ���2 As String: ���2 = .Cells(i, my���2Col)
            Dim ���Ӑ�1 As String: ���Ӑ�1 = .Cells(i, my���Ӑ�1Col)
            Dim ���1 As String: ���1 = .Cells(i, my���1Col)
            Dim ���Ӑ�2 As String: ���Ӑ�2 = .Cells(i, my���Ӑ�2Col)
            Dim ���2 As String: ���2 = .Cells(i, my���2Col)
'            Dim JointG As String: JointG = .Cells(i, myJointGCol)
'            Dim �d���i�� As String: �d���i�� = .Cells(i, myPVSW�i��col)
'            Dim �d���T�C�Y As String: �d���T�C�Y = .Cells(i, myPVSW�T�C�Ycol)
'            Dim �d���F As String: �d���F = .Cells(i, myPVSW�Fcol)
'            Dim �}���}11 As String: �}���}11 = .Cells(i, my�}���}11Col)
'            Dim �}���}12 As String: �}���}12 = .Cells(i, my�}���}12Col)
'            Dim �}���}21 As String: �}���}21 = .Cells(i, my�}���}21Col)
'            Dim �}���}22 As String: �}���}22 = .Cells(i, my�}���}22Col)
'            Dim AB�敪 As String: AB�敪 = .Cells(i, myAB�敪Col)
'            Dim �d��YBM As String: �d��YBM = .Cells(i, my�d��YBMCol)
            
            Dim ���葤1 As String, ���葤2 As String
            If Len(cav2) < 4 Then ���葤1 = �[��2 & "_" & String(3 - Len(cav2), " ") & cav2 & "_" & ��2
            If Len(Cav1) < 4 Then ���葤2 = �[��1 & "_" & String(3 - Len(Cav1), " ") & Cav1 & "_" & ��1
            'RLTF����̃f�[�^
            Dim �i�� As String: �i�� = .Cells(i, my�i��Col)
            Dim �T�C�Y As String: �T�C�Y = .Cells(i, my�T�C�YCol)
            Dim �T�C�Y�� As String: �T�C�Y�� = .Cells(i, my�T�C�Y��Col)
            Dim �F As String: �F = .Cells(i, my�FCol)
            Dim �F�� As String: �F�� = .Cells(i, my�F��Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim RLTFtoPVSW As String: RLTFtoPVSW = .Cells(i, myRLTFtoPVSW)
            '�T�u�}�f�[�^_Ver181�̒ǉ��f�[�^
            Dim �T�u As String: �T�u = .Cells(i, my�T�uCol)
            Dim �n��1 As String: �n��1 = .Cells(i, my�n�_�n��Col)
            Dim �n��2 As String: �n��2 = .Cells(i, my�I�_�n��Col)
            Dim AB As String: AB = .Cells(i, myABcol)
        End With
line10:
        With Workbooks(myBookName).Sheets(newSheetName)
        Dim �D��1 As Long, �D��2 As Long, �D��3 As Long, addCol As Long
        Dim addRow As Long: addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row + 1
            If .Cells(1, 1) = "" Then
                For X = 1 To ���i�_���v
                    If ���i�o��(X) = 1 Then
                        addCol = addCol + 1
                        .Cells(1, addCol).NumberFormat = "@"
                        .Cells(1, addCol) = ���i�g����(X, 1)
                        ���i�g����(X, 2) = addCol
                    End If
                Next X
                    
                .Cells(1, addCol + 1) = "�d�����ʖ�"
                .Cells(1, addCol + 2) = "��H����"
                .Cells(1, addCol + 3) = "�[�����ʎq": �D��1 = addCol + 3
                .Cells(1, addCol + 4) = "�L���r�e�BNo.": �D��3 = addCol + 4
                '.Cells(1, addCol + 5) = "����No"
                '.Cells(1, addCol + 6) = "�����i��"
'                .Cells(1, addCol + 7) = "Joint���"
                .Cells(1, addCol + 5) = "�_�u����H����"
'                .Cells(1, addCol + 9) = "�[�q�i��"
'                .Cells(1, addCol + 10) = "�S����i��"
                .Cells(1, addCol + 6) = "��햼��"
                .Cells(1, addCol + 7) = "�[�����Ӑ�i��"
                .Cells(1, addCol + 8) = "�[�����i��": �D��2 = addCol + 13
'                .Cells(1, addCol + 14) = "�W���C���g�O���[�v"
                
'                .Cells(1, addCol + 15) = "�d���i��": Columns(addCol + 15).NumberFormat = "@"
'                .Cells(1, addCol + 16) = "�d���T�C�Y": Columns(addCol + 16).NumberFormat = "@"
'                .Cells(1, addCol + 17) = "�d���F": Columns(addCol + 17).NumberFormat = "@"
'                .Cells(1, addCol + 18) = "�}���}�F�P": Columns(addCol + 18).NumberFormat = "@"
'                .Cells(1, addCol + 19) = "�}���}�F�Q": Columns(addCol + 19).NumberFormat = "@"

'                .Cells(1, addCol + 20) = "A/B�EB/C�敪"
'                .Cells(1, addCol + 21) = "�d���x�a�l"
                .Cells(1, addCol + 9) = "RLTFtoPVSW_"
                .Cells(1, addCol + 10) = "�i��_"
                .Cells(1, addCol + 11) = "�T�C�Y_"
                .Cells(1, addCol + 12) = "�T��_"
                .Cells(1, addCol + 13) = "�F_"
                .Cells(1, addCol + 14) = "�F��_"
                .Cells(1, addCol + 15) = "����_"
                .Cells(1, addCol + 16) = "����_"
                .Cells(1, addCol + 17) = "JCDF_"
                .Cells(1, addCol + 18) = "�[�q_"
                .Cells(1, addCol + 19) = "�}_"
                .Cells(1, addCol + 20) = "����_"
                .Cells(1, addCol + 21) = "���葤"
                .Cells(1, addCol + 22) = "��_"
                .Cells(1, addCol + 23) = "LED_"
                .Cells(1, addCol + 24) = "�|�C���g1_"
                .Cells(1, addCol + 25) = "�|�C���g2_"
                .Cells(1, addCol + 26) = "FUSE_"
                .Cells(1, addCol + 27) = "�R�����g_"
                .Cells(1, addCol + 28) = "PVSWtoPOINT_"
                .Cells(1, addCol + 29) = "�T�u"
                .Cells(1, addCol + 30) = "�n����"
                .Cells(1, addCol + 31) = "AB_"
                .Range(.Columns(1), .Columns(31)).NumberFormat = "@"
                .Columns(addCol + 20).NumberFormat = 0
            Else
                '.Range(.Cells(addRow, 1), .Cells(addRow + 1, addCol)) = ���i�i��Ran.Value
                For X = 1 To ���i�_���v
                    If ���i�o��(X) = 1 Then
                    If ���i�_���v <> 1 Then
                        .Range(.Cells(addRow, CLng(���i�g����(X, 2))), .Cells(addRow + 1, CLng(���i�g����(X, 2)))) = ���i�g����(X, 1)
                    Else
                        .Range(.Cells(addRow, CLng(���i�g����(X, 2))), .Cells(addRow + 1, CLng(���i�g����(X, 2)))) = ���i�g����(X, 1)
                    End If
                    End If
                Next X
                .Range(.Cells(addRow, addCol + 1), .Cells(addRow + 1, addCol + 1)) = �d�����ʖ�
                .Range(.Cells(addRow, addCol + 5), .Cells(addRow + 1, addCol + 5)) = ����
                .Range(.Cells(addRow, addCol + 6), .Cells(addRow + 1, addCol + 6)).Value = �����i��.Value
                .Range(.Cells(addRow, addCol + 6), .Cells(addRow + 1, addCol + 6)).Interior.color = �����i��.Interior.color
                .Range(.Cells(addRow, addCol + 14), .Cells(addRow + 1, addCol + 14)) = JointG
                .Range(.Cells(addRow, addCol + 20), .Cells(addRow + 1, addCol + 20)) = AB�敪
                .Range(.Cells(addRow, addCol + 21), .Cells(addRow + 1, addCol + 21)) = �d��YBM
                .Range(.Cells(addRow, addCol + 22), .Cells(addRow + 1, addCol + 22)) = �i��
                .Range(.Cells(addRow, addCol + 23), .Cells(addRow + 1, addCol + 23)) = �T�C�Y
                .Range(.Cells(addRow, addCol + 24), .Cells(addRow + 1, addCol + 24)) = �T�C�Y��
                .Range(.Cells(addRow, addCol + 25), .Cells(addRow + 1, addCol + 25)) = �F
                .Range(.Cells(addRow, addCol + 26), .Cells(addRow + 1, addCol + 26)) = �F��
                .Range(.Cells(addRow, addCol + 27), .Cells(addRow + 1, addCol + 27)) = ����
                .Range(.Cells(addRow, addCol + 28), .Cells(addRow + 1, addCol + 28)) = PVSWtoNMB
                .Range(.Cells(addRow, addCol + 30), .Cells(addRow + 1, addCol + 30)) = �V�[���h�t���O
                .Range(.Cells(addRow, addCol + 38), .Cells(addRow + 1, addCol + 38)) = �T�u
                .Range(.Cells(addRow, addCol + 38), .Cells(addRow + 1, addCol + 40)) = AB
                .Cells(addRow, addCol + 2) = ��1
                .Cells(addRow + 1, addCol + 2) = ��2
                .Cells(addRow, addCol + 3) = �[��1
                .Cells(addRow + 1, addCol + 3) = �[��2
                .Cells(addRow, addCol + 4) = Cav1
                .Cells(addRow + 1, addCol + 4) = cav2
                .Cells(addRow, addCol + 14) = Joint1
                .Cells(addRow + 1, addCol + 14) = Joint2
                .Cells(addRow, addCol + 8) = �_�u����1
                .Cells(addRow + 1, addCol + 8) = �_�u����2
                .Cells(addRow, addCol + 9) = ���i11
                .Cells(addRow + 1, addCol + 9) = ���i21
                .Cells(addRow, addCol + 10) = ���i12
                .Cells(addRow + 1, addCol + 10) = ���i22
                .Cells(addRow, addCol + 11) = ���1
                .Cells(addRow + 1, addCol + 11) = ���2
                .Cells(addRow, addCol + 12) = ���Ӑ�1
                .Cells(addRow + 1, addCol + 12) = ���Ӑ�2
                .Cells(addRow, addCol + 13) = ���1
                .Cells(addRow + 1, addCol + 13) = ���2
                .Cells(addRow, addCol + 15) = �d���i��
                .Cells(addRow + 1, addCol + 15) = �d���i��
                .Cells(addRow, addCol + 16) = �d���T�C�Y
                .Cells(addRow + 1, addCol + 16) = �d���T�C�Y
                .Cells(addRow, addCol + 17) = �d���F
                .Cells(addRow + 1, addCol + 17) = �d���F
                .Cells(addRow, addCol + 18) = �}���}11
                .Cells(addRow + 1, addCol + 18) = �}���}21
                .Cells(addRow, addCol + 19) = �}���}12
                .Cells(addRow + 1, addCol + 19) = �}���}22
                .Cells(addRow, addCol + 29) = ���葤1
                .Cells(addRow + 1, addCol + 29) = ���葤2
                .Cells(addRow, addCol + 31) = "�n"
                .Cells(addRow + 1, addCol + 31) = "�I"
                .Cells(addRow, addCol + 39) = �n��1
                .Cells(addRow + 1, addCol + 39) = �n��2
            End If
        End With
line20:
    Next i
    '���בւ�
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            '.Add Key:=Range(Cells(1, �D��2).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow + 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
End Sub

Sub �n���}�쐬_Ver2001(�I��, �O���[�v���, �O���[�v��)
    Call Init2
    
    Dim sTime As Single: sTime = Timer
    Debug.Print "0= " & Round(Timer - sTime, 2)

    Call �œK��
    
    If �O���[�v��� = "���C���i��" Then ��n�����i�i�� = �O���[�v�� Else ��n�����i�i�� = ""  '�w�肵���琻�i�g�������쐬���Ȃ�_���̐��i�i�Ԃ̒l���g�p���Ă��Ȃ�
    
    Dim ����� As String: ����� = �O���[�v���
    Dim ����G As String: ����G = �O���[�v��
    Dim p As Long
    �I��s = Split(�I��, ",")
    
    Dim step0T As Long, step0 As Long
    
    ProgressBar.Show vbModeless

    '�n���}�^�C�v = "�\��" '0:�쐬���Ȃ� or ��� or �`�F�b�J�[�p or ��H���� or �\�� or ����[��
    Select Case �I��s(0)
    Case "0"
    �n���}�^�C�v = "0"
    Case "1"
    �n���}�^�C�v = "���"
    Case "2"
    �n���}�^�C�v = "�`�F�b�J�[�p"
    Case "3"
    �n���}�^�C�v = "��H����"
    Case "4"
    �n���}�^�C�v = "�\��"
    Case "5"
    �n���}�^�C�v = "����[��"
    End Select
    
    '�v���O���X�o�[��STEP��
    If �n���}�^�C�v = "�`�F�b�J�[�p" Then
        step0T = 10
    Else
        step0T = 9
    End If
        
    '�n���\�� = "1" '0:�����A1:��n���}�A2:��n���}(��n���͏������j�A3:��n���}(��n���̓p�^�[��)�A4:��n���͕\�����Ȃ�
    �n���\�� = �I��s(1)
    
    '�������i = "0" '0:�\�����Ȃ��A40:��n�����i�A50:��n�����i
    Select Case �I��s(2)
    Case "0"
    �������i = "0"
    Case "1"
    �������i = "40"
    End Select
    
    Dim ��ƕ\���ϊ� As String ': ��ƕ\���ϊ� = "1" '0:�ϊ����Ȃ��A1:�T�C�Y����ƕ\���L���ɕϊ�����
    ��ƕ\���ϊ� = �I��s(3)
    
    'MAX��H�\�� = "0"
    MAX��H�\�� = �I��s(4)
    
    '�n����ƕ\��
    With wb(0).Sheets("�ݒ�")
        Dim myKey As Variant
        Set myKey = .Cells.Find("�n���F_", , , 1)
        If CLng(�I��s(5)) <> -1 Then
            �n����ƕ\�� = myKey.Offset(CLng(�I��s(5)), 1)
        End If
    End With
    
    myFont = "�l�r �S�V�b�N"
    Dim minW�w�� As Long
    Select Case �n���}�^�C�v
    Case "�`�F�b�J�[�p"
        minW�w�� = 24 '24
    Case "��H����", "�\��", "����[��"
        minW�w�� = 28
    Case Else
        minW�w�� = 18 '18
    End Select
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF���["
    Dim newSheetName As String: newSheetName = "�n���}_" & �O���[�v��� & "_" & Replace(�O���[�v��, " ", "")
    
    'PVSW_RLTF����[�������擾
    With wb(0).Sheets("�ݒ�")
        Dim i As Long
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
    'Call ���i�i��RAN_set2(���i�i��RAN, ����G, �����, ��n�����i�i��)
    
    step0 = step0 + 1
    Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF]��[�[���ꗗ]���Q�Ƃ��ăn�����̏W�v��", step0T, step0, 100, 100)
    Call PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬_2017
    
    Dim ws As Worksheet
    
    Debug.Print "1= " & Round(Timer - sTime, 2): sTime = Timer
    
    step0 = step0 + 1
    'Sheet���i���X�g�̃f�[�^���Z�b�g
    If �������i <> 0 Then
        With wb(0).Sheets("���i���X�g")
            Dim ���i���X�gkey As Range: Set ���i���X�gkey = .Cells.Find("���i�i��", , , 1)
            Dim ���i���X�gtitle As Range: Set ���i���X�gtitle = .Rows(���i���X�gkey.Row)
            Dim ���i���X�g���i�i��Col As Long: ���i���X�g���i�i��Col = ���i���X�gkey.Column
            Dim ���i���X�g�\��Col As Long: ���i���X�g�\��Col = ���i���X�gtitle.Find("�\����", , , 1).Column
            Dim ���i���X�g���i�i��Col As Long: ���i���X�g���i�i��Col = ���i���X�gtitle.Find("���i�i��", , , 1).Column
            Dim ���i���X�g�H��Col As Long: ���i���X�g�H��Col = ���i���X�gtitle.Find("�H��a", , , 1).Column
            Dim ���i���X�g���Col As Long: ���i���X�g���Col = ���i���X�gtitle.Find("���", , , 1).Column
            Dim ���i���X�g�[��Col As Long
            ���i���X�g�[��Col = ���i���X�gtitle.Find(��n�����i�i��, , , 1).Column
            Dim ���i���X�gT�ď�Col As Long: ���i���X�gT�ď�Col = ���i���X�gtitle.Find("���ޏڍ�", , , 1).Column
            Dim ���i���X�glastRow As Long: ���i���X�glastRow = .Cells(.Rows.count, ���i���X�gkey.Column).End(xlUp).Row
            Dim ���i���X�g() As String: ���i���X�gc = 0
            Dim ���i���X�gs() As String
            For i = ���i���X�gkey.Row To ���i���X�glastRow
                'If Replace(��n�����i�i��, " ", "") = Replace(.Cells(i, ���i���X�g���i�i��Col), " ", "") Then
                If .Cells(i, ���i���X�g�[��Col) <> "" Then
                    ReDim Preserve ���i���X�g(5, ���i���X�gc)
                    ReDim Preserve ���i���X�gs(���i�i��RANc, ���i���X�gc)
                    ���i���X�g(0, ���i���X�gc) = .Cells(i, ���i���X�g�\��Col)
                    ���i���X�g(1, ���i���X�gc) = .Cells(i, ���i���X�g���i�i��Col)
                    ���i���X�g(2, ���i���X�gc) = .Cells(i, ���i���X�g�H��Col)
                    ���i���X�g(3, ���i���X�gc) = .Cells(i, ���i���X�g���Col)
                    If ���i���X�g(2, ���i���X�gc) = "40" And ���i���X�g(3, ���i���X�gc) = "T" Then
                        ���i���X�g(4, ���i���X�gc) = Mid(.Cells(i, ���i���X�gT�ď�Col), 6)
                    Else
                        ���i���X�g(4, ���i���X�gc) = .Cells(i, ���i���X�gT�ď�Col)
                    End If
                    If ���i���X�g�[��Col <> 0 Then ���i���X�g(5, ���i���X�gc) = .Cells(i, ���i���X�g�[��Col) '��n���}�ŕ\������ɂ͐��i�i�Ԗ��̒[�������K�v
                    
                    For n = 1 To ���i�i��RANc
                        ���i���X�gs(n, ���i���X�gc) = .Cells(i, ���i���X�g�H��Col + n)
                    Next n
                    ���i���X�gc = ���i���X�gc + 1
                End If
                'End If
                Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[���i���X�g]����f�[�^�擾��", step0T, step0, ���i���X�glastRow, i)
            Next i
        End With
    End If
    
    Debug.Print "2= " & Round(Timer - sTime, 2): sTime = Timer
    
    Dim �I���o�� As String
    Dim �{�����[�h As Long: �{�����[�h = 1 '0(�����{) or 1(Cav��{)
    Dim �{�� As Single
    Dim frameWidth As Long, frameWidth1 As Long, frameWidth2 As Long, frameHeight1 As Long, frameHeight2 As Long, cornerSize As Single
    Dim pp As Long
    'Call PVSWcsv�ɓd�������擾_FromNMB_Ver1931
    step0 = step0 + 1
    Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF]����[PVSW_RLTF���[]���쐬", step0T, step0, 100, 100)
    Call PVSWcsv���[�̃V�[�g�쐬_Ver2001
    
    If �n���}�^�C�v = "�`�F�b�J�[�p" Then
        step0 = step0 + 1
        Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF���[]�Ƀ|�C���g�i���o�[�̎擾", step0T, step0, 100, 100)
        Call PVSWcsv���[�Ƀ|�C���g�擾
    End If

    Dim �n���}��� As String: �n���}��� = "�ʐ^" ' �ʐ^(�ʐ^���������͗��}) or ���}�B�g���q�̓n���}��ނɉ�����(�Œ�)PVSW_RLTF���[�Ƀn���}��ނ��o�͂��鎞�ɍs���B
    Dim �n���}�g���q As String
    'Dim �{�� As Single: If �n���}�^�C�v = "�`�F�b�J�[�p" Then �{�� = 2 Else �{�� = 1.4
    'PVSW_RLTF
    '2��16�i��_�ϊ�
    Dim ex As Long
    Dim varBinary As Variant
    Dim colHValue As New Collection  '�A�z�z��ACollection�I�u�W�F�N�g�̍쐬
    Dim lngNu() As Long
    varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
    Set colHValue = New Collection '������
    For ex = 0 To 15 '�A�z�z���varBinary�̊e�l���L�[�Ƃ��āA16�i�@�u0�`F�v�̒l���i�[
        colHValue.add CStr(Hex$(ex)), varBinary(ex)
    Next
    'PVSW_RLTF���[�̃f�[�^
    With wb(0).Sheets(mySheetName)
        Dim my�^�C�g��Row As Long: my�^�C�g��Row = .Cells.Find("�i��_").Row
        Dim my�^�C�g��Col As Long: my�^�C�g��Col = .Cells.Find("�i��_").Column
        Dim my�^�C�g��Ran As Range: Set my�^�C�g��Ran = Rows(my�^�C�g��Row) '.Range(.Cells(my�^�C�g��Row, 1), .Cells(my�^�C�g��Row, my�^�C�g��Col))
        Dim my�d�����ʖ�Col As Long: my�d�����ʖ�Col = .Cells.Find("�d�����ʖ�").Column
        Dim my��Col As Long: my��Col = .Cells.Find("��H����").Column
        Dim myCavCol As Long: myCavCol = .Cells.Find("�L���r�e�B").Column
        Dim my�[��Col As Long: my�[��Col = .Cells.Find("�[�����ʎq").Column
'        Dim my����Col As Long: my����Col = .Cells.Find("����No").Column
'        Dim my�����i��Col As Long: my�����i��Col = .Cells.Find("�����i��").Column
'        Dim myJointCol As Long: myJointCol = .Cells.Find("JOINT���").Column
        Dim my�_�u����Col As Long: my�_�u����Col = .Cells.Find("��_").Column
'        Dim my���i1Col As Long: my���i1Col = .Cells.Find("�[�q�i��").Column
'        Dim my���i2Col As Long: my���i2Col = .Cells.Find("�S����i��").Column
        Dim my���Col As Long: my���Col = .Cells.Find("��햼��").Column
        Dim my���Ӑ�Col As Long: my���Ӑ�Col = .Cells.Find("�[�����Ӑ�i��").Column
        Dim my���Col As Long: my���Col = .Cells.Find("�[�����i��").Column
'        Dim myJointGCol As Long: myJointGCol = .Cells.Find("�W���C���g�O���[�v").Column
'        Dim my�d���i��Col As Long: my�d���i��Col = .Cells.Find("�d���i��").Column
'        Dim my�d���T�C�YCol As Long: my�d���T�C�YCol = .Cells.Find("�d���T�C�Y").Column
'        Dim my�d���FCol As Long: my�d���FCol = .Cells.Find("�d���F").Column
        Dim my�}���}1Col As Long: my�}���}1Col = .Cells.Find("�}_").Column
        'Dim my�}���}2Col As Long: my�}���}2Col = .Cells.Find("�}���}�F�Q").Column
        'Dim myAB�敪Col As Long: myAB�敪Col = .Cells.Find("A/B�EB/C�敪").Column
        'Dim my�d��YBMCol As Long: my�d��YBMCol = .Cells.Find("�d���x�a�l").Column
        Dim my���葤Col As Long: my���葤Col = .Cells.Find("����_").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(my�^�C�g��Row, .Columns.count).End(xlToLeft).Column
'        Dim myPVSW�}���}1Col As Long: myPVSW�}���}1Col = .Cells.Find("�}���}�F�P").Column
        Dim my��col As Long: my��col = .Cells.Find("��_").Column
        Set my�^�C�g��Ran = Nothing
        'PVSW_RLTF���[�ɂ���NMB����̃f�[�^
        Dim my�i��Col As Long: my�i��Col = .Cells.Find("�i��_").Column
        Dim my�T�C�YCol As Long: my�T�C�YCol = .Cells.Find("�T�C�Y_").Column
        Dim my�T�C�Y��Col As Long: my�T�C�Y��Col = .Cells.Find("�T��_").Column
        Dim my�FCol As Long: my�FCol = .Cells.Find("�F_").Column
        Dim my�F��Col As Long: my�F��Col = .Cells.Find("�F��_", , , 1).Column
        Dim my����Col As Long: my����Col = .Cells.Find("����_", , , 1).Column
        Dim my����Col As Long: my����Col = .Cells.Find("����_", , , 1).Column
        Dim myJCDFcol As Long: myJCDFcol = .Cells.Find("JCDF_", , , 1).Column
        Dim my����Col As Long: my����Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim my�[�qCol As Long: my�[�qCol = .Cells.Find("�[�q_", , , 1).Column
        Dim my�}Col As Long: my�}Col = .Cells.Find("�}_", , , 1).Column
        
        Dim myRLTFtoPVSW As Long: myRLTFtoPVSW = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        'PVSW_RLTF���[�ɂ���T�u�}�f�[�^_Ver181�̒ǉ��f�[�^
        Dim my�T�uCol As Long: my�T�uCol = .Cells.Find("�T�u", , , 1).Column
        Dim my�n��Col As Long: my�n��Col = .Cells.Find("�n��", , , 1).Column
        Dim my���[�n��Col As Long: my���[�n��Col = .Cells.Find("���[�n��", , , 1).Column
        Dim my���[�qCol  As Long: my���[�qCol = .Cells.Find("���[���[�q", , , 1).Column
        my�n���i���o�[Col = .Cells.Find("#", , , 1).Column
        Dim my�F��2Col As Long: my�F��2Col = .Cells.Find("�F��", , , 1).Column
        Dim my�F��SIcol As Long: my�F��SIcol = .Cells.Find("�F��SI_", , , 1).Column
'        Dim my�n����Col As Long: my�n����Col = .Cells.Find("�n����", , , 1).Column
        
        'PVSW_RLTF���[�ɂ���|�C���g�̃f�[�^
        Dim my�|�C���g1Col As Long: my�|�C���g1Col = .Cells.Find("�|�C���g1_", , , 1).Column
        Dim my�|�C���g2Col As Long: my�|�C���g2Col = .Cells.Find("�|�C���g2_", , , 1).Column: Dim �|�C���g2 As String
        Dim my��d�W�~col As Long: my��d�W�~col = .Cells.Find("��d�W�~_", , , 1).Column
        Dim my���b�LCol As Long: my���b�LCol = .Cells.Find("��_", , , 1).Column
        Dim ��w As String
        Dim my�|�C���gResultCol As Long: my�|�C���gResultCol = .Cells.Find("PVSWtoPOINT_").Column: Dim �|�C���gResult As String
        Dim xx, c, myFlag, b As Long
        
        Dim kaiGyo As Long
        Select Case ���i�i��RANc
        Case 1, 2, 3, 4
            kaiGyo = ���i�i��RANc
        Case 3, 5, 6, 9
            kaiGyo = 3
        Case Else
            kaiGyo = 4
        End Select
        
        Dim myPartNameCol As Long: myPartNameCol = myLastCol + 1: .Cells(my�^�C�g��Row, myPartNameCol) = "PartName"
        Dim myX As Long: myX = myLastCol + 2: .Cells(my�^�C�g��Row, myX) = "x"
        Dim myY As Long: myY = myLastCol + 3: .Cells(my�^�C�g��Row, myY) = "y"
        Dim myW As Long: myW = myLastCol + 4: .Cells(my�^�C�g��Row, myW) = "width"
        Dim myH As Long: myH = myLastCol + 5: .Cells(my�^�C�g��Row, myH) = "height"
        Dim my�`��Col As Long: my�`��Col = myLastCol + 6: .Cells(my�^�C�g��Row, my�`��Col) = "�`��"
        Dim my�g�p�ԍ�Col As Long: my�g�p�ԍ�Col = myLastCol + 7: .Cells(my�^�C�g��Row, my�g�p�ԍ�Col) = "�g�p�ԍ�"
        Dim myWcol As Long: myWcol = myLastCol + 8: .Cells(my�^�C�g��Row, myWcol) = "Width"
        Dim my�n���}���Col As Long: my�n���}���Col = myLastCol + 9: .Cells(my�^�C�g��Row, my�n���}���Col) = "�n���}���"
        Dim my�n���}�g���qCol As Long: my�n���}�g���qCol = myLastCol + 10: .Cells(my�^�C�g��Row, my�n���}�g���qCol) = "�n���}�g���q"
        Dim myEmptyPlugCol As Long: myEmptyPlugCol = myLastCol + 11: .Cells(my�^�C�g��Row, myEmptyPlugCol) = "EmptyPlug"
        Dim myPlugColorCol As Long: myPlugColorCol = myLastCol + 12: .Cells(my�^�C�g��Row, myPlugColorCol) = "PlugColor"
        '.Cells(my�^�C�g��Row, myLastCol + 13) = "�n����"
        '���W�f�[�^�̓Ǎ���(�C���|�[�g�t�@�C��)
        Dim Target As New FileSystemObject
        Dim TargetDir As String: TargetDir = �A�h���X(1) & "\200_CAV���W"
        
        If Dir(TargetDir, vbDirectory) = "" Then MsgBox "���L�̃t�@�C���������ׁA�e�L���r�e�B�̍��W��������܂���B" & vbCrLf & "���ވꗗ+�ō��W�̏o�͂��s���Ă�����s���ĉ������B" & vbCrLf & vbCrLf & �A�h���X(1) & "\CAV���W.txt"
        
        Dim outY As Long: outY = 1
        Dim outX As Long
        Dim lastgyo As Long: lastgyo = 1
        Dim fileCount As Long: fileCount = 0
        Dim inX As Long
        Dim temp
        Dim �g�p���istr As String
        Dim �g�p���i_�[�� As String
        Dim Make���sflag As Long
        
        step0 = step0 + 1
        For i = my�^�C�g��Row + 1 To myLastRow
            If InStr(�g�p���i_�[��, .Cells(i, my���Col) & "_" & .Cells(i, my�[��Col)) = 0 Then
                �g�p���i_�[�� = �g�p���i_�[�� & "," & .Cells(i, my���Col) & "_" & .Cells(i, my�[��Col)
                Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF���[]����g�p���i�f�[�^�̎擾", step0T, step0, myLastRow, i)
            End If
        Next i
        
        Dim �g�p���i_�[��s As Variant
        Dim �g�p���i_�[��c As Variant
        Dim aa As Variant
        Dim ���W����Flag As Boolean
        Dim �g�p���i_�[��s_count As Long
        '�g�p���iStr�ɁA����g�p������W�f�[�^������
        Dim intFino As Variant
        intFino = FreeFile
        Dim ���r(1) As String
        step0 = step0 + 1
        �g�p���i_�[��s = Split(�g�p���i_�[��, ",")
        For Each �g�p���i_�[��c In �g�p���i_�[��s
            If �g�p���i_�[��c <> "" Then
                c = Split(�g�p���i_�[��c, "_")
                ���i�i��str = c(0)
                If Len(���i�i��str) = 10 Then ���i�i��str = Left(���i�i��str, 4) & "-" & Mid(���i�i��str, 5, 4) & "-" & Mid(���i�i��str, 9, 2) Else ���i�i��str = Left(���i�i��str, 4) & "-" & Mid(���i�i��str, 5, 4)
                ���W����Flag = False
                ���r(0) = "png": ���r(1) = "emf"
                For ss = 0 To 1
                    '�ʐ^,���}�̏��ŒT��
                    URL = �A�h���X(1) & "\200_CAV���W\" & ���i�i��str & "_1_001_" & ���r(ss) & ".txt"
                    If Dir(URL) <> "" Then
                        intFino = FreeFile
                        Open URL For Input As #intFino
                        Do Until EOF(intFino)
                            Line Input #intFino, aa
                            temp = Split(aa, ",")
                            If Replace(temp(0), "-", "") = c(0) Then
                                �g�p���istr = �g�p���istr & "," & temp(0) & "_" & temp(1) & "_" & temp(2) & "_" & temp(3) & "_" & temp(4) & "_" & temp(5) & "_" & temp(6) & "_" & temp(7) & "_" & temp(8) & "_" & temp(9) & "_" & c(1) & "_" & temp(10)
                            End If
                        Loop
                        Close #intFino
                        Exit For
                    End If
                Next ss
            End If
            Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "200_CAV���W����g�p����CAV���W���擾", step0T, step0, UBound(�g�p���i_�[��s), �g�p���i_�[��s_count)
            �g�p���i_�[��s_count = �g�p���i_�[��s_count + 1
        Next �g�p���i_�[��c
        Dim �g�p���i As Variant, �g�p���is As Variant, �g�p���ic As Variant

        step0 = step0 + 1
        For p = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            �g�p���is = Split(�g�p���istr, ",")
            For Each �g�p���ic In �g�p���is
                If �g�p���ic <> "" Then
                    myFlag = 0
                    temp = Split(�g�p���ic, "_")
                    For i = my�^�C�g��Row + 1 To myLastRow
                        If .Cells(i, p) <> "" Then
                            If .Cells(i, my���Col) = Replace(temp(0), "-", "") Then
                                If .Cells(i, my�[��Col) = Val(temp(10)) Then
                                    If .Cells(i, myCavCol) = Val(temp(1)) Then
                                        .Cells(i, myPartNameCol) = temp(0)
                                        .Cells(i, myX) = temp(2)
                                        .Cells(i, myY) = temp(3)
                                        .Cells(i, myW) = temp(4)
                                        .Cells(i, myH) = temp(5)
                                        .Cells(i, my�`��Col) = temp(7)
                                        .Cells(i, my�g�p�ԍ�Col) = temp(9)
                                        .Cells(i, myWcol) = temp(11)
                                        .Cells(i, my�n���}���Col) = temp(8)
                                        If temp(8) = "�ʐ^" Then
                                            .Cells(i, my�n���}�g���qCol) = ".png"
                                        Else
                                            .Cells(i, my�n���}�g���qCol) = ".emf"
                                        End If
                                        myFlag = 1
                                    End If
                                End If
                            End If
                        End If
                    Next i
                    '�Y���f�[�^����
                    If myFlag = 0 Then
                        Dim last����Row As Long: last����Row = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row + 1
                        Dim last�[��Row As Long: last�[��Row = .Cells(.Rows.count, my�[��Col).End(xlUp).Row + 1
                        Dim addLastRow As Long: If last����Row > last�[��Row Then addLastRow = last����Row Else addLastRow = last�[��Row
                        '.Range(.Cells(addLastRow, my���i�i��Ran0), .Cells(addLastRow, my���i�i��Ran1)) = 0
                        .Cells(addLastRow, p) = "0"
                        .Cells(addLastRow, my�[��Col) = temp(10)
                        .Cells(addLastRow, myCavCol) = temp(1)
                        .Cells(addLastRow, myPartNameCol) = temp(0)
                        .Cells(addLastRow, myX) = temp(2)
                        .Cells(addLastRow, myY) = temp(3)
                        .Cells(addLastRow, myW) = temp(4)
                        .Cells(addLastRow, myH) = temp(5)
                        .Cells(addLastRow, my�`��Col) = temp(7)
                        .Cells(addLastRow, my�g�p�ԍ�Col) = temp(9)
                        .Cells(addLastRow, myWcol) = temp(11)
                        .Cells(addLastRow, my�n���}���Col) = temp(8)
                        If temp(8) = "�ʐ^" Then
                            .Cells(addLastRow, my�n���}�g���qCol) = ".png"
                        Else
                            .Cells(addLastRow, my�n���}�g���qCol) = ".emf"
                        End If
                        .Cells(addLastRow, my���Col) = Replace(temp(0), "-", "")
                    End If
                End If
            Next �g�p���ic
            Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF���[]��CAV���W���Z�b�g", step0T, step0, UBound(���i�i��RAN, 2), p)
        Next p
    
'PartName���u�����N�̎��ɒ[�����i�Ԃ���擾_���߂�
        Dim ���a As String
        For i = my�^�C�g��Row + 1 To myLastRow
            If .Cells(i, myPartNameCol) = "" Then
                ���a = .Cells(i, my���Col)
                If Len(���a) <> 0 Then
                    If Len(���a) = 8 Then
                        ���a = Left(���a, 4) & "-" & Mid(���a, 5, 4)
                    Else
                        ���a = Left(���a, 4) & "-" & Mid(���a, 5, 4) & "-" & Mid(���a, 9)
                    End If
                    .Cells(i, myPartNameCol) = ���a
                End If
            End If
        Next i
    
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, my�[��Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, myPartNameCol).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, myCavCol).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        If addLastRow = 0 Then addLastRow = myLastRow '�󂫂�Cav��������
        .Sort.SetRange Range(Rows(2), Rows(addLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
    step0 = step0 + 1
    Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF���[]��[CAV�ꗗ]�̋�������Z�b�g", step0T, step0, 100, 100)
    'CAV�ꗗ�̃V�[�g�������EmptyPlug���擾����
    Dim tempFlg As Boolean
    Dim myRow As Long, myCol(5) As Long
    Dim �h���R�l�N�^v(4) As String
    For Each ws In Worksheets
        If ws.Name = "CAV�ꗗ" Then
            With wb(0).Sheets("CAV�ꗗ")
                Set myKey = .Cells.Find("���i�i��", , , 1)
                myCol(0) = myKey.Column
                myCol(1) = .Cells.Find("�[����", , , 1).Column
                myCol(2) = .Cells.Find("Cav", , , 1).Column
                myCol(3) = .Cells.Find("EmptyPlug", , , 1).Column
                myCol(4) = .Cells.Find("PlugColor", , , 1).Column
                myCol(5) = .Cells.Find(��n�����i�i��, , , 1).Column
                myRow = myKey.Row + 1
                cav�ꗗlastrow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
                cav�ꗗrow = myKey.Row + 1
                Do Until .Cells(myRow, myCol(0)) = ""
                    �h���R�l�N�^v(0) = .Cells(myRow, myCol(0))
                    �h���R�l�N�^v(1) = .Cells(myRow, myCol(1))
                    �h���R�l�N�^v(2) = .Cells(myRow, myCol(2))
                    �h���R�l�N�^v(3) = .Cells(myRow, myCol(3))
                    �h���R�l�N�^v(4) = .Cells(myRow, myCol(4))
                    With wb(0).Sheets(mySheetName)
                        For i = my�^�C�g��Row + 1 To .Cells(.Rows.count, my�[��Col).End(xlUp).Row
                            If CStr(.Cells(i, my���Col)) = Replace(�h���R�l�N�^v(0), "-", "") Then
                                If CStr(.Cells(i, myCavCol)) = �h���R�l�N�^v(2) Then
                                    If CStr(.Cells(i, my�[��Col)) = �h���R�l�N�^v(1) Then
                                        If .Cells(i, my�d�����ʖ�Col) = "" Then
                                            .Cells(i, myEmptyPlugCol) = �h���R�l�N�^v(3)
                                            .Cells(i, myPlugColorCol) = �h���R�l�N�^v(4)
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End With
                    myRow = myRow + 1
                Loop
            End With
        End If
    Next ws
    Set myKey = Nothing
    
    If �n���}�^�C�v = "�`�F�b�J�[�p" Then
        step0 = step0 + 1
        Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[PVSW_RLTF���[]��CAV���W���Z�b�g", step0T, step0, 100, 100)
        Call PVSWcsv���[�Ƀ|�C���g�擾
    End If
    
    '���[�N�V�[�g�̒ǉ�
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Tab.color = False
    
    'ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule.AddFromFile �A�h���X(0) & "\002_��A���쐬_�}���}.txt"
    
    Dim X, Y, w As Single, h As Single, minW As Single: minW = -1
    Dim minH As Single: minH = -1
    Dim cc As Long ', ccc As Long
    If addLastRow > myLastRow Then myLastRow = addLastRow
    Dim �d���f�[�^() As String: ReDim �d���f�[�^(2, 1) As String
    
    Call �œK��
    
    '�n���}
    step0 = step0 + 1
    For i = my�^�C�g��Row To myLastRow
        With wb(0).Sheets(mySheetName)
            'Set ���i�i��RAN = .Range(.Cells(i, my���i�i��Ran0), .Cells(i, my���i�i��Ran1))
            Set ���i�i��v = .Range(.Cells(i, 1), .Cells(i, ���i�i��RANc))
            If i = 1 Then GoTo line10
            Dim �d�����ʖ� As String: �d�����ʖ� = .Cells(i, my�d�����ʖ�Col)
            Dim �� As String: �� = .Cells(i, my��Col)
            Dim �[�� As String: �[�� = .Cells(i, my�[��Col)
            Dim ��� As String: ��� = .Cells(i, my���Col)
            If �[�� = "" Then GoTo line20
            If ��� = "" Then GoTo line20
            Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "[�n���}] �[��" & �[�� & " �̍쐬", step0T, step0, myLastRow, i)
            cav = .Cells(i, myCavCol)
'            Dim ���� As String: ���� = .Cells(i, my����Col)
'            Dim �����i�� As String: �����i�� = .Cells(i, my�����i��Col)
'            Dim �����i��co As Long: �����i��co = .Cells(i, my�����i��Col).Interior.Color
            'Dim Joint As String: Joint = .Cells(i, myJointCol)
            Dim �_�u���� As String: �_�u���� = .Cells(i, my�_�u����Col)
'            Dim ���i1 As String: ���i1 = .Cells(i, my���i1Col)
'            Dim ���i2 As String: ���i2 = .Cells(i, my���i2Col)
            Dim ��� As String: ��� = .Cells(i, my���Col)
            Dim ���Ӑ� As String: ���Ӑ� = .Cells(i, my���Ӑ�Col)
            'Dim JointG As String: JointG = .Cells(i, myJointGCol)
            Dim �}���}1 As String: �}���}1 = Replace(.Cells(i, my�}���}1Col), " ", "")
            '�}���}2 = .Cells(i, my�}���}2Col)
            'Dim AB�敪 As String: AB�敪 = .Cells(i, myAB�敪Col)
            'Dim �d��YBM As String: �d��YBM = .Cells(i, my�d��YBMCol)
            Dim ���葤 As String: ���葤 = .Cells(i, my���葤Col)
            'Dim �V�[���h�t���O As String: �V�[���h�t���O = " "
            Dim �� As String: �� = .Cells(i, my��col)
            'NMB����̃f�[�^
            Dim �i�� As String: �i�� = .Cells(i, my�i��Col)
            Dim �T�C�Y As String: �T�C�Y = .Cells(i, my�T�C�YCol)
            Dim �T�C�Y�� As String: �T�C�Y�� = .Cells(i, my�T�C�Y��Col)
            Dim �F As String: �F = .Cells(i, my�FCol)
            Dim �F�� As String: �F�� = .Cells(i, my�F��Col)
            If �F�� = "SI" And .Cells(i, my�F��SIcol) <> "" Then �F�� = �F�� & "_" & .Cells(i, my�F��SIcol)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim JCDF As String: JCDF = .Cells(i, myJCDFcol)
            Dim �[�q As String: �[�q = .Cells(i, my�[�qCol)
            Dim �} As String: �} = .Cells(i, my�}Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim �T�u As String: �T�u = .Cells(i, my�T�uCol)                  '�����͉��L�Əd��
            Dim �|�C���g1 As String: �|�C���g1 = .Cells(i, my�|�C���g1Col)   '�����͉��L�Əd��
            
            Dim �[��bak As String, �[��firstRow As Long, �[��firstRow2 As Long, ���bak As String, PartNamenext As String, PartNamebak As String
            Dim RLTFtoPVSW As String, partName As String, �`�� As String, �g�p�ԍ� As String, �� As String, �[��next As String, ���next As String
            RLTFtoPVSW = .Cells(i, myRLTFtoPVSW)
            partName = .Cells(i, myPartNameCol)
            
            X = .Cells(i, myX)
            Y = .Cells(i, myY)
            If .Cells(i, myW) = "" Then
                w = 0
            Else
                w = .Cells(i, myW)
                If w < minW Or minW = -1 Then minW = w
            End If
            If .Cells(i, myH) = "" Then
                h = 0
            Else
                h = .Cells(i, myH)
                If w < minH Or minH = -1 Then minH = h
            End If

            �`�� = .Cells(i, my�`��Col)
            �g�p�ԍ� = .Cells(i, my�g�p�ԍ�Col)
            �� = .Cells(i, myWcol)
            �n���}��� = .Cells(i, my�n���}���Col)
            �n���}�g���q = .Cells(i, my�n���}�g���qCol)
            �[��next = .Cells(i + 1, my�[��Col) '�[���̕`�悪�Ōォ�m�F
            ���next = .Cells(i + 1, my���Col) '�[���̕`�悪�Ōォ�m�F
            PartNamenext = .Cells(i + 1, myPartNameCol)
        End With
line10:
        
        With wb(0).Sheets(newSheetName)
            Dim ������r() As String: ReDim ������r(0, 2)
            Dim ���F As String
            If i = 1 Then
                For p = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                    '.Range(.Cells(1, my���i�i��Ran0), .Cells(1, my���i�i��Ran1)).Value = ���i�i��Ran.Value
                    .Cells(1, p) = ���i�i��RAN(1, p)
                Next p
                .Range(.Cells(1, 1), .Cells(1, ���i�i��RANc)).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 1) = "�[�����i��": .Columns(���i�i��RANc + 1).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 2) = "�\��": .Columns(���i�i��RANc + 2).NumberFormat = "@"
                '.Cells(1, ���i�i��ranc + 3) = "�\��": .Columns(���i�i��ranc + 3).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 3) = "�i��": .Columns(���i�i��RANc + 3).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 4) = "�T�C�Y": .Columns(���i�i��RANc + 4).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 5) = "�F�ď�": .Columns(���i�i��RANc + 5).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 6) = "�[����": .Columns(���i�i��RANc + 6).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 7) = "Cav": .Columns(���i�i��RANc + 7).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 8) = "�F": .Columns(���i�i��RANc + 8).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 9) = "��": .Columns(���i�i��RANc + 9).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 10) = "��": .Columns(���i�i��RANc + 10).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 11) = "�}": .Columns(���i�i��RANc + 11).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 12) = "�}1": .Columns(���i�i��RANc + 12).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 13) = "���葤": .Columns(���i�i��RANc + 13).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 14) = "��": .Columns(���i�i��RANc + 14).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 15) = "Point": .Columns(���i�i��RANc + 15).NumberFormat = "@"
                .Cells(1, ���i�i��RANc + 16) = "Sub": .Columns(���i�i��RANc + 16).NumberFormat = "@"
                Dim myColPoint As Single: myColPoint = .Cells(1, ���i�i��RANc + 18).Left
                Dim myRowPoint As Single: myRowPoint = .Rows(2).Top
                Dim myRowSel As Long: myRowSel = 2
                Dim myRowHeight As Single: myRowHeight = .Rows(1).Height
            Else
                If �i�� = "" Then GoTo line15
                '.Range(.Cells(myRowSel, my���i�i��Ran0), .Cells(myRowSel, my���i�i��Ran1)) = ���i�i��Ran.Value
                If partName = "" Then
                    If Len(���) = 8 Then
                        partName = Left(���, 4) & "-" & Mid(���, 5, 4)
                    Else
                        partName = Left(���, 4) & "-" & Mid(���, 5, 4) & "-" & Mid(���, 9, 4)
                    End If
                End If
                
                �d���f�[�^a = partName & "," & Left(�d�����ʖ�, 4) & "," & �i�� & "," & �T�C�Y�� & "," & �F�� & "," & _
                                              �[�� & "," & cav & ",," & �� & "," & �_�u���� & "," & �}���}1 & "," & _
                                              �}���}1 & "," & ���葤 & "," & �� & "," & �|�C���g1 & "," & �T�u
                '�����������o�^����ĂȂ����m�F
                c = 1: �d���f�[�^1�܂Ƃ� = ""
                
                For p = 1 To UBound(�d���f�[�^, 2)
                    If �d���f�[�^(2, p) = �d���f�[�^a Then
                        �d���f�[�^s = Split(�d���f�[�^(1, p), ",")
                        For Each �d���f�[�^ss In �d���f�[�^s
                            If ���i�i��v(c) <> "" Then
                                �d���f�[�^1�܂Ƃ� = �d���f�[�^1�܂Ƃ� & ���i�i��v(c) & ","
                            Else
                                �d���f�[�^1�܂Ƃ� = �d���f�[�^1�܂Ƃ� & �d���f�[�^ss & ","
                            End If
                            c = c + 1
                        Next
                        �d���f�[�^1�܂Ƃ� = Left(�d���f�[�^1�܂Ƃ�, Len(�d���f�[�^1�܂Ƃ�) - 1)
                        �d���f�[�^(1, p) = �d���f�[�^1�܂Ƃ�
                    End If
                Next p
                
                '�V�K�ǉ�
                If c = 1 Then
                    pp = pp + 1
                    ReDim Preserve �d���f�[�^(2, pp)
                    For Each ���i�i��vv In ���i�i��v
                        �d���f�[�^(1, pp) = �d���f�[�^(1, pp) & ���i�i��vv & ","
                    Next
                    �d���f�[�^(1, pp) = Left(�d���f�[�^(1, pp), Len(�d���f�[�^(1, pp)) - 1)
                    �d���f�[�^(2, pp) = �d���f�[�^a
                End If
line15:

                'If �[�� = 7 Then Stop
                '���i&�[�����ω�������������f�[�^�Ɛ}���o�͂���
                If partName <> "" Then
                    If �[�� & "_" & partName <> �[��bak & "_" & PartNamebak Then �[��firstRow = i: �[��firstRow2 = myRowSel
                    If �[�� & "_" & partName <> �[��next & "_" & PartNamenext Then
                    '�d���f�[�^�o��
                    For p = 1 To pp
                        If �d���f�[�^(1, p) <> "" Then
                            �d���f�[�^1s = Split(�d���f�[�^(1, p), ",")
                            c = 1
                            For Each �d���f�[�^1ss In �d���f�[�^1s
                                .Cells(myRowSel, c).NumberFormat = "@"
                                .Cells(myRowSel, c) = �d���f�[�^1ss
                                c = c + 1
                            Next
                            �d���f�[�^2s = Split(�d���f�[�^(2, p), ",")
                            For Each �d���f�[�^2ss In �d���f�[�^2s
                                .Cells(myRowSel, c).NumberFormat = "@"
                                .Cells(myRowSel, c) = �d���f�[�^2ss
                                If .Cells(1, c) = "�F" Then
                                    ���F = CStr(.Cells(myRowSel, ���i�i��RANc + 5))
                                    '�V�[���hSI�̎��̓`���[�u�F�ɕύX
                                    If InStr(���F, "_") > 0 Then
                                        ���F = Mid(���F, InStr(���F, "_") + 1)
                                    End If
                                    Call �d���F�ŃZ����h��(myRowSel, CLng(c), ���F)
                                End If
                                c = c + 1
                            Next
                            myRowSel = myRowSel + 1
                        End If
                    Next p
                    'Erase �d���f�[�^
                    ReDim �d���f�[�^(2, 1) As String
                        With wb(0).Sheets(mySheetName)
                            '�}�̏���
                            Dim ���i��r() As String: ReDim ���i��r(���i�i��RANc, 4) '0=���i�i��,1=�d������,2=�킩���,3=MAX��H�ɂ��A���}�b�`,4=MAX��H�̏���
                            For p = 1 To ���i�i��RANc
                                ���i��r(p, 2) = 0
                                ���i��r(p, 3) = 0
                                For b = �[��firstRow To i
                                    If .Cells(b, p) <> "" Then
                                        �\�� = Left(.Cells(b, my�d�����ʖ�Col), 4)
                                        If �n����ƕ\�� <> "" And �\�� = "" Then GoTo line155
                                        �n���i���o�[ = .Cells(b, my�n���i���o�[Col)
                                        If �n���i���o�[ > �n����ƕ\�� And �n����ƕ\�� <> "" Then GoTo line155
                                        �|�C���g1 = .Cells(b, my�|�C���g1Col)
                                        �|�C���g2 = .Cells(b, my�|�C���g2Col)
                                        �|�C���gResult = .Cells(b, my�|�C���gResultCol)
                                        �� = .Cells(b, my��Col)
                                        ��w = Left(.Cells(b, my�_�u����Col), 4)
                                        �F�� = Replace(.Cells(b, my�F��Col), " ", "")
                                        '�V�[���hSI�̎��A�F�Ă��`���[�u�F�ɕύX
                                        If �F�� = "SI" And .Cells(b, my�F��SIcol) <> "" Then �F�� = .Cells(b, my�F��SIcol)
                                        �T�C�Y�� = Replace(.Cells(b, my�T�C�Y��Col), "F", "")
                                        �}���}1 = Replace(.Cells(b, my�}���}1Col), " ", "")
                                        �V�[���h�t���O = " "
                                        ��ƋL�� = .Cells(b, my�F��2Col + 1)
                                        ����[�� = .Cells(b, my���葤Col)
                                        If ����[�� <> "" Then ����[�� = Left(.Cells(b, my���葤Col), InStr(.Cells(b, my���葤Col), "_") - 1)
                                        'PVSW_RLTF���[�ɂ���T�u�}�f�[�^_Ver181�̒ǉ��f�[�^
                                        �T�u = .Cells(b, my�T�uCol)
                                        �n�� = .Cells(b, my�n��Col) & "!" & .Cells(b, my���[�n��Col) & "!" & .Cells(b, my���[�qCol) & "!" & .Cells(b, my�n���i���o�[Col) & "!" & .Cells(b, my��d�W�~col) & "!" & .Cells(b, my���b�LCol) '���[�n���̓n���}�ŗ��[����n���̎���1�A
                                        
                                        Select Case �n���}�^�C�v
                                        Case "�`�F�b�J�[�p"
                                            If �|�C���g2 = "" Then
                                                �I���o�� = �|�C���g1
                                            Else
                                                �I���o�� = �|�C���g1 & "!" & �|�C���g2
                                            End If
                                        Case "��H����"
                                            �I���o�� = ��
                                        Case "�\��"
                                            �I���o�� = �\��
                                        Case "����[��"
                                            �I���o�� = ����[��
                                        End Select
                                        
                                        If ��ƕ\���ϊ� = "1" And ��ƋL�� <> "" Then
                                            �T�C�Y�� = ��ƋL��
                                        End If
                                        
                                        '�f�[�^�����ʉ����������
                                        If ���i��r(p, 1) = "" Then
                                            ���i��r(p, 0) = .Cells(1, p)
                                            With wb(0).Sheets(mySheetName)
                                                ���i��r(p, 1) = .Cells(b, myX) & "_" & .Cells(b, myY) & "_" & .Cells(b, myW) & "_" & .Cells(b, myH) & "_" & �F�� & "_" & _
                                                                �}���}1 & "_" & �V�[���h�t���O & "_" & �I���o�� & "_" & Left(�T�C�Y��, 3) & "_" & �n�� & "_" & _
                                                                .Cells(b, myEmptyPlugCol) & "_" & .Cells(b, myPlugColorCol) & "_" & ��w & "_" & .Cells(b, myCavCol)
                                                If �F�� <> "" Then ���i��r(p, 2) = 1
                                            End With
                                        Else
                                            ���i��r(p, 0) = .Cells(1, p)
                                            With wb(0).Sheets(mySheetName)
                                                ���i��r(p, 1) = ���i��r(p, 1) & "," & .Cells(b, myX) & "_" & .Cells(b, myY) & "_" & .Cells(b, myW) & "_" & .Cells(b, myH) & "_" & �F�� & "_" & _
                                                                 �}���}1 & "_" & �V�[���h�t���O & "_" & �I���o�� & "_" & Left(�T�C�Y��, 3) & "_" & �n�� & "_" & _
                                                                 .Cells(b, myEmptyPlugCol) & "_" & .Cells(b, myPlugColorCol) & "_" & ��w & "_" & .Cells(b, myCavCol)
                                                If �F�� <> "" Then ���i��r(p, 2) = 1
                                            End With
                                            
'                                            If �[�� = 4 Then Debug.Print ���i��r(p, 0), ���i��r(p, 1), ���i��r(p, 2)
'                                            If �[�� = 5 Then Stop
                                        End If
                                    Else

                                    End If
line155:
                                Next b
                            Next p
                        End With
                      
                        '��r����������狤�ʉ��Ɏg�p���Ȃ��������폜����
'                        Dim ���i��rc As Variant
'                        For p = 1 To ���i�i��RANc
'                            jog2 = ""
'                            ���i��rc = Split(���i��r(p, 1), ",")
'                            For c = LBound(���i��rc) To UBound(���i��rc)
'                                jog = ""
'                                ���i��rcc = Split(���i��rc(c), "_")
'                                For cc = LBound(���i��rcc) To UBound(���i��rcc)
'                                    Debug.Print ���i��rcc(cc)
'                                    If cc <> 12 Then '�����폜
'                                        jog = jog & "_" & ���i��rcc(cc)
'                                    Else
'                                        jog = jog & "_" & ""
'                                    End If
'                                Next cc
'                                jog2 = jog2 & "," & Mid(jog, 2)
'                            Next c
'                            ���i��r(p, 1) = Mid(jog2, 2)
'                        Next p
                        
                        '�����������������ꍇ�A�u�����N�̕����폜(�_�u���Ƃ��{���_�[)
                        Dim sp1 As Variant, sp2 As Variant
                        Dim c2 As Long, cTemp As Long, cCav As String, c2temp As Long, c2Cav As String, temp___ As Long
                        Dim c1��w As String, c2��w As String
                        For p = 1 To ���i�i��RANc
                            ���i��rc = Split(���i��r(p, 1), ",")
                            For c = LBound(���i��rc) To UBound(���i��rc)
                                For c2 = LBound(���i��rc) To UBound(���i��rc)
                                    If c <> c2 Then
                                        sp1 = Split(���i��rc(c), "_")
                                        cCav = sp1(13)
                                        'c1��w = sp1(12)
                                        sp2 = Split(���i��rc(c2), "_")
                                        c2Cav = sp2(13)
                                        'c2��w = sp2(12)
                                        If cCav = c2Cav Then
                                            'temp___ = Replace(���i��rc(c), "_", "")
                                            'If temp___ = "" Then ���i��r(p, 1) = Replace(���i��r(p, 1), ���i��rc(c), "")
                                            If Replace(���i��r(p, 1), ",", "") = "" Then ���i��r(p, 1) = ���i��rc(c)
                                        End If
                                    End If
                                Next c2
                            Next c
                        Next p
                        Dim p2 As Long, pp2 As Long
                        
                        '���i���̏����������Ȃ琻�i�i�Ԃ�����
                        If MAX��H�\�� = "1" Then
                           '�_�u�������̎��A����������w��t����
                            For p = 1 To ���i�i��RANc
                                If ���i��r(p, 2) = 1 Then
                                    ���i��rc = Split(���i��r(p, 1), ",")
                                    koshin = ""
                                    For ppp = LBound(���i��rc) To UBound(���i��rc)
                                        cav1s = Split(���i��rc(ppp), "_")
                                        Cav1 = cav1s(13)
                                        flg = False
                                            For ppp2 = LBound(���i��rc) To UBound(���i��rc)
                                                If ppp <> ppp2 Then
                                                    cav2s = Split(���i��rc(ppp2), "_")
                                                    cav2 = cav2s(13)
                                                    If Cav1 = cav2 Then
                                                        flg = True
                                                    End If
                                                End If
                                            Next ppp2
                                        If flg = True Then
                                            koshin = koshin & ���i��rc(ppp) & "_w,"
                                        Else
                                            koshin = koshin & ���i��rc(ppp) & "_,"
                                        End If
                                    Next ppp
                                    ���i��r(p, 1) = Left(koshin, Len(koshin) - 1)
                                End If
                            Next p

                            '�����������Ȃ琻�i�i�Ԃ�����
                            For p = 1 To ���i�i��RANc
                                If ���i��r(p, 2) = 1 Then
                                    For p2 = 1 To ���i�i��RANc
                                        If ���i��r(p2, 2) = 1 Then
                                            If p <> p2 Then
                                                flg = False: max1 = 0: max2 = 0: kari = ""
                                                If ���i��r(p, 4) = "" Then
                                                    ���i��rc = Split(���i��r(p, 1), ",")
                                                Else
                                                    ���i��rc = Split(���i��r(p, 4), ",")
                                                End If
                                                If ���i��r(p2, 4) = "" Then
                                                    ���i��rc2 = Split(���i��r(p2, 1), ",")
                                                Else
                                                    ���i��rc2 = Split(���i��r(p2, 4), ",")
                                                End If
                                                For ppp = LBound(���i��rc) To UBound(���i��rc)
                                                    ���i��rcc = Split(���i��rc(ppp), "_")
                                                    Cav1 = ���i��rcc(13)
                                                    iro1 = ���i��rcc(4)
                                                    kai1 = ���i��rcc(12)
                                                    mei1 = ���i��rcc(9)
                                                    w1 = ���i��rcc(14)
                                                    
                                                    For ppp2 = LBound(���i��rc2) To UBound(���i��rc2)
                                                            ���i��rcc2 = Split(���i��rc2(ppp2), "_")
                                                            cav2 = ���i��rcc2(13)
                                                            iro2 = ���i��rcc2(4)
                                                            kai2 = ���i��rcc2(12)
                                                            mei2 = ���i��rcc2(9)
                                                            w2 = ���i��rcc2(14)
                                                        If Cav1 = cav2 Then
                                                            'If Cav1 = 52 Then Stop
                                                            '�_�u���ȊO
                                                            If w1 = "" And w2 = "" Then
                                                                If iro1 = iro2 And maj1 = maj2 Then
                                                                    kari = kari & ���i��rc(ppp) & ","
                                                                ElseIf iro1 = "" And iro2 <> "" Then
                                                                    kari = kari & ���i��rc2(ppp2) & ","
                                                                    max1 = 1
                                                                ElseIf iro1 <> "" And iro2 = "" Then
                                                                    kari = kari & ���i��rc(ppp) & ","
                                                                    max2 = 1
                                                                Else
                                                                    flg = True
                                                                End If
                                                            '�_�u��
                                                            ElseIf Left(mei1, 5) <> "Bonda" Then
                                                                If w1 = "w" And w2 = "w" Then
                                                                    If iro1 = iro2 And maj1 = maj2 Then
                                                                        kari = kari & ���i��rc(ppp) & ","
                                                                    ElseIf iro1 = "" And iro2 <> "" Then
                                                                        kari = kari & ���i��rc2(ppp2) & ","
                                                                        max1 = 1
                                                                    ElseIf iro1 <> "" And iro2 = "" Then
                                                                        kari = kari & ���i��rc(ppp) & ","
                                                                        max2 = 1
                                                                    Else
                                                                        'flg = true��L���ɂ�����_�u�舳�������邾���ŋ��ʉ�����Ȃ�
                                                                        kari = kari & ���i��rc(ppp) & ","
                                                                        'flg = True
                                                                    End If
                                                                End If
                                                            'Bonda
                                                            ElseIf Left(mei1, 5) = "Bonda" And Left(mei2, 5) = "Bonda" Then
                                                                kari = kari & ���i��rc(ppp) & ","
                                                                If InStr(kari, ",") <> InStrRev(kari, ",") Then Exit For
                                                            End If
                                                        End If
                                                    Next ppp2
                                                Next ppp
                                                If flg = False Then
                                                    '�����̓o�^
                                                    If kari <> "" Then
                                                        ���i��r(p, 4) = Left(kari, Len(kari) - 1)
                                                        ���i��r(p, 0) = ���i��r(p, 0) & "_" & ���i��r(p2, 0)
                                                        ���i��r(p2, 0) = ""
                                                        ���i��r(p2, 2) = 0
                                                        ���i��r(p, 3) = ���i��r(p, 3) & "0"
                                                    Else
                                                        ���i��r(p, 4) = ���i��r(p, 1)
                                                    End If
                                                End If
                                                
                                            End If
                                        End If
                                    Next p2
                                Else
                                    '���i��r(p, 3) = ���i��r(p, 3) & "0"
                                End If
                            Next p
                            '�������������`�F�b�N
                            For p = 1 To ���i�i��RANc
                                For pp4 = LBound(���i��r, 1) To UBound(���i��r, 1)
                                    If ���i��r(pp4, 4) <> "" Then
                                        If ���i��r(pp4, 0) Like "*" & ���i�i��RAN(1, p) & "*" Then
                                            ���i��r0s = Split(���i��r(p, 1), ",")
                                            ���i��r4s = Split(���i��r(pp4, 4), ",")
                                            flg = False
                                            If UBound(���i��r0s) = UBound(���i��r4s) Then
                                                For pp5 = LBound(���i��r4s) To UBound(���i��r4s)
                                                    ���i��r0ss = Split(���i��r0s(pp5), "_")
                                                    ���i��r4ss = Split(���i��r4s(pp5), "_")
                                                    If ���i��r0ss(4) <> ���i��r4ss(4) Or _
                                                       ���i��r0ss(5) <> ���i��r4ss(5) Or _
                                                       ���i��r0ss(12) <> ���i��r4ss(12) Then '4=�F��,5=�}���},12=cav
                                                        flg = True
                                                        Exit For
                                                    End If
                                                Next pp5
                                            Else
                                                flg = True
                                            End If
                                            If flg = True Then '�o�^���Ă�������ƈقȂ�Ȃ�1
                                                bb = InStr(���i��r(pp4, 0), ���i�i��RAN(1, p))
                                                bbb = (bb \ 16) + 1
                                                kari = ""
                                                For p2 = 1 To Len(���i��r(pp4, 3))
                                                    If p2 = bbb Then
                                                        kari = kari & "1"
                                                    Else
                                                        kari = kari & Mid(���i��r(pp4, 3), p2, 1)
                                                    End If
                                                Next p2
                                                ���i��r(pp4, 3) = kari
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next pp4
                            Next p
                            '�����̍X�V
                            For pp4 = LBound(���i��r, 1) To UBound(���i��r, 1)
                                If ���i��r(pp4, 4) <> "" Then
                                    ���i��r(pp4, 1) = ���i��r(pp4, 4)
                                End If
                            Next pp4
                        Else
                            For p = 1 To ���i�i��RANc
                                For p2 = 1 To ���i�i��RANc
                                    If p <> p2 Then
                                        If ���i��r(p, 0) <> "" Then
                                            If ���i��r(p, 1) = ���i��r(p2, 1) Then
                                                ���i��r(p, 0) = ���i��r(p, 0) & "_" & ���i��r(p2, 0)
                                                ���i��r(p2, 0) = ""
                                            End If
                                        End If
                                    End If
                                Next p2
                            Next p
                        End If
                        
                        '�����������i�i�Ԗ��ɐ}���쐬_1.941
                        If �n���}�^�C�v = "0" Then GoTo line17
                        Dim �� As Long, ���� As Long, �摜URL As String
                        �� = 0: ���� = 0: �n��count = 0
                        '���̒[���̃n����Ɛ����J�E���g
                        For p = 1 To ���i�i��RANc
                            ���i��rs = Split(���i��r(p, 1), ",")
                            For e = LBound(���i��rs) To UBound(���i��rs)
                                ���i��rss = Split(���i��rs(e), "_")
                                ���i��rsss = Split(���i��rss(9), "!")
                                If ���i��rsss(3) <= �n����ƕ\�� Then
                                    �n��count = �n��count + 1
                                End If
                            Next e
                        Next p

                        If �n��count = 0 And �F�Ŕ��f = True And �n����ƕ\�� <> "" Then GoTo line17
                        If �n��count = 0 And �F�Ŕ��f = False And �n����ƕ\�� <> "" Then GoTo line17
                        
                        For p = 1 To ���i�i��RANc
                            If ���i��r(p, 0) <> "" And ���i��r(p, 2) = 1 Then
                                
                                '�g�������m�F
                                Dim �g�������� As String: �g�������� = ""
                                Dim ���i�i��c As Variant, ���i�i�Ԗ�v As Variant, flag As Long
                                For o = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                                    ���i�i�Ԗ�v = ���i�i��RAN(1, o)
                                    ���i�i��c = Split(���i��r(p, 0), "_")
                                    flag = 0
                                    For Each c In ���i�i��c
                                        If Replace(���i�i�Ԗ�v, " ", "") = Replace(c, " ", "") Then
                                            �g�������� = �g�������� & 1
                                            flag = 1
                                        End If
                                    Next c
                                    If flag = 0 Then �g�������� = �g�������� & 0
                                Next

                                Dim BtoH As String
                                Dim strB As String
                                strB = �g��������
                                Dim myLen As Long
                                myLen = RoundUp(Len(strB) / 4, 0)
                                strB = String((myLen * 4) - Len(strB), "0") & strB '����������Ȃ��ꍇ,0������
                                ReDim strBtoH(1 To myLen)
                                For ex = 1 To myLen '2�i�@(4bit��)��16�i�@�ɕϊ�
                                    strBtoH(ex) = colHValue.Item(Mid$(strB, (ex - 1) * 4 + 1, 4))
                                Next
                                BtoH = Join$(strBtoH, vbNullString)
                                �[���} = �[�� & "_" & BtoH
                                
                                '�摜�̔z�u
                                ReDim ���\�L(2, 0): ���c = 0
                                Dim �摜����flg As Boolean: �摜����flg = False
                                '�ʐ^
                                �摜URL = �A�h���X(1) & "\���ވꗗ+_�ʐ^\" & partName & "_1_" & Format(1, "000") & ".png"
                                If Dir(�摜URL) = "" Then
                                    '���}
                                    �摜URL = �A�h���X(1) & "\���ވꗗ+_���}\" & partName & "_1_" & Format(1, "000") & ".emf"
                                    If Dir(�摜URL) = "" Then
                                        �摜����flg = True 'GoTo line18
                                    End If
                                End If
                                
                                'If minW = -1 Then GoTo line18 'Cav���W��������Ώ������Ȃ�
                                If �摜����flg = True Then 'CAV���W�Ƀf�[�^��������
                                    With ActiveSheet
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 60).Name = �[���}
                                        On Error Resume Next
                                            .Shapes.Range(�[���}).Adjustments.Item(1) = 0.1
                                        On Error GoTo 0
                                        .Shapes.Range(�[���}).Line.Weight = 1.6
                                        .Shapes.Range(�[���}).TextFrame2.TextRange.Text = ""
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 35, 10, 80, 40).Name = �[���} & "_1"
                                        .Shapes.Range(�[���} & "_1").Adjustments.Item(1) = 0.1
                                        .Shapes.Range(�[���} & "_1").Line.Weight = 1.6
                                        .Shapes.Range(�[���} & "_1").TextFrame2.TextRange.Text = "no picture"
                                        .Shapes.Range(�[���}).Select
                                        .Shapes.Range(�[���} & "_1").Select False
                                        Selection.Group.Select
                                        Selection.Name = �[���}
                                    End With
                                ElseIf minW = -1 Then  '�摜��������
                                    With ActiveSheet
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 60).Name = �[���}
                                        On Error Resume Next
                                            .Shapes.Range(�[���}).Adjustments.Item(1) = 0.1
                                        On Error GoTo 0
                                        .Shapes.Range(�[���}).Line.Weight = 1.6
                                        .Shapes.Range(�[���}).TextFrame2.TextRange.Text = ""
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 35, 10, 80, 40).Name = �[���} & "_1"
                                        .Shapes.Range(�[���} & "_1").Adjustments.Item(1) = 0.1
                                        .Shapes.Range(�[���} & "_1").Line.Weight = 1.6
                                        .Shapes.Range(�[���} & "_1").TextFrame2.TextRange.Text = "Cav���W������"
                                        .Shapes.Range(�[���}).Select
                                        .Shapes.Range(�[���} & "_1").Select False
                                        Selection.Group.Select
                                        Selection.Name = �[���}
                                    End With
                                Else
                                    With ActiveSheet.Pictures.Insert(�摜URL)
                                        .Name = �[���}
                                        .ShapeRange(�[���}).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft '�摜���傫���ƃT�C�Y������������邩���̃T�C�Y�ɖ߂�
                                        If �{�����[�h = 1 Then '�ϔ{
                                            'Debug.Print �[�� & "_" & minW & "_" & minH
                                            If minW < minH Then
                                                my�� = (minW�w�� / minW)
                                            Else
                                                my�� = (minW�w�� / minH)
                                            End If
                                            If �`�� = "Cir" Then my�� = my�� * 1.2
                                        Else
                                            my�� = .Width / (.Width / 3.08) * ��
                                            my�� = my�� / .Width * �{��
                                        End If
                                        .ShapeRange(�[���}).ScaleHeight my��, msoTrue, msoScaleFromTopLeft
                                        .CopyPicture
                                        .Delete
                                    End With
                                    DoEvents
                                    Sleep 70  '��2.191.06��10��70_��������
                                    DoEvents
                                    .Paste
                                    Selection.Name = �[���}
                                End If
                                .Shapes(�[���}).Left = 0
                                .Shapes(�[���}).Top = 0
                                Dim myPicHeight As Single: myPicHeight = .Shapes(�[���}).Height
                                
                                '�F�̔z�u
                                If minW <> -1 And �摜����flg = False Then 'CAV���W�Ƀf�[�^��������
                                    '���^�p�x
                                    With wb(0).Sheets("�[���ꗗ")
                                        �[��Col = .Cells.Find("���^�p�x", , , 1).Column
                                        �[��row = .Cells.Find(�[��, , , 1).Row
                                        ���^�p�x = .Cells(�[��row, �[��Col)
                                    End With
                                    �[��cav�W�� = ""
                                    Dim RowStr As Variant, myStr As Variant, V As Variant
                                    Dim ��n��count As Long: ��n��count = 0
                                    Dim ��n��count As Long: ��n��count = 0
                                    Dim cavBak As Long, skipFlg As Boolean
                                    RowStr = Split(���i��r(p, 1), ",")
                                    cavCount = 0
                                    For n = LBound(RowStr) To UBound(RowStr)
                                        If RowStr(n) <> "" Then
                                            skipFlg = False
                                            V = Split(RowStr(n), "_")
                                            �n��s = Split(V(9), "!")
                                            cav = V(13)
                                            If cav = cavBak Then
                                                cavCount = cavCount + 1
                                            Else
                                                cavCount = 1
                                            End If
                                            '�n����ƕ\����I�����Ă���ꍇ
                                            If �n����ƕ\�� <> "" Then
                                                If �n����ƕ\�� < �n��s(3) Then
                                                    V(4) = ""
                                                    V(7) = ""
                                                End If
                                            End If
                                            '�z���}�̌�n���}�A��n���d���͕\�����Ȃ�
                                            If Left(V(9), 1) = "��" And �n���\�� = 4 Then skipFlg = True
                                            If V(4) = "" And �n���\�� = 4 Then
                                                skipFlg = True
                                            End If
                                            '��cav��2�𒴂���ꍇ�͏������΂�_MAX���ʂ̎��{���_�[��2�𒴂���
                                            If (cavCount <= 2 And Not (Left(V(9), 5) = "Bonda")) Or (cavCount = 1 And (Left(V(9), 5) = "Bonda")) Then
                                                If V(0) <> "" And V(0) <> 0 Then
                                                    If skipFlg = False Then
                                                        Call ColorMark3(�[��, CSng(V(0)), CSng(V(1)), CSng(V(2)), CSng(V(3)), Replace(V(4), " ", ""), �n���}���, �`��, Replace(CStr(V(5)), " ", ""), V(6), V(7), V(8), V(9), V(10), V(11), RowStr)
                                                    End If
                                                End If
                                            End If
                                            cavBak = cav
                                        End If
                                        '��n�����A�v���b�N�̕\���p
                                        If cavCount = 1 Then
                                            If Left(�n��s(0), 1) = "��" Then
                                                ��n��count = ��n��count + 1
                                            ElseIf Left(�n��s(0), 1) = "��" Then
                                                ��n��count = ��n��count + 1
                                            End If
                                        End If
                                    Next n
                                    '���^�p�x
                                    For n = LBound(RowStr) To UBound(RowStr)
                                        If RowStr(n) <> "" Then
                                            V = Split(RowStr(n), "_")
                                            cav = V(13)
                                            If ���^�p�x <> "" Then
                                                On Error Resume Next
                                                Select Case ���^�p�x
                                                Case "90"
                                                    'election.ShapeRange.TextFrame2.Orientation = msoTextOrientationUpward
                                                Case "180"
                                                    .Shapes.Range(�[���} & "_" & cav).Rotation = ���^�p�x
                                                Case "270"
                                                    'Selection.ShapeRange.TextFrame2.Orientation = msoTextOrientationDownward
                                                End Select
                                                On Error GoTo 0
                                            End If
                                        End If
                                    Next n
                                    
                                    Dim �[��cav�W��s As Variant, �[��cav�W��c As Variant
                                    �[��cav�W��s = Split(�[��cav�W��, ",")
                                    For Each �[��cav�W��c In �[��cav�W��s
                                        On Error Resume Next '�_�u���̏ꍇ���O�ς���Ă邩��
                                        .Shapes.Range(�[��cav�W��c).Select False
                                        On Error GoTo 0
                                    Next
                                    .Shapes.Range(�[���}).Select False
                                    If Selection.ShapeRange.count > 1 Then
                                        Selection.Group.Select
                                        Selection.Name = �[���}
                                    End If
                                    If ���^�p�x <> "" Then
                                        Dim Large As Long
                                        If Selection.Width > Selection.Height Then
                                            Large = Selection.Width
                                        Else
                                            Large = Selection.Height
                                        End If
                                        Selection.Left = Large
                                        Selection.Top = Large
                                        .Shapes(�[���}).LockAspectRatio = msoTrue
                                        .Shapes(�[���}).Rotation = ���^�p�x
                                        Selection.Left = 0
                                        Selection.Top = 0
                                    End If
                                End If
                                    
                                frameWidth1 = .Shapes(�[���}).Width
                                frameHeight1 = .Shapes(�[���}).Height
                                If �[���i���o�[�\�� = True Then
                                    '�[�����^�C�g��
                                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 60, 32).Name = �[���} & "_b"
                                    .Shapes.Range(�[���} & "_b").Adjustments.Item(1) = 0.2
                                    cornerSize = .Shapes.Range(�[���} & "_b").Height * 0.2
                                    .Shapes.Range(�[���} & "_b").Fill.ForeColor.RGB = RGB(250, 250, 250)
                                    .Shapes.Range(�[���} & "_b").TextFrame2.TextRange.Font.Size = 30
                                    .Shapes.Range(�[���} & "_b").TextFrame2.TextRange.Font.Bold = msoTrue
                                    .Shapes.Range(�[���} & "_b").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                                    .Shapes.Range(�[���} & "_b").TextFrame2.TextRange.Text = �[��
                                    
                                    .Shapes.Range(�[���} & "_b").Line.ForeColor.RGB = RGB(0, 0, 0)
                                    .Shapes.Range(�[���} & "_b").Line.Weight = 1.6
                                    .Shapes.Range(�[���} & "_b").TextFrame2.MarginLeft = 0
                                    .Shapes.Range(�[���} & "_b").TextFrame2.MarginRight = 0
                                    .Shapes.Range(�[���} & "_b").TextFrame2.MarginTop = 0
                                    .Shapes.Range(�[���} & "_b").TextFrame2.MarginBottom = 0
                                    .Shapes.Range(�[���} & "_b").TextFrame2.VerticalAnchor = msoAnchorMiddle
                                    .Shapes.Range(�[���} & "_b").TextFrame2.HorizontalAnchor = msoAnchorNone
                                    .Shapes.Range(�[���} & "_b").TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                                    '.Shapes.Range(�[���} & "_b") = "���Make"
                                    Dim myTagTerminal As Variant: myTagTerminal = �[���} & "_b"
    
                                    '���i�i�Ԃ̕\��
                                    ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 60, 8, Len(partName) * 7.5, 20).Select
                                    'Selection.Text = Left(���i�i�Ԗ�Ran(1).Value, 7)
                                    Selection.Text = partName
                                    Selection.Font.Size = 12
                                    Selection.Font.Bold = msoTrue
                                    'Selection.ShapeRange.Width = 88
                                    Selection.ShapeRange.TextFrame2.MarginLeft = 3
                                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                    Dim myTagProduct As String: myTagProduct = Selection.Name
                                    If Len(�g��������) > 1 And ��n�����i�i�� = "" Then
                                        Selection.Top = 0
                                        '�g�������i�i�Ԃ̕\��
                                        Dim xLeft As Long, yTop As Long, myWidth As Long, myHeight As Long
                                        xLeft = 60.8
                                        yTop = 13
                                        myWidth = 20
                                        myHeight = 11
                                        Dim myLabel() As String: ReDim myLabel(���i�i��RANc) As String
                                        For r = 1 To Len(�g��������)
                                            ActiveSheet.Shapes.AddShape(msoShapeRectangle, xLeft, yTop, myWidth, myHeight).Select
                                            Selection.ShapeRange.Line.Weight = 1
                                            Selection.ShapeRange.TextFrame2.MarginLeft = 2
                                            Selection.ShapeRange.TextFrame2.MarginRight = 0
                                            Selection.ShapeRange.TextFrame2.MarginTop = 0
                                            Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                            Selection.Text = Right(Replace(���i�i��RAN(1, r), " ", ""), 3)
                                            Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                                            'MAX��H�Ȃ�t�H���g��
    '                                        If �[�� = "304" Then Stop
                                            If MAX��H�\�� = "1" Then
                                                bb = InStr(���i��r(p, 0), ���i�i��RAN(1, r))
                                                If bb > 0 Then
                                                    If Mid(���i��r(p, 3), (bb \ 16) + 1, 1) = "1" Then
                                                        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 255
                                                    End If
                                                End If
                                            End If
                                            Selection.Font.Name = myFont
                                            Selection.ShapeRange.Line.ForeColor.RGB = 0
                                            Selection.Font.Bold = msoTrue
                                            Selection.Font.Size = 9
                                            If Mid(�g��������, r, 1) = 1 Then
                                                If ���i�i��RAN(1, r).Interior.color = 16777215 Then
                                                    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 230, 0)
                                                Else
                                                    Selection.ShapeRange.Fill.ForeColor.RGB = ���i�i��RAN(1, r).Interior.color
                                                End If
                                            Else
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(200, 200, 200)
                                                Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                            End If
                                            myLabel(r) = Selection.Name
                                            xLeft = xLeft + myWidth
                                            If r Mod kaiGyo = 0 Then yTop = yTop + myHeight: xLeft = 60.8
                                        Next
                                        '�O���[�v��
                                        For r = 1 To Len(�g��������) - 1
                                            .Shapes.Range(myLabel(r)).Select False
                                        Next r
                                        Selection.ShapeRange.ZOrder msoSendToBack
                                    End If
                                    .Shapes.Range(myTagProduct).Select False
                                    .Shapes.Range(myTagTerminal).Select False
                                    Selection.Group.Select
                                    Selection.Name = �[���} & "_t"
                                    frameWidth2 = Selection.Width
                                    frameHeight2 = Selection.Height
                                End If
                                '�t���[���̒ǉ�
                                If frameWidth1 < frameWidth2 Then frameWidth = frameWidth2 Else frameWidth = frameWidth1
                                ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, frameWidth + 0.3, frameHeight1 + frameHeight2 + 2).Select
                                If frameWidth < frameHeight1 + frameHeight2 Then
                                    cornerSize = cornerSize / frameWidth
                                Else
                                    cornerSize = cornerSize / (frameHeight1 + frameHeight2)
                                End If
                                Selection.ShapeRange.Adjustments.Item(1) = cornerSize
                                Selection.ShapeRange.Line.Weight = 1.6
                                If �[���i���o�[�\�� = False Then
                                    Selection.Border.LineStyle = 0
                                End If
                                On Error Resume Next
                                mycheck = V(9)
                                If Err.Number = 13 Then GoTo line16
                                On Error GoTo 0
                                '��n��-��n�����̕\��_1.991
                                If Left(V(9), 5) <> "Bonda" And Left(V(9), 5) <> "Earth" Then
                                    Selection.Text = ��n��count & " - " & ��n��count
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
                                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
                                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorBottom
                                    Selection.ShapeRange.TextFrame2.MarginRight = 3.5
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 80, 80)
                                    Selection.ShapeRange.ZOrder msoBringToFront
                                End If
line16:
'                                If �T���v���쐬���[�h = False Then
                                    Selection.ShapeRange.Fill.Visible = msoFalse
'                                End If
                                'Selection.ShapeRange.ZOrder msoSendToBack
                                'Selection.ShapeRange.ZOrder msoBringToFront
                                If �[���i���o�[�\�� = True Then
                                    .Shapes.Range(�[���} & "_t").Select False
                                    'Selection.ShapeRange.ZOrder msoSendToBack
                                    Selection.Group.Select
                                End If
                                'Selection.ShapeRange.ZOrder msoSendToBack
                                Selection.Name = �[���} & "_t"
                                'Selection.OnAction = "���Make"
                                .Shapes.Range(�[���}).Top = frameHeight2 + 1.5
                                .Shapes.Range(�[���}).Left = (frameWidth2 - frameWidth1) / 2
                                
                                '������̕\��_1.926
                                Dim yAdd As Long: yAdd = 0
                                If �������i <> 0 Then
                                    yTop = Selection.Top + Selection.Height
                                    ccFlg = False
                                    ���c = 0
                                    ReDim ���\�L(2, 0)
                                    For cc = cav�ꗗrow To cav�ꗗlastrow
                                        If �[�� = Sheets("CAV�ꗗ").Cells(cc, myCol(1)) Then
                                            If partName = Sheets("CAV�ꗗ").Cells(cc, myCol(0)) Then
                                                ccFlg = True
                                                If Sheets("CAV�ꗗ").Cells(cc, myCol(3)) = "" Then GoTo Nextcc
                                                If Sheets("CAV�ꗗ").Cells(cc, myCol(5)) = "" Then
                                                    For cc2 = LBound(���\�L, 2) To UBound(���\�L, 2)
                                                        If ���\�L(0, cc2) = Sheets("CAV�ꗗ").Cells(cc, myCol(3)) Then
                                                            ���\�L(1, cc2) = ���\�L(1, cc2) + 1
                                                            GoTo Nextcc
                                                        End If
                                                    Next cc2
                                                    '�V�K�ǉ�
                                                    ���c = ���c + 1
                                                    ReDim Preserve ���\�L(2, ���c)
                                                    ���\�L(0, ���c) = Sheets("CAV�ꗗ").Cells(cc, myCol(3))
                                                    ���\�L(1, ���c) = 1
                                                    ���\�L(2, ���c) = Sheets("CAV�ꗗ").Cells(cc, myCol(4))
                                                End If
                                            End If
                                        Else
                                            If ccFlg = True Then Exit For
                                        End If
Nextcc:
                                    Next cc
                                    
                                    If ���c > 0 Then
                                        For aa = 1 To ���c
                                            ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 5, yTop + yAdd, frameWidth - 10, 15.7).Select
                                            Selection.ShapeRange.Fill.Visible = msoFalse
                                            Selection.Text = "* " & ���\�L(0, aa) & " �~" & ���\�L(1, aa)
                                            
                                            Selection.Font.Size = 12
                                            Selection.Font.Name = myFont
                                            'Selection.Characters(1, 1).Font.Size = 20
                                            Call �F�ϊ�(���\�L(2, aa), clocode1, clocode2, clofont)
                                            Selection.Characters(1, 1).Font.color = clocode1
                                            Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Line.Visible = True
                                            If �������� = True Then
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.color.RGB = 16777215
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.Radius = 10
                                            Else
                                                Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Line.ForeColor.RGB = 0
                                            End If
                                            Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Line.Weight = 0.1
                                            Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Size = 16
                                            'Selection.Characters(0, 2).Font.Name = "Calibri"
                                            Selection.Font.Bold = True
                                            Selection.ShapeRange.Left = 0
                                            Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                            Selection.ShapeRange.TextFrame2.MarginRight = 0
                                            Selection.ShapeRange.TextFrame2.MarginTop = 0
                                            Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                            Selection.Name = �[���} & "_" & ���\�L(0, aa)
                                            yAdd = yAdd + Selection.Height
                                        Next aa
                                    End If
                                    
                                    �d����� = True 'temp
                                    If �d����� = True Then
                                        Dim �d�����RAN
                                        
                                        �d�����val = SQL_�d�����RANset(�d�����RAN, ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), 1), ActiveWorkbook, �[��)
                                        If �d�����val <> 0 Then
                                        ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 5, yTop + yAdd, frameWidth - 10, 15.7).Select
                                        Selection.ShapeRange.ZOrder msoSendToBack
                                        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                        Selection.Font.Size = 12
                                        Selection.Font.Name = myFont
                                        �d����񒅐F = ""
                                        For rr = LBound(�d�����RAN, 2) To UBound(�d�����RAN, 2)
                                            ����[�� = Left(�d�����RAN(9, rr), InStr(�d�����RAN(9, rr), "_") - 1)
                                            Selection.Text = Selection.Text & vbLf & �d�����RAN(7, rr) & "" & String(5 - Len(�d�����RAN(4, rr)), " ") & �d�����RAN(4, rr) & "mm" & String(4 - Len(����[��), " ") & ����[�� & " "
                                            If �d�����RAN(8, rr) <> "" Then
                                                �d����񒅐F = �d����񒅐F & "," & �d�����RAN(8, rr) & "_" & Len(Selection.Text)
                                                Selection.Text = Selection.Text & "��"
                                            End If
                                        Next rr
                                        Selection.Text = Mid(Selection.Text, 2)
                                        If �d����񒅐F <> "" Then
                                            �d����񒅐Fsp = Split(�d����񒅐F, ",")
                                            For rr = LBound(�d����񒅐Fsp) + 1 To UBound(�d����񒅐Fsp)
                                                �d����񒅐Fspsp = Split(�d����񒅐Fsp(rr), "_")
                                                Call �F�ϊ�(�d����񒅐Fspsp(0), clocode1, clocode2, clofont)
                                                Selection.Characters(Val(�d����񒅐Fspsp(1)), 1).Font.color = clocode1
                                                Selection.ShapeRange.TextFrame2.TextRange.Characters(�d����񒅐Fspsp(1), 1).Font.Line.Visible = True
                                                Selection.ShapeRange.TextFrame2.TextRange.Characters(�d����񒅐Fspsp(1), 1).Font.Size = 13
                                            Next rr
                                        End If
                                       
                                        'Selection.Characters(0, 2).Font.Name = "Calibri"
                                        Selection.Font.Bold = True
                                        Selection.ShapeRange.Left = 0
                                        Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                        Selection.ShapeRange.TextFrame2.MarginRight = 0
                                        Selection.ShapeRange.TextFrame2.MarginTop = 0
                                        Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                        Selection.ShapeRange.Width = 128
                                        Selection.Name = �[���} & "_" & "�d�����"
                                        yAdd = yAdd + Selection.Height
                                        End If
                                    End If
                                    
                                    Dim VO�ꗗ As String: VO�ꗗ = ""
                                    Dim VO�ꗗtemp As Variant
                                    For aa = 0 To ���i���X�gc - 1
                                        If ���i���X�g(5, aa) = �[�� And ���i���X�g(2, aa) = "40" Then
                                            ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 5, yTop + yAdd, frameWidth - 10, 15.7).Select
                                            Selection.Name = �[���} & "_�ڍ�" & aa
                                            Selection.Font.Size = 12
                                            Selection.Font.Bold = True
                                            Selection.Font.Name = myFont
                                            Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                            Selection.ShapeRange.TextFrame2.MarginRight = 0
                                            Selection.ShapeRange.TextFrame2.MarginTop = 0
                                            Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                            Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                                            If ���i���X�g(2, aa) = "40" And ���i���X�g(3, aa) = "T" Then 'VO�̎��ɃC���X�g
                                                Selection.Text = ���i���X�g(1, aa) & vbCrLf & "(" & ���i���X�g(4, aa) & ")"
                                                Selection.Height = 36
                                                Selection.Left = 27
                                                �摜URL_VO = �A�h���X(0) & "\VO�C���X�g2.png"
                                                Set ob = ActiveSheet.Shapes.AddPicture(�摜URL_VO, False, True, 0, yTop + yAdd + 2, 27, 27)
                                                'ob.Name = �[���} & "_VO�C���X�g" & aa
                                                'ActiveSheet.Pictures.Insert(�摜URL_VO).Name = �[���} & "_VO�C���X�g" & aa
'                                                ActiveSheet.Shapes(�[���} & "_VO�C���X�g" & aa).Top = yTop + yAdd + 2
'                                                ActiveSheet.Shapes(�[���} & "_VO�C���X�g" & aa).Left = 0
'                                                ActiveSheet.Shapes(�[���} & "_VO�C���X�g" & aa).ScaleHeight 0.3, msoTrue, msoScaleFromTopLeft
'                                                ActiveSheet.Shapes(�[���} & "_VO�C���X�g" & aa).Select False
                                                ob.Select False
                                                Selection.Group.Select
                                                Set ob = Nothing
                                            ElseIf ���i���X�g(2, aa) = "40" And ���i���X�g(4, aa) = "�X���[�N���b�v" Then
                                                Selection.Text = ���i���X�g(1, aa)
                                                Selection.Height = 17.5
                                                Selection.Left = 27
                                                �摜URL_�X���[ = �A�h���X(0) & "\�X���[�C���X�g.png"
                                                Set ob = ActiveSheet.Shapes.AddPicture(�摜URL_�X���[, False, True, 0, yTop + yAdd + 2, 27, 27)
'                                                ActiveSheet.Pictures.Insert(�摜URL_�X���[).Name = �[���} & "_�X���[�C���X�g" & aa
'                                                ActiveSheet.Shapes(�[���} & "_�X���[�C���X�g" & aa).Top = yTop + yAdd + 2
'                                                ActiveSheet.Shapes(�[���} & "_�X���[�C���X�g" & aa).Left = 0
'                                                ActiveSheet.Shapes(�[���} & "_�X���[�C���X�g" & aa).ScaleHeight 0.2, msoTrue, msoScaleFromTopLeft
'                                                ActiveSheet.Shapes(�[���} & "_�X���[�C���X�g" & aa).Select False
                                                ob.Select False
                                                Selection.Group.Select
                                                Set ob = Nothing
                                            Else
                                                Selection.Text = ���i���X�g(1, aa) & "_" & Left(StrConv(���i���X�g(4, aa), vbNarrow), 12)
                                                Selection.Characters(Len(���i���X�g(1, aa)) + 1, 20).Font.Size = 8
                                                Selection.Height = 17.5
                                                Selection.Left = 14
                                                �摜URL_���̑� = �A�h���X(0) & "\���̑�.png"
                                                Set ob = ActiveSheet.Shapes.AddPicture(�摜URL_���̑�, False, True, 0, yTop + yAdd + 0.5, 13, 13)
'                                                ActiveSheet.Pictures.Insert(�摜URL_���̑�).Name = �[���} & "_���̑�" & aa
'                                                ActiveSheet.Shapes(�[���} & "_���̑�" & aa).Top = yTop + yAdd + 0.5
'                                                ActiveSheet.Shapes(�[���} & "_���̑�" & aa).Left = 0
'                                                ActiveSheet.Shapes(�[���} & "_���̑�" & aa).ScaleHeight 0.15, msoTrue, msoScaleFromTopLeft
'                                                ActiveSheet.Shapes(�[���} & "_���̑�" & aa).Select False
                                                ob.Select False
                                                Selection.Group.Select
                                                Set ob = Nothing
                                            End If
                                            If �������� = True Then
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.color.RGB = 16777215
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.Radius = 10
                                            End If
                                            Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeTextToFitShape
                                            Selection.Name = �[���} & "_v" & aa
                                            VO�ꗗ = VO�ꗗ & Selection.Name & ","
                                            yAdd = yAdd + Selection.Height
                                        End If
                                    Next aa
                                    
                                    If QR��� = True Then
                                        myQR = �[�� & "-"
                                        Call QR�R�[�h���N���b�v�{�[�h�Ɏ擾(myQR)
                                        ActiveSheet.PasteSpecial Format:="�} (JPEG)", Link:=False, DisplayAsIcon:=False
                                        Selection.Height = 40
                                        Selection.Top = 0
                                        Selection.Left = ActiveSheet.Shapes.Range(�[���}).Width + 2
                                        Selection.Name = �[���} & "_qr"
                                    End If
                                
                                    '���^����
                                    With wb(0).Sheets("�[���ꗗ")
                                        Dim ���^���� As String
                                        Set �[��key = .Cells.Find("�[����", , , 1)
                                        ���^Col = .Cells.Find("���^����", , , 1).Column
                                        �[��row = .Columns(�[��key.Column).Cells.Find(�[��, , , 1).Row
                                        ���^���� = .Cells(�[��row, ���^Col)
                                        If ���^���� <> "" Then
                                            �摜URL_seikei = �A�h���X(0) & "\seikei.png"
                                            Set ob = ActiveSheet.Shapes.AddPicture(�摜URL_seikei, False, True, frameWidth - 30, frameHeight2 + frameHeight1 + 1, 30, 27)
                                            ob.ZOrder msoSendToBack
                                            ob.Rotation = CInt(���^����)
'                                            ActiveSheet.Pictures.Insert(�摜URL_seikei).Name = �[���} & "_seikei"
'                                            ActiveSheet.Shapes(�[���} & "_seikei").ZOrder msoSendToBack
'                                            ActiveSheet.Shapes(�[���} & "_seikei").Rotation = CInt(���^����)
'                                            ActiveSheet.Shapes(�[���} & "_seikei").Width = 30
'                                            ActiveSheet.Shapes(�[���} & "_seikei").Top = frameHeight2 + frameHeight1 + 1
'                                            ActiveSheet.Shapes(�[���} & "_seikei").Left = frameWidth - 30
'                                            ActiveSheet.Shapes(�[���} & "_seikei").Select False
                                            ob.Select False
'                                            Selection.Group.Select
                                            Set ob = Nothing
                                        End If
                                    End With
                                    
                                    For aa = 1 To ���c
                                        .Shapes.Range(�[���} & "_" & ���\�L(0, aa)).Select False
                                    Next aa
                                    
                                    If VO�ꗗ <> "" Then
                                        VO�ꗗtemp = Split(Left(VO�ꗗ, Len(VO�ꗗ) - 1), ",")
                                        For Each vo In VO�ꗗtemp
                                            .Shapes.Range(vo).Select False
                                        Next vo
                                    End If
                                End If
                                .Shapes.Range(�[���} & "_t").Select False
                                .Shapes.Range(�[���}).Select False
                                If QR��� = True Then .Shapes.Range(�[���} & "_qr").Select False
                                If �d����� = True And �d�����val > 0 Then .Shapes.Range(�[���} & "_�d�����").Select False
                                Selection.Group.Select
                                Selection.Name = �[���}
                                Selection.Placement = xlMove '�Z���ɍ��킹�Ĉړ��͂��邪�T�C�Y�ύX�͂��Ȃ�
                                '�}�̍Ō�̏���
                                If partName = "" Then
                                    myRowPoint = myRowPoint
                                Else
                                    .Shapes(�[���}).Left = myColPoint '+ (.Shapes(�[���}).Width * ��)
                                    ���� = .Shapes(�[���}).Height
                                    myRowPoint = Rows(�[��firstRow2).Top + (���� * ��)
                                    .Shapes(�[���}).Top = myRowPoint
                                End If
                            �� = �� + 1
                            End If
                        Next p
                        'myRowPoint = (myRowPoint - 1) + (���� * ��)
line17:
                        If myRowSel * myRowHeight < myRowPoint + (����) Then
                            myRowSel = WorksheetFunction.RoundUp((myRowPoint + (����)) / myRowHeight, 0) + 2
                        Else
                            myRowSel = myRowSel + 1
                        End If
line175:
                        minW = -1
                        minH = -1
                        pp = 0
line18:
                    End If
                End If
                If partName = "" And �[�� & "_" & partName <> �[��next & "_" & PartNamenext Then myRowSel = myRowSel + 1:  myRowPoint = myRowSel * myRowHeight
                �[��bak = �[��
                PartNamebak = partName
                'If �i�� <> "" Then myRowSel = myRowSel + 1
            End If
        End With
line20:
    Next i
    Set Target = Nothing
    
    Debug.Print "9= " & Round(Timer - sTime, 2): sTime = Timer
    
    '�^�C�g���𐮂���
    With wb(0).Sheets(newSheetName)
        .Range(.Rows(1), Rows(2)).Insert
        .Range(.Rows(1), Rows(2)).NumberFormat = "@"
        .Cells(4, 1).Activate
        ActiveWindow.FreezePanes = True
        Dim myCount As Long: myCount = 1
        Dim ���i�i�� As String, ���i�^�C�g�� As String, ���i�^�C�g��bak As String
        For X = 1 To ���i�i��RANc
            ���i�i�� = Replace(.Cells(3, X), " ", "")
            Select Case Len(Replace(���i�i��, " ", ""))
                Case 8
                    ���i�^�C�g�� = Left(���i�i��, 4)
                    If ���i�^�C�g�� <> ���i�^�C�g��bak Then .Cells(2, X) = ���i�^�C�g��
                    .Cells(3, X) = Mid(���i�i��, 5, 4)
                    .Columns(X).ColumnWidth = 5.2
                    ���i�^�C�g��bak = ���i�^�C�g��
                Case 10
                    ���i�^�C�g�� = Left(���i�i��, 7)
                    If ���i�^�C�g�� <> ���i�^�C�g��bak Then .Cells(2, X) = ���i�^�C�g��
                    .Cells(3, X) = Mid(���i�i��, 8, 3)
                    .Columns(X).ColumnWidth = 3.9
                    ���i�^�C�g��bak = ���i�^�C�g��
                Case Else
                
            End Select
            .Cells(1, X).Font.Size = 8
            .Cells(1, X) = ���i�i��RAN(8, X)
            .Cells(1, X).NumberFormat = "mm/dd"
        Next X
        '���i�i�Ԃ̔z�u�����l��
        .Range(.Columns(1), .Columns(X - 1)).HorizontalAlignment = xlLeft
        '�񕝂̐ݒ�
        .Columns(X).AutoFit
        .Range(.Columns(X), .Columns(X + 8)).AutoFit
        .Columns(X + 9).ColumnWidth = 6.4
        .Columns(X + 10).ColumnWidth = 3.6
        .Columns(X + 11).ColumnWidth = 3.6
        .Columns(X + 12).ColumnWidth = 11
        .Columns(X + 13).AutoFit
        .Columns(X + 14).ColumnWidth = 4
        '�^�C�g���Ƃ��\��
        If �n���\�� = "0" Then
            �n���\�� = ""
        ElseIf �n���\�� = "1" Then
            �n���\�� = "_��n��"
        ElseIf �n���\�� = "2" Then
            �n���\�� = "_��n��"
        ElseIf �n���\�� = "4" Then
            �n���\�� = "��n���͕\�����Ȃ�"
        End If
        .Cells(2, X).Value = ����G & "_" & �n���}�^�C�v & �n���\��
        
        '.Cells(1, myCount + 1).Value = "Ver" & Sheets("�J��").Cells(Sheets("�J��").Cells(Rows.Count, 2).End(xlUp).Row, 2).Value
        If Left(myBookName, 5) = "���Y����+" Then
            .Cells(, X).Value = Left(myBookName, InStrRev(myBookName, ".") - 1) '& "_Ver" & Mid(myBookName, 7, InStr(myBookName, "_") - 7)
        Else
            Stop '�t�@�C�����ύX����?
        End If
        '����͈͂̐ݒ�
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .Zoom = 100
            .PaperSize = xlPaperA3
            .Orientation = xlLandscape
        End With
        '�}���}��A���̏o�̓V���[�g�J�b�g
        .Cells.Find("�}1", , , 1).AddComment
        .Cells.Find("�}1", , , 1).Comment.Text "Ctrl+ENTER�Ŗ�A�����쐬"
        .Cells.Find("�}1", , , 1).Comment.Shape.TextFrame.AutoSize = True
        .Cells.Find("�}1", , , 1).Comment.Shape.TextFrame.Characters.Font.Size = 11
    End With
    
    Call �}�W�b�N�t���̗���
    Call �œK�����ǂ�
    
    Unload ProgressBar
    
End Sub

Sub ���j���[()
    UserForm2.Show
End Sub

Sub �R�l�N�^���̔�r_PVSW_RLTF���[to���ވꗗ()

    'Call �����[���̃V�[�g�쐬
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim newSheetName As String: newSheetName = "�R�l�N�^����rto���ވꗗ"
    Dim comBookName As String: comBookName = "���ވꗗ�쐬�V�X�e��_Ver1.2.xlsm"
    Dim comSheetName As String: comSheetName = "���ވꗗ"
        
    '���[�N�V�[�g�̒ǉ�
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName

    Stop
    newSheet.Tab.color = False
    
    Dim myCount As Long
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim ���i�i��0Col As Long: ���i�i��0Col = 1
        Dim ���i�i��1Col As Long
        Do
            If Len(.Cells(1, myCount + 1)) = 15 Then
                myCount = myCount + 1
            Else
                Exit Do
            End If
        Loop
        ���i�i��1Col = myCount
        Dim �^�C�g��Ran As Range: Set �^�C�g��Ran = .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.count).End(xlToLeft).Column))
        Dim �[��Col As Long: �[��Col = �^�C�g��Ran.Find("�[�����ʎq").Column
        Dim �[�����i��Col As Long: �[�����i��Col = �^�C�g��Ran.Find("�[�����i��").Column
        Dim �d�����ʖ�Col As Long: �d�����ʖ�Col = �^�C�g��Ran.Find("�d�����ʖ�").Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ���i�i��1Col).End(xlUp).Row
        Dim ���i�L��() As String: ReDim ���i�L��(���i�i��1Col - 1)
        Dim addRow As Long
    End With
    
    For i = 1 To lastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            Dim ���i�i�Ԗ�Ran As Range: Set ���i�i�Ԗ�Ran = .Range(.Cells(1, ���i�i��0Col), .Cells(1, ���i�i��1Col))
            Dim ���i�i��RAN As Range: Set ���i�i��RAN = .Range(.Cells(i, ���i�i��0Col), .Cells(i, ���i�i��1Col))
            Dim �[��  As String: �[�� = .Cells(i, �[��Col)
            Dim �[��Nxt As String: �[��Nxt = .Cells(i + 1, �[��Col)
            Dim �[�����i�� As String: �[�����i�� = .Cells(i, �[�����i��Col)
            Dim �d�����ʖ� As String: �d�����ʖ� = .Cells(i, �d�����ʖ�Col)
        End With
        
        With Workbooks(myBookName).Sheets(newSheetName)
            myCount = 0
            For Each ���i�i�� In ���i�i��RAN
                myCount = myCount + 1
                If ���i�i��.Value = "1" Then
                    ���i�L��(myCount - 1) = "1"
                End If
            Next
            If �[�� <> �[��Nxt Then
                    If i = 1 Then
                        myCount = 0
                        For Each ���i�i�� In ���i�i��RAN
                            myCount = myCount + 1
                            .Cells(1, myCount) = ���i�i��.Value
                        Next
                        .Cells(1, ���i�i��1Col + 1) = "�d�����ʖ�" '�Ō�ɗ�폜
                        .Cells(1, ���i�i��1Col + 2) = "�[�����ʎq"
                        .Cells(1, ���i�i��1Col + 3) = "�[�����i��"
                    Else
                        myCount = 0: addRow = .Cells(.Rows.count, ���i�i��1Col + 1).End(xlUp).Row + 1
                        For Each c In ���i�L��
                            myCount = myCount + 1
                            .Cells(addRow, myCount) = c
                        Next
                        .Cells(addRow, myCount + 1) = �d�����ʖ�
                        .Cells(addRow, myCount + 2) = �[��
                        .Cells(addRow, myCount + 3) = �[�����i��
                        ReDim ���i�L��(���i�i��1Col - 1)
                    End If
            Else
            End If
        End With
        �[��bak = �[��
    Next i
    
    '�[�����i��&���ʂ̈ꗗ�쐬
    With Workbooks(myBookName).Sheets(newSheetName)
        myCount = 0
        For Each ���i�i�Ԗ� In ���i�i�Ԗ�Ran
            myCount = myCount + 1
            .Cells(1, ���i�i��1Col + 5 + myCount) = ���i�i�Ԗ�
        Next
            
        Dim myDic As Object, myKey, myItem: Dim myVal, myVal2, myVal3
        ' ---myDic�փf�[�^���i�[
        myVal = .Range(.Cells(2, ���i�i��1Col + 1), .Cells(addRow, ���i�i��1Col + 3)).Value
        '�����f�[�^��z��Ɋi�[
        For Y = ���i�i��0Col To ���i�i��1Col
            Set myDic = CreateObject("Scripting.Dictionary")
            For i = 1 To UBound(myVal, 1)
                If .Cells(i + 1, Y) = "1" Then
                    myVal2 = myVal(i, 3)
                    If Not myDic.exists(myVal2) Then
                        myDic.add myVal2, 1
                    Else
                        myDic(myVal2) = myDic(myVal2) + 1
                    End If
                Else
                    myVal2 = myVal(i, 3)
                    If Not myDic.exists(myVal2) Then
                        myDic.add myVal2, 0
                    Else
                        myDic(myVal2) = myDic(myVal2) + 0
                    End If
                End If
            Next i
            '��Key,Item�̏����o��
            myKey = myDic.keys
            myItem = myDic.items
                For i = 0 To UBound(myKey)
                    myVal3 = Split(myKey(i), "_")
                    .Cells(i + 2, ���i�i��1Col + 5).Value = myVal3(0)
                    .Cells(i + 2, ���i�i��1Col + Y + 5).Value = myItem(i)
                    If .Cells(i + 2, ���i�i��1Col + Y + 5).Value = 0 Then .Cells(i + 2, ���i�i��1Col + Y + 5).Value = ""
                Next i
            Set myDic = Nothing
        Next Y
    End With
    
    '���ވꗗ�̒l���擾
    For Y = ���i�i��0Col To ���i�i��1Col
        With Workbooks(myBookName).Sheets(newSheetName)
            '�������i�i�ԂŋN�������V����Col��I��
            lastRow = .Cells(.Rows.count, ���i�i��1Col + 5).End(xlUp).Row
            Dim myProduct As String: myProduct = .Cells(1, ���i�i��1Col + 5 + Y)
        End With
        With Workbooks(comBookName).Sheets(comSheetName)
            Dim �V����r As Range
            Dim ����3Col As Long: ����3Col = .Cells.Find("����3_").Column
            Dim keyCell As Range: Set keyCell = .Cells.Find("���i�i��_")
            Dim �N����Ran As Range: Set �N����Ran = .Cells.Find("�N����_")
            Dim �N����new As String: �N����new = ""
            Dim ���i�i��Ran As Range: Set ���i�i��Ran = .Range(.Cells(keyCell.Row + 1, keyCell.Column), .Cells(.Cells.SpecialCells(xlLastCell).Row, keyCell.Column))
            Dim firstFoundCell As Range: Set firstFoundCell = .Range(.Cells(keyCell.Row, 1), .Cells(keyCell.Row, .Columns.count)).Find(Replace(myProduct, " ", ""))
            Set FoundCell = Nothing
            Do
                If FoundCell Is Nothing Then Set FoundCell = firstFoundCell
                Set FoundCell = .Range(.Cells(keyCell.Row, 1), .Cells(keyCell.Row, .Columns.count)).FindNext(FoundCell)
                Dim ���i�i��Col As Long
                �N���� = .Cells(�N����Ran.Row, FoundCell.Column)
                If �N����new = "" Or �N����new < �N���� Then
                    ���i�i��Col = FoundCell.Column
                    �N����new = �N����
                End If
                If firstFoundCell.address = FoundCell.address Then Exit Do
            Loop
        End With
        
        For i = 2 To lastRow
            With Workbooks(myBookName).Sheets(newSheetName)
                Dim myPartName As String: myPartName = .Cells(i, ���i�i��1Col + 5)
                '�i��2���ڂ��A���t�@�x�b�g(�G�A�o�b�N��p)
                Dim flag�ϊ� As Long: flag�ϊ� = 0
                If Mid(myPartName, 2, 1) Like "[A-Z]" Then
                    Select Case Mid(myPartName, 2, 1)
                    Case "A"
                    str2 = 0
                    Case "B"
                    str2 = 1
                    Case "C"
                    str2 = 2
                    Case "D"
                    str2 = 3
                    Case Else
                    Stop
                    End Select
                    myPartName = Left(myPartName, 1) & str2 & Mid(myPartName, 3, 20)
                    flag�ϊ� = 1
                End If
                If flag�ϊ� = 1 Then If .Cells(i, ���i�i��1Col + 5).Comment Is Nothing Then .Cells(i, ���i�i��1Col + 5).AddComment Text:=myPartName & " �Ƃ��Č���"
                If Len(myPartName) = 8 Then
                    myPartName = Left(myPartName, 4) & "-" & Mid(myPartName, 5, 4)
                ElseIf Len(myPartName) = 10 Then
                    myPartName = Left(myPartName, 4) & "-" & Mid(myPartName, 5, 4) & "-" & Mid(myPartName, 9, 2)
                Else
                    Stop
                End If
                Dim my���� As String: my���� = .Cells(i, ���i�i��1Col + 5 + Y)
                Set FoundCell = ���i�i��Ran.Find(myPartName)
                If FoundCell Is Nothing Then
                    .Cells(i, ���i�i��1Col + 5 + ���i�i��1Col + 2) = "NotFound"
                Else
                    Dim com���� As String
                    With Workbooks(comBookName).Sheets(comSheetName)
                        com���� = .Cells(FoundCell.Row, ���i�i��Col).Value
                        ����3 = .Cells(FoundCell.Row, ����3Col).Value
                    End With
                    If my���� <> com���� Then
                        .Cells(i, ���i�i��1Col + 5 + Y) = .Cells(i, ���i�i��1Col + 5 + Y) & "_" & com����
                        .Cells(i, ���i�i��1Col + 5 + Y).Interior.color = RGB(200, 100, 100)
                    End If
                        .Cells(i, ���i�i��1Col + 5 + ���i�i��1Col + 1) = ����3
                End If
            End With
        Next i
    Next Y
End Sub

Function �}�W�b�N�t���̗���()
    Call �œK��
    
    Dim myCol As Long, myRow As Long, myCol2 As Long, i As Long, i2 As Long, i3 As Long, i4 As Long, ii As Long, myCount As Long
    Dim ���iCol As Long, �T�C�YCol As Long, �Fcol As Long, cavCol As Long, �}Col As Long, �}1Col As Long, firstRow As Long, ��col As Long
    Dim �_�u����Col As Long, �\��Col As Long, �FCol2 As Long, �i��col As Long, ���葤col As Long, �[��Col As Long
    Dim ���i As String, �T�C�Y As String, �F As String, cav As String, �} As String, �� As String, �_�u���� As String, ���i�_�� As Long, �[�� As String
    Dim �T�C�Y2 As String, �F2 As String, cav2 As String, �}2 As String, ��2 As String, �_�u����2 As String, ���i�_��2 As Long
    Dim �T�C�Y3 As String, �F3 As String, cav3 As String, �}3 As String, ��3 As String, �_�u����3 As String, ���i�_��3 As Long
    Dim ���ibak As String
    Dim �}�W�b�N��� As String, �}�W�b�Ns As Variant, �}�W�b�Nc As Variant
    
    With Sheets("�ݒ�")
        Set key = .Cells.Find("�}�W���_", , , 1)
        For X = key.Column + 1 To .Cells(key.Row, .Columns.count).End(xlToLeft).Column
            �}�W�b�N��` = �}�W�b�N��` & "_" & key.Offset(0, X)
        Next X
        �}�W�b�N��` = Mid(�}�W�b�N��`, 2)
    End With
        
    With ActiveSheet
        myCol = .Cells.Find("�}", , , xlWhole).Column
        myCol2 = .Cells.Find("�i��", , , 1).Column
        ���iCol = .Cells.Find("�[�����i��", , , 1).Column
        �\��Col = .Cells.Find("�\��", , , 1).Column
        �T�C�YCol = .Cells.Find("�T�C�Y", , , 1).Column
        �Fcol = .Cells.Find("�F�ď�", , , 1).Column
        �[��Col = .Cells.Find("�[����", , , 1).Column
        �FCol2 = .Cells.Find("�F", , , 1).Column
        �i��col = .Cells.Find("�i��", , , 1).Column
        cavCol = .Cells.Find("Cav", , , 1).Column
        �}Col = .Cells.Find("�}", , , 1).Column
        �}1Col = .Cells.Find("�}1", , , 1).Column
        ��col = .Cells.Find("��", , , 1).Column
        �_�u����Col = .Cells.Find("��", , , 1).Column
        ���葤col = .Cells.Find("���葤", , , 1).Column
        myRow = .Cells.Find("�}", , , 1).Row
        
        .Range(.Cells(myRow + 1, �}1Col), .Cells(.Cells(.Rows.count, myCol2).End(xlUp).Row, �}1Col)).Value = .Range(.Cells(myRow + 1, �}Col), .Cells(.Cells(.Rows.count, myCol2).End(xlUp).Row, �}Col)).Value
        .Range(.Cells(myRow + 1, �}1Col), .Cells(.Cells(.Rows.count, myCol2).End(xlUp).Row, �}1Col)).Interior.Pattern = xlNone
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            ���i = .Cells(i, ���iCol)
            '�O���[�v�̐擪�s�擾
            If ���i <> "" And ���ibak = "" Then
                firstRow = i
                ���ibak = ���i
            End If
            '�O���[�v����o��
            If ���i = "" And ���ibak <> "" Then
                For i2 = firstRow To i - 1
                    ���i�_�� = WorksheetFunction.Sum(.Range(.Cells(i2, 1), .Cells(i2, ���iCol - 1)))
                    �T�C�Y = Replace(Replace(.Cells(i2, �T�C�YCol), " ", ""), "F", "")
If �T�C�Y <= 0.5 Then �T�C�Y = 0.5
                    �F = Replace(.Cells(i2, �Fcol), " ", "")
                    cav = .Cells(i2, cavCol)
                    �} = Replace(.Cells(i2, �}1Col), " ", ""): If �} = "" Then �} = "null"
                    �� = Replace(.Cells(i2, ��col), " ", "")
                    �_�u���� = Replace(.Cells(i2, �_�u����Col), " ", "")
                    �[�� = .Cells(i2, �[��Col)
                    '������艺�s�ɓ����������������ׂ�
                    myCount = 1
                    'For i3 = firstRow + myCount To i - 1
                    For i3 = firstRow To i - 1
                        ���i�_��2 = WorksheetFunction.Sum(.Range(.Cells(i3, 1), .Cells(i3, ���iCol - 1)))
                        �T�C�Y2 = Replace(Replace(.Cells(i3, �T�C�YCol), " ", ""), "F", "")
If �T�C�Y2 <= 0.5 Then �T�C�Y2 = 0.5
                        �F2 = Replace(.Cells(i3, �Fcol), " ", "")
                        cav2 = .Cells(i3, cavCol)
                        �}2 = Replace(.Cells(i3, �}1Col), " ", ""): If �}2 = "" Then �}2 = "null"
                        ��2 = Replace(.Cells(i3, ��col), " ", "")
                        �_�u����2 = Replace(.Cells(i3, �_�u����Col), " ", "")
                        If cav = cav2 And �F = �F2 Then
                            '.Cells(i3, �}1Col) = .Cells(i2, �}1Col)
                        Else
                            If �T�C�Y & "_" & �F & "_" & �} = �T�C�Y2 & "_" & �F2 & "_" & �}2 Then
                                '�g���ĂȂ��}�W�b�N�F��T��
                                �}�W�b�N��� = �}�W�b�N��`
                                For i4 = firstRow To i - 1
                                    ���i�_��3 = WorksheetFunction.Sum(.Range(.Cells(i4, 1), .Cells(i4, ���iCol - 1)))
                                    �T�C�Y3 = Replace(Replace(.Cells(i4, �T�C�YCol), " ", ""), "F", "")
If �T�C�Y3 <= 0.5 Then �T�C�Y3 = 0.5
                                    �F3 = Replace(.Cells(i4, �Fcol), " ", "")
                                    cav3 = .Cells(i4, cavCol)
                                    �}3 = Replace(.Cells(i4, �}1Col), " ", ""): If �}3 = "" Then �}3 = "null"
                                    ��3 = Replace(.Cells(i4, ��col), " ", "")
                                    �_�u����3 = Replace(.Cells(i4, �_�u����Col), " ", "")
                                    If cav2 <> cav3 Then
                                        If �T�C�Y & "_" & �F = �T�C�Y3 & "_" & �F3 Then
                                            
                                            '�g�p����Ă���}�W�b�N���폜
                                            �}�W�b�N���s = Split(�}�W�b�N���, "_")
                                            �}�W�b�N��� = ""
                                            For X = LBound(�}�W�b�N���s) To UBound(�}�W�b�N���s)
                                                If �}�W�b�N���s(X) <> �}3 Then
                                                    �}�W�b�N��� = �}�W�b�N��� & "_" & �}�W�b�N���s(X)
                                                End If
                                            Next X
                                            �}�W�b�N��� = Mid(�}�W�b�N���, 2)
'                                            �}�W�b�N��� = Replace(�}�W�b�N���, �}3 & "_", "")
                                        End If
                                    End If
                                Next i4
                                '�}�W�b�N���Ɏc���Ă���F�̍��[���g��
                                �}�W�b�Ns = Split(�}�W�b�N���, "_")
                                For Each �}�W�b�Nc In �}�W�b�Ns
                                    If �}�W�b�Nc <> "" Then Exit For
                                Next �}�W�b�Nc
                                If �}�W�b�Nc = "" Then
                                    If InStr(�}���}�s��, �[��) = 0 Then
                                        �}���}�s�� = �}���}�s�� & "_" & �[��
                                    End If
                                End If
                                
                                If �}�W�b�Nc = "null" Then �}�W�b�Nc = ""
                                If ���i�_�� > ���i�_��2 Then
                                    .Cells(i3, �}1Col).Value = �}�W�b�Nc
                                    '.Cells(i2, �}1Col).Interior.Color = rgbRed
                                Else
                                    .Cells(i2, �}1Col).Value = �}�W�b�Nc
                                    '.Cells(i2, �}1Col).Interior.Color = rgbRed
                                End If
                            End If
                        End If
                        myCount = myCount + 1
                    Next i3
''                    'cav�ňقȂ�}�V�J��t�����ꍇ�͓����}�V�J�ɂ���
                    For ii = firstRow To i - 1
                        If .Cells(ii, cavCol) & .Cells(ii, �Fcol) = .Cells(ii + 1, cavCol) & .Cells(ii + 1, �Fcol) Then
                            .Cells(ii + 1, �}1Col) = .Cells(ii, �}1Col)
                        End If
                    Next ii
                Next i2
                ���ibak = ""
            End If
'
'            For ii = myRow + 1 To .Cells(.Rows.Count, myCol2).End(xlUp).Row
'                If .Cells(ii, �}1Col).Interior.Color <> vbRed Then
'                    .Cells(ii, �}1Col).Value = ""
'                End If
'            Next ii
        Next i
                
        '���_�u���̎��̓}�W�b�N��Ă��Ȃ��̂ō폜
        Dim cavBak As String, ��Bak As String, cavNext As String, ��Next As String, startRow As Long
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            cav = .Cells(i, cavCol).Value
            �� = .Cells(i, ��col).Value
            �_�u���� = .Cells(i, �_�u����Col).Value
            cavNext = .Cells(i + 1, cavCol).Value
            ��Next = .Cells(i + 1, ��col).Value
            If �_�u���� <> "" Then
'                If cav & �� <> cavBak & ��Bak Then startRow = i
'
'                If cav & �� <> cavNext & ��Next Then
'                    For i2 = startRow To i
'                        For i3 = startRow To i
'                            If i2 <> i3 Then
'                                If .Cells(i2, �\��Col) = Left(.Cells(i3, �_�u����Col), 4) Then
                                    .Cells(i, �}1Col) = .Cells(i, �}Col)
                                    .Cells(i, �}1Col).Interior.Pattern = xlNone
'                                 End If
'                             End If
'                         Next i3
'                     Next i2
'                 End If
                cavBak = cav
                ��Bak = ��
            End If
        Next i
        '�d���Z���F�̉��Ɍr��������
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            If .Cells(i, cavCol).Value <> .Cells(i + 1, cavCol).Value Then
                .Cells(i, �FCol2).Borders(xlEdgeBottom).LineStyle = xlContinuous
            End If
        Next i
        
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            '��Ă����}�W�b�N�̉ӏ���ԐF�h��Ԃ�
            If .Cells(i, �}Col).Value <> .Cells(i, �}1Col).Value Then
                .Cells(i, �}1Col).Interior.color = rgbRed
            Else
                '.Cells(i, �}1Col) = ""
            End If
            '����cav�Ȃ̂ɓd���F���قȂ�̂Ńn�C���C�g
            If .Cells(i, �_�u����Col).Value = "" Then
                If .Cells(i, cavCol).Value = .Cells(i + 1, cavCol).Value And .Cells(i, ���葤col).Value = .Cells(i + 1, ���葤col).Value Then
                    If .Cells(i, �Fcol).Value <> .Cells(i + 1, �Fcol).Value Then
                        .Cells(i, �Fcol).Interior.color = rgbRed
                        .Cells(i + 1, �Fcol).Interior.color = rgbRed
                    End If
                End If
            End If
            '����cav�Ȃ̂ɓd���T�C�Y���قȂ�̂Ńn�C���C�g
            If .Cells(i, �_�u����Col).Value = "" Then
                If .Cells(i, cavCol).Value = .Cells(i + 1, cavCol).Value And .Cells(i, ���葤col).Value = .Cells(i + 1, ���葤col).Value Then
                    If .Cells(i, �T�C�YCol).Value <> .Cells(i + 1, �T�C�YCol).Value Then
                        .Cells(i, �T�C�YCol).Interior.color = rgbRed
                        .Cells(i + 1, �T�C�YCol).Interior.color = rgbRed
                    End If
                End If
            End If
            '����cav�Ȃ̂ɓd���i�킪�قȂ�̂Ńn�C���C�g
            If .Cells(i, �_�u����Col).Value = "" Then
                If .Cells(i, cavCol).Value = .Cells(i + 1, cavCol).Value And .Cells(i, ���葤col).Value = .Cells(i + 1, ���葤col).Value Then
                    If .Cells(i, �i��col).Value <> .Cells(i + 1, �i��col).Value Then
                        .Cells(i, �i��col).Interior.color = rgbRed
                        .Cells(i + 1, �i��col).Interior.color = rgbRed
                    End If
                End If
            End If
        Next i
        
    End With
    Call �œK�����ǂ�
End Function

Function testColor()
    With Range("o10").Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 45
        .Gradient.ColorStops.Clear
        .Gradient.ColorStops.add(0).color = rgbRed
        .Gradient.ColorStops.add(0.4).color = rgbRed
        .Gradient.ColorStops.add(0.401).color = rgbBlue
        .Gradient.ColorStops.add(0.599).color = rgbBlue
        .Gradient.ColorStops.add(0.6).color = rgbRed
        .Gradient.ColorStops.add(1).color = rgbRed
    End With
End Function

Function PVSWcsv���[����ݒ�2�ɒ[���ꗗ��n��()
    With ActiveWorkbook.Sheets("PVSW_RLTF���[")
        Dim �d������Col As Long: �d������Col = .Cells.Find("�d�����ʖ�").Column
        Dim �d������Row As Long: �d������Row = .Cells.Find("�d�����ʖ�").Row
        Dim �[�����ʎqCol As Long: �[�����ʎqCol = .Cells.Find("�[�����ʎq").Column
        Dim ��햼��Col As Long: ��햼��Col = .Cells.Find("��햼��").Column
        Dim �[�����i��Col As Long: �[�����i��Col = .Cells.Find("�[�����i��").Column
        Dim PVSWtoNMBCol As Long: PVSWtoNMBCol = .Cells.Find("PVSWtoNMB_").Column
        Dim �V�[���hCol As Long: �V�[���hCol = .Cells.Find("�V�[���h").Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d������Col).End(xlUp).Row
        Dim i As Long
        Dim myDic As Object, myKey, myItem
        Dim myVal, myVal2, myVal3
        Set myDic = CreateObject("Scripting.dictionary")
        Dim maxCol As Long: maxCol = WorksheetFunction.Max(�[�����ʎqCol, ��햼��Col, �[�����i��Col, �V�[���hCol, PVSWtoNMBCol)
        myVal = .Range(.Cells(�d������Row + 1, 1), .Cells(lastRow, maxCol))
        For i = 1 To UBound(myVal, 1)
            If myVal(i, �V�[���hCol) <> "S" Or myVal(i, PVSWtoNMBCol) = "found" Then
                myVal2 = myVal(i, ��햼��Col) & "," & myVal(i, �[�����i��Col) & "," & myVal(i, �[�����ʎqCol)
                If Replace(myVal2, ",", "") <> "" Then
                    If Not myDic.exists(myVal2) Then
                        myDic.add myVal2, 1
                    End If
                End If
            End If
        Next i
    End With
    With ActiveWorkbook.Sheets("�ݒ�2")
        Dim out��햼��Row As Long: out��햼��Row = .Cells.Find("��햼��").Row
        Dim out��햼��Col As Long: out��햼��Col = .Cells.Find("��햼��").Column
        Dim out�[�����i��Col As Long: out�[�����i��Col = .Cells.Find("���i�i��").Column
        Dim out�[�����ʎqCol As Long: out�[�����ʎqCol = .Cells.Find("�[��").Column
        Dim out�T�uCol As Long: out�T�uCol = .Cells.Find("�T�u��").Column
        .Range(.Cells(out��햼��Row + 1, out��햼��Col), .Cells(.Rows.count, out��햼��Col)) = ""
        .Range(.Cells(out��햼��Row + 1, out�[�����i��Col), .Cells(.Rows.count, out�[�����i��Col)) = ""
        .Range(.Cells(out��햼��Row + 1, out�[�����ʎqCol), .Cells(.Rows.count, out�[�����ʎqCol)) = ""
        .Range(.Cells(out��햼��Row + 1, out�T�uCol), .Cells(.Rows.count, out�T�uCol)) = ""
        myKey = myDic.keys
        myItem = myDic.items
        For i = 0 To UBound(myKey)
            myVal3 = Split(myKey(i), ",")
            .Cells(out��햼��Row + 1 + i, out��햼��Col) = myVal3(0)
            .Cells(out��햼��Row + 1 + i, out�[�����i��Col) = myVal3(1)
            .Cells(out��햼��Row + 1 + i, out�[�����ʎqCol) = myVal3(2)
        Next i
    End With
End Function
Public Function �n���}�̈���p�f�[�^�쐬(Optional �p���T�C�Y���� As String, Optional newBookName)

    Dim i As Long, myLeft As Single, myTop As Single
    Dim ���� As Long
    Dim �p�� As String
    Dim ���� As String
    ���� = 1
    
    newBookName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & "_" & newBookName
    
    Dim ����^�C�g�� As String
    ����^�C�g�� = ActiveSheet.Name
    Dim �N����
    Set �N���� = Range("a1")
    
    Dim �v�����g�T�C�Y
    Dim �v�����g�z�E�R�E
    
    �p���T�C�Y����s = Split(�p���T�C�Y����, "-")
    �p�� = �p���T�C�Y����s(0)
    ���� = �p���T�C�Y����s(1)
    
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim newSheetName As String: newSheetName = mySheetName & "p"
    
    Dim maxHeight As Long, maxWidth As Long
    Dim breakHeight As Long: breakHeight = 0
    Select Case �p��
        Case "A4"
            �v�����g�T�C�Y = xlPaperA4
            If ���� = "��" Then
                �v�����g�z�E�R�E = xlLandscape
                maxHeight = 623
                maxWidth = 880
            Else
                �v�����g�z�E�R�E = xlPortrait
                maxHeight = 880
                maxWidth = 623
            End If
        Case "A3"
            �v�����g�T�C�Y = xlPaperA3
            If ���� = "��" Then
                �v�����g�z�E�R�E = xlLandscape
                maxHeight = 880
                maxWidth = 1246
            Else
                �v�����g�z�E�R�E = xlPortrait
                maxHeight = 1246
                maxWidth = 880
            End If
        Case Else
            MsgBox "����T�C�Y���Ή����Ă��܂���"
            Exit Function
    End Select
    
    Call �œK��
    Dim objShp As Shape
    '���[�N�u�b�N�쐬
    myBookpath = ActiveWorkbook.Path
    
    '�o�͐�f�B���N�g����������΍쐬
    If Dir(myBookpath & "\50_��n���}", vbDirectory) = "" Then
        MkDir myBookpath & "\50_��n���}"
    End If
    
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(myBookpath & "\50_��n���}\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    
    Workbooks.add
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=myBookpath & "\50_��n���}\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
        
    ActiveSheet.Name = newSheetName
    
    '����͈͂̐ݒ�
    With ActiveSheet
        .Range("a1").NumberFormat = "@"
        .Range("a1") = ����^�C�g��
        .Range("a2") = "�N����_" & CStr(�N����.Value)
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0.9)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .Zoom = 100
            .PaperSize = �v�����g�T�C�Y
            .Orientation = �v�����g�z�E�R�E
        End With
    End With
    
    myTop = 27
    For Each objShp In myBook.Sheets(mySheetName).Shapes
        If objShp.Type = 4 Then GoTo nextOBJSHP
        objShp.Copy
        myLeft = 3
        For i = 1 To ����
            DoEvents
            Sleep 10
            DoEvents
            Sheets(newSheetName).Paste
            Selection.Left = myLeft
            Selection.Top = myTop
            myLeft = myLeft + Selection.Width + 3
        Next i
        If myTop + Selection.Height - breakHeight > maxHeight Then
            Sheets(newSheetName).HPageBreaks.add before:=Cells(RoundUp((myTop - 2) / 13.5, 0), 1)
            breakHeight = Cells(RoundUp((myTop - 2) / Rows(1).Height, 0), 1).Top
        End If
        myTop = myTop + Selection.Height + 12
nextOBJSHP:
    Next objShp
    Call �œK�����ǂ�

End Function

Public Sub ���������V�X�e���p�f�[�^�쐬v2182(Optional CB6)
    If IsMissing(CB6) Then CB6 = "8216658233390"

    CB6 = Replace(CB6, " ", "")
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim myTestURL As String
    Set ws(0) = myBook.ActiveSheet
    
    dirName = "\70_�ėp���������V�X�e��point\"
    
    Call �œK��
    Call Init2
    Dim objShp As Shape
    '���[�N�u�b�N�쐬
    myBookpath = ActiveWorkbook.Path
    '�o�͐�f�B���N�g����������΍쐬
    If Dir(myBookpath & dirName, vbDirectory) = "" Then
        MkDir myBookpath & dirName
    End If
    '�o�͐�f�B���N�g����������΍쐬_���i�i��
    If Dir(myBookpath & dirName & CB6, vbDirectory) = "" Then
        MkDir myBookpath & dirName & CB6
        'FileCopy �A�h���X(0) & "\�ėp���������V�X�e��\myBlink.js", myBookpath & dirName & CB6 & "\myBlink.js"
    End If
    '�o�͐�f�B���N�g����������΍쐬_���i�i��\img
    If Dir(myBookpath & dirName & CB6 & "\img", vbDirectory) = "" Then
        MkDir myBookpath & dirName & CB6 & "\img"
    End If
    '�o�͐�f�B���N�g����������΍쐬_���i�i��\css
    If Dir(myBookpath & dirName & CB6 & "\css", vbDirectory) = "" Then
        MkDir myBookpath & dirName & CB6 & "\css"
    End If
    
    With myBook.Sheets("PVSW_RLTF���[")
        'html�̏o��
        Set myKey = .Cells.Find("�|�C���g1_", , , 1)
        Dim ����col(6) As Long
        ����col(0) = myKey.Column
        ����col(1) = .Cells.Find("�\��_", , , 1).Column
        ����col(2) = .Cells.Find("�F��", , , 1).Column
        ����col(3) = 1 '�T�u
        ����col(4) = .Cells.Find("�[�����ʎq", , , 1).Column
        ����col(5) = .Cells.Find("�n��", , , 1).Column
        ����col(6) = .Cells.Find("�L���r�e�B", , , 1).Column
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            �\�� = .Cells(Y, ����col(1))
            �F�� = .Cells(Y, ����col(2))
            �T�u = .Cells(Y, ����col(3))
            �T�u = Replace(�T�u, "*", "") '�T�u�i���o�[����*������
            point = .Cells(Y, ����col(0))
            �[�� = .Cells(Y, ����col(4))
            ��ƍH�� = .Cells(Y, ����col(5))
            cav = .Cells(Y, ����col(6))
            'html�쐬
            myPath = myBookpath & dirName & CB6
            myTestURL = TEXT�o��_�ėp���������V�X�e��html(myPath, �\��, �F��, �T�u, point, �[��, ��ƍH��, cav)
        Next Y
        
        '�[��cav�̈ꗗ�Z�b�g
        Dim cssran() As String, myCount As Long
        ReDim cssran(8, 0) As String
        For Y = myKey.Row + 1 To lastRow
            point = .Cells(Y, ����col(0))
            �[�� = .Cells(Y, ����col(4))
            cav = .Cells(Y, ����col(6))
            �F�� = .Cells(Y, ����col(2))
            If point <> "" Then
                ReDim Preserve cssran(8, myCount)
                cssran(0, myCount) = point
                cssran(1, myCount) = �[��
                cssran(2, myCount) = cav
                cssran(3, myCount) = �F��
                myCount = myCount + 1
            End If
        Next Y
    End With
    
    Dim �[��temp As Object
    �o�͍ςݒ[�� = ""
    For Y = LBound(cssran, 2) To UBound(cssran, 2)
        �[�� = cssran(1, Y)
        cav = cssran(2, Y)
        '�[��.png�̏o��
        If InStr(�o�͍ςݒ[��, "_" & �[�� & "_") = 0 Then
            If Not (�[��temp Is Nothing) Then �[��temp.Delete
            '�[���摜�̔{�������߂�
            Set objShp = myBook.Sheets(mySheetName).Shapes(�[�� & "_1")
            objShp.Copy
            ws(0).Paste
            Set �[��temp = Selection.ShapeRange
            �[��temp.Top = 0
            �[��temp.Left = 0
            ActiveWindow.Zoom = 100
            '�T�C�Y���s�N�Z���Ŏw��
            ��lx = 1280
            ��ly = 700
            �䗦xy = ��lx / ��ly
            myW = �[��temp.Width
            myH = �[��temp.Height
            If myW > myH * �䗦xy Then �{�� = ��lx / myW Else �{�� = ��ly / myH
            �{�� = �{�� / 96 * 72 '�|�C���g���s�N�Z���ɕϊ�
             '�[���̏o��
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Set cht = ActiveSheet.ChartObjects.add(0, 0, �[��temp.Width * �{��, �[��temp.Height * �{��).Chart
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            cht.Export fileName:=myBookpath & dirName & CB6 & "\img\" & �[�� & ".png", filtername:="PNG"
            cht.Parent.Delete
            �o�͍ςݒ[�� = �o�͍ςݒ[�� & "_" & �[�� & "_"
        End If
   
        '�[��cav.png�̏o��
        For Each obj In �[��temp.GroupItems
            If obj.Name = �[�� & "_1_" & cav Then
                obj.Copy
                Sleep 10
                ws(0).Paste
                Selection.Left = obj.Left
                Selection.Top = obj.Top
                '�_�ŗp�ɃI�[�g�V�F�C�v��ύX
                Selection.ShapeRange.Fill.Visible = msoTrue
                Selection.ShapeRange.Fill.Transparency = 0
                Selection.ShapeRange.Fill.Solid
                tempcolor = Selection.ShapeRange.Fill.ForeColor
                Selection.ShapeRange.Fill.ForeColor.RGB = tempcolor
                Selection.ShapeRange.Line.Visible = False
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                Selection.ShapeRange.Glow.color.RGB = tempcolor
                Selection.ShapeRange.Glow.Transparency = 0
                Selection.ShapeRange.Glow.Radius = 13
                
                With ws(0).Shapes.AddShape(1, 0, 0, �[��temp.Width, �[��temp.Height)
                    .Left = 0
                    .Top = 0
                    .Fill.Visible = msoFalse
                    .Line.Visible = msoFalse
                    .Select False
                End With
                Selection.Group.Name = "Cavtemp"
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                Set cht = ActiveSheet.ChartObjects.add(0, 0, objShp.Width * �{��, objShp.Height * �{��).Chart
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                DoEvents '�x���Ȃ邩��
                Sleep 10
                DoEvents
                cht.Paste
                cht.Export fileName:=myBookpath & dirName & CB6 & "\img\" & �[�� & "_1_" & cav & ".png", filtername:="PNG"
                cht.Parent.Delete
                ws(0).Shapes("Cavtemp").Delete
                Exit For
            End If
        Next obj
nextY:
    Next Y
        
    'css
    Dim box2l As Single, box2t As Single, box2w As Single, box2h As Single
    With myBook.Sheets(mySheetName)
        Set myKey = .Cells.Find("Cav", , , 1)
        Col1 = myKey.Column
        col2 = .Cells.Find("�[����", , , 1).Column
        col3 = .Cells.Find("Point", , , 1).Column
        '���W�̊������擾
        For Y = LBound(cssran, 2) To UBound(cssran, 2)
            �[�� = cssran(1, Y)
            
            Set objshp1 = .Shapes(�[�� & "_1")
            
            On Error Resume Next
            Set objShp2 = .Shapes(�[�� & "_1_" & cav)
            If Err.Number = 438 Or Err.Number = -2147024809 Then  '�Ώۂ�Cav��Shapes�������ꍇ
                If cav <> 1 Then
                    Set objShp2 = .Shapes(�[�� & "_1_" & 1)
                Else
                    'Stop '���m�F_bonda�Ƃ�
                End If
            End If
            On Error GoTo 0
            
            box2l = (objShp2.Left - objshp1.Left) / objshp1.Width
            box2t = (objShp2.Top - objshp1.Top) / objshp1.Height
            box2w = objShp2.Width / objshp1.Width
            box2h = objShp2.Height / objshp1.Height
            
            cssran(4, Y) = box2l
            cssran(5, Y) = box2t
            cssran(6, Y) = box2w
            cssran(7, Y) = box2h
            '�d���R�[�h
            �F�� = cssran(3, Y)
            If InStr(�F��, "/") > 0 Then �F�� = Left(�F��, InStr(�F��, "/") - 1)
            If cssran(3, Y) = "" Then
                clocode1 = "EEEEEE" '��|�C���g
                clofont = "000000"
            Else
                Call �F�ϊ�css(�F��, clocode1, clocode2, clofont)
line20:
            End If
        Next Y
        'css�ɏo��
        For Y = LBound(cssran, 2) To UBound(cssran, 2)
            myPath = myBookpath & dirName & CB6 & "\css" & "\wh" & Format(cssran(0, Y), "0000") & ".css"
            �F�� = cssran(3, Y)
            If InStr(�F��, "/") > 0 Then �F�� = Left(�F��, InStr(�F��, "/") - 1)
            If cssran(3, Y) = "" Then
                clocode1 = "EEEEEE" '��|�C���g
                clofont = "000000"
            Else
                Call �F�ϊ�css(�F��, clocode1, clocode2, clofont)
            End If
            Call TEXT�o��_�ėp���������V�X�e��css(myPath, clocode1, clofont)
        Next Y
    End With
    
    'myBlink�쐬
    myPath = myBookpath & dirName & CB6 & "\myBlink.js"
    Call TEXT�o��_�ėp���������V�X�e��js(myPath)
    Call �œK�����ǂ�
    
    Shell "EXPLORER.EXE  " & myTestURL
    ActiveWindow.Zoom = 100
End Sub


Public Function ���������V�X�e���p�f�[�^�쐬_�|�C���g��(Optional CB6 As String)
    CB6 = Replace(CB6, " ", "")
'    Dim i As Long, myLeft As Single, myTop As Single
'    Dim ���� As Long
'    Dim �p�� As String
'    Dim ���� As String
'    ���� = 1
'
'    Dim ����^�C�g�� As String
'    ����^�C�g�� = ActiveSheet.Name
'    Dim �N����
'    Set �N���� = Range("a1")
'
'    Dim �v�����g�T�C�Y
'    Dim �v�����g�z�E�R�E
'
'    �p���T�C�Y����s = Split(�p���T�C�Y����, "-")
'    �p�� = �p���T�C�Y����s(0)
'    ���� = �p���T�C�Y����s(1)
'
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
'    Dim newSheetName As String: newSheetName = mySheetName & "p"
'
'    Dim maxHeight As Long, maxWidth As Long
'    Dim breakHeight As Long: breakHeight = 0
'    Select Case �p��
'        Case "A4"
'            �v�����g�T�C�Y = xlPaperA4
'            If ���� = "��" Then
'                �v�����g�z�E�R�E = xlLandscape
'                maxHeight = 623
'                maxWidth = 880
'            Else
'                �v�����g�z�E�R�E = xlPortrait
'                maxHeight = 880
'                maxWidth = 623
'            End If
'        Case "A3"
'            �v�����g�T�C�Y = xlPaperA3
'            If ���� = "��" Then
'                �v�����g�z�E�R�E = xlLandscape
'                maxHeight = 880
'                maxWidth = 1246
'            Else
'                �v�����g�z�E�R�E = xlPortrait
'                maxHeight = 1246
'                maxWidth = 880
'            End If
'        Case Else
'            MsgBox "����T�C�Y���Ή����Ă��܂���"
'            Exit Function
'    End Select
    
    Call �œK��
    Dim objShp As Shape
    '���[�N�u�b�N�쐬
    myBookpath = ActiveWorkbook.Path
    
    '�o�͐�f�B���N�g����������΍쐬
    If Dir(myBookpath & "\80_�ėp���������V�X�e���ppoint", vbDirectory) = "" Then
        MkDir myBookpath & "\80_�ėp���������V�X�e���ppoint"
    End If
    '�o�͐�f�B���N�g����������΍쐬_���i�i��
    If Dir(myBookpath & "\80_�ėp���������V�X�e���ppoint\" & CB6, vbDirectory) = "" Then
        MkDir myBookpath & "\80_�ėp���������V�X�e���ppoint\" & CB6
    End If
    '�o�͐�f�B���N�g����������΍쐬_���i�i��\img
    If Dir(myBookpath & "\80_�ėp���������V�X�e���ppoint\" & CB6 & "\img", vbDirectory) = "" Then
        MkDir myBookpath & "\80_�ėp���������V�X�e���ppoint\" & CB6 & "\img"
    End If
    
    With myBook.Sheets(mySheetName)
        Set myKey = .Cells.Find("Cav", , , 1)
        Col1 = myKey.Column
        col2 = .Cells.Find("�[����", , , 1).Column
        col3 = .Cells.Find("Point", , , 1).Column
        lastRow = .Cells(.Rows.count, col3).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            �|�C���g = Format(.Cells(Y, col3), "0000")
            �[�� = .Cells(Y, col2)
            cav = .Cells(Y, Col1)
            If �|�C���g <> "" Then
                Set objShp = .Shapes(�[�� & "_1_" & cav)
                '�n�C���C�g
                objShp.SoftEdge.Radius = 1
                With objShp.Glow
                    .color.RGB = RGB(250, 5, 5)
                    .Transparency = 0.15
                    .Radius = 5
                End With
                Set objshp1 = .Shapes(�[�� & "_1")
                �[�� = Left(objshp1.Name, InStr(objshp1.Name, "_") - 1)
                '�I��͈͂��擾
                'Set rg = Selection
                 '�I�������͈͂��摜�`���ŃR�s�[
                objshp1.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '�摜�\��t���p�̖��ߍ��݃O���t���쐬
                Set cht = ActiveSheet.ChartObjects.add(0, 0, objshp1.Width, objshp1.Height).Chart
                 '���ߍ��݃O���t�ɓ\��t����
                cht.Paste
                 'JPEG�`���ŕۑ�
                'Selection
                '�T�C�Y����
                ActiveWindow.Zoom = 100
                ��l = 434
                myW = Selection.Width
                myH = Selection.Height
                If myW > myH Then
                    �{�� = ��l / myW
                Else
                    �{�� = ��l / myH
                End If
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
                'ActiveSheet.Shapes("�O���t 1").ScaleHeight 0.87, msoFalse, msoScaleFromTopLeft
                'Selection.ShapeRange.Width = 444
                'cht.ScaleWidth 444, msoFalse, msoScaleFromTopLeft
                'Debug.Print �[��, Selection.Width, Selection.Height
                cht.Export fileName:=myBookpath & "\80_�ėp���������V�X�e���ppoint\" & CB6 & "\img\" & �|�C���g & ".jpg", filtername:="JPG"
                 '���ߍ��݃O���t���폜
                cht.Parent.Delete
                '�n�C���C�g�����ɖ߂�
                objShp.SoftEdge.Radius = 0
                With objShp.Glow
                    .Radius = 0
                End With
            End If
        Next Y
    End With
    
    With myBook.Sheets("PVSW_RLTF���[")
        Set myKey = .Cells.Find("�|�C���g1_", , , 1)
        Dim ����col(5) As Long
        ����col(0) = myKey.Column
        ����col(1) = .Cells.Find("�\��_", , , 1).Column
        ����col(2) = .Cells.Find("�F��", , , 1).Column
        ����col(3) = 1 '�T�u
        ����col(4) = .Cells.Find("�[��", , , 1).Column
        ����col(5) = .Cells.Find("�n��", , , 1).Column
        lastRow = .Cells(.Rows.count, ����col(1)).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            �\�� = .Cells(Y, ����col(1))
            �F�� = .Cells(Y, ����col(2))
            �T�u = Replace(.Cells(Y, ����col(3)), "*", "")
            point = .Cells(Y, ����col(0))
            �[�� = .Cells(Y, ����col(4))
            ��ƍH�� = .Cells(Y, ����col(5))
            'html�쐬
            myPath = myBookpath & "\80_�ėp���������V�X�e���ppoint\" & CB6
            Call TEXT�o��_�ėp���������V�X�e��(myPath, �\��, �F��, �T�u, point, �[��, ��ƍH��)
        Next Y
    End With
    
    Call �œK�����ǂ�
    ActiveWindow.Zoom = 100
End Function

Public Function �|�C���g�ꗗ�̃V�[�g�쐬_2190()

    Dim sTime As Single: sTime = Timer
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�|�C���g�ꗗ"
    Call �A�h���X�Z�b�g(myBook)
    my�� = 20
        
    With Workbooks(myBookName).Sheets(mySheetName)
        'PVSW_RLTF����̃f�[�^
        Dim my�^�C�g��Row As Long: my�^�C�g��Row = .Cells.Find("�i��_").Row
        Dim my�^�C�g��Col As Long: my�^�C�g��Col = .Cells.Find("�i��_").Column
        Dim my�^�C�g��Ran As Range: Set my�^�C�g��Ran = .Range(.Cells(my�^�C�g��Row, 1), .Cells(my�^�C�g��Row, my�^�C�g��Col))
        Dim my�d�����ʖ�Col As Long: my�d�����ʖ�Col = .Cells.Find("�d�����ʖ�").Column
        Dim my��1Col As Long: my��1Col = .Cells.Find("�n�_����H����").Column
        Dim my�[��1Col As Long: my�[��1Col = .Cells.Find("�n�_���[�����ʎq").Column
        Dim myCav1Col As Long: myCav1Col = .Cells.Find("�n�_���L���r�e�B").Column
        Dim my��2Col As Long: my��2Col = .Cells.Find("�I�_����H����").Column
        Dim my�[��2Col As Long: my�[��2Col = .Cells.Find("�I�_���[�����ʎq").Column
        Dim myCav2Col As Long: myCav2Col = .Cells.Find("�I�_���L���r�e�B").Column
        Dim my����Col As Long: my����Col = .Cells.Find("����No").Column
        Dim my�����i��Col As Long: my�����i��Col = .Cells.Find("�����i��").Column
        Dim myJoint1Col As Long: myJoint1Col = .Cells.Find("�n�_��JOINT���").Column
        Dim myJoint2Col As Long: myJoint2Col = .Cells.Find("�I�_��JOINT���").Column
        Dim my�_�u����1Col As Long: my�_�u����1Col = .Cells.Find("�n�_���_�u����H����").Column
        Dim my�_�u����2Col As Long: my�_�u����2Col = .Cells.Find("�I�_���_�u����H����").Column
        
        Dim myPVSW�i��col As Long: myPVSW�i��col = .Cells.Find("�d���i��").Column
        Dim myPVSW�T�C�Ycol As Long: myPVSW�T�C�Ycol = .Cells.Find("�d���T�C�Y").Column
        Dim myPVSW�Fcol As Long: myPVSW�Fcol = .Cells.Find("�d���F").Column
        Dim my�}���}11Col As Long: my�}���}11Col = .Cells.Find("�n�_���}���}�F�P").Column
        Dim my�}���}12Col As Long: my�}���}12Col = .Cells.Find("�n�_���}���}�F�Q").Column
        Dim my�}���}21Col As Long: my�}���}21Col = .Cells.Find("�I�_���}���}�F�P").Column
        Dim my�}���}22Col As Long: my�}���}22Col = .Cells.Find("�I�_���}���}�F�Q").Column
        
        Dim my���i11Col As Long: my���i11Col = .Cells.Find("�n�_���[�q�i��").Column
        Dim my���i21Col As Long: my���i21Col = .Cells.Find("�I�_���[�q�i��").Column
        Dim my���i12Col As Long: my���i12Col = .Cells.Find("�n�_���S����i��").Column
        Dim my���i22Col As Long: my���i22Col = .Cells.Find("�I�_���S����i��").Column
        Dim my���1Col As Long: my���1Col = .Cells.Find("�n�_����햼��").Column
        Dim my���2Col As Long: my���2Col = .Cells.Find("�I�_����햼��").Column
        Dim my���Ӑ�1Col As Long: my���Ӑ�1Col = .Cells.Find("�n�_���[�����Ӑ�i��").Column
        Dim my���1Col As Long: my���1Col = .Cells.Find("�n�_���[�����i��").Column
        Dim my���Ӑ�2Col As Long: my���Ӑ�2Col = .Cells.Find("�I�_���[�����Ӑ�i��").Column
        Dim my���2Col As Long: my���2Col = .Cells.Find("�I�_���[�����i��").Column
        Dim myJointGCol As Long: myJointGCol = .Cells.Find("�W���C���g�O���[�v").Column
        Dim myAB�敪Col As Long: myAB�敪Col = .Cells.Find("A/B�EB/C�敪").Column
        Dim my�d��YBMCol As Long: my�d��YBMCol = .Cells.Find("�d���x�a�l").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(my�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Set my�^�C�g��Ran = Nothing
        'RLTF����̃f�[�^
        Dim my�i��Col As Long: my�i��Col = .Cells.Find("�i��_").Column
        Dim my�T�C�YCol As Long: my�T�C�YCol = .Cells.Find("�T�C�Y_").Column
        Dim my�T�C�Y��Col As Long: my�T�C�Y��Col = .Cells.Find("�T��_").Column
        Dim my�FCol As Long: my�FCol = .Cells.Find("�F_").Column
        Dim my�F��Col As Long: my�F��Col = .Cells.Find("�F��_").Column
        Dim my����Col As Long: my����Col = .Cells.Find("�ؒf��_").Column
        Dim myPVSWtoNMB As Long: myPVSWtoNMB = .Cells.Find("RLTFtoPVSW_").Column
        
        Dim my���i�i��Ran0 As Long, my���i�i��Ran1 As Long, X As Long
        For X = 1 To myLastCol
            If Len(.Cells(my�^�C�g��Row, X)) = 15 Then
                If my���i�i��Ran0 = 0 Then my���i�i��Ran0 = X
            Else
                If my���i�i��Ran0 <> 0 Then my���i�i��Ran1 = X - 1: Exit For
            End If
        Next X
        
        'Dictionary
        Dim myDic As Object, myKey, myItem
        Dim myVal, myVal2, myVal3
        Set myDic = CreateObject("Scripting.Dictionary")
        myVal = .Range(.Cells(1, 1), .Cells(myLastRow, myLastCol))
    End With
    
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
    If newSheet.Name = "�|�C���g�ꗗ" Then
        newSheet.Tab.color = 14470546
    End If
    
line11:
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile �A�h���X(0) & "\onKey\004_CodeModule_�|�C���g�ꗗ.txt"
    If Err.Number <> 0 Then GoTo line11
    On Error GoTo 0
    
    'PVSW_RLTF to �|�C���g�ꗗ
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant, �|�C���g�ꗗRAN As Variant
    For i = my�^�C�g��Row To myLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            Set ���i�i��RAN = .Range(.Cells(i, my���i�i��Ran0), .Cells(i, my���i�i��Ran1))
            Dim �d�����ʖ� As String: �d�����ʖ� = .Cells(i, my�d�����ʖ�Col)
            Dim ��1 As String: ��1 = .Cells(i, my��1Col)
            Dim �[��(1) As String
            �[��(0) = .Cells(i, my�[��1Col)
            �[��(1) = .Cells(i, my�[��2Col)
            Dim cav(1) As String
            cav(0) = .Cells(i, myCav1Col)
            cav(1) = .Cells(i, myCav2Col)
            Dim ��2 As String: ��2 = .Cells(i, my��2Col)
            Dim �[��2 As String: �[��2 = .Cells(i, my�[��2Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim �����i�� As Range: Set �����i�� = .Cells(i, my�����i��Col)
            Dim �V�[���h�t���O As String: If �����i��.Interior.color = 9868950 Then �V�[���h�t���O = "S" Else �V�[���h�t���O = ""
            Dim Joint1 As String: Joint1 = .Cells(i, myJoint1Col)
            Dim Joint2 As String: Joint2 = .Cells(i, myJoint2Col)
            Dim �_�u����1 As String: �_�u����1 = .Cells(i, my�_�u����1Col)
            Dim �_�u����2 As String: �_�u����2 = .Cells(i, my�_�u����2Col)
            Dim ���i11 As String: ���i11 = .Cells(i, my���i11Col)
            Dim ���i21 As String: ���i21 = .Cells(i, my���i21Col)
            Dim ���i12 As String: ���i12 = .Cells(i, my���i12Col)
            Dim ���i22 As String: ���i22 = .Cells(i, my���i22Col)
            Dim ���1 As String: ���1 = .Cells(i, my���1Col)
            Dim ���2 As String: ���2 = .Cells(i, my���2Col)
            Dim ���Ӑ�1 As String: ���Ӑ�1 = .Cells(i, my���Ӑ�1Col)
            Dim ���(1) As String
            ���(0) = .Cells(i, my���1Col)
            ���(1) = .Cells(i, my���2Col)
            Dim ���Ӑ�2 As String: ���Ӑ�2 = .Cells(i, my���Ӑ�2Col)
            Dim JointG As String: JointG = .Cells(i, myJointGCol)
            Dim �d���i�� As String: �d���i�� = .Cells(i, myPVSW�i��col)
            Dim �d���T�C�Y As String: �d���T�C�Y = .Cells(i, myPVSW�T�C�Ycol)
            Dim �d���F As String: �d���F = .Cells(i, myPVSW�Fcol)
            Dim �}���}11 As String: �}���}11 = .Cells(i, my�}���}11Col)
            Dim �}���}12 As String: �}���}12 = .Cells(i, my�}���}12Col)
            Dim �}���}21 As String: �}���}21 = .Cells(i, my�}���}21Col)
            Dim �}���}22 As String: �}���}22 = .Cells(i, my�}���}22Col)
            Dim AB�敪 As String: AB�敪 = .Cells(i, myAB�敪Col)
            Dim �d��YBM As String: �d��YBM = .Cells(i, my�d��YBMCol)
            
            Dim ���葤1 As String, ���葤2 As String
            If Len(cav2) < 4 Then ���葤1 = �[��2 & "_" & String(3 - Len(cav2), " ") & cav2 & "_" & ��2
            If Len(Cav1) < 4 Then ���葤2 = �[��1 & "_" & String(3 - Len(Cav1), " ") & Cav1 & "_" & ��1
            'NMB����̃f�[�^
            Dim �i�� As String: �i�� = .Cells(i, my�i��Col)
            Dim �T�C�Y As String: �T�C�Y = .Cells(i, my�T�C�YCol)
            Dim �T�C�Y�� As String: �T�C�Y�� = .Cells(i, my�T�C�Y��Col)
            Dim �F As String: �F = .Cells(i, my�FCol)
            Dim �F�� As String: �F�� = .Cells(i, my�F��Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim PVSWtoNMB As String: PVSWtoNMB = .Cells(i, myPVSWtoNMB)
        End With
        
        With Workbooks(myBookName).Sheets(newSheetName)
            Dim �D��1 As Long, �D��2 As Long, �D��3 As Long
            If .Cells(1, 1) = "" Then
                Dim addCol As Long, ���i�i�� As Variant
                Dim addRow As Long: addRow = .Cells(.Rows.count, addCol + 2).End(xlUp).Row + 1
                For Each ���i�i�� In ���i�i��RAN
                    addCol = addCol + 1
                    .Cells(1, addCol) = ���i�i��
                Next
                .Cells(1, addCol + 1) = "�[�����i��": Columns(addCol + 1).NumberFormat = "@": �D��2 = addCol + 1
                .Cells(1, addCol + 2) = "�[����": Columns(addCol + 2).NumberFormat = 0: �D��1 = addCol + 2
                .Cells(1, addCol + 3) = "Cav": Columns(addCol + 3).NumberFormat = 0: �D��3 = addCol + 3
                .Cells(1, addCol + 4) = "LED": Columns(addCol + 4).NumberFormat = 0
                .Cells(1, addCol + 5) = "�|�C���g1": Columns(addCol + 5).NumberFormat = 0: .Cells(1, addCol + 5).Interior.color = RGB(255, 255, 0)
                .Cells(1, addCol + 6) = "�|�C���g2": Columns(addCol + 6).NumberFormat = 0
                .Cells(1, addCol + 7) = "FUSE": Columns(addCol + 7).NumberFormat = 0
                .Cells(1, addCol + 8) = "��d�W�~": Columns(addCol + 8).NumberFormat = 0: .Cells(1, addCol + 8).Interior.color = RGB(255, 255, 0)
                .Cells(1, addCol + 9) = "�ȈՃ|�C���g": Columns(addCol + 9).NumberFormat = 0
                .Cells(1, addCol + 10) = "���}_�\�ʎ�": Columns(addCol + 10).NumberFormat = "@"
            Else
                For r = 0 To 1
                    '�o�^�̊m�F
                    For Y = 2 To addRow
                        If .Cells(Y, addCol + 1) = ���(r) Then
                            If .Cells(Y, addCol + 2) = �[��(r) Then
                                If .Cells(Y, addCol + 3) = cav(r) Then
                                    For X = my���i�i��Ran0 To my���i�i��Ran1
                                        �l = ���i�i��RAN(X)
                                        If �l <> "" Then �l = 1 Else �l = 0
                                        �lb = .Cells(Y, 0 + X)
                                        .Cells(Y, 0 + X) = �l Or �lb
                                    Next X
                                    GoTo line30
                                End If
                            End If
                        End If
                    Next Y
                    '�V�K�o�^
                    addRow = .Cells(.Rows.count, addCol + 2).End(xlUp).Row + 1
                    .Cells(addRow, addCol + 1) = ���(r)
                    .Cells(addRow, addCol + 2) = �[��(r)
                    .Cells(addRow, addCol + 3) = cav(r)
                    For X = my���i�i��Ran0 To my���i�i��Ran1
                        �l = ���i�i��RAN(X)
                        If �l <> "" Then �l = 1 Else �l = 0
                        �lb = .Cells(addRow, 0 + X)
                        .Cells(addRow, 0 + X) = �l Or �lb
                    Next X
line30:
                Next r
            End If
        End With
    Next i
    
    '���בւ�
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    
    With Workbooks(myBookName).Sheets(newSheetName)
        Dim ���i�i�� As String
        Dim ���i�i��bak As String
        Dim ���i�i��next As String
        Dim �[��bak As String
        Dim �[��next As String
        Dim cav�� As String
        Dim iCav As Long
        Dim startRow As Long
        Dim endRow As Long
        myLastRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row
        For i = 2 To myLastRow
            ���i�i�� = .Cells(i, addCol + 1)
            ���i�i��next = .Cells(i + 1, addCol + 1)
            �[��(0) = .Cells(i, addCol + 2)
            �[��next = .Cells(i + 1, addCol + 2)
            cav(0) = .Cells(i, addCol + 3)
            If ���i�i�� <> ���i�i��bak Or �[��(0) <> �[��bak Then startRow = i

            If ���i�i�� <> ���i�i��next Or �[��(0) <> �[��next Then
                'Cav���𒲂ׂ�
                cav�� = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(���i�i��), "�R�l�N�^�ɐ�_")
                If cav�� = "" Then cav�� = 1 '�A�[�X�[�q�̏ꍇ
                For iCav = 1 To CLng(cav��)
                    For i2 = startRow To i
                        If iCav = .Cells(i2, addCol + 3) Then GoTo line20
                    Next i2
                    addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row + 1
                    .Cells(addRow, addCol + 1) = ���i�i��
                    .Cells(addRow, addCol + 2) = �[��(0)
                    .Cells(addRow, addCol + 3) = iCav
line20:
                Next iCav
            End If
            ���i�i��bak = ���i�i��
            �[��bak = �[��(0)
        Next i
    End With
    '���בւ�
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(2), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(2, 1).Select
        ActiveWindow.FreezePanes = True
        '�}�g���N�X��0���u�����N�ɂ���
        .Range(.Cells(2, 1), .Cells(addRow, my���i�i��Ran1)).Replace "0", ""
        .Rows(1).Insert
        For X = my���i�i��Ran0 To my���i�i��Ran1
            .Cells(1, X) = Mid(.Cells(2, X), 8, 3)
        Next X
        .Range(.Columns(1), .Columns(addCol + 8)).AutoFit
        .Range(.Cells(2, 1), .Cells(addRow, my���i�i��Ran1)).ColumnWidth = 3.2
    End With
lineTemp:
    '���}_�\�ʎ��̒ǉ�_2.189.93
    With Workbooks(myBookName).Sheets(newSheetName)
        .Activate
        Dim �[�����i��str As String, �[�����i��strNext As String
        Dim �[��str As String, �[��strNext As String
        Dim �z��() As String
        Set mykey1 = .Cells.Find("�[�����i��", , , 1)
        Set mykey2 = .Cells.Find("�[����", , , 1)
        Set mykey3 = .Cells.Find("Cav", , , 1)
        Set mykey4 = .Cells.Find("�|�C���g1", , , 1)
        Set mykey5 = .Cells.Find("��d�W�~", , , 1)
        addRow = mykey1.End(xlDown).Row
        Dim ryakuCol As Long: ryakuCol = .Cells.Find("���}_�\�ʎ�", , , 1).Column
        Dim topRow As Long: topRow = mykey1.Row + 1
'        For i = mykey1.Row + 1 To addRow + 1
'            �[�����i��str = .Cells(i, mykey1.Column)
'            �[�����i��str = �[�����i�ԕϊ�(�[�����i��str)
'            �[��str = .Cells(i, mykey2.Column)
'
'            �[�����i��strNext = .Cells(i + 1, mykey1.Column)
'            �[�����i��strNext = �[�����i�ԕϊ�(�[�����i��strNext)
'            �[��strNext = .Cells(i + 1, mykey2.Column)
'            If �[�����i��str <> �[�����i��strNext Or �[��str <> �[��strNext Then
'                ReDim �z��(7, 0)
'                myCount = 0
'                For y = topRow To i
'                    addc = UBound(�z��, 2) + 1
'                    ReDim Preserve �z��(7, addc)
'                    �z��(0, addc) = .Cells(y, mykey3.Column)
'                    �z��(1, addc) = .Cells(y, mykey4.Column)
'                    �z��(2, addc) = .Cells(y, mykey5.Column)
'                Next y
'                Set �摜��v = �|�C���g�i���o�[�}�쐬(�[�����i��str, �[��str, �z��, i)
'                �摜��v.Select
'                �摜��v.Left = .Cells(topRow, ryakuCol).Left
'                �摜��v.Top = .Cells(topRow, ryakuCol).Top
'                myHeight = Rows(i + 1).Top - Rows(topRow).Top
'                �摜��v.Height = myHeight
'                topRow = i + 1
'            End If
'        Next i
    End With
    
    �|�C���g�ꗗ�̃V�[�g�쐬_2190 = Round(Timer - sTime, 2)
    
End Function

Public Function �S���̃Z���G���^�[()
    Dim lastRow As Long: lastRow = ActiveSheet.UsedRange.Rows.count
    Dim lastCol As Long: lastCol = ActiveSheet.UsedRange.Columns.count
    Dim startRow As Long: startRow = 1
    Call �œK��
    For X = 1 To lastCol
        For Y = startRow To lastRow
            Cells(Y, X).Value = Cells(Y, X).Value
        Next Y
    Next X
    Call �œK�����ǂ�
End Function

Public Function �[���ʐ����ꗗ�쐬()

    ���type = "C"
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim newSheetName As String: newSheetName = "�����ꗗ_" & ���type
    
    'PVSW_RLTF���R�s�[���ă��l�[��
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = "PVSW_RLTF_temp" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    Workbooks(myBookName).Sheets("PVSW_RLTF").Copy after:=Sheets("PVSW_RLTF")
    ActiveSheet.Name = "PVSW_RLTF_temp"
    Call PVSWcsv�̋��ʉ�_Ver1944_�����ύX
    
    Call ���i�i��RAN_set2(���i�i��RAN, ���type, "����", "")
    Call SQL_�[���ꗗ_2(���i�i��RAN, �d���ꗗRAN, �[���ꗗran, myBookName)
    
    '�����m�F�p�̃V�[�g�ǉ�
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    Worksheets.add after:=Worksheets("PVSW_RLTF_temp")
    ActiveSheet.Cells.NumberFormat = "@"
    ActiveSheet.Name = newSheetName
    Dim �[�� As String
    For i = LBound(�[���ꗗran) To UBound(�[���ꗗran)
        With Workbooks(myBookName).Sheets(newSheetName)
            If i = LBound(�[���ꗗran) Then
                .Cells(1, 1) = "�[���ʐ����ꗗ_" & ���type
                .Cells(2, 1) = "�\��"
                .Cells(2, 2) = "�[��"
                .Cells(2, 3) = "��"
                .Cells(2, 4) = "cav"
                .Cells(2, 5) = "����_"
                .Cells(2, 6) = "������_"
                .Cells(2, 7) = "�[��"
                .Cells(2, 8) = "��"
                .Cells(2, 9) = "cav"
                .Cells(2, 10) = "���l"
                For X = 1 To ���i�i��RANc
                        .Cells(2, 10 + X) = ���i�i��RAN(1, X - 1)
                        .Cells(1, 10 + X) = Mid(���i�i��RAN(1, X - 1), 8, 3)
                    Next X
                addRow = 3
            End If
            If IsNull(�[���ꗗran(i)) Then GoTo line20
            �[�� = �[���ꗗran(i)
            For k = LBound(�d���ꗗRAN, 2) To UBound(�d���ꗗRAN, 2)
                '�n�_
                If �[�� = �d���ꗗRAN(���i�i��RANc + 3, k) Then
                    .Cells(addRow, 1) = �d���ꗗRAN(���i�i��RANc + 0, k)
                    .Cells(addRow, 2) = �d���ꗗRAN(���i�i��RANc + 3, k)
                    .Cells(addRow, 3) = �d���ꗗRAN(���i�i��RANc + 1, k)
                    .Cells(addRow, 4) = �d���ꗗRAN(���i�i��RANc + 5, k)
                    .Cells(addRow, 5) = �d���ꗗRAN(���i�i��RANc + 7, k)
                    .Cells(addRow, 6) = �d���ꗗRAN(���i�i��RANc + 8, k)
                    .Cells(addRow, 7) = �d���ꗗRAN(���i�i��RANc + 4, k)
                    .Cells(addRow, 8) = �d���ꗗRAN(���i�i��RANc + 2, k)
                    .Cells(addRow, 9) = �d���ꗗRAN(���i�i��RANc + 6, k)
                    .Cells(addRow, 10) = �d���ꗗRAN(���i�i��RANc + 9, k)
                    For X = 1 To ���i�i��RANc
                        .Cells(addRow, 10 + X) = �d���ꗗRAN(X - 1, k)
                    Next X
                    addRow = addRow + 1
                End If
                '�n�_
                If �[�� = �d���ꗗRAN(���i�i��RANc + 4, k) Then
                    .Cells(addRow, 1) = �d���ꗗRAN(���i�i��RANc + 0, k)
                    .Cells(addRow, 2) = �d���ꗗRAN(���i�i��RANc + 4, k)
                    .Cells(addRow, 3) = �d���ꗗRAN(���i�i��RANc + 2, k)
                    .Cells(addRow, 4) = �d���ꗗRAN(���i�i��RANc + 6, k)
                    .Cells(addRow, 5) = �d���ꗗRAN(���i�i��RANc + 7, k)
                    .Cells(addRow, 6) = �d���ꗗRAN(���i�i��RANc + 8, k)
                    .Cells(addRow, 7) = �d���ꗗRAN(���i�i��RANc + 3, k)
                    .Cells(addRow, 8) = �d���ꗗRAN(���i�i��RANc + 1, k)
                    .Cells(addRow, 9) = �d���ꗗRAN(���i�i��RANc + 5, k)
                    .Cells(addRow, 10) = �d���ꗗRAN(���i�i��RANc + 9, k)
                    For X = 1 To ���i�i��RANc
                        .Cells(addRow, 10 + X) = �d���ꗗRAN(X - 1, k)
                    Next X
                    addRow = addRow + 1
                End If
            Next k
        End With
line20:
    Next i
    
    '���בւ�
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 7).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(3), Rows(addRow - 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    
End Function

Public Function �z���}�쐬one(���i�i��RAN, ��n���摜Sheet)

    Call �œK��
    Call �A�h���X�Z�b�g(myBook)

    ��� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), 1)
    ���� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), 1)
    ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), 1)
    
    Set myBook = ActiveWorkbook
    Dim newBookName As String: newBookName = Left(myBook.Name, InStr(myBook.Name, "_")) & "�z���}_" & Replace(���i�i��str, " ", "")
        
    'Call Ver181_PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬
    'Call �n���}�쐬_Ver2001
    
    '�g�p����T�u�ꗗ���쐬
    Call SQL_�z���T�u�擾(�z���T�uRAN, ���i�i��str)
    
    If Dir(myBook.Path & "\55_�z���}", vbDirectory) = "" Then
        MkDir (myBook.Path & "\55_�z���}")
    End If
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(myBook.Path & "\55_�z���}\" & newBookName & "_" & Format(i, "000") & ".xlsm") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsm"
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    '������ǂݎ���p�ŊJ��
    baseBookName = "����_�z���}.xlsm"
    On Error Resume Next
    Workbooks.Open fileName:=�A�h���X(0) & "\genshi\" & baseBookName, ReadOnly:=True
    If Err = 1004 Then
        MsgBox "System+ �̃A�h���X��������܂���B�V�[�g[�ݒ�]���������Ă��������B"
        End
    End If
    On Error GoTo 0
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBook.Path & "\55_�z���}\" & newBookName
    Set wb(1) = ActiveWorkbook
    Application.DisplayAlerts = True
    On Error GoTo 0
    Call Init
    
    With Workbooks(newBookName)
        For i = LBound(�z���T�uRAN, 2) To UBound(�z���T�uRAN, 2)
            Dim �T�u As String
            �T�u = �z���T�uRAN(0, i)
            .Sheets("genshi").Copy before:=Sheets("genshi")
            ActiveSheet.Name = �T�u
            With .Sheets(CStr(�T�u))
                Workbooks(myBook.Name).Activate
                Call �œK��
                Call �z���}�쐬(���i�i��str, �T�u, 0, ���, ��n���摜Sheet)
                Call �œK�����ǂ�
                ActiveSheet.Shapes.SelectAll
                'If Selection.count <= 1 Then Stop
                'Selection.Group.Select
                Selection.Cut
                .Activate
                .Paste
                Selection.Left = 3
                Selection.Top = 65
                Selection.Ungroup
                .Range("aa2") = Replace(���i�i��str, " ", "")
                .Range("ad2") = �T�u
                .Range("a2") = ���
                
                Dim ���i�i��HeaderBak As String
                Y = 5: X = 0: ���iHeaderBak = ""

                .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBook.Name, 6, 5)
                '.PageSetup.RightHeader = "&R" & "&14 " & ���i�i��str & "&14 �z���}-" & "&14 " & �T�u & "   " & "&P/&N"
                .Cells(1, 1).Select
            End With
        Next i
        Set ws(1) = Worksheets.add(before:=Worksheets("base"))
        ws(1).Name = "�\��-SUB"
        ws(1).Cells.NumberFormat = "@"

        Call SQL_�z���}�p_���i�i��_�\��_SUB(RAN, ���i�i��str, myBook)

        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            For X = LBound(RAN) To UBound(RAN)
                ws(1).Cells(Y, X + 1) = Replace(RAN(X, Y), " ", "")
            Next X
        Next Y
        Application.DisplayAlerts = False
        wb(1).Save
        Application.DisplayAlerts = True
    End With
    Call �œK�����ǂ�
End Function

Public Function �z���}�쐬one3(Optional ���i�i��RAN, Optional ��n���摜Sheet)

    Call �œK��
    Call �A�h���X�Z�b�g(myBook)
    
    ��� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), 1)
    ���� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), 1)
    ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), 1)
    ��zstr = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "��z"), 1)
    
    Set myBook = ActiveWorkbook
    Dim newBookName As String: newBookName = Left(myBook.Name, InStr(myBook.Name, "_")) & "�z���}_" & Replace(���i�i��str, " ", "") & "_" & ��zstr
    Dim footSize
    'Call Ver181_PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬
    'Call �n���}�쐬_Ver2001
    
    '�g�p����T�u�ꗗ���쐬
    Call SQL_�z���T�u�擾(�z���T�uRAN, ���i�i��str)
    
    If Dir(myBook.Path & "\56_�z���}_�U��", vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_�z���}_�U��")
    End If
    If Dir(myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr, vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr)
    End If
    If Dir(myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img", vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img")
    End If
    If Dir(myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css", vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css")
    End If
    
    Call ��n���}�Ăяo���pQR����f�[�^�쐬(���)
    If �z���}�쐬temp = 0 Then
        Call �U�����j�^�̈ړ��f�[�^�쐬_��n���}csv(���i�i��str, ��zstr, ���)
        Call �U�����j�^�̈ړ��f�[�^�쐬_�\��_�\���̒��Scsv(���i�i��str, ��zstr, ���)
        Call �U�����j�^�̈ړ��f�[�^�쐬_�\��_�T�u�̒��Scsv(���i�i��str, ��zstr, ���)
    End If
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(myBook.Path & "\56_�z���}_�U��\" & newBookName & "_" & Format(i, "000") & ".xlsm") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsm"
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    '������ǂݎ���p�ŊJ��
    baseBookName = "����_�z���}.xlsm"
    On Error Resume Next
    Workbooks.Open fileName:=�A�h���X(0) & "\genshi\" & baseBookName, ReadOnly:=True
    If Err = 1004 Then
        MsgBox "System+ �̃A�h���X��������܂���B�V�[�g[�ݒ�]���������Ă��������B"
        End
    End If
    On Error GoTo 0
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBook.Path & "\56_�z���}_�U��\" & newBookName
    Set wb(1) = ActiveWorkbook
    Set ws(2) = myBook.Sheets("���_" & ���)
    Application.DisplayAlerts = True
    On Error GoTo 0
    'index�̏o��
    FileCopy �A�h���X(0) & "\�z���U��\index.html", myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\index.html"
    FileCopy �A�h���X(0) & "\�z���U��\css\index.css", myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css" & "\index.css"
    FileCopy �A�h���X(0) & "\�z���U��\img\index.png", myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img" & "\index.png"
    'change�̏o��
    FileCopy �A�h���X(0) & "\�z���U��\change.html", myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\change.html"
    FileCopy �A�h���X(0) & "\�z���U��\css\change.css", myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css" & "\change.css"
    FileCopy �A�h���X(0) & "\�z���U��\img\change.png", myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img" & "\change.png"
    
    mypath0 = myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\myBlink.js"
    Call TEXT�o��_�z���o�H_�[��js(mypath0)
    mypath0 = myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\myBlink2.js"
    Call TEXT�o��_�z���o�H_�[��js2(mypath0)
    If �z���}�쐬temp = 1 Then GoTo line77
'   �o�H�z��
    ReDim �z���T�usize(2, 0)
    With Workbooks(newBookName)
        For i = LBound(�z���T�uRAN, 2) To UBound(�z���T�uRAN, 2)
            Dim �T�u As String
            �T�u = �z���T�uRAN(0, i)
            .Activate
            .Sheets("genshi").Copy before:=Sheets("genshi")
            wb(1).ActiveSheet.Name = �T�u
            With .Sheets(CStr(�T�u))
                'WS(2).Activate
                footSize = �z���}�쐬3(���i�i��str, ��zstr, �T�u, 0, ���, ��n���摜Sheet)
                Call �œK��
                ws(2).Activate
                ActiveWindow.ScrollColumn = 1
                ActiveWindow.ScrollRow = 1
                ws(2).Shapes.SelectAll
                Selection.Group.Name = "����"
                ws(2).Shapes("����").Select
                
                ReDim Preserve �z���T�usize(2, UBound(�z���T�usize, 2) + 1)
                mybasewidth = Selection.Width
                mybaseheight = Selection.Height
'                �z���T�usize(0, UBound(�z���T�usize, 2)) = �T�u
'                �z���T�usize(1, UBound(�z���T�usize, 2)) = mybasewidth
'                �z���T�usize(2, UBound(�z���T�usize, 2)) = mybaseheight
                
                '�o��
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '�摜�\��t���p�̖��ߍ��݃O���t���쐬
                Set cht = ws(2).ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
                 '���ߍ��݃O���t�ɓ\��t����
                DoEvents
                Sleep 10
                DoEvents
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                '�T�C�Y����
                ActiveWindow.Zoom = 100
                '��l = 1000
                �{�� = 1
                ws(2).Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
                ws(2).Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
                '
                cht.Export fileName:=wb(0).Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �T�u & ".png", filtername:="PNG"
                
                 '���ߍ��݃O���t���폜
                ws(2).Activate
                cht.Parent.Delete
                ws(2).Shapes.SelectAll
                Selection.Cut 'Copy Workbooks(newBookName).Sheets(CStr(�T�u)).Cells(1, 1)
                wb(1).Activate
                wb(1).Sheets(CStr(�T�u)).Activate
                .Cells(1, 1).Activate
                DoEvents
                Sleep 10
                DoEvents
                .Paste
                Selection.Left = 3
                Selection.Top = 65
                'Selection.ShapeRange.Ungroup
                .Range("aa2") = Replace(���i�i��str, " ", "")
                .Range("ad2") = �T�u
                .Range("a2") = ���
                
                Dim ���i�i��HeaderBak As String
                Y = 5: X = 0: ���iHeaderBak = ""

                .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBook.Name, 6, 5) & "_" & ��zstr
                '.PageSetup.RightHeader = "&R" & "&14 " & ���i�i��str & "&14 �z���}-" & "&14 " & �T�u & "   " & "&P/&N"
                .Cells(1, 1).Select
            End With
nextii:
        Next i
'       ���[���o�H
        ws(2).Activate
        'Base
        Call �z���}�쐬3(���i�i��str, ��zstr, "Base", 0, ���, ��n���摜Sheet)
        Call �œK��
        '�[������left�̒l�Z�b�g
        Dim �[��leftRAN() As String
        ReDim �[��leftRAN(1, 0)
        For i = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
            ReDim Preserve �[��leftRAN(1, UBound(�[��leftRAN, 2) + 1)
            �[��leftRAN(0, UBound(�[��leftRAN, 2)) = �[���ꗗran(1, i)
            �[��leftRAN(1, UBound(�[��leftRAN, 2)) = ws(0).Shapes(�[���ꗗran(1, i)).Left
        Next i
        Workbooks(myBook.Name).Activate
        ReDim Preserve �z���T�usize(2, UBound(�z���T�usize, 2) + 1)
        'ActiveSheet.Shapes("��a").Ungroup
        ActiveSheet.Shapes("���").Ungroup
        ActiveSheet.Shapes.SelectAll
        Selection.Group.Name = "����"
        ActiveSheet.Shapes("����").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height
        Selection.Ungroup

        Call SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, myBook.Name)
        
        With Workbooks(newBookName)
            For Y = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
                �[��str = �[���ꗗran(1, Y)
                Call SQL_�z��_�[���o�H�擾(�[���o�HRAN, ���i�i��str, �[��str)
                For i = LBound(�[���o�HRAN, 2) To UBound(�[���o�HRAN, 2)
                    �[��from = �[���o�HRAN(0, i)
                    �[��to = �[���o�HRAN(1, i)

                    Set �[��from = Nothing: Set �[��to = Nothing
                    If �[���o�HRAN(0, i) <> "" Then Set �[��from = ws(2).Cells.Find(�[���o�HRAN(0, i), , , 1)
                    If �[���o�HRAN(1, i) <> "" Then Set �[��to = ws(2).Cells.Find(�[���o�HRAN(1, i), , , 1)
                    On Error Resume Next
                    If �[��from = "" Then Set �[��from = Nothing
                    If �[��to = "" Then Set �[��to = Nothing
                    On Error GoTo 0
                    If Not (�[��from Is Nothing) Then ws(2).Shapes(�[���o�HRAN(0, i)).Select False
                    If Not (�[��to Is Nothing) Then ws(2).Shapes(�[���o�HRAN(1, i)).Select False
                    
                    If �[��from Is Nothing And �[��to Is Nothing Then GoTo nextI '���[����Nothing�Ȃ珈�����Ȃ�
                    If Not (�[��from Is Nothing) And Not (�[��to Is Nothing) Then '�ǂ��炩�̒[����Nothing�Ȃ�I�����Ȃ�
                        If �[��from <> �[��to Then '�[���������Ȃ�I�����Ȃ�
                            '���z������[���Ԃ̃��C���ɐF�t��
                            If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                                
                            Set �[��1 = �[��from
                            Set �[��2 = Nothing
                            Do Until �[��1.Row = �[��to.Row
line10:
                                '-X�����ɓ���
                                Do Until �[��1.Column = 1
                                    Set �[��2 = �[��1.Offset(0, -2)
                                    On Error Resume Next
                                        ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                        ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                    On Error GoTo 0
                                   
                                    Set �[��1 = �[��2
                                    If Left(�[��1.Value, 1) = "U" Then ws(2).Shapes(�[��1.Value).Select False
                                    If �[��1 = �[��1.Offset(myStep, 0) Then Exit Do '��܂��͉��������[���Ȃ�Y�����֓���
                                Loop
                                'Y�����ɓ���
                                Do Until �[��1.Row = �[��to.Row
                                    Set �[��2 = �[��1.Offset(myStep, 0)
                                    If �[��1 <> �[��2 Then
                                        On Error Resume Next
                                            ws(2).Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                            ws(2).Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                        On Error GoTo 0
                                    End If
                                    Set �[��1 = �[��2
                                    If Left(�[��1.Value, 1) = "U" Then ws(2).Shapes(�[��1.Value).Select False
                                    If �[��1.Offset(myStep, 0) = "" Then GoTo line10 '�i�ސ悪�󗓂Ȃ�X�����ړ��ɖ߂�
                                    If Left(�[��1.Offset(myStep, 0), 1) <> "U" Then GoTo line10 '�i�ސ悪U����Ȃ����X�����ړ��ɖ߂�
                                    If �[��1 <> �[��1.Offset(myStep, 0) And �[��1.Column <> 1 Then Exit Do  '��܂��͉��������[���Ȃ�Y�����֓���
                                Loop
                            Loop
                                
                            'to�̍s��[��to�ɐi��
                            Do Until �[��1.Column = �[��to.Column
                                '1�s�ɒ[����2�ӏ��ȏ゠��ꍇ��z�肵�Đi�s�����𔻒f
                                If �[��1.Column > �[��to.Column Then myStepX = -2 Else myStepX = 2
                                Set �[��2 = �[��1.Offset(0, myStepX)
                                On Error Resume Next
                                    ws(2).Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                    ws(2).Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                On Error GoTo 0
                                Set �[��1 = �[��2
                                If Left(�[��1.Value, 1) = "U" Then ws(2).Shapes(�[��1.Value).Select False
                            Loop
                        End If
                    End If
nextI:
                Next i
                '�o�H�̍��W���擾����ׂɃO���[�v��
                Sleep 10
                ws(2).Activate
                If Selection.ShapeRange.count > 1 Then
                    Selection.Group.Name = "temp"
                    ws(2).Shapes("temp").Select
                Else
                'Selection.Name = "temp"
                End If
                myLeft = Selection.Left
                myTop = Selection.Top
                myWidth = Selection.Width
                myHeight = Selection.Height
                Sleep 10
                Selection.Copy
                If Selection.ShapeRange.Type = msoGroup Then
                    ws(2).Shapes("temp").Select
                    Selection.Ungroup
                End If
                DoEvents
                Sleep 10
                DoEvents
                ws(2).Paste
                If Selection.ShapeRange.Type <> msoGroup Then Selection.ShapeRange.Name = "temp"
    
                Selection.Left = myLeft
                Selection.Top = myTop
                '�o�H�ɐF��h��
                rootColor = RGB(0, 255, 102)
                'Call �F�ϊ�(�[���o�HRAN(3, i), clocode1, clocode2, clofont)
                If Selection.ShapeRange.Type = msoGroup Then
                    For Each ob In Selection.ShapeRange.GroupItems
                        If InStr(ob.Name, "to") > 0 Then
                            ob.Line.ForeColor.RGB = rootColor
                            ob.Line.Weight = 8
                        Else
                            If ob.Name = �[��from Then
                                ob.Fill.ForeColor.RGB = rootColor
                            Else
                                ob.Line.ForeColor.RGB = rootColor
                            End If
                            ob.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                        End If
                    Next
                    ws(2).Shapes("temp").Select
                Else
                    Selection.ShapeRange.Fill.ForeColor.RGB = rootColor
                    On Error Resume Next
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                    On Error GoTo 0
                End If
            
                wb(0).Sheets("���_" & ���).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, 480).Select
                Selection.Name = "��f"
                wb(0).Sheets("���_" & ���).Shapes("��f").Adjustments.Item(1) = 0
                wb(0).Sheets("���_" & ���).Shapes("��f").Fill.Transparency = 1
                wb(0).Sheets("���_" & ���).Shapes("��f").Line.Visible = msoFalse
                wb(0).Sheets("���_" & ���).Shapes("temp").Select False
        
                Selection.Group.Name = "temp�[���摜"
                wb(0).Sheets("���_" & ���).Shapes("temp�[���摜").Select
                myfootwidth = Selection.Width
                myfootleft = Selection.Left
                myfootheight = Selection.Height
            
                Selection.Name = "temp"
        
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '�摜�\��t���p�̖��ߍ��݃O���t���쐬
                Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, 480).Chart
        
                 '���ߍ��݃O���t�ɓ\��t����
                 DoEvents
                 Sleep 10
                DoEvents
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                '�T�C�Y����
                ActiveWindow.Zoom = 100
                '��l = 1000
                �{�� = 1
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
                If Not �[��from Is Nothing Then
                    cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �[��from & "_2.png", filtername:="PNG"
                    mypath3 = ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\" & �[��from & "-.html"
                    Call TEXT�o��_�z���o�H_�[���o�Hhtml_UTF8(mypath3, �[��from, "", ���i�i��str, "Base", "", "")
                    
                End If
                
                For u = LBound(�z���T�usize, 2) To UBound(�z���T�usize, 2)
                    If �T�u = �z���T�usize(0, u) Then
                        uu = u
                        Exit For
                    End If
                Next u

                cht.Parent.Delete
                ws(2).Shapes("temp").Delete
nextY:
'                Application.DisplayAlerts = False
'                WB(1).Save
'                Application.DisplayAlerts = True
            Next Y
        End With

'       ���[���o�H�p�̃n���}
        cb�I�� = "5,1,1,1,0,-1"
        �}���}�`�� = 160
        �[���i���o�[�\�� = True
        Call �n���}�쐬_Ver2001(cb�I��, "���C���i��", ���i�i��str)
        ws(2).Activate
        For Y = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
            �[��str = �[���ꗗran(1, Y)
            Call SQL_�z��_�[���o�H�擾(�[���o�HRAN, ���i�i��str, �[��str)
            wb(0).Sheets("�n���}_���C���i��_" & Replace(���i�i��str, " ", "")).Shapes(�[��str & "_" & 1).Copy
            DoEvents
            Sleep 10
                DoEvents
            ws(2).Paste
            Dim RANtemp() As String
            ReDim RANtemp(2, 0)
            For Each ob In ActiveSheet.Shapes(�[��str & "_" & 1).GroupItems
                '�{�̂̔w�i�F
                If ob.Name = �[��str & "_1" Then
                    ob.Glow.color.RGB = RGB(255, 255, 255)
                    ob.Glow.Radius = 4
                    ob.Glow.Transparency = 0.3
                End If
            Next ob
            Selection.Width = Selection.Width * 1
            Selection.Height = Selection.Height * 1
            left2 = ws(2).Shapes(�[��str).Left
            height2 = ws(2).Shapes(�[��str & "_1").Height
            If left2 + Selection.Width - 1220 > 0 Then
                left2 = 1220 - Selection.Width
            End If
            Selection.Left = left2
            Selection.Top = 0
            wb(0).Sheets("���_" & ���).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, height2).Select
            Selection.Name = "��f"
            wb(0).Sheets("���_" & ���).Shapes("��f").Adjustments.Item(1) = 0
            wb(0).Sheets("���_" & ���).Shapes("��f").Fill.Transparency = 1
            wb(0).Sheets("���_" & ���).Shapes("��f").Line.Visible = msoFalse
            wb(0).Sheets("���_" & ���).Shapes(�[��str & "_1").Select False
    
            Selection.Group.Name = "temp�[���摜"
            wb(0).Sheets("���_" & ���).Shapes("temp�[���摜").Select
            Selection.Name = "temp"
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '�摜�\��t���p�̖��ߍ��݃O���t���쐬
            Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, height2).Chart
    
             '���ߍ��݃O���t�ɓ\��t����
             DoEvents
             Sleep 10
                DoEvents
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            
            '�T�C�Y����
            ActiveWindow.Zoom = 100
            '��l = 1000
            �{�� = 1
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
            
            cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �[��str & "_2_foot.png", filtername:="PNG"
            
            cht.Parent.Delete
            ws(2).Shapes("temp").Delete
            
            Dim �Fv As String, �Tv As String, �[��v As String, �}v As String, �n��v As String
            For i = LBound(�[���o�HRAN, 2) To UBound(�[���o�HRAN, 2)
                �[������ = �[���o�HRAN(1, i)
                If IsNull(�[������) Then GoTo line13
                �[��v = �[������
                �Tv = �[���o�HRAN(2, i)
                �Fv = �[���o�HRAN(3, i)
                If IsNull(�[���o�HRAN(4, i)) Then �[���o�HRAN(4, i) = ""
                �}v = �[���o�HRAN(4, i)
                ��v = �[���o�HRAN(6, i)
                If ��v <> "" Then
                    If ��v = "#" Or ��v = "*" Or ��v = "=" Then
                        �Tv = "Tw"
                    ElseIf ��v = "E" Then
                        �Tv = "S"
                    Else
                        �Tv = ��v
                    End If
                End If
                ���Oc = 0
                For Each objShp In ActiveSheet.Shapes
                    If objShp.Name = �[��v & "_!" Then
                        ���Oc = ���Oc + 1
                    End If
                Next objShp
                
                '�\����_�e�[���̉��̌�n���\��
                With ActiveSheet.Shapes(�[��v)
                    .Select
                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (���Oc * 15), Selection.Top, 15, 15).Select
                    Call �F�ϊ�(�Fv, clocode1, clocode2, clofont)
                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(�Tv, "F", ""), 3)
                    Selection.ShapeRange.Adjustments.Item(1) = 0.15
                    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                    Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                    Selection.ShapeRange.Name = �[��v & "_!"

                    myFontColor = clofont '�t�H���g�F���x�[�X�F�Ō��߂�
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                    Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
                    Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 8.5
                    Selection.Font.Name = myFont
                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                    Selection.ShapeRange.TextFrame2.MarginLeft = 0
                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                    '�X�g���C�v�͌��ʂ��g��
                    If clocode1 <> clocode2 Then
                        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                            .color = clocode1
                            .color.TintAndShade = 0
                            .color.Brightness = 0
                            .Transparency = 0#
                            .Radius = 8
                        End With
                    End If
                    '�}���}
                    If �}v <> "" Then
                        myLeft = Selection.Left
                        myTop = Selection.Top
                        myHeight = Selection.Height
                        myWidth = Selection.Width
                        For Each objShp In Selection.ShapeRange
                            Set objShpTemp = objShp
                        Next objShp
                        ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                        Call �F�ϊ�(�}v, clocode1, clocode2, clofont)
                        myFontColor = clofont
                        Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                        Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                        objShpTemp.Select False
                        Selection.Group.Select
                        Selection.Name = �[��v & "_!"
                    End If
                End With
line13:
            Next i
            
            wb(0).Sheets("���_" & ���).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, 480).Select
            Selection.Name = "��f"
            wb(0).Sheets("���_" & ���).Shapes("��f").Adjustments.Item(1) = 0
            wb(0).Sheets("���_" & ���).Shapes("��f").Fill.Transparency = 1
            wb(0).Sheets("���_" & ���).Shapes("��f").Line.Visible = msoFalse
            
            For Each ob In wb(0).Sheets("���_" & ���).Shapes
                If Right(ob.Name, 2) = "_!" Then
                    ob.Select False
                End If
            Next ob
            Selection.Group.Name = "temp�[���摜"
            wb(0).Sheets("���_" & ���).Shapes("temp�[���摜").Select
            Selection.Name = "temp"
    
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '�摜�\��t���p�̖��ߍ��݃O���t���쐬
            Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, 480).Chart
    
             '���ߍ��݃O���t�ɓ\��t����
             DoEvents
             Sleep 10
                DoEvents
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            
            '�T�C�Y����
            ActiveWindow.Zoom = 100
            '��l = 1000
            �{�� = 1
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
            
            cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �[��str & "_2_tansen.png", filtername:="PNG"
            
            cht.Parent.Delete
            ws(2).Shapes("temp").Delete
        Next Y
'       ���o�H
        ws(2).Activate
        'Base
        Call �z���}�쐬3(���i�i��str, ��zstr, "Base", 0, ���, ��n���摜Sheet)
        Workbooks(myBook.Name).Activate
        ReDim Preserve �z���T�usize(2, UBound(�z���T�usize, 2) + 1)
        'ActiveSheet.Shapes("��a").Ungroup
        ActiveSheet.Shapes("���").Ungroup
        ActiveSheet.Shapes.SelectAll
        Selection.Group.Name = "����"
        ActiveSheet.Shapes("����").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height
        Selection.Ungroup
        
        Call SQL_�z���}�p_��H(�z���[��RAN, ���i�i��str, myBook)
        For i = LBound(�z���[��RAN, 2) + 1 To UBound(�z���[��RAN, 2)
            '���[��
            �\�� = �z���[��RAN(2, i)
'            If InStr("0125_0900_1301", �\��) > 0 Then Stop
            Set �[��from = Nothing: Set �[��to = Nothing
            If �z���[��RAN(4, i) <> "" Then Set �[��from = ws(2).Cells.Find(�z���[��RAN(4, i), , , 1)
            If �z���[��RAN(5, i) <> "" Then Set �[��to = ws(2).Cells.Find(�z���[��RAN(5, i), , , 1)
            If Not (�[��from Is Nothing) Then ws(2).Shapes(�z���[��RAN(4, i)).Select
            If Not (�[��to Is Nothing) Then ws(2).Shapes(�z���[��RAN(5, i)).Select False
   
            If �[��from Is Nothing And �[��to Is Nothing Then GoTo nextiii '���[����Nothing�Ȃ珈�����Ȃ�
            If Not (�[��from Is Nothing) And Not (�[��to Is Nothing) Then '�ǂ��炩�̒[����Nothing�Ȃ�I�����Ȃ�
                If �[��from <> �[��to Then '�[���������Ȃ�I�����Ȃ�
                    '���z������[���Ԃ̃��C���ɐF�t��
                    If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                        
                    Set �[��1 = �[��from
                    Set �[��2 = Nothing
'                    For y = �[��from.Row To �[��to.Row Step myStep
                    Do Until �[��1.Row = �[��to.Row
line11:
                        '-X�����ɓ���
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                            On Error GoTo 0
                           
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then ws(2).Shapes(�[��1.Value).Select False
                            If �[��1 = �[��1.Offset(myStep, 0) Then Exit Do '��܂��͉��������[���Ȃ�Y�����֓���
                        Loop
                        'Y�����ɓ���
                        Do Until �[��1.Row = �[��to.Row
                            Set �[��2 = �[��1.Offset(myStep, 0)
                            If �[��1 <> �[��2 Then
                                On Error Resume Next
                                    ws(2).Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                    ws(2).Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                On Error GoTo 0
                            End If
                            Set �[��1 = �[��2
                            Debug.Print �[��1.Row, �[��1.Column
                            If Left(�[��1.Value, 1) = "U" Then ws(2).Shapes(�[��1.Value).Select False
                            If �[��1.Row = �[��to.Row Then Exit Do '�[��to�Ɠ����s�Ȃ�[��to�ɐi��
                            If �[��1.Offset(myStep, 0) = "" Then GoTo line11 '�i�ސ悪�󗓂Ȃ�X�����ړ��ɖ߂�
                            If Left(�[��1.Offset(myStep, 0), 1) <> "U" Then GoTo line11 '�i�ސ悪U����Ȃ����X�����ړ��ɖ߂�
                            If �[��1 <> �[��1.Offset(myStep, 0) And �[��1.Column <> 1 Then Exit Do  '��܂��͉��������[���Ȃ�Y�����֓���
                        Loop
                    Loop
                        
                    'to�̍s��[��to�ɐi��
                    Do Until �[��1.Column = �[��to.Column
                        '1�s�ɒ[����2�ӏ��ȏ゠��ꍇ��z�肵�Đi�s�����𔻒f
                        If �[��1.Column > �[��to.Column Then myStepX = -2 Else myStepX = 2
                        Set �[��2 = �[��1.Offset(0, myStepX)
                        On Error Resume Next
                            ws(2).Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                            ws(2).Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                        On Error GoTo 0
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then ws(2).Shapes(�[��1.Value).Select False
                    Loop
'                    Next y
                End If
            End If

            '�o�H�̍��W���擾����ׂɃO���[�v��
            Sleep 10
            ws(2).Activate
            If Selection.ShapeRange.count > 1 Then
                Selection.Group.Name = "temp"
                ws(2).Shapes("temp").Select
            Else
                'Selection.Name = "temp"
            End If
            myLeft = Selection.Left
            myTop = Selection.Top
            myWidth = Selection.Width
            myHeight = Selection.Height
            Sleep 10
            Selection.Copy
            If Selection.ShapeRange.Type = msoGroup Then
                ws(2).Shapes("temp").Select
                Selection.Ungroup
            Else
            End If
            DoEvents
            Sleep 70
            DoEvents
            ws(2).Paste
            If Selection.ShapeRange.Type <> msoGroup Then Selection.ShapeRange.Name = "temp"

            Selection.Left = myLeft
            Selection.Top = myTop
            '�o�H�ɐF��h��
            Call �F�ϊ�(�z���[��RAN(3, i), clocode1, clocode2, clofont)
            If Selection.ShapeRange.Type = msoGroup Then
                For Each ob In Selection.ShapeRange.GroupItems
                    If InStr(ob.Name, "to") > 0 Then
                        ob.Line.ForeColor.RGB = clocode1
                        ob.Line.Weight = 8
                        If �z���[��RAN(3, i) = "B" Or �z���[��RAN(3, i) = "GY" Then
                            ob.Glow.color.RGB = RGB(255, 255, 255)
                            ob.Glow.Radius = 8
                            ob.Glow.Transparency = 0.5
                        End If
                    ElseIf InStr(ob.Name, "U") > 0 Then
                        ob.Fill.ForeColor.RGB = rootColor
                    Else
                        ob.Fill.ForeColor.RGB = clocode1
                        ob.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                    End If
                Next
                ws(2).Shapes("temp").Select
            Else
                Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0 '2.191.12�Ŏb��ύX���g�p���Ȃ�
            End If
        
        wb(0).Sheets("���_" & ���).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, 480).Select
        Selection.Name = "��f"
        wb(0).Sheets("���_" & ���).Shapes("��f").Adjustments.Item(1) = 0
        wb(0).Sheets("���_" & ���).Shapes("��f").Fill.Transparency = 1
        wb(0).Sheets("���_" & ���).Shapes("��f").Line.Visible = msoFalse
        wb(0).Sheets("���_" & ���).Shapes("temp").Select False

        Selection.Group.Name = "temp�[���摜"
        wb(0).Sheets("���_" & ���).Shapes("temp�[���摜").Select
        myfootwidth = Selection.Width
        myfootleft = Selection.Left
        myfootheight = Selection.Height
    
        Selection.Name = "temp"
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '�摜�\��t���p�̖��ߍ��݃O���t���쐬
        Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, 480).Chart

         '���ߍ��݃O���t�ɓ\��t����
         DoEvents
         Sleep 10
         DoEvents
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        
        '�T�C�Y����
        ActiveWindow.Zoom = 100
        '��l = 1000
        �{�� = 1
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
        Dim �F�� As String, �F��b As String
        �F�� = �z���[��RAN(3, i)
        If InStr(�F��, "/") > 0 Then
            �F��b = Left(�F��, InStr(�F��, "/") - 1)
        Else
            �F��b = �F��
        End If
        cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �z���[��RAN(4, i) & "to" & �z���[��RAN(5, i) & "_" & �F��b & ".png", filtername:="PNG"

        mypath1 = wb(0).Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\" & �z���[��RAN(2, i) & ".html"
        
        If Dir(wb(0).Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �z���[��RAN(1, i) & ".png") = "" Then
            �T�u2 = "Base"
        Else
            �T�u2 = �z���[��RAN(1, i)
        End If
       
        '2.191.00
        Call TEXT�o��_�z���o�Hhtml_UTF8(mypath1, �z���[��RAN(4, i), �z���[��RAN(5, i), �z���[��RAN(0, i), �z���[��RAN(1, i), �T�u2, �z���[��RAN(2, i), �F��b, �z���[��RAN(7, i), �z���[��RAN(8, i), �z���[��RAN(9, i), �z���[��RAN(10, i), �[��leftRAN)
        
        For u = LBound(�z���T�usize, 2) To UBound(�z���T�usize, 2)
            If �T�u = �z���T�usize(0, u) Then
                uu = u
                Exit For
            End If
        Next u
        
        Call �F�ϊ�css(�z���[��RAN(3, i), clocode1, clocode2, clofont)
        mypath2 = wb(0).Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css\" & �z���[��RAN(2, i) & ".css"
        myEx = myTop / mybaseheight * 100
        'mytopEx = (0.0000007 * myEx ^ 3) + (0.00006 * myEx ^ 2) + (0.9978 * myEx) + 0.0512
        myTopEx = 0.9861 * myEx ^ 1.001
        
        Call TEXT�o��_�z���o�Hcss(mypath2, myLeft / mybasewidth * 100, myTopEx, myWidth / mybasewidth, myHeight / mybaseheight, clocode1, clofont)
        
        'Shell "EXPLORER.EXE " & mypath1
        jjj = "0011_0565_0626_0569_0674_0607_0497"
        jjjs = Split(jjj, "_")
        For jj = 0 To UBound(jjjs)
            If jjjs(jj) = �z���[��RAN(2, i) Then
                'Stop
                Debug.Print �z���[��RAN(2, i), myTopEx, myEx
                Shell "EXPLORER.EXE " & mypath1
            End If
        Next jj
        'Stop
        cht.Parent.Delete
        ws(2).Shapes("temp").Delete
nextiii:
        Next i
        Application.DisplayAlerts = False
        wb(1).Save
        Application.DisplayAlerts = True
    End With
    
line77:
    '�z���o�H�p_��n���}.png�̏o��
    ��lx = 1440
    ��ly = 900
    �䗦xy = ��lx / ��ly
    cb�I�� = "4,1,1,1,0,-1"
    �}���}�`�� = 160
    �[���i���o�[�\�� = False
    Call �n���}�쐬_Ver2001(cb�I��, "���C���i��", ���i�i��str)
    Call SQL_�z����n���_�Ŏ擾(��n���_��ran, ���i�i��str)
    
    Call �œK��
    Dim Width0 As Single, height0 As Single
    �{��0 = 2
    Set ws(3) = wb(0).Sheets("�n���}_���C���i��_" & Replace(���i�i��str, " ", ""))
    �[�����i��row = ws(3).Cells.Find("�[�����i��", , , 1).Row
    �[�����i��Col = ws(3).Cells.Find("�[�����i��", , , 1).Column
    �[��Col = ws(3).Rows(�[�����i��row).Find("�[����", , , 1).Column
    For Each objShp In ws(3).Shapes
        '�o��
        �[��str = objShp.Name
        If InStr(�[��str, "Comment") > 0 Then GoTo line90
        myW = objShp.Width
        myH = objShp.Height
        If myW > myH * �䗦xy Then �{�� = ��lx / myW Else �{�� = ��ly / myH
        �{�� = �{�� / 96 * 72 '�|�C���g���s�N�Z���ɕϊ�
        If InStr(�[��str, "_") > 0 Then
            �[��0 = Left(�[��str, InStr(�[��str, "_") - 1)
            �[��row = ws(3).Columns(�[��Col).Find(�[��0, , , 1).Row
            ���i�i��str = ws(3).Cells(�[��row, �[�����i��Col)
        End If

        '�w�i�������̂ŃR�l�N�^�ʐ^����glow
        For Each ob In objShp.GroupItems
            If ob.Name = �[��str And Left(���i�i��str, 4) <> "7009" Then
                ob.Glow.color.RGB = RGB(255, 255, 255)
                ob.Glow.Radius = 3.5
                ob.Glow.Transparency = 0.4
                Exit For
            End If
        Next ob
        objShp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '�摜�\��t���p�̖��ߍ��݃O���t���쐬
        Set cht = ws(3).ChartObjects.add(0, 0, objShp.Width * �{��, objShp.Height * �{��).Chart
         '���ߍ��݃O���t�ɓ\��t����
        DoEvents
        Sleep 10
        DoEvents
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        '�T�C�Y����
'        WS(3).Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��0, False, msoScaleFromTopLeft
'        WS(3).Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��0, False, msoScaleFromTopLeft
        '�摜�T�C�Y����������cht�Ɠ����T�C�Y�ɂȂ�Ȃ�?�̂ō��킹��
        If Selection.Width <> objShp.Width * �{�� Then
            On Error Resume Next
            Selection.Width = objShp.Width * �{��
            On Error GoTo 0
        End If
        
        cht.Export fileName:=wb(0).Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �[��str & ".png", filtername:="PNG"
        cht.Parent.Delete
        
        mypath0 = ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\" & �[��0 & ".html"
        
        Call TEXT�o��_�z���o�H_�[��html_UTF8(mypath0, �[��str, �[��0, ���i�i��str)
        If �z���}�쐬temp = 1 Then GoTo line90
        '��n���_�ŗp�̉摜�o��
        For Each objShp2 In objShp.GroupItems
            If InStr(objShp2.Name, �[��str & "_") > 0 Then
                �[��temp = Mid(objShp2.Name, Len(�[��str) + 2)
                If IsNumeric(�[��temp) = True Then
                    If ��n���_�� = True Then GoTo line70
                    '��n������Ȃ��ꍇ�摜�o�͂��Ȃ�
                    For pp = LBound(��n���_��ran, 2) To UBound(��n���_��ran, 2)
                        If �[��str = ��n���_��ran(0, pp) & "_1" Then
                            If Left(��n���_��ran(2, pp), "1") = "��" Then
                                If �[��temp = ��n���_��ran(1, pp) Then GoTo line70
                            End If
                        End If
                    Next pp
                    GoTo line80
line70:
                    ws(3).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, objShp.Width, objShp.Height).Name = "��f"
                    ws(3).Shapes("��f").Adjustments.Item(1) = 0
                    ws(3).Shapes("��f").Fill.Transparency = 1
                    ws(3).Shapes("��f").Line.Visible = msoFalse
                    objShp2.Copy
                    DoEvents
                    Sleep 10
                    DoEvents
                    ws(3).Paste
                    Selection.Left = objShp2.Left - objShp.Left + 1
                    Selection.Top = objShp2.Top - objShp.Top
                    
                    '�_�ŗp��CAV��ύX
                    On Error Resume Next
                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                    On Error GoTo 0
                    
                    Selection.ShapeRange.Fill.Visible = msoTrue
                    Selection.ShapeRange.Fill.Transparency = 0
                    Selection.ShapeRange.Fill.Solid
                    tempcolor = Selection.ShapeRange.Fill.ForeColor
                    Selection.ShapeRange.Fill.ForeColor.RGB = tempcolor
                    Selection.ShapeRange.Line.Visible = False
                    Selection.ShapeRange.Glow.color.RGB = tempcolor
                    Selection.ShapeRange.Glow.Transparency = 0
                    Selection.ShapeRange.Glow.Radius = 13
                    
                    ws(3).Shapes("��f").Select False
                    Selection.Group.Select
                    Selection.Name = "cavTemp"
                    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                    Set cht = ws(3).ChartObjects.add(0, 0, Selection.Width * �{��, Selection.Height * �{��).Chart
                     '���ߍ��݃O���t�ɓ\��t����
                    DoEvents
                    Sleep 10
                    DoEvents
                    cht.Paste
                    '�摜�T�C�Y����������cht�Ɠ����T�C�Y�ɂȂ�Ȃ�?�̂ō��킹��
                    If Selection.Width <> objShp.Width * �{�� Then
                        Selection.Width = objShp.Width * �{��
                    End If
                    cht.PlotArea.Fill.Visible = mesofalse
                    cht.ChartArea.Fill.Visible = msoFalse
                    cht.ChartArea.Border.LineStyle = 0
                    cht.Export fileName:=wb(0).Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �[��str & "_" & �[��temp & ".png", filtername:="PNG"
                    cht.Parent.Delete
                    ws(3).Shapes("cavTemp").Delete
                End If
            End If
line80:
        Next objShp2
line90:
    Next objShp
    
line99:
    mypath0 = myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css\atohame.css"
    Call TEXT�o��_�z���o�H_�[��css(mypath0)
    
    mypath0 = myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\css\tanmatukeiro.css"
    Call TEXT�o��_�z���o�H_�[���o�Hcss(mypath0)

    mypath0 = myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\myBlink_end.js"
    Call TEXT�o��_�z���o�H_�[��js2(mypath0)

    mypath0 = myBook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\version.txt"
    Call TEXT�o��_�z���o�H_ver(mypath0)
    '�d�������o�����z���}��ۑ������ɕ���
    Application.DisplayAlerts = False
    wb(1).Close , savechanges = False
    Application.DisplayAlerts = True
    
    Call �œK�����ǂ�
    
End Function

Public Function CAV�ꗗ�쐬()
    Call �œK��

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "CAV�ꗗ"
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call �A�h���X�Z�b�g(myBook)
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
      
    '�������O�̃V�[�g�����邩�m�F
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
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        If newSheet.Name = "CAV�ꗗ" Then
            newSheet.Tab.color = 14470546
        End If
    End If
    
    '�g�p�R�l�N�^�ꗗ���쐬
    Call SQL_���i�ʒ[���ꗗ_�h��(�R�l�N�^�ꗗRAN, ���i�i��RAN, myBook)
    '�g�p�R�l�N�^�ꗗ�ɖh���敪������
    Call SQL_���i�ʒ[���ꗗ_�h���敪(RAN, �R�l�N�^�ꗗRAN)
    
    With myBook.Sheets(newSheetName)
        Set key = .Cells.Find("�[����", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 2
            .Cells(1, 1) = "�R�l�N�^�ꗗ"
            .Cells(keyRow, keyCol - 1) = "�h���敪"
            .Cells(keyRow, keyCol - 1).AddComment.Text "1=�h���^�C�v" & vbCrLf & "2=��h���^�C�v" & vbCrLf & "3=�h���A��h���̋敪����"
            .Cells(keyRow, keyCol + 0) = "�[����"
            .Cells(keyRow, keyCol + 1) = "���i�i��"
            .Cells(keyRow, keyCol + 2) = "Cav"
            .Cells(keyRow, keyCol + 3) = "Width"
            .Cells(keyRow, keyCol + 4) = "Height"
            .Cells(keyRow, keyCol + 5) = "EmptyPlug"
            .Cells(keyRow, keyCol + 6) = "PlugColor"
            lastRow = keyRow
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Cells.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(�R�l�N�^�ꗗRAN, 2) + 1 To UBound(�R�l�N�^�ꗗRAN, 2)
            ��� = �R�l�N�^�ꗗRAN(0, Y)
            �[�� = �R�l�N�^�ꗗRAN(1, Y)
            �h���敪 = �R�l�N�^�ꗗRAN(2, Y)
            If InStr(���, "-") = 0 Then
                Select Case Len(���)
                Case 8
                    ��� = Left(���, 4) & "-" & Mid(���, 5, 4)
                Case 10
                    ��� = Left(���, 4) & "-" & Mid(���, 5, 4) & "-" & Mid(���, 9, 2)
                End Select
            End If
            '�o�^�����邩�m�F
            For i = keyRow To lastRow
                flg = False
                If �[�� = .Cells(i, keyCol) And ��� = .Cells(i, keyCol + 1) Then
                    flg = True
                    addRow = i
                    Exit For
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                ���W�t�@�C���m�� = ""
                ���W�t�@�C�� = �A�h���X(1) & "\200_CAV���W\" & ��� & "_1_001_png.txt"
                If Dir(���W�t�@�C��) <> "" Then
                    ���W�t�@�C���m�� = ���W�t�@�C��
                Else
                    ���W�t�@�C�� = �A�h���X(1) & "\200_CAV���W\" & ��� & "_1_001_emf.txt"
                    If Dir(���W�t�@�C��) <> "" Then ���W�t�@�C���m�� = ���W�t�@�C��
                End If
                
                If ���W�t�@�C���m�� <> "" Then
                    Dim buf As String
                    Dim cc As Long
                    cc = 0
                    Open ���W�t�@�C���m�� For Input As #1
                        Do Until EOF(1)
                            Line Input #1, buf
                            If cc = 0 Then GoTo linenext
                            bufsp = Split(buf, ",")
                            addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                            .Cells(addRow, keyCol - 1) = �h���敪
                            .Cells(addRow, keyCol + 0) = �[��
                            .Cells(addRow, keyCol + 1) = bufsp(0)
                            .Cells(addRow, keyCol + 2) = bufsp(1)
                            .Cells(addRow, keyCol + 3) = bufsp(2)
                            .Cells(addRow, keyCol + 4) = bufsp(3)
                            .Cells(addRow, keyCol + 5) = bufsp(13)
                            .Cells(addRow, keyCol + 6) = bufsp(14)
linenext:
                        cc = 1
                        Loop
                    Close #1
                End If
            End If
        Next Y
        
        '���i�i�Ԗ��Ɏg�p�������m�F
        For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            Call SQL_���i�ʒ[���ꗗ_�g�p�d���m�F(�g�p�d��ran, ���i�i��str)
            Set aKey = .Cells.Find(���i�i��str, , , 1)
            If aKey Is Nothing Then
                addCol = .Cells(keyRow, .Columns.count).End(xlToLeft).Column + 1
                .Cells(keyRow - 0, addCol) = ���i�i��str
                .Cells(keyRow - 1, addCol) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), r)
                .Cells(keyRow - 2, addCol).Font.Size = 10
                .Cells(keyRow - 2, addCol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, addCol) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�N����"), r)
                .Columns(addCol).ColumnWidth = 4.6
            Else
                addCol = aKey.Column
            End If
            .Range(.Cells(keyRow + 1, addCol), .Cells(.Rows.count, addCol)).ClearContents
            
            For Y = LBound(�g�p�d��ran, 2) To UBound(�g�p�d��ran, 2)
                If IsNull(�g�p�d��ran(1, Y)) Then GoTo nextY
                �[�� = �g�p�d��ran(1, Y)
                ��� = �g�p�d��ran(2, Y)
                acav = �g�p�d��ran(3, Y)
                If �[�� = "" Or ��� = "" Or acav = "" Then GoTo nextY
                �T�u = �g�p�d��ran(0, Y)
                For i = keyRow + 1 To addRow
                    If �[�� = .Cells(i, keyCol + 0) Then
                        If ��� = Replace(.Cells(i, keyCol + 1), "-", "") Then
                            If CStr(acav) = CStr(.Cells(i, keyCol + 2)) Then
                                If .Cells(i, addCol) = "" Then
                                    .Cells(i, addCol) = "1"
                                End If
                                GoTo nextY
                            End If
                        End If
                    End If
                Next i
nextY:
            Next Y
            '�d����1�_�ȏ����[���Ŗh���^�C�v�̒[���͐F�t��
            firstRow = keyRow + 1
            flg = False
            For i = keyRow + 1 To addRow
                �T�u = .Cells(i, addCol)
                If �T�u <> "" Then flg = True
                �[�� = .Cells(i, keyCol + 0)
                ��� = .Cells(i, keyCol + 1)
                cav = CStr(.Cells(i, keyCol + 2))
                �h���敪 = Left(.Cells(i, keyCol - 1), 1)
                �[��next = .Cells(i + 1, keyCol + 0)
                ���next = .Cells(i + 1, keyCol + 1)
                cavNext = CStr(.Cells(i, keyCol + 2))
                
                If �[�� & ��� <> �[��next & ���next Then
                    If flg = True And �h���敪 <> "2" Then
                        For i2 = firstRow To i
                            
                            If .Cells(i2, addCol) = "" Then
                                .Cells(i2, keyCol + 5).Interior.color = RGB(146, 204, 255)
                            End If
                        Next i2
                    End If
                    firstRow = i + 1
                    flg = False
                End If
nextI:
            Next i
        Next r
        
        'MD�f�[�^������ꍇ�A���i�Ԃ��擾
        For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            �ݕ�str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "��z"), r)
            myCount = 0
            If MD = True Then myCount = SQL_MD�t�@�C���ǂݍ���_���(���i�i��str, �ݕ�str, ���RAN)
            Dim ���str2 As String, ���i�i��str2 As String, cavStr2 As String, �[��str2 As String
            If myCount <> Empty Then
                lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
                For i = LBound(���RAN, 2) To UBound(���RAN, 2)
                    ���str2 = ���RAN(0, i)
                    ���str2 = �[�����i�ԕϊ�(���str2)
                    cavStr2 = ���RAN(2, i)
                    �H��str2 = ���RAN(3, i)
                    �[��str2 = ���RAN(4, i)
                    ���str2 = ���RAN(5, i)
                    ���str2 = �[�����i�ԕϊ�(���str2)
                    
                    For Y = keyRow + 1 To lastRow
                        If .Cells(Y, keyCol) = �[��str2 Then
                            If .Cells(Y, keyCol + 1) = ���str2 Then
                                If .Cells(Y, keyCol + 2) = cavStr2 Then
                                    If .Cells(Y, keyCol + 5).Value <> "" And .Cells(Y, keyCol + 5).Value <> ���str2 Then Stop '���i�ɂ���ċ�����قȂ�?
                                    .Cells(Y, keyCol + 5).Value = ���str2
                                    Exit For
                                End If
                            End If
                        End If
                    Next Y
                Next i
                '���i�i�Ԃ�MD�ɂ��ċL��
                With myBook.Sheets("���i�i��")
                    Dim ���C���i�� As Variant: Set ���C���i�� = .Cells.Find("���C���i��", , , 1)
                    Dim seihinRow As Long: seihinRow = .Columns(���C���i��.Column).Find(���i�i��str & String(15 - Len(���i�i��str), " "), , , 1).Row
                    .Cells(seihinRow, .Rows(���C���i��.Row).Find("MD", , , 1).Column).Value = �ݕ�str
                End With
            End If
        Next r
    End With
    
     '�\�[�g
    With myBook.Sheets(newSheetName)
        .Select
        .Range(Columns(keyCol - 1), Columns(keyCol + 6)).AutoFit
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call �œK�����ǂ�

End Function

Public Function CAV�ꗗ�쐬2190()
    Dim sTime As Single: sTime = Timer
    Call �œK��

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "CAV�ꗗ"
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call �A�h���X�Z�b�g(myBook)
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
      
    '�������O�̃V�[�g�����邩�m�F
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
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        If newSheet.Name = "CAV�ꗗ" Then
            newSheet.Tab.color = 14470546
        End If
    End If
    
    '�g�p�R�l�N�^�ꗗ���쐬
    Call SQL_���i�ʒ[���ꗗ_�h��(�R�l�N�^�ꗗRAN, ���i�i��RAN, myBook)
'    '�g�p�R�l�N�^�ꗗ�ɖh���敪������
   'Call SQL_���i�ʒ[���ꗗ_�h���敪(RAN, �R�l�N�^�ꗗRAN)
    
    With myBook.Sheets(newSheetName)
        Set key = .Cells.Find("�[����", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 2
            .Cells(1, 1) = "�R�l�N�^�ꗗ"
            .Cells(keyRow, keyCol - 1) = "�h���敪"
            .Cells(keyRow, keyCol - 1).AddComment.Text "1=�h���^�C�v" & vbCrLf & "2=��h���^�C�v" & vbCrLf & "3=�h���A��h���̋敪����"
            .Cells(keyRow, keyCol + 0) = "�[����"
            .Cells(keyRow, keyCol + 1) = "���i�i��"
            .Cells(keyRow, keyCol + 2) = "Cav"
            .Cells(keyRow, keyCol + 3) = "Width"
            .Cells(keyRow, keyCol + 4) = "Height"
            .Cells(keyRow, keyCol + 5) = "EmptyPlug"
            .Cells(keyRow, keyCol + 6) = "PlugColor"
            lastRow = keyRow
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Cells.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(�R�l�N�^�ꗗRAN, 2) + 1 To UBound(�R�l�N�^�ꗗRAN, 2)
            ��� = �R�l�N�^�ꗗRAN(0, Y)
            �[�� = �R�l�N�^�ꗗRAN(1, Y)
            �h���敪 = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(���), "�h���敪_")
            If InStr(���, "-") = 0 Then
                Select Case Len(���)
                Case 8
                    ��� = Left(���, 4) & "-" & Mid(���, 5, 4)
                Case 10
                    ��� = Left(���, 4) & "-" & Mid(���, 5, 4) & "-" & Mid(���, 9, 2)
                End Select
            End If
            '�o�^�����邩�m�F
            For i = keyRow To lastRow
                flg = False
                If �[�� = .Cells(i, keyCol) And ��� = .Cells(i, keyCol + 1) Then
                    flg = True
                    addRow = i
                    Exit For
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                ���W�t�@�C���m�� = ""
                ���W�t�@�C�� = �A�h���X(1) & "\200_CAV���W\" & ��� & "_1_001_png.txt"
                If Dir(���W�t�@�C��) <> "" Then
                    ���W�t�@�C���m�� = ���W�t�@�C��
                Else
                    ���W�t�@�C�� = �A�h���X(1) & "\200_CAV���W\" & ��� & "_1_001_emf.txt"
                    If Dir(���W�t�@�C��) <> "" Then ���W�t�@�C���m�� = ���W�t�@�C��
                End If
                
                If ���W�t�@�C���m�� <> "" Then
                    Dim buf As String
                    Dim cc As Long
                    cc = 0
                    Open ���W�t�@�C���m�� For Input As #1
                        Do Until EOF(1)
                            Line Input #1, buf
                            If cc = 0 Then GoTo linenext
                            bufsp = Split(buf, ",")
                            addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                            .Cells(addRow, keyCol - 1) = �h���敪
                            .Cells(addRow, keyCol + 0) = �[��
                            .Cells(addRow, keyCol + 1) = bufsp(0)
                            .Cells(addRow, keyCol + 2) = bufsp(1)
                            .Cells(addRow, keyCol + 3) = bufsp(2)
                            .Cells(addRow, keyCol + 4) = bufsp(3)
                            .Cells(addRow, keyCol + 5) = bufsp(13)
                            .Cells(addRow, keyCol + 6) = ""
linenext:
                        cc = 1
                        Loop
                    Close #1
                End If
            End If
        Next Y
        
        '���i�i�Ԗ��Ɏg�p�������m�F
        For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            Call SQL_���i�ʒ[���ꗗ_�g�p�d���m�F(�g�p�d��ran, ���i�i��str)
            Set aKey = .Cells.Find(���i�i��str, , , 1)
            If aKey Is Nothing Then
                addCol = .Cells(keyRow, .Columns.count).End(xlToLeft).Column + 1
                .Cells(keyRow - 0, addCol) = ���i�i��str
                .Cells(keyRow - 1, addCol) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), r)
                .Cells(keyRow - 2, addCol).Font.Size = 10
                .Cells(keyRow - 2, addCol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, addCol) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�N����"), r)
                .Columns(addCol).ColumnWidth = 4.6
            Else
                addCol = aKey.Column
            End If
            .Range(.Cells(keyRow + 1, addCol), .Cells(.Rows.count, addCol)).ClearContents
            
            For Y = LBound(�g�p�d��ran, 2) To UBound(�g�p�d��ran, 2)
                If IsNull(�g�p�d��ran(1, Y)) Then GoTo nextY
                �[�� = �g�p�d��ran(1, Y)
                ��� = �g�p�d��ran(2, Y)
                acav = �g�p�d��ran(3, Y)
                If �[�� = "" Or ��� = "" Or acav = "" Then GoTo nextY
                �T�u = �g�p�d��ran(0, Y)
                For i = keyRow + 1 To addRow
                    If �[�� = .Cells(i, keyCol + 0) Then
                        If ��� = Replace(.Cells(i, keyCol + 1), "-", "") Then
                            If CStr(acav) = CStr(.Cells(i, keyCol + 2)) Then
                                If .Cells(i, addCol) = "" Then
                                    .Cells(i, addCol) = "1"
                                End If
                                GoTo nextY
                            End If
                        End If
                    End If
                Next i
nextY:
            Next Y
            '�d����1�_�ȏ����[���Ŗh���^�C�v�̒[���͐F�t��
            firstRow = keyRow + 1
            flg = False
            For i = keyRow + 1 To addRow
                �T�u = .Cells(i, addCol)
                If �T�u <> "" Then flg = True
                �[�� = .Cells(i, keyCol + 0)
                ��� = .Cells(i, keyCol + 1)
                cav = CStr(.Cells(i, keyCol + 2))
                �h���敪 = Left(.Cells(i, keyCol - 1), 1)
                �[��next = .Cells(i + 1, keyCol + 0)
                ���next = .Cells(i + 1, keyCol + 1)
                cavNext = CStr(.Cells(i, keyCol + 2))
                
                If �[�� & ��� <> �[��next & ���next Then
                    If flg = True And �h���敪 <> "2" Then
                        For i2 = firstRow To i
                            
                            If .Cells(i2, addCol) = "" Then
                                .Cells(i2, keyCol + 5).Interior.color = RGB(146, 204, 255)
                            End If
                        Next i2
                    End If
                    firstRow = i + 1
                    flg = False
                End If
nextI:
            Next i
        Next r
        
        'MD�f�[�^������ꍇ�A���i�Ԃ��擾
        For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            �ݕ�str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "��z"), r)
            myCount = 0
            myCount = SQL_MD�t�@�C���ǂݍ���_���(���i�i��str, �ݕ�str, ���RAN)
            Dim ���str2 As String, ���i�i��str2 As String, cavStr2 As String, �[��str2 As String
            If myCount <> Empty Then
                lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
                For i = LBound(���RAN, 2) To UBound(���RAN, 2)
                    ���str2 = ���RAN(0, i)
                    ���str2 = �[�����i�ԕϊ�(���str2)
                    cavStr2 = ���RAN(2, i)
                    �H��str2 = ���RAN(3, i)
                    �[��str2 = ���RAN(4, i)
                    ���str2 = ���RAN(5, i)
                    ���str2 = �[�����i�ԕϊ�(���str2)
                    
                    For Y = keyRow + 1 To lastRow
                        If .Cells(Y, keyCol) = �[��str2 Then
                            If .Cells(Y, keyCol + 1) = ���str2 Then
                                If .Cells(Y, keyCol + 2) = cavStr2 Then
                                    If .Cells(Y, keyCol + 5).Value <> "" And .Cells(Y, keyCol + 5).Value <> ���str2 Then Stop '���i�ɂ���ċ�����قȂ�?
                                    PlugColor = ���ޏڍׂ̓ǂݍ���(���str2, "�F_")
                                    .Cells(Y, keyCol + 5).Value = ���str2
                                    .Cells(Y, keyCol + 6).Value = PlugColor
                                    Exit For
                                End If
                            End If
                        End If
                    Next Y
                Next i
                '���i�i�Ԃ�MD�ɂ��ċL��
                With myBook.Sheets("���i�i��")
                    Dim ���C���i�� As Variant: Set ���C���i�� = .Cells.Find("���C���i��", , , 1)
                    Dim seihinRow As Long: seihinRow = .Columns(���C���i��.Column).Find(���i�i��str & String(15 - Len(���i�i��str), " "), , , 1).Row
                    .Cells(seihinRow, .Rows(���C���i��.Row).Find("MD", , , 1).Column).Value = �ݕ�str
                End With
            End If
        Next r
    End With
    
     '�\�[�g
    With myBook.Sheets(newSheetName)
        .Select
        .Range(Columns(keyCol - 1), Columns(keyCol + 6)).AutoFit
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call �œK�����ǂ�

    CAV�ꗗ�쐬2190 = Round(Timer - sTime)
End Function


Public Function �ؒf�����ꗗ()

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�ؒf�����ꗗ"
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call �A�h���X�Z�b�g(myBook)
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
      
    '�������O�̃V�[�g�����邩�m�F
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
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        If newSheet.Name = newSheetName Then
            newSheet.Tab.color = 14470546
        End If
    End If
    
    With myBook.Sheets(newSheetName)
        Set key = .Cells.Find("�\����", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 2
            .Cells(1, 1) = "�ؒf�����ꗗ"
            .Cells(keyRow, keyCol - 1) = "����"
            .Cells(keyRow, keyCol - 1).AddComment.Text "²��(*#=)" & vbCrLf & "�����(E)"
            .Cells(keyRow, keyCol + 0) = "�[����"
            .Cells(keyRow, keyCol + 1) = "���i�i��"
            .Cells(keyRow, keyCol + 2) = "Cav"
            .Cells(keyRow, keyCol + 3) = "Width"
            .Cells(keyRow, keyCol + 4) = "Height"
            .Cells(keyRow, keyCol + 5) = "EmptyPlug"
            .Cells(keyRow, keyCol + 6) = "PlugColor"
            lastRow = keyRow
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Cells.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(�R�l�N�^�ꗗRAN, 2) + 1 To UBound(�R�l�N�^�ꗗRAN, 2)
            ��� = �R�l�N�^�ꗗRAN(0, Y)
            �[�� = �R�l�N�^�ꗗRAN(1, Y)
            �h���敪 = �R�l�N�^�ꗗRAN(2, Y)
            If InStr(���, "-") = 0 Then
                Select Case Len(���)
                Case 8
                    ��� = Left(���, 4) & "-" & Mid(���, 5, 4)
                Case 10
                    ��� = Left(���, 4) & "-" & Mid(���, 5, 4) & "-" & Mid(���, 9, 2)
                End Select
            End If
            '�o�^�����邩�m�F
            For i = keyRow To lastRow
                flg = False
                If �[�� = .Cells(i, keyCol) And ��� = .Cells(i, keyCol + 1) Then
                    flg = True
                    addRow = i
                    Exit For
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                'Stop
                Call SQL_���i�ʒ[���ꗗ_CAV���W(���WRAN, ���, myBook)
                For r = LBound(���WRAN, 2) + 1 To UBound(���WRAN, 2)
                    addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                    .Cells(addRow, keyCol - 1) = �h���敪
                    .Cells(addRow, keyCol + 0) = �[��
                    .Cells(addRow, keyCol + 1) = ���WRAN(0, r)
                    .Cells(addRow, keyCol + 2) = ���WRAN(1, r)
                    .Cells(addRow, keyCol + 3) = ���WRAN(2, r)
                    .Cells(addRow, keyCol + 4) = ���WRAN(3, r)
                    .Cells(addRow, keyCol + 5) = ���WRAN(4, r)
                    .Cells(addRow, keyCol + 6) = ���WRAN(5, r)
                Next r
            End If
        Next Y
        
        '���i�i�Ԗ��Ɏg�p�������m�F
        For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            Call SQL_���i�ʒ[���ꗗ_�g�p�d���m�F(�g�p�d��ran, ���i�i��str)
            Set aKey = .Cells.Find(���i�i��str, , , 1)
            If aKey Is Nothing Then
                addCol = .Cells(keyRow, .Columns.count).End(xlToLeft).Column + 1
                .Cells(keyRow - 0, addCol) = ���i�i��str
                .Cells(keyRow - 1, addCol) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), r)
                .Cells(keyRow - 2, addCol).Font.Size = 10
                .Cells(keyRow - 2, addCol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, addCol) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�N����"), r)
                .Columns(addCol).ColumnWidth = 4.6
            Else
                addCol = aKey.Column
            End If
            .Range(.Cells(keyRow + 1, addCol), .Cells(.Rows.count, addCol)).ClearContents
            
            For Y = LBound(�g�p�d��ran, 2) To UBound(�g�p�d��ran, 2)
                If IsNull(�g�p�d��ran(1, Y)) Then GoTo nextY
                �[�� = �g�p�d��ran(1, Y)
                ��� = �g�p�d��ran(2, Y)
                acav = �g�p�d��ran(3, Y)
                If �[�� = "" Or ��� = "" Or acav = "" Then GoTo nextY
                �T�u = �g�p�d��ran(0, Y)
                For i = keyRow + 1 To addRow
                    If �[�� = .Cells(i, keyCol + 0) Then
                        If ��� = Replace(.Cells(i, keyCol + 1), "-", "") Then
                            If CStr(acav) = CStr(.Cells(i, keyCol + 2)) Then
                                If .Cells(i, addCol) = "" Then
                                    .Cells(i, addCol) = "1"
                                End If
                                GoTo nextY
                            End If
                        End If
                    End If
                Next i
nextY:
            Next Y
            '�d����1�_�ȏ����[���Ŗh���^�C�v�̒[���͐F�t��
            
            firstRow = keyRow + 1
            flg = False
            For i = keyRow + 1 To addRow
                �T�u = .Cells(i, addCol)
                If �T�u <> "" Then flg = True
                �[�� = .Cells(i, keyCol + 0)
                ��� = .Cells(i, keyCol + 1)
                cav = CStr(.Cells(i, keyCol + 2))
                �h���敪 = Left(.Cells(i, keyCol - 1), 1)
                �[��next = .Cells(i + 1, keyCol + 0)
                ���next = .Cells(i + 1, keyCol + 1)
                cavNext = CStr(.Cells(i, keyCol + 2))
                
                If �[�� & ��� <> �[��next & ���next Then
                    If flg = True And �h���敪 <> "2" Then
                        For i2 = firstRow To i
                            
                            If .Cells(i2, addCol) = "" Then
                                .Cells(i2, keyCol + 5).Interior.color = RGB(146, 204, 255)
                            End If
                        Next i2
                    End If
                    firstRow = i + 1
                    flg = False
                End If
nextI:
            Next i
        Next r
        
    End With
    
     '�\�[�g
    With myBook.Sheets(newSheetName)
        .Select
        .Range(Columns(keyCol - 1), Columns(keyCol + 6)).AutoFit
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call �œK�����ǂ�

End Function

Public Function ���_�A����_�}���}()
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim mySheetName2 As String: mySheetName2 = "PVSW_RLTF"
    Dim mySheetName3 As String: mySheetName3 = "��A��_�}���}"
    
    Dim i As Long

    With Workbooks(myBookName).Sheets(mySheetName2)
        Dim ���i�i��RAN As Range
        Dim my�^�C�g��Col As Long
        Dim my�^�C�g��Row As Long: my�^�C�g��Row = .Cells.Find("�d�����ʖ�", , , xlWhole).Row
        Dim my���i�g����Ran0 As Long, my���i�g����Ran1 As Long
        For i = 1 To .Columns.count
            If Len(.Cells(my�^�C�g��Row, i)) = 15 Then
                If my���i�g����Ran0 = 0 Then my���i�g����Ran0 = i
            Else
                If my���i�g����Ran0 <> 0 Then my���i�g����Ran1 = i - 1: Exit For
            End If
        Next i
        Set ���i�i��RAN = .Range(.Cells(my�^�C�g��Row, my���i�g����Ran0), .Cells(my�^�C�g��Row, my���i�g����Ran1))
    End With
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim �ύX�Oc As Long: �ύX�Oc = .Cells.Find("�}", , , xlWhole).Column
        Dim �ύX�Or As Long: �ύX�Or = .Cells.Find("�}", , , xlWhole).Row
        Dim �ύX��c As Long: �ύX��c = .Cells.Find("�}1", , , xlWhole).Column
        Dim �[��c As Long: �[��c = .Cells.Find("�[����", , , xlWhole).Column
        Dim �\��c As Long: �\��c = .Cells.Find("�\��", , , xlWhole).Column
        Dim �T�C�Yc As Long: �T�C�Yc = .Cells.Find("�T�C�Y", , , xlWhole).Column
        Dim �Fc As Long: �Fc = .Cells.Find("�F�ď�", , , xlWhole).Column
        Dim ��c As Long: ��c = .Cells.Find("��", , , xlWhole).Column
        Dim ��s As Variant
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �[��c).End(xlUp).Row
        Dim �ύX��� As String
        
        Dim �ύX�O As String, �ύX�� As String, �[�� As String, �\�� As String, �� As String, �T�C�Y As String, �F As String
        Dim ���i�g����Ran As Range
        For i = �ύX�Or + 1 To lastRow
            �ύX�O = .Cells(i, �ύX�Oc)
            �ύX�� = .Cells(i, �ύX��c)
            If �ύX�O <> �ύX�� Then
                �ύX��� = ""
                If �ύX�O = "" Then
                    �ύX��� = "ADD"
                ElseIf �ύX�� = "" Then
                    �ύX��� = "DEL"
                Else
                    �ύX��� = "CH"
                End If
                �\�� = .Cells(i, �\��c)
                �T�C�Y = .Cells(i, �T�C�Yc)
                �F = .Cells(i, �Fc)
                �� = .Cells(i, ��c)
                Set ���i�g����Ran = .Range(.Cells(i, my���i�g����Ran0), .Cells(i, my���i�g����Ran1))
            
            With Workbooks(myBookName).Sheets(mySheetName2)
                Dim �n�_��c As Long: �n�_��c = .Cells.Find("�n�_����H����", , , xlWhole).Column
                Dim �n�_�[��c As Long: �n�_�[��c = .Cells.Find("�n�_���[�����ʎq", , , xlWhole).Column
                Dim �n�_cavC As Long: �n�_cavC = .Cells.Find("�n�_���L���r�e�BNo.", , , xlWhole).Column
                Dim �I�_��c As Long: �I�_��c = .Cells.Find("�I�_����H����", , , xlWhole).Column
                Dim �I�_�[��c As Long: �I�_�[��c = .Cells.Find("�I�_���[�����ʎq", , , xlWhole).Column
                Dim �I�_cavC As Long: �I�_cavC = .Cells.Find("�I�_���L���r�e�BNo.", , , xlWhole).Column
                Dim addRow As Long
                Dim �n�_�� As String, �n�_�[�� As String, �n�_cav As String, �I�_�� As String, �I�_�[�� As String, �I�_cav As String
                �n�_�� = .Cells(i, �n�_��c)
                �n�_�[�� = .Cells(i, �n�_�[��c)
                �n�_cav = .Cells(i, �n�_cavC)
                �I�_�� = .Cells(i, �I�_��c)
                �I�_�[�� = .Cells(i, �I�_�[��c)
                �I�_cav = .Cells(i, �I�_cavC)
            End With
            With Workbooks(myBookName).Sheets(mySheetName3)
                Dim outFirstRow As Long
                Dim out�\��r As Long: out�\��r = .Cells.Find("�\��" & Chr(10) & "W-No.", , , xlWhole).Row
                Dim out�\��c As Long: out�\��c = .Cells.Find("�\��" & Chr(10) & "W-No.", , , xlWhole).Column
                If outFirstRow = 0 Then outFirstRow = .Cells(.Rows.count, out�\��c).End(xlUp).Row + 1
                Dim out������c As Long: out������c = .Cells.Find("������_", , , 1).Column
                Dim out�T�C�Yc As Long: out�T�C�Yc = .Cells.Find("�T�C�Y" & Chr(10) & "Size", , , xlWhole).Column
                Dim out�Fc As Long: out�Fc = .Cells.Find("�F" & Chr(10) & "Color", , , xlWhole).Column
                Dim out�n�_��c As Long: out�n�_��c = .Cells.Find("�n�_��", , , 1).Column
                Dim out�n�_�[��c As Long: out�n�_�[��c = .Cells.Find("�[��" & Chr(10) & "Tno", , , xlWhole).Column
                Dim out�n�_��c As Long: out�n�_��c = .Cells.Find("��" & Chr(10) & "Cno", , , xlWhole).Column
                Dim out�n�_��c As Long: out�n�_��c = .Cells.Find("��H����" & Chr(10) & "Circuit", , , xlWhole).Column
                Dim out�n�_�}���}�Oc As Long: out�n�_�}���}�Oc = .Cells.Find("�}���}" & Chr(10) & "�ύX�O", , , xlWhole).Column
                Dim out�n�_����c As Long: out�n�_����c = .Cells.Find("����", , , xlWhole).Column
                Dim out�n�_�}���}��c As Long: out�n�_�}���}��c = .Cells.Find("�}���}" & Chr(10) & "�ύX��", , , xlWhole).Column
                Dim out�I�_��c As Long: out�I�_��c = .Cells.Find("�I�_��", , , 1).Column
                Dim out�I�_�[��c As Long: out�I�_�[��c = .Cells.Find("�[��" & Chr(10) & "Tno_", , , xlWhole).Column
                Dim out�I�_��c As Long: out�I�_��c = .Cells.Find("��" & Chr(10) & "Cno_", , , xlWhole).Column
                Dim out�I�_��c As Long: out�I�_��c = .Cells.Find("��H����" & Chr(10) & "Circuit_", , , xlWhole).Column
                Dim out�I�_�}���}�Oc As Long: out�I�_�}���}�Oc = .Cells.Find("�}���}" & Chr(10) & "�ύX�O_", , , xlWhole).Column
                Dim out�I�_����c As Long: out�I�_����c = .Cells.Find("����_", , , xlWhole).Column
                Dim out�I�_�}���}��c As Long: out�I�_�}���}��c = .Cells.Find("�}���}" & Chr(10) & "�ύX��_", , , xlWhole).Column
                Dim outKeyc As Long: outKeyc = .Cells.Find("key_", , , xlWhole).Column
                Dim out���i�i��c As Long: out���i�i��c = .Cells.Find("���i�i��", , , xlWhole).Column
'                Dim key As Range: Set key = .Columns(outKeyc).Find(Val(��s(0)), , , xlWhole)
'                If key Is Nothing Then
'                    addRow = .Cells(.Rows.Count, out�\��c).End(xlUp).Row + 1
'                Else
'                    addRow = key.Row
'                End If
                If �\�� = "0181" Then Stop
                addRow = .Cells(.Rows.count, out�\��c).End(xlUp).Row + 1
                Dim FoundCell As Range: Set FoundCell = .Range(.Cells(outFirstRow, out�\��c), .Cells(addRow, out�\��c)).Find(�\��, , , 1)
                Dim FirstCell As Range: Set FirstCell = FoundCell
                Dim foundCells As Range: Set foundCells = FoundCell
                If Not (FoundCell Is Nothing) Then
                    Do
                        Set FoundCell = .Range(.Cells(outFirstRow, out�\��c), .Cells(addRow, out�\��c)).FindNext(FoundCell)
                        If FoundCell.address = FirstCell.address Then
                            Exit Do
                        Else
                            Set foundCells = Union(foundCells, FoundCell)
                        End If
                    Loop
                End If
                
                For yy = 1 To foundCells.count
                    For aa = my���i�g����Ran0 To my���i�g����Ran1
                        '.cells(���i�g����ran(a)
                    Next aa
                Next yy
                
                If FoundCell Is Nothing Then
                    addRow = .Cells(.Rows.count, out�\��c).End(xlUp).Row + 1
                Else
                    'addRow = .Row
                End If
                
                'addRow = .Cells(.Rows.Count, out�\��c).End(xlUp).Row + 1
                Dim a As Long
                For a = my���i�g����Ran0 To my���i�g����Ran1
                    .Cells(out�\��r, out���i�i��c + a - 1) = ���i�i��RAN(a)
                    .Cells(addRow, out���i�i��c + a - 1) = ���i�g����Ran(a)
                Next
                '.Cells(addRow, outKeyc) = ��s(0)
                .Cells(addRow, out�\��c).NumberFormat = "@"
                .Cells(addRow, out�\��c).Value = �\��
                .Cells(addRow, out�T�C�Yc) = �T�C�Y
                .Cells(addRow, out�Fc) = �F
                .Cells(addRow, out�n�_�[��c) = �n�_�[��
                .Cells(addRow, out�n�_��c) = �n�_cav
                .Cells(addRow, out�n�_��c) = �n�_��
                .Cells(addRow, out�I�_�[��c) = �I�_�[��
                .Cells(addRow, out�I�_��c) = �I�_cav
                .Cells(addRow, out�I�_��c) = �I�_��
                If �� = "�n" Then
                    .Cells(addRow, out�n�_�}���}�Oc) = �ύX�O
                    .Cells(addRow, out�n�_����c) = �ύX���
                    .Cells(addRow, out�n�_�}���}��c) = �ύX��
                    .Cells(addRow, out�n�_�}���}��c).Font.Bold = True
                    .Cells(addRow, out�n�_�}���}��c).Interior.color = vbRed
                End If
                If �� = "�I" Then
                    .Cells(addRow, out�I�_�}���}�Oc) = �ύX�O
                    .Cells(addRow, out�I�_����c) = �ύX���
                    .Cells(addRow, out�I�_�}���}��c) = �ύX��
                    .Cells(addRow, out�I�_�}���}��c).Font.Bold = True
                    .Cells(addRow, out�I�_�}���}��c).Interior.color = vbRed
                End If
                .Cells(addRow, out������c) = Date
            End With
            End If
        Next i
    End With
    
    With Workbooks(myBookName).Sheets(mySheetName3)
        '�r��
        With .Range(.Cells(out�\��r, 1), .Cells(addRow, out���i�i��c + my���i�g����Ran1 - 1))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        .Range(.Cells(out�\��r - 1, out�n�_��c), .Cells(addRow, out�n�_��c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out�\��r - 1, out�I�_��c), .Cells(addRow, out�I�_��c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out�\��r - 1, out���i�i��c), .Cells(addRow, out���i�i��c)).Borders(1).Weight = xlMedium
        '�\�[�g
        With .Sort.SortFields
            .Clear
            .add key:=Cells(out�\��r, out������c), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(out�\��r, out�\��c), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '       .Add key:=Cells(out�\��r, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 4), Order:=xlAscending, DataOption:=0
    '            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Activate
        .Sort.SetRange .Range(.Rows(out�\��r), Rows(addRow))
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
End Function

Public Function ���_�A����_�}���}_Ver2002()

    aa = MsgBox("���̃V�[�g��[�}]��[�}1]�ɈႢ������ӏ����A���}���}�ύX�Ƃ��č쐬���܂��B" & vbCrLf, vbYesNo, "�}���}��A���̍쐬")
    If aa <> vbYes Then End
    
    '���i�i�Ԃ̃V�[�g���Z�b�g
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")

    'Stop '�}���}���Ă̑Ώې��i�i�Ԃ��擾
        
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim mySheetName3 As String: mySheetName3 = "��A��_�}���}"
    Dim i As Long
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim myKey As Range: Set myKey = .Cells.Find("�[�����i��", , , 1)
        '���i�i�Ԃ��Z�b�g
        ReDim �}���}���i�i��(myKey.Column - 2, 1): ���i�i��head = ""
        For X = 0 To myKey.Column - 2
            If .Cells(myKey.Row - 1, X + 1) <> "" Then ���i�i��head = .Cells(myKey.Row - 1, X + 1)
            �}���}���i�i��(X, 0) = ���i�i��head & .Cells(myKey.Row, X + 1)
            For x2 = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
                If �}���}���i�i��(X, 0) = Replace(���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), x2), " ", "") Then
                    �}���}���i�i��(X, 1) = x2
                    Exit For
                End If
            Next x2
        Next X
        Dim �\��c As Long: �\��c = .Cells.Find("�\��", , , xlWhole).Column
        Dim �ύX�Oc As Long: �ύX�Oc = .Cells.Find("�}", , , xlWhole).Column
        Dim �ύX�Or As Long: �ύX�Or = .Cells.Find("�}", , , xlWhole).Row
        Dim �ύX��c As Long: �ύX��c = .Cells.Find("�}1", , , xlWhole).Column
        Dim �[��c As Long: �[��c = .Cells.Find("�[����", , , xlWhole).Column
        Dim �T�C�Yc As Long: �T�C�Yc = .Cells.Find("�T�C�Y", , , xlWhole).Column
        Dim �Fc As Long: �Fc = .Cells.Find("�F�ď�", , , xlWhole).Column
        Dim ��c As Long: ��c = .Cells.Find("��", , , xlWhole).Column
        Dim ��s As Variant
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �[��c).End(xlUp).Row
        Dim �ύX��� As String
line00:
        'mySQLtemp0�AmySQLtemp1�̍쐬
        mySQL0 = " SELECT * from [" & mySheetName & "$]"
        Call SQL_JUNK(mySQL0, mySheetName, 2, 1, �\��c - 1)
        If myErrFlg = True Then GoTo line00
        '��A��_�}���}�ւ̏o��
        mysql = " SELECT Products,�\��,�T�C�Y,�F�ď�,�[����,Cav,��,�},�}1,�[����_,Cav_,��_,�}_,�}1_ from [" & "SQLtemp1" & "$] "
        Call SQL_�}���}�ύX�˗�(mysql)
        Application.DisplayAlerts = False
        Sheets("SQLtemp0").Delete
        Sheets("SQLtemp1").Delete
        Application.DisplayAlerts = True
    End With
    
    MsgBox "�������������܂���"
    
End Function
Public Function ���_�A����_����_Ver2001()
    
    ���� = "C"
    ���� = "543B_test"
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim mySheetName2 As String: mySheetName2 = "PVSW_RLTF"
    Dim mySheetName3 As String: mySheetName3 = "��A��_����"
    Dim i As Long
    
    Call ���i�i��RAN_set2(���i�i��RAN, ����, "����", "")
    
    Call SQL_�ύX�˗�_����(���i�i��RAN, �����ύXRAN, myBookName)
    
    With Workbooks(myBookName).Sheets(mySheetName3)
        Dim key As Range: Set key = .Cells.Find("����", , , 1)
        Dim keyCol As Long: keyCol = .Cells.Find("����", , , 1).Column
        Dim ���lCol As Long: ���lCol = .Cells.Find("���l" & vbLf & "remarks" & vbLf & vbLf, , , 1).Column
        Dim addRow As Long: addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
        Dim ���i�i��RANCol() As Long
        ReDim ���i�i��RANCol(���i�i��RANc - 1)
        '���i�i�Ԃ̗�ԍ����Z�b�g
        For i = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
            Set myfind = .Rows(key.Row).Find(���i�i��RAN(1, i), , , 1)
            If myfind Is Nothing Then
                .Columns(���lCol).Insert
                .Cells(key.Row, ���lCol) = ���i�i��RAN(1, i)
                ���i�i��RANCol(i) = ���lCol
                ���lCol = ���lCol + 1
            Else
                ���i�i��RANCol(i) = myfind.Column
            End If
        Next i
    End With
    
    For i = LBound(�����ύXRAN, 2) To UBound(�����ύXRAN, 2)
        With Workbooks(myBookName).Sheets(mySheetName3)
            .Cells(addRow, keyCol + 0) = ����
            .Cells(addRow, keyCol + 3) = �����ύXRAN(���i�i��RANc + 0, i)
            .Cells(addRow, keyCol + 4) = �����ύXRAN(���i�i��RANc + 1, i)
            .Cells(addRow, keyCol + 5) = �����ύXRAN(���i�i��RANc + 2, i)
            .Cells(addRow, keyCol + 7) = �����ύXRAN(���i�i��RANc + 3, i)
            .Cells(addRow, keyCol + 9) = �����ύXRAN(���i�i��RANc + 4, i)
            For ii = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
                .Cells(addRow, ���i�i��RANCol(ii)) = �����ύXRAN(ii, i)
            Next ii
            .Cells(addRow, keyCol + 11 + UBound(���i�i��RAN, 2)) = �����ύXRAN(���i�i��RANc + 6, i)
        End With
    Next i
    
    MsgBox "�������������܂���"
    
End Function

Public Function ���i�ʒ[���ꗗ�̃V�[�g�쐬_1800()
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "���i�ʒ[���ꗗ"
    
    
    With Workbooks(myBookName).Sheets("���i�i��")
        �n���}�A�h���X = .Cells.Find("System+", , , 1).Offset(0, 1).Value
    End With
    
    With Workbooks(myBookName).Sheets(mySheetName)
        'PVSW_RLTF����̃f�[�^
        Dim my�^�C�g��Row As Long: my�^�C�g��Row = .Cells.Find("�i��_").Row
        Dim my�^�C�g��Col As Long: my�^�C�g��Col = .Cells.Find("�i��_").Column
        Dim my�^�C�g��Ran As Range: Set my�^�C�g��Ran = .Range(.Cells(my�^�C�g��Row, 1), .Cells(my�^�C�g��Row, my�^�C�g��Col))
        Dim my�d�����ʖ�Col As Long: my�d�����ʖ�Col = .Cells.Find("�d�����ʖ�").Column
        Dim my��1Col As Long: my��1Col = .Cells.Find("�n�_����H����").Column
        Dim my�[��1Col As Long: my�[��1Col = .Cells.Find("�n�_���[�����ʎq").Column
        Dim myCav1Col As Long: myCav1Col = .Cells.Find("�n�_���L���r�e�BNo.").Column
        Dim my��2Col As Long: my��2Col = .Cells.Find("�I�_����H����").Column
        Dim my�[��2Col As Long: my�[��2Col = .Cells.Find("�I�_���[�����ʎq").Column
        Dim myCav2Col As Long: myCav2Col = .Cells.Find("�I�_���L���r�e�BNo.").Column
        Dim my����Col As Long: my����Col = .Cells.Find("����No").Column
        Dim my�����i��Col As Long: my�����i��Col = .Cells.Find("�����i��").Column
        Dim myJoint1Col As Long: myJoint1Col = .Cells.Find("�n�_��JOINT���").Column
        Dim myJoint2Col As Long: myJoint2Col = .Cells.Find("�I�_��JOINT���").Column
        Dim my�_�u����1Col As Long: my�_�u����1Col = .Cells.Find("�n�_���_�u����H����").Column
        Dim my�_�u����2Col As Long: my�_�u����2Col = .Cells.Find("�I�_���_�u����H����").Column
        
        Dim myPVSW�i��col As Long: myPVSW�i��col = .Cells.Find("�d���i��").Column
        Dim myPVSW�T�C�Ycol As Long: myPVSW�T�C�Ycol = .Cells.Find("�d���T�C�Y").Column
        Dim myPVSW�Fcol As Long: myPVSW�Fcol = .Cells.Find("�d���F").Column
        Dim my�}���}11Col As Long: my�}���}11Col = .Cells.Find("�n�_���}���}�F�P").Column
        Dim my�}���}12Col As Long: my�}���}12Col = .Cells.Find("�n�_���}���}�F�Q").Column
        Dim my�}���}21Col As Long: my�}���}21Col = .Cells.Find("�I�_���}���}�F�P").Column
        Dim my�}���}22Col As Long: my�}���}22Col = .Cells.Find("�I�_���}���}�F�Q").Column
        
        Dim my���i11Col As Long: my���i11Col = .Cells.Find("�n�_���[�q�i��").Column
        Dim my���i21Col As Long: my���i21Col = .Cells.Find("�I�_���[�q�i��").Column
        Dim my���i12Col As Long: my���i12Col = .Cells.Find("�n�_���S����i��").Column
        Dim my���i22Col As Long: my���i22Col = .Cells.Find("�I�_���S����i��").Column
        Dim my���1Col As Long: my���1Col = .Cells.Find("�n�_����햼��").Column
        Dim my���2Col As Long: my���2Col = .Cells.Find("�I�_����햼��").Column
        Dim my���Ӑ�1Col As Long: my���Ӑ�1Col = .Cells.Find("�n�_���[�����Ӑ�i��").Column
        Dim my���1Col As Long: my���1Col = .Cells.Find("�n�_���[�����i��").Column
        Dim my���Ӑ�2Col As Long: my���Ӑ�2Col = .Cells.Find("�I�_���[�����Ӑ�i��").Column
        Dim my���2Col As Long: my���2Col = .Cells.Find("�I�_���[�����i��").Column
        Dim myJointGCol As Long: myJointGCol = .Cells.Find("�W���C���g�O���[�v").Column
        Dim myAB�敪Col As Long: myAB�敪Col = .Cells.Find("A/B�EB/C�敪").Column
        Dim my�d��YBMCol As Long: my�d��YBMCol = .Cells.Find("�d���x�a�l").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(my�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Set my�^�C�g��Ran = Nothing
        'NMB����̃f�[�^
        Dim my�i��Col As Long: my�i��Col = .Cells.Find("�i��_").Column
        Dim my�T�C�YCol As Long: my�T�C�YCol = .Cells.Find("�T�C�Y_").Column
        Dim my�T�C�Y��Col As Long: my�T�C�Y��Col = .Cells.Find("�T��_").Column
        Dim my�FCol As Long: my�FCol = .Cells.Find("�F_").Column
        Dim my�F��Col As Long: my�F��Col = .Cells.Find("�F��_").Column
        Dim my����Col As Long: my����Col = .Cells.Find("����_").Column
        Dim myPVSWtoNMB As Long: myPVSWtoNMB = .Cells.Find("RLTFtoPVSW_").Column
        
        Dim my���i�i��Ran0 As Long, my���i�i��Ran1 As Long, X As Long
        For X = 1 To myLastCol
            If Len(.Cells(my�^�C�g��Row, X)) = 15 Then
                If my���i�i��Ran0 = 0 Then my���i�i��Ran0 = X
            Else
                If my���i�i��Ran0 <> 0 Then my���i�i��Ran1 = X - 1: Exit For
            End If
        Next X
        
        'Dictionary
        Dim myDic As Object, myKey, myItem
        Dim myVal, myVal2, myVal3
        Set myDic = CreateObject("Scripting.Dictionary")
        myVal = .Range(.Cells(1, 1), .Cells(myLastRow, myLastCol))
    End With
    
    '���[�N�V�[�g�̒ǉ�
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            'MsgBox "���� " & newSheetName & " �̃V�[�g�������݂��܂��B" & vbCrLf _
                   & vbCrLf & _
                   "�����̃V�[�g���폜���邩�A�V�[�g����ύX���Ă�����s���ĉ������B"
            'Exit Function
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Cells.NumberFormat = "@"
    
    '�o�^�p�z��̐錾
    Dim �o�^dB As Variant
    ReDim �o�^dB(3, 0)
    Dim �o�^dBcount As Long
    Dim �o�^fLag As Boolean, xx As Long
    Dim my���Col As Long, my�[��Col As Long, db As Long
    
    'PVSW_RLTF to ���i�ʒ[���ꗗ
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    For i = my�^�C�g��Row To myLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            If i = my�^�C�g��Row Then Set ���i�i��RAN = .Range(.Cells(i, my���i�i��Ran0), .Cells(i, my���i�i��Ran1))
            Dim ���i�g����str As String: ���i�g����str = ""
            For i2 = 1 To my���i�i��Ran1
                If .Cells(i, i2) = "" Then
                    ���i�g����str = ���i�g����str & "0"
                Else
                    ���i�g����str = ���i�g����str & "1"
                End If
            Next i2
            Dim �d�����ʖ� As String: �d�����ʖ� = .Cells(i, my�d�����ʖ�Col)
            Dim ��1 As String: ��1 = .Cells(i, my��1Col)
            Dim �[��1 As String: �[��1 = .Cells(i, my�[��1Col)
            Dim Cav1 As String: Cav1 = .Cells(i, myCav1Col)
            Dim ��2 As String: ��2 = .Cells(i, my��2Col)
            Dim �[��2 As String: �[��2 = .Cells(i, my�[��2Col)
            Dim cav2 As String: cav2 = .Cells(i, myCav2Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim �����i�� As Range: Set �����i�� = .Cells(i, my�����i��Col)
            Dim �V�[���h�t���O As String: If �����i��.Interior.color = 9868950 Then �V�[���h�t���O = "S" Else �V�[���h�t���O = ""
            Dim Joint1 As String: Joint1 = .Cells(i, myJoint1Col)
            Dim Joint2 As String: Joint2 = .Cells(i, myJoint2Col)
            Dim �_�u����1 As String: �_�u����1 = .Cells(i, my�_�u����1Col)
            Dim �_�u����2 As String: �_�u����2 = .Cells(i, my�_�u����2Col)
            Dim ���i11 As String: ���i11 = .Cells(i, my���i11Col)
            Dim ���i21 As String: ���i21 = .Cells(i, my���i21Col)
            Dim ���i12 As String: ���i12 = .Cells(i, my���i12Col)
            Dim ���i22 As String: ���i22 = .Cells(i, my���i22Col)
            Dim ���1 As String: ���1 = .Cells(i, my���1Col)
            Dim ���2 As String: ���2 = .Cells(i, my���2Col)
            Dim ���Ӑ�1 As String: ���Ӑ�1 = .Cells(i, my���Ӑ�1Col)
            Dim ���1 As String: ���1 = .Cells(i, my���1Col)
            Dim ���Ӑ�2 As String: ���Ӑ�2 = .Cells(i, my���Ӑ�2Col)
            Dim ���2 As String: ���2 = .Cells(i, my���2Col)
            Dim JointG As String: JointG = .Cells(i, myJointGCol)
            Dim �d���i�� As String: �d���i�� = .Cells(i, myPVSW�i��col)
            Dim �d���T�C�Y As String: �d���T�C�Y = .Cells(i, myPVSW�T�C�Ycol)
            Dim �d���F As String: �d���F = .Cells(i, myPVSW�Fcol)
            Dim �}���}11 As String: �}���}11 = .Cells(i, my�}���}11Col)
            Dim �}���}12 As String: �}���}12 = .Cells(i, my�}���}12Col)
            Dim �}���}21 As String: �}���}21 = .Cells(i, my�}���}21Col)
            Dim �}���}22 As String: �}���}22 = .Cells(i, my�}���}22Col)
            Dim AB�敪 As String: AB�敪 = .Cells(i, myAB�敪Col)
            Dim �d��YBM As String: �d��YBM = .Cells(i, my�d��YBMCol)
            
            Dim ���葤1 As String, ���葤2 As String
            If Len(cav2) < 4 Then ���葤1 = �[��2 & "_" & String(3 - Len(cav2), " ") & cav2 & "_" & ��2
            If Len(Cav1) < 4 Then ���葤2 = �[��1 & "_" & String(3 - Len(Cav1), " ") & Cav1 & "_" & ��1
            'NMB����̃f�[�^
            Dim �i�� As String: �i�� = .Cells(i, my�i��Col)
            Dim �T�C�Y As String: �T�C�Y = .Cells(i, my�T�C�YCol)
            Dim �T�C�Y�� As String: �T�C�Y�� = .Cells(i, my�T�C�Y��Col)
            Dim �F As String: �F = .Cells(i, my�FCol)
            Dim �F�� As String: �F�� = .Cells(i, my�F��Col)
            Dim ���� As String: ���� = .Cells(i, my����Col)
            Dim PVSWtoNMB As String: PVSWtoNMB = .Cells(i, myPVSWtoNMB)
        End With
        
        With Workbooks(myBookName).Sheets(newSheetName)
            Dim �D��1 As Long, �D��2 As Long, �D��3 As Long
            Dim addRow As Long: addRow = .Cells(.Rows.count, my�d�����ʖ�Col).End(xlUp).Row + 1
            Dim ���i�g���� As Variant
            Dim hh As Long, ���i�g����val As String, �g���� As String
            If .Cells(2, 1) = "" Then
                Dim addCol As Long, ���i�i�� As Variant
                addCol = 0
                .Cells(2, addCol + 1) = "�[�����i��": �D��2 = addCol + 1
                .Cells(2, addCol + 2) = "�[����": �D��1 = addCol + 2
                .Rows(1).NumberFormat = "@"
            Else
                'NMB�̗L���m�F
                If PVSWtoNMB = "Found" Then
                    For xx = 1 To 2
                        Select Case xx
                        Case 1
                            my���Col = my���1Col
                            my�[��Col = my�[��1Col
                        Case 2
                            my���Col = my���2Col
                            my�[��Col = my�[��2Col
                        End Select
                        '�z��̓o�^�L�����m�F
                        �o�^fLag = 0
                        For db = 1 To �o�^dBcount
                            If CStr(�o�^dB(1, db)) = CStr(myVal(i, my���Col)) And CStr(�o�^dB(2, db)) = CStr(myVal(i, my�[��Col)) Then
'                            '�L��̂Ŏg�p���i�i�Ԃ�ǉ�����
'                                ���i�g����str = ""
'                                For Each ���i�g���� In ���i�i��v
'                                    If ���i�g���� = "" Then ���i�g���� = 0
'                                    ���i�g����str = ���i�g����str & ���i�g����
'                                Next ���i�g����
                                ���i�g����val = ""
                                For hh = 1 To Len(���i�g����str)
                                    �g���� = Mid(String(Len(���i�g����str) - Len(�o�^dB(3, db)), "0") & �o�^dB(3, db), hh, 1) Or Mid(���i�g����str, hh, 1)
                                   ���i�g����val = ���i�g����val & �g����
                                Next hh
                                �o�^dB(3, db) = ���i�g����val
                                �o�^fLag = 1
                                Exit For
                            End If
                        Next
                        '����������o�^
                        If �o�^fLag = 0 Then
                            �o�^dBcount = �o�^dBcount + 1
                            'ReDim Preserve �o�^dB(3) As Integer
                            ReDim Preserve �o�^dB(3, �o�^dBcount) As Variant
                            �o�^dB(1, �o�^dBcount) = myVal(i, my���Col)
                            �o�^dB(2, �o�^dBcount) = myVal(i, my�[��Col)
'                            ���i�g����str = ""
'                            For Each ���i�g���� In ���i�i��v
'                                If ���i�g���� = "" Then ���i�g���� = 0
'                                ���i�g����str = ���i�g����str & ���i�g����
'                            Next ���i�g����
                            �o�^dB(3, �o�^dBcount) = ���i�g����str
                        End If
                    Next xx
                End If
            End If
        End With
    Next i
    
    With Workbooks(myBookName).Sheets(newSheetName)
        For db = 1 To �o�^dBcount
            .Cells(db + 2, addCol + 1) = �o�^dB(1, db)
            .Cells(db + 2, addCol + 2) = �o�^dB(2, db)
            �o�^dB(3, db) = String(Len(���i�g����str) - Len(�o�^dB(3, db)), "0") & �o�^dB(3, db)
            For i = 1 To Len(���i�g����str)
                If Mid(�o�^dB(3, db), i, 1) <> 0 Then
                    .Cells(db + 2, addCol + 2 + i) = Mid(�o�^dB(3, db), i, 1)
                End If
            Next i
        Next db
        Set myDic = Nothing
        
        For Each ���i�i�� In ���i�i��RAN
            addCol = addCol + 1
            .Cells(2, addCol + 2) = ���i�i��
            .Cells(1, addCol + 2) = Mid(���i�i��, 8, 3)
        Next
    End With
    
    '���בւ�
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, �D��2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            '.Add key:=Range(Cells(1, �D��3).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(3), Rows(�o�^dBcount + 2))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    '���X���[�N���b�v
    
    With Workbooks(myBookName).Sheets(newSheetName)
        addRow = 3
        addCol = .Cells(2, .Columns.count).End(xlToLeft).Column + 2
        .Cells(addRow - 1, addCol) = "CLIP"
    End With
    
    '���h���R�l�N�^�ꗗ�쐬
    Dim �g�p���i_�[�� As String
    With Workbooks(myBookName).Sheets(newSheetName)
        For i = 3 To �o�^dBcount + 2
            If InStr(�g�p���i_�[��, .Cells(i, 1) & "_" & .Cells(i, 2)) = 0 Then
                �g�p���i_�[�� = �g�p���i_�[�� & "," & .Cells(i, 1) & "_" & .Cells(i, 2)
            End If
        Next i
    End With

    '���W�f�[�^�̓Ǎ���(�C���|�[�g�t�@�C��)
    Dim TargetName As String: TargetName = "CAV���W.txt"
    Dim Target As New FileSystemObject
    Dim TargetFile As String
    TargetFile = �n���}�A�h���X & "\00_�V�X�e���p�[�c\" & TargetName
    Dim intFino As Variant
    intFino = FreeFile
    Open TargetFile For Input As #intFino
    Dim outY As Long: outY = 1
    Dim outX As Long
    Dim lastgyo As Long: lastgyo = 1
    Dim fileCount As Long: fileCount = 0
    Dim inX As Long
    Dim temp
    Dim �g�p���i_�[��s As Variant
    Dim �g�p���i_�[��c As Variant
    Dim aa As Variant
    Dim ���W����Flag As Boolean
    Dim c As Variant, �g�p���istr As String
    
    '�g�p���iStr�ɁA����g�p���镔�i�i�ԍ��W�f�[�^��S�ē����
    �g�p���i_�[��s = Split(�g�p���i_�[��, ",")
    For Each �g�p���i_�[��c In �g�p���i_�[��s
        If �g�p���i_�[��c <> "" Then
            c = Split(�g�p���i_�[��c, "_")
            ���W����Flag = False
            '�ʐ^��T��
            intFino = FreeFile
            Open TargetFile For Input As #intFino
            Do Until EOF(intFino)
                Line Input #intFino, aa
                temp = Split(aa, ",")
                If "�ʐ^" = temp(8) Then
                    If Replace(temp(0), "-", "") = c(0) Then
                        If temp(7) = "Cir" Then
                            �g�p���istr = �g�p���istr & "," & temp(0) & "_" & c(1) & "_" & temp(1) & "_" & temp(4) & "_" & temp(5)
                        End If
                        ���W����Flag = True
                    Else
                        If ���W����Flag = True Then Exit Do
                    End If
                End If
            Loop
            Close #intFino
            
            '�ʐ^�������̂ŗ��}��T��
            If ���W����Flag = False Then
                intFino = FreeFile
                Open TargetFile For Input As #intFino
                Do Until EOF(intFino)
                    Line Input #intFino, aa
                    temp = Split(aa, ",")
                    If "���}" = temp(8) Then
                        If Replace(temp(0), "-", "") = c(0) Then
                            If temp(7) = "Cir" Then
                                �g�p���istr = �g�p���istr & "," & temp(0) & "_" & c(1) & "_" & temp(1) & "_" & temp(4) & "_" & temp(5)
                            End If
                            ���W����Flag = True
                        Else
                            If ���W����Flag = True Then Exit Do
                        End If
                    End If
                Loop
                Close #intFino
            End If
        End If
    Next �g�p���i_�[��c
    
    Dim �g�p���is As String, �g�p���ic As Variant, �g�p As Variant, �g�p���i As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        addRow = 3
        addCol = .Cells(2, .Columns.count).End(xlToLeft).Column + 2
        .Cells(addRow - 1, addCol + 0) = "�h���R�l�N�^�i��"
        .Cells(addRow - 1, addCol + 1) = "�[����_"
        .Cells(addRow - 1, addCol + 2) = "cav"
        .Cells(addRow - 1, addCol + 3) = "width"
        .Cells(addRow - 1, addCol + 4) = "height"
        .Cells(addRow - 1, addCol + 5) = "EmptyPlug"
        .Cells(addRow - 1, addCol + 6) = "PlugColor"
        �g�p���ic = Split(�g�p���istr, ",")
        For Each �g�p���i In �g�p���ic
            If �g�p���i <> "" Then
                �g�p = Split(�g�p���i, "_")
                .Cells(addRow, addCol + 0) = �g�p(0)
                .Cells(addRow, addCol + 1) = �g�p(1)
                .Cells(addRow, addCol + 2) = �g�p(2)
                .Cells(addRow, addCol + 3) = �g�p(3)
                .Cells(addRow, addCol + 4) = �g�p(4)
                addRow = addRow + 1
            End If
        Next �g�p���i
    End With
    
End Function

Public Function ���i�ʒ[���ꗗ�̃V�[�g�쐬_2009()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�[���ꗗ"
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Call SQL_���i�ʒ[���ꗗ(RAN, ���i�i��RAN, myBook)
      
    '�V�[�g��:���i�ʒ[���ꗗ��������΍쐬
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    '�t�B�[���h���̃Z�b�g
    With ActiveWorkbook.Sheets("�t�B�[���h��")
        Set myKey = .Cells.Find("�t�B�[���h��_�[���ꗗ", , , 1)
        Set myArea = .Range(myKey.Offset(1, 0).address, myKey.Offset(2, 0).End(xlToRight).address)
    End With
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set myKey = .Cells.Find("�[�����i��", , , 1)
        'setup
        If myKey Is Nothing Then '�V�K�쐬�̎�
            Set myKey = .Cells(3, 1)
            Call �t�B�[���h���̒ǉ�(myBook.Sheets(newSheetName), myKey, myArea, "l")
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Else '���������鎞
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            ��� = RAN(0, Y)
            �[�� = RAN(1, Y)
            ���i = RAN(2, Y)
            �N�� = RAN(3, Y)
            If ��� & �[�� = "" Then GoTo line20
            '���i�i�Ԃ̗�
            Set fnd = .Rows(myKey.Row).Find(���i, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(myKey.Row, myKey.Column).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(���i, " ", "")) = 10 Then
                    ���iA = Mid(���i, 8, 3)
                Else
                    ���iA = Mid(���i, 5, 4)
                End If
                .Cells(myKey.Row - 0, incol) = ���i
                .Cells(myKey.Row - 1, incol) = ���iA
                .Cells(myKey.Row - 1, incol).ColumnWidth = Len(���iA) * 1.05
                .Cells(myKey.Row - 2, incol).NumberFormat = "mm/dd"
                .Cells(myKey.Row - 2, incol).ShrinkToFit = True
                .Cells(myKey.Row - 2, incol) = �N��
            Else
                incol = fnd.Column
                .Cells(myKey.Row - 2, incol).NumberFormat = "mm/dd"
                .Cells(myKey.Row - 2, incol) = �N��
            End If
            
            '�o�^�����邩�m�F
            For i = myKey.Row + 1 To lastRow
                flg = False
                If ��� = .Cells(i, myKey.Column) Then
                    If �[�� = .Cells(i, myKey.Column + 1) Then
                        flg = True
                        addRow = i
                        Exit For
                    End If
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                addRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, myKey.Column + 0) = ���
                .Cells(addRow, myKey.Column + 1) = �[��
            End If
            If .Cells(addRow, incol) = "" Then
                .Cells(addRow, incol) = "0"
            End If
line20:
        Next Y
    End With
    '�\�[�g
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(myKey.Row + 1, myKey.Column + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(myKey.Row + 1, myKey.Column + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(myKey.Row + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(myKey.Row + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Columns(1).ColumnWidth = 2
        .Cells(1, myKey.Column) = "�[���ꗗ"
        Set mykey0 = .Cells.Find("���^�p�x", , , 1)
        If mykey0 Is Nothing Then
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Offset(0, 1) = "���^�p�x"
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Interior.color = RGB(255, 255, 0)
        End If
        Set mykey0 = .Cells.Find("���^����", , , 1)
        If mykey0 Is Nothing Then
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Offset(0, 1) = "���^����"
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Interior.color = RGB(255, 255, 0)
        End If
        Set mykey0 = .Cells.Find("���^����", , , 1)
        '�r��������
        .Range(.Cells(myKey.Row, myKey.Column), .Cells(addRow, incol + 2)).Borders.LineStyle = True
        '.Range(.Cells(myKey.Row - 1, myKey.Column + 2), .Cells(addRow, myKey.Column)).Borders.LineStyle = True
    End With
    
    If RLTF�T�u = True Then
        With myBook.Sheets(newSheetName)
            addRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            �[��Col = .Cells.Find("�[����", , , 1).Column
            For X = myKey.Column + 2 To mykey0.Column
                ���i�i��str = .Cells(myKey.Row, X)
                For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                    If ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r) Then
                        �Ώۃt�@�C�� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "SUB"), r) & ".csv"
                        If Dir(myBook.Path & "\07_SUB\" & �Ώۃt�@�C��) <> "" Then
                            Call SUB�f�[�^�擾(SUB�f�[�^RAN, myBook.Path & "\07_SUB\" & �Ώۃt�@�C��)
                            Call SQL_�[���T�u�ꗗ(�[���T�uran, ���i�i��str, myBook)
                            For Y = myKey.Row + 1 To addRow
                                �[�����i��str = .Cells(Y, myKey.Column).Value
                                �[��str = .Cells(Y, �[��Col).Value
                                If .Cells(Y, X) = "" Then GoTo line30
                                For i = LBound(SUB�f�[�^RAN) + 1 To UBound(SUB�f�[�^RAN)
                                    '�t�B�[���h���̊m�F
                                    Dim SUB�f�[�^RANsp As Variant
                                    SUB�f�[�^RANsp = Split(SUB�f�[�^RAN(i), ",")
                                    If i = 1 Then
                                        For ii = LBound(SUB�f�[�^RANsp) To UBound(SUB�f�[�^RANsp)
                                            If SUB�f�[�^RANsp(ii) = "���i�i��" Then �[�����i��lng = ii
                                            If SUB�f�[�^RANsp(ii) = "�[��No." Then �[��lng = ii
                                            If SUB�f�[�^RANsp(ii) = "�T�uNo." Then �T�ulng = ii
                                        Next ii
                                    End If
                                    If �[��str = SUB�f�[�^RANsp(�[��lng) Then
                                        If �[�����i��str = SUB�f�[�^RANsp(�[�����i��lng) Then
                                            If SUB�f�[�^RANsp(�T�ulng) <> "" Then
                                                .Cells(Y, X) = SUB�f�[�^RANsp(�T�ulng)
                                                GoTo line30
                                            End If
                                        End If
                                    End If
                                Next i
                                'SUB�f�[�^�ɂȂ��ꍇ
                                For ii = LBound(�[���T�uran, 2) + 1 To UBound(�[���T�uran, 2)
                                    If �[��str = �[���T�uran(0, ii) Then
                                        If �[�����i��str = �[���T�uran(1, ii) Then
                                            
                                            .Cells(Y, X) = �[���T�uran(2, ii)
                                            GoTo line30
                                        End If
                                    End If
                                Next ii
                                '����ł�����������R�l�N�^�����̃T�u�₩��c
                                .Cells(Y, X) = "c"
line30:
                            Next Y
                        End If
                    End If
                Next r
            Next X
        End With
    End If
    
    MsgBox "�쐬���܂����B"
    
End Function

Public Function A_�d���ꗗ�̃V�[�g�쐬()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�d���ꗗ"
    
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    
    Call SQL_�d���ꗗ(RAN, ���i�i��RAN, myBook)
      
    '�V�[�g��:���i�ʒ[���ꗗ��������΍쐬
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("�i��", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "�i��"
            .Cells(keyRow, keyCol + 1) = "�T�C�Y"
            .Cells(keyRow, keyCol + 2) = "�T�C�Y��"
            .Cells(keyRow, keyCol + 3) = "�F"
            .Cells(keyRow, keyCol + 4) = "�F��"
            .Cells(keyRow, keyCol + 5) = "����"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
            
            .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            �i�� = RAN(0, Y)
            �T�C�Y = RAN(1, Y)
            �T�C�Y�� = RAN(2, Y)
            �F = RAN(3, Y)
            �F�� = RAN(4, Y)
            ���� = RAN(5, Y)
            ���i = RAN(6, Y)
            �ӏ��� = RAN(7, Y)
            If �i�� & �T�C�Y = "" Then GoTo line20
            '���i�i�Ԃ̗�
            Set fnd = .Rows(keyRow).Find(���i, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(���i, " ", "")) = 10 Then
                    ���iA = Mid(���i, 8, 3)
                Else
                    ���iA = Mid(���i, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = ���i
                .Cells(keyRow - 1, incol) = ���iA
                .Cells(keyRow - 1, incol).ColumnWidth = Len(���iA) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = �N��
            Else
                incol = fnd.Column
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol) = �N��
            End If
            
            '�o�^�����邩�m�F
            For i = keyRow + 1 To lastRow
                flg = False
                If �i�� = .Cells(i, keyCol) Then
                    If �T�C�Y = .Cells(i, keyCol + 1) Then
                        If �T�C�Y�� = .Cells(i, keyCol + 2) Then
                            If �F = .Cells(i, keyCol + 3) Then
                                If �F�� = .Cells(i, keyCol + 4) Then
                                    If ���� = .Cells(i, keyCol + 5) Then
                                        flg = True
                                        addRow = i
                                        .Cells(addRow, incol) = .Cells(addRow, incol) + �ӏ���
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = �i��
                .Cells(addRow, keyCol + 1) = �T�C�Y
                .Cells(addRow, keyCol + 2) = �T�C�Y��
                .Cells(addRow, keyCol + 3) = �F
                .Cells(addRow, keyCol + 4) = �F��
                .Cells(addRow, keyCol + 5) = ����
                .Cells(addRow, incol) = �ӏ���
            End If
line20:
        Next Y
    End With
    '�\�[�g
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�r��������
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = "�d���ꗗ"
    End With
    
    MsgBox "�쐬���܂����B"
    
End Function


Public Function A_�[�q�ꗗ�̃V�[�g�쐬()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�[�q�ꗗ"
    
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Call SQL_�[�q�ꗗ(RAN, ���i�i��RAN, myBook)
    
    With myBook.Sheets("�ݒ�")
        Set aKey = .Cells.Find("�[�q�t�@�~���[_", , , 1)
        �[�q�t�@�~���[ran = .Cells(aKey.Row, aKey.Column)
    End With
    
    With myBook.Sheets("PVSW_RLTF")
        Dim aCol As Long
        aCol = .Cells.Find("�n�_���[�q_", , , 1).Column
        Set �n�_�[�qRan = .Columns(aCol)
        aCol = .Cells.Find("�I�_���[�q_", , , 1).Column
        Set �I�_�[�qRan = .Columns(aCol)
    End With
    
    '�V�[�g��:���i�ʒ[���ꗗ��������΍쐬
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("�[�q�i��", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "�[�q�i��"
            .Cells(keyRow, keyCol + 1) = "�t���i��"
            .Cells(keyRow, keyCol + 2) = "Family"
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "����"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            �[�q�i�� = RAN(0, Y)
            �t���i�� = RAN(1, Y)
            ���b�L = RAN(2, Y)
            ���� = RAN(3, Y)
            ���i = RAN(4, Y)
            �ӏ��� = RAN(5, Y)
            If �[�q�i�� & ���� = "" Then GoTo line20
            '���i�i�Ԃ̗�
            Set fnd = .Rows(keyRow).Find(���i, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(���i, " ", "")) = 10 Then
                    ���iA = Mid(���i, 8, 3)
                Else
                    ���iA = Mid(���i, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = ���i
                .Cells(keyRow - 1, incol) = ���iA
                .Cells(keyRow - 1, incol).ColumnWidth = Len(���iA) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = �N��
            Else
                incol = fnd.Column
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol) = �N��
            End If
            
            '�o�^�����邩�m�F
            For i = keyRow + 1 To lastRow
                flg = False
                If �[�q�i�� = .Cells(i, keyCol) Then
                    If �t���i�� = .Cells(i, keyCol + 1) Then
                        If ���� = .Cells(i, keyCol + 5) Then
                            flg = True
                            addRow = i
                            .Cells(addRow, incol) = .Cells(addRow, incol) + �ӏ���
                            Exit For
                        End If
                    End If
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = �[�q�i��
                '�[�q�t�@�~���[�̎擾(����)
                Set myColor = �n�_�[�qRan.Find(�[�q�i��, , , 1)
                    If myColor Is Nothing Then Set myColor = �I�_�[�qRan.Find(�[�q�i��, , , 1)
                        myColor = myColor.Interior.color
                �[�q�t�@�~���[ = ""
                If myColor <> 16777215 Then
                    b = 0
                    Do Until aKey.Offset(b, 1) = ""
                        If myColor = aKey.Offset(b, 1).Interior.color Then
                            �[�q�t�@�~���[ = aKey.Offset(b, 1) & "_" & aKey.Offset(b, 2)
                            Exit Do
                        End If
                        b = b + 1
                    Loop
                End If
                .Cells(addRow, keyCol + 1) = �t���i��
                .Cells(addRow, keyCol + 2) = �[�q�t�@�~���[
                If myColor <> 16777215 Then .Cells(addRow, keyCol + 2).Interior.color = myColor
                .Cells(addRow, keyCol + 5) = ����
                .Cells(addRow, incol) = �ӏ���
            End If
line20:
        Next Y
    End With
    '�\�[�g
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�r��������
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = "�[�q�ꗗ"
    End With
    
    MsgBox "�쐬���܂����B"
    
End Function


Public Function A_�R�l�N�^�ꗗ�̃V�[�g�쐬()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�R�l�N�^�ꗗ"
    
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Call SQL_�R�l�N�^�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�V�[�g��:���i�ʒ[���ꗗ��������΍쐬
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("�[�����i��", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "�[�����i��"
            .Cells(keyRow, keyCol + 1) = "�[����"
            .Cells(keyRow, keyCol + 2) = ""
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "����"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            �[����� = RAN(0, Y)
            �[�� = RAN(1, Y)
            ���� = RAN(2, Y)
            ���i = RAN(3, Y)
            �ӏ��� = RAN(4, Y)
            If �[����� & ���� = "" Then GoTo line20
            '���i�i�Ԃ̗�
            Set fnd = .Rows(keyRow).Find(���i, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(���i, " ", "")) = 10 Then
                    ���iA = Mid(���i, 8, 3)
                Else
                    ���iA = Mid(���i, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = ���i
                .Cells(keyRow - 1, incol) = ���iA
                .Cells(keyRow - 1, incol).ColumnWidth = Len(���iA) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = �N��
            Else
                incol = fnd.Column
            End If
            
            '�o�^�����邩�m�F
            flg = False
            For i = keyRow + 1 To lastRow
                If �[����� = .Cells(i, keyCol + 0) Then
                    If �[�� = .Cells(i, keyCol + 1) Then
                        If ���� = .Cells(i, keyCol + 5) Then
                            flg = True
                            addRow = i
                            .Cells(addRow, incol) = .Cells(addRow, incol) + �ӏ���
                            Exit For
                        End If
                    End If
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = �[�����
                .Cells(addRow, keyCol + 1) = �[��
                .Cells(addRow, keyCol + 5) = ����
                .Cells(addRow, incol) = �ӏ���
            End If
line20:
        Next Y
    End With
    '�\�[�g
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�r��������
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        '�E�B���h�E�g�̌Œ�
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = newSheetName
    End With
    
    MsgBox "�쐬���܂����B"
    
End Function

Public Function B_�}���K�C�h�o�^�ꗗ�̃V�[�g�쐬()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�}���K�C�h�o�^�ꗗ"
    
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Call SQL_�}���K�C�h�o�^�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�V�[�g��:���i�ʒ[���ꗗ��������΍쐬
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    
    Dim newSheet As Worksheet, ���� As String
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("�[�����i��", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "�[�����i��"
            .Cells(keyRow, keyCol + 1) = "�}���K�C�h"
            .Cells(keyRow, keyCol + 2) = "�[����"
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "����"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            �[����� = RAN(0, Y)
            �[�� = RAN(1, Y)
            ���� = RAN(2, Y)
            ���i = RAN(3, Y)
            �ӏ��� = RAN(4, Y)
            �[�q = RAN(5, Y)
            If IsNull(�[�q) Then �[�q = ""
            If �[����� & ���� = "" Then GoTo line20
            '���i�i�Ԃ̗�
            Set fnd = .Rows(keyRow).Find(���i, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(���i, " ", "")) = 10 Then
                    ���iA = Mid(���i, 8, 3)
                Else
                    ���iA = Mid(���i, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = ���i
                .Cells(keyRow - 1, incol) = ���iA
                .Cells(keyRow - 1, incol).ColumnWidth = Len(���iA) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = �N��
            Else
                incol = fnd.Column
            End If
            
            '�o�^�����邩�m�F
            flg = False
            For i = keyRow + 1 To lastRow
                If �[����� = .Cells(i, keyCol + 0) Then
                    If �[�q = .Cells(i, keyCol + 1) Then
                        If �[�� = .Cells(i, keyCol + 2) Then
                            If ���� = .Cells(i, keyCol + 5) Then
                                flg = True
                                addRow = i
                                .Cells(addRow, incol) = .Cells(addRow, incol) + �ӏ���
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = �[�����
                .Cells(addRow, keyCol + 1) = �[�q
                .Cells(addRow, keyCol + 2) = �[��
                .Cells(addRow, keyCol + 5) = ����
                .Cells(addRow, incol) = �ӏ���
            End If
line20:
        Next Y
    End With
    '�\�[�g
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�r��������
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Activate
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        '�E�B���h�E�g�̌Œ�
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = newSheetName
    End With
    
    MsgBox "�쐬���܂����B"
    
End Function

Public Function A_�}���K�C�h�ꗗ�̃V�[�g�쐬()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "�}���K�C�h�ꗗ"
    
    
    Dim i As Long, i2 As Long, ���i�i��RAN As Variant
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    Call SQL_�}���K�C�h�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�V�[�g��:���i�ʒ[���ꗗ��������΍쐬
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet, ���� As String
    '�V�[�g�������ꍇ�쐬
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("�[�����i��", , , 1)
        'setup
        If key Is Nothing Then '�V�K�쐬�̎�
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "�}���K�C�h"
            .Cells(keyRow, keyCol + 1) = ""
            .Cells(keyRow, keyCol + 2) = ""
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "����"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '���������鎞
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            �[����� = RAN(0, Y)
            �[�� = RAN(1, Y)
            ���� = RAN(2, Y)
            ���i = RAN(3, Y)
            �ӏ��� = RAN(4, Y)
            �[�q = RAN(5, Y)
            If IsNull(�[�q) Then �[�q = ""
            If �[����� & ���� = "" Then GoTo line20
            '���i�i�Ԃ̗�
            Set fnd = .Rows(keyRow).Find(���i, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(���i, " ", "")) = 10 Then
                    ���iA = Mid(���i, 8, 3)
                Else
                    ���iA = Mid(���i, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = ���i
                .Cells(keyRow - 1, incol) = ���iA
                .Cells(keyRow - 1, incol).ColumnWidth = Len(���iA) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = �N��
            Else
                incol = fnd.Column
            End If
            
            '�o�^�����邩�m�F
            flg = False
            For i = keyRow + 1 To lastRow
'                If �[����� = .Cells(i, keyCol + 0) Then
                    If �[�q = .Cells(i, keyCol + 0) Then
'                        If �[�� = .Cells(i, keyCol + 2) Then
                            If ���� = .Cells(i, keyCol + 5) Then
                                flg = True
                                addRow = i
                                .Cells(addRow, incol) = .Cells(addRow, incol) + �ӏ���
                                Exit For
                            End If
'                        End If
                    End If
'                End If
            Next i
            '�����̂Œǉ�
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = �[�q
                .Cells(addRow, keyCol + 1) = ""
                .Cells(addRow, keyCol + 2) = ""
                .Cells(addRow, keyCol + 5) = ����
                .Cells(addRow, incol) = �ӏ���
            End If
line20:
        Next Y
    End With
    '�\�[�g
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�r��������
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Activate
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        '�E�B���h�E�g�̌Œ�
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = newSheetName
    End With
    
    MsgBox "�쐬���܂����B"
    
End Function

Public Function ���i���X�g�̍쐬_Ver1940()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "���i���X�g"
    
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    
    With Workbooks(myBookName).Sheets("�ݒ�")
        �n���}�A�h���X = .Cells.Find("���i���_", , , 1).Offset(0, 1).Value
    End With
    
    '���i�i�Ԃ̃��C���i�Ԃ�RLTF��Ǎ���
    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�i��key As Range: Set ���i�i��key = .Cells.Find("���C���i��", , , 1)
        Dim RLTFkey As Range: Set RLTFkey = .Cells.Find("RLTF", , , 1)
        Dim ���i�i��lastRow As Long: ���i�i��lastRow = .Cells(.Rows.count, ���i�i��key.Column).End(xlUp).Row
        Dim ��������() As String: ReDim ��������(���i�i��lastRow - ���i�i��key.Row, 2)
        Dim ���i�_�� As Long: ���i�_�� = ���i�i��lastRow - ���i�i��key.Row
        Dim n As Long
        For n = 1 To ���i�_��
            ��������(n, 1) = .Cells(���i�i��key.Row + n, ���i�i��key.Column)
            ��������(n, 2) = .Cells(RLTFkey.Row + n, RLTFkey.Column)
        Next n
        Set ���i�i��key = Nothing
        Set RLTFkey = Nothing
    End With
    
    '���ޏڍ�txt�̓Ǎ���
    Dim ���ޏڍ�() As String
    Dim TargetFile As String: TargetFile = �n���}�A�h���X & "\���ޏڍ�" & ".txt"
    Dim intFino As Integer
    Dim aRow As String, aCel As Variant, ���ޏڍ�c As Long: ���ޏڍ�c = -1
    Dim ���ޏڍ�v As String
    intFino = FreeFile
    Open TargetFile For Input As #intFino
    Do Until EOF(intFino)
        Line Input #intFino, aRow
        aCel = Split(aRow, ",")
        ���ޏڍ�c = ���ޏڍ�c + 1
        For a = LBound(aCel) To UBound(aCel)
            ReDim Preserve ���ޏڍ�(UBound(aCel), ���ޏڍ�c)
            ���ޏڍ�(a, ���ޏڍ�c) = aCel(a)
        Next a
    Loop
    Close #intFino
    
    Dim �i�[V() As Variant: ReDim �i�[V(0)
    Dim �i�[L() As Variant: ReDim �i�[L(���i�_��, 0)
    Dim V() As String: ReDim V(15 + ���i�_��)
    Dim c As Long
    '�^�C�g���s
    �i�[V(c) = "�\����,���i�i��,�ď�,����1,����2,�F,�ؒf��,,,���ޏڍ�,���,�H��"
    For n = 1 To ���i�_��
        �i�[V(c) = �i�[V(c) & "," & Replace(��������(n, 1), " ", "")
    Next n
    
    '���i�i�Ԗ���RLTF����ǂݍ���
    For n = 1 To ���i�_��
        '���͂̐ݒ�(�C���|�[�g�t�@�C��)
        TargetFile = myBookpath & "\05_RLTF_A\" & ��������(n, 2) & ".txt"
        
        intFino = FreeFile
        Open TargetFile For Input As #intFino
        Do Until EOF(intFino)
            Line Input #intFino, aRow
            If Replace(��������(n, 1), " ", "") = Replace(Mid(aRow, 1, 15), " ", "") Then
                If Mid(aRow, 27, 1) = "T" Then '�`���[�u
                    V(0) = Mid(aRow, 1, 15) '���i�i��
                    V(1) = Mid(aRow, 19, 3)   '�ݕ�
                    V(2) = "" 'Mid(aRow, 27, 4)   'T�\����
                    V(3) = Replace(Mid(aRow, 375, 8), " ", "") '���i�i��
                    Select Case Len(V(3))
                        Case 8
                            V(3) = Left(V(3), 3) & "-" & Mid(V(3), 4, 3) & "-" & Mid(V(3), 7, 3)
                        Case Else
                            Stop
                    End Select
                    V(4) = Mid(aRow, 383, 6)  'T�ď�
                    V(5) = Mid(aRow, 389, 4)  'T����1
                    V(6) = Mid(aRow, 393, 4)  'T����2
                    V(7) = Replace(Mid(aRow, 397, 6), " ", "") 'T�F
                    V(8) = CLng(Mid(aRow, 403, 5))  'T�ؒf��
                    V(9) = "" 'Mid(aRow, 544, 1) '�Ȃ�1
                    V(10) = "" 'Mid(aRow, 544, 4) '�Ȃ�2
                    V(11) = Mid(aRow, 153, 2)  '�H��
                    V(12) = "T"
                    V(13) = 1 '����
                If V(5) <> "    " And V(6) <> "    " Then 'VO
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "�~" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & " L=" & String(4 - Len(Format(Mid(aRow, 403, 5), 0)), " ") & Format(Mid(aRow, 403, 5), 0)
                ElseIf V(5) <> "    " Then 'COT
                    V(15) = Left(V(3), 3) & "-D" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "�~" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                ElseIf V(6) <> "    " Then 'VS
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & "�~" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                End If
                    GoSub �i�[���s
                ElseIf Mid(aRow, 27, 1) = "B" Then '40�H���ȍ~�̕��i
                    For X = 0 To 9
                        If Mid(aRow, 175 + (X * 20) + 10, 3) = "ATO" Then
                            V(0) = Mid(aRow, 1, 15) '���i�i��
                            V(1) = Mid(aRow, 19, 3)   '�ݕ�
                            V(2) = ""                 'T�\����
                            V(3) = Replace(Mid(aRow, 175 + (X * 20), 10), " ", "") '���i�i��
                            Select Case Len(V(3))
                                Case 8
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4)
                                Case 9, 10
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4) & "-" & Mid(V(3), 9, 2)
                                Case Else
                                    Stop
                            End Select
                            '���ޏڍׂ̎擾
                            ���ޏڍ�v = ""
                            For a = 0 To ���ޏڍ�c
                                If ���ޏڍ�(0, a) = V(3) Then
                                    If Left(���ޏڍ�(2, a), 2) = "F1" Then '�N���b�v�̎�
                                        ���ޏڍ�v = Mid(���ޏڍ�(4, a), 4)
                                    Else
                                        ���ޏڍ�v = Mid(���ޏڍ�(3, a), 5)
                                    End If
                                    Exit For
                                End If
                            Next a
                            V(4) = ""  'T�ď�
                            V(5) = ""  'T����1
                            V(6) = ""  'T����2
                            V(7) = ""  'T�F
                            V(8) = ""  'T�ؒf��
                            V(9) = "" '�Ȃ�1
                            V(10) = "" '�Ȃ�2
                            V(11) = Mid(aRow, 558 + (X * 2), 2) '�H��
                            V(12) = "B"
                            V(13) = CLng(Mid(aRow, 189 + (X * 20), 4)) '����
                            V(15) = ���ޏڍ�v
                            GoSub �i�[���s
                        End If
                    Next X
                End If
            End If
        Loop
        Close #intFino
    Next n
    
    '�V�[�g�ǉ�
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Tab.color = 14470546
    '�o��
    Dim Val As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        .Cells.NumberFormat = "@"
        .Columns("I").NumberFormat = 0
        For a = LBound(�i�[V) To UBound(�i�[V)
            Val = Split(�i�[V(a), ",")
            If a = LBound(�i�[V) Then '�t�B�[���h��
                For b = LBound(Val) To UBound(Val)
                    .Cells(a + 1, b + 1) = Val(b)
                Next b
            Else
                max�� = 0
                For n = 1 To ���i�_��
                   If �i�[L(n, a) > CLng(max��) Then max�� = CLng(�i�[L(n, a))
                Next n
                For i = 1 To max��
                    addRow = .Cells(.Rows.count, 2).End(xlUp).Row + 1
                    For b = LBound(Val) To UBound(Val)
                        .Cells(addRow, b + 1) = Val(b)
                    Next b
                    For n = 1 To ���i�_��
                        If �i�[L(n, a) <> 0 Then
                            If �i�[L(n, a) <> "" Then
                                �i�[L(n, a) = �i�[L(n, a) - 1
                                .Cells(addRow, UBound(Val) + n + 1) = "0"
                            End If
                        End If
                    Next n
                Next i
            End If
        Next a
        'T�ď̂̃t�H���g�ݒ�
        .Columns("l").Font.Name = "�l�r �S�V�b�N"
        '�H��a�̒ǉ�
        .Columns("m").Insert
        .Range("m1") = "�H��a"
        '�t�B�b�g
        .Columns("A:p").AutoFit
        '�s�̒ǉ�
        '.Rows("1:2").Insert
        
        '�E�B���h�E�g�̌Œ�
        .Range("a2").Select
        ActiveWindow.FreezePanes = True
        '�r��
        With .Range(.Cells(1, 1), .Cells(addRow, UBound(Val) + ���i�_�� + 2))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        '�\�[�g
        With .Sort.SortFields
            .Clear
            .add key:=Cells(1, 11), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 12), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange .Range(.Rows(2), Rows(addRow))
        With .Sort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
Exit Function

�i�[���s:
    �i�[temp = V(2) & "," & V(3) & "," & V(4) & "," & V(5) & "," & V(6) & "," & V(7) & "," & V(8) & "," & V(9) & "," & V(10) & "," & V(15) & "," & V(12) & "," & V(11)
    '��������������
    For cc = 1 To c
        If �i�[V(cc) = �i�[temp Then
            For nn = 1 To ���i�_��
                If ��������(nn, 1) = V(0) Then
                    �i�[L(nn, cc) = CLng(�i�[L(nn, cc)) + CLng(V(13))
                    Return
                End If
            Next nn
        End If
    Next cc
    '�V�K�o�^
    For nn = 1 To ���i�_��
        If ��������(nn, 1) = V(0) Then
            c = c + 1
            ReDim Preserve �i�[V(c)
            ReDim Preserve �i�[L(���i�_��, c)
            �i�[V(c) = �i�[temp
            �i�[L(nn, c) = V(13)
        End If
    Next nn
Return
        
End Function

Public Function ���i���X�g�̍쐬_Ver2040()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "���i���X�g"
    
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    
    Call �A�h���X�Z�b�g(myBook)
        
    '���i�i�Ԃ̃��C���i�Ԃ�RLTF��Ǎ���
    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�i��key As Range: Set ���i�i��key = .Cells.Find("���C���i��", , , 1)
        Dim RLTFkey As Range: Set RLTFkey = .Cells.Find("RLTF-A", , , 1)
        Dim ����Col As Long: ����Col = .Cells.Find("����", , , 1).Column
        Dim ���i�i��lastRow As Long: ���i�i��lastRow = .Cells(.Rows.count, ���i�i��key.Column).End(xlUp).Row
        Dim ��������() As String: ReDim ��������(���i�i��lastRow - ���i�i��key.Row, 4)
        Dim ���i�_�� As Long: ���i�_�� = ���i�i��lastRow - ���i�i��key.Row
        Dim n As Long
        For n = 1 To ���i�_��
            ��������(n, 1) = .Cells(���i�i��key.Row + n, ���i�i��key.Column)
            ��������(n, 2) = .Cells(RLTFkey.Row + n, RLTFkey.Column)
            ��������(n, 4) = .Cells(RLTFkey.Row + n, ����Col)
        Next n
        Set ���i�i��key = Nothing
        Set RLTFkey = Nothing
    End With
       
    Dim �i�[V() As Variant: ReDim �i�[V(0)
    Dim �i�[L() As Variant: ReDim �i�[L(���i�_��, 0)
    Dim V() As String: ReDim V(15 + ���i�_��)
    Dim c As Long
    '�^�C�g���s
    �i�[V(c) = "�\����,���i�i��,�ď�,����1,����2,�F,�ؒf��,,,���ޏڍ�,���,�H��"
    For n = 1 To ���i�_��
        �i�[V(c) = �i�[V(c) & "," & ��������(n, 1)
    Next n
    
    '���i�i�Ԗ���RLTF����ǂݍ���
    For n = 1 To ���i�_��
        '���͂̐ݒ�(�C���|�[�g�t�@�C��)
        TargetFile = myBookpath & "\05_RLTF_A\" & ��������(n, 2) & ".txt"
        If Dir(TargetFile) <> "" Then
            intFino = FreeFile
            Open TargetFile For Input As #intFino
            Do Until EOF(intFino)
                Line Input #intFino, aRow
                If Replace(��������(n, 1), " ", "") = Replace(Mid(aRow, 1, 15), " ", "") Then
                    If Mid(aRow, 27, 1) = "T" Then '�`���[�u
                        V(0) = Mid(aRow, 1, 15) '���i�i��
                        If ��������(n, 1) = V(0) Then ��������(n, 3) = CDate("20" & Mid(aRow, 482, 2) & "/" & Mid(aRow, 484, 2) & "/" & Mid(aRow, 486, 2))
                        V(1) = Mid(aRow, 19, 3)   '�ݕ�
                        V(2) = "" 'Mid(aRow, 27, 4)   'T�\����
                        V(3) = Replace(Mid(aRow, 375, 8), " ", "") '���i�i��
                        Select Case Len(V(3))
                            Case 8
                                V(3) = Left(V(3), 3) & "-" & Mid(V(3), 4, 3) & "-" & Mid(V(3), 7, 3)
                            Case Else
                                Stop
                        End Select
                        V(4) = Mid(aRow, 383, 6)  'T�ď�
                        V(5) = Mid(aRow, 389, 4)  'T����1
                        V(6) = Mid(aRow, 393, 4)  'T����2
                        V(7) = Replace(Mid(aRow, 397, 6), " ", "") 'T�F
                        V(8) = CLng(Mid(aRow, 403, 5))  'T�ؒf��
                        V(9) = "" 'Mid(aRow, 544, 1) '�Ȃ�1
                        V(10) = "" 'Mid(aRow, 544, 4) '�Ȃ�2
                        V(11) = Mid(aRow, 153, 2)  '�H��
                        V(12) = "T"
                        V(13) = 1 '����
                    If V(5) <> "    " And V(6) <> "    " Then 'VO
                        V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                                & "�~" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                                & " L=" & String(4 - Len(Format(Mid(aRow, 403, 5), 0)), " ") & Format(Mid(aRow, 403, 5), 0)
                    ElseIf V(5) <> "    " Then 'COT
                        V(15) = Left(V(3), 3) & "-D" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                                & "�~" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                    ElseIf V(6) <> "    " Then 'VS
                        V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                                & "�~" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                    End If
                        GoSub �i�[���s
                    ElseIf Mid(aRow, 27, 1) = "B" Then '40�H���ȍ~�̕��i
                        For X = 0 To 9
                            If Mid(aRow, 175 + (X * 20) + 10, 3) = "ATO" Then
                                V(0) = Mid(aRow, 1, 15) '���i�i��
                                V(1) = Mid(aRow, 19, 3)   '�ݕ�
                                V(2) = ""                 'T�\����
                                V(3) = Replace(Mid(aRow, 175 + (X * 20), 10), " ", "") '���i�i��
                                Select Case Len(V(3))
                                    Case 8
                                        V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4)
                                    Case 9, 10
                                        V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4) & "-" & Mid(V(3), 9, 2)
                                    Case Else
                                        Stop
                                End Select
                                V(4) = ""  'T�ď�
                                V(5) = ""  'T����1
                                V(6) = ""  'T����2
                                V(7) = ""  'T�F
                                V(8) = ""  'T�ؒf��
                                V(9) = "" '�Ȃ�1
                                V(10) = "" '�Ȃ�2
                                V(11) = Mid(aRow, 558 + (X * 2), 2) '�H��
                                V(12) = "B"
                                V(13) = CLng(Mid(aRow, 189 + (X * 20), 4)) '����
                                If Left(���ޏڍׂ̓ǂݍ���(V(3), "���i���_"), 2) = "F1" Then
                                    V(15) = Mid(���ޏڍׂ̓ǂݍ���(V(3), "�N�����v�^�C�v_"), 4)
                                Else
                                    V(15) = Mid(���ޏڍׂ̓ǂݍ���(V(3), "���i����_"), 5)
                                End If
                                GoSub �i�[���s
                            End If
                        Next X
                    End If
                End If
            Loop
            Close #intFino
        End If
    Next n
    
    '�������O�̃V�[�g�����邩�m�F
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
    newSheet.Cells(1, 1) = newSheetName
    newSheet.Cells(1, 3) = "�T�u�}�A���i���X�g�쐬�Ɏg�p���܂��B"
    newSheet.Cells(2, 1) = "VO��X���[�N���b�v���A�T�u�}�ɍڂ��镔�i�̒[��������͂��Ă��������B�R�l�N�^�A�h������0�̂܂ܕύX���Ȃ��ŉ������B"
    If newSheet.Name = "���i���X�g" Then
        newSheet.Tab.color = 14470546
    End If
    
    '�o��
    Dim Val As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        .Cells.NumberFormat = "@"
        .Columns("I").NumberFormat = 0
        For a = LBound(�i�[V) To UBound(�i�[V)
            Val = Split(�i�[V(a), ",")
            If a = LBound(�i�[V) Then '�t�B�[���h��
                For b = LBound(Val) To UBound(Val)
                    .Cells(a + 3, b + 1) = Val(b)
                    For X = LBound(��������, 1) + 1 To UBound(��������, 1)
                        If ��������(X, 1) = Val(b) Then
                            .Cells(a + 1, b + 1).NumberFormat = "mm/dd"
                            .Cells(a + 1, b + 1) = ��������(X, 3)
                            .Cells(a + 1, b + 1).ShrinkToFit = True
                            .Cells(a + 2, b + 1).NumberFormat = "@"
                            .Cells(a + 2, b + 1) = ��������(X, 4)
                            .Columns(b + 1).ColumnWidth = Len(��������(X, 4)) * 1.05
                        End If
                    Next X
                Next b
            Else
                max�� = 0
                For n = 1 To ���i�_��
                   If �i�[L(n, a) > CLng(max��) Then max�� = CLng(�i�[L(n, a))
                Next n
                For i = 1 To max��
                    addRow = .Cells(.Rows.count, 2).End(xlUp).Row + 1
                    For b = LBound(Val) To UBound(Val)
                        .Cells(addRow, b + 1) = Val(b)
                    Next b
                    If Val(10) = "T" And Val(11) = "40" Then .Cells(addRow, 12).Interior.color = RGB(255, 255, 0)
                    If Val(9) = "�X���[�N���b�v" Then .Cells(addRow, 12).Interior.color = RGB(255, 255, 0)
                    If Val(9) = "�[�q�W�~���i" Then .Cells(addRow, 12).Interior.color = RGB(255, 255, 0)
                    For n = 1 To ���i�_��
                        If �i�[L(n, a) <> 0 Then
                            If �i�[L(n, a) <> "" Then
                                �i�[L(n, a) = �i�[L(n, a) - 1
                                .Cells(addRow, UBound(Val) + n + 1) = "0"
                            End If
                        End If
                    Next n
                Next i
            End If
        Next a
        'T�ď̂̃t�H���g�ݒ�
        .Columns("l").Font.Name = "�l�r �S�V�b�N"
        '�H��a�̒ǉ�
        .Columns("m").Insert
'        .Columns("m").Interior.Pattern = xlNone
        .Range("m3") = "�H��a"
        .Range("m3").AddComment
        .Range("m3").Comment.Text "��n���ŕt�����镔�i��40�����"
        .Range("m3").Comment.Shape.TextFrame.AutoSize = True
        .Range("m3").Interior.color = RGB(255, 255, 0)
        .Range(.Range("m3"), .Cells(3, .Columns.count).End(xlToLeft)).Interior.color = RGB(255, 255, 0)
        '�t�B�b�g
        .Columns("A:l").AutoFit
        .Columns(1).ColumnWidth = 7
        .Columns(3).ColumnWidth = 11
        .Columns("h:i").ColumnWidth = 2
        
        '�s�̒ǉ�
        '.Rows("1:2").Insert
        
        '�E�B���h�E�g�̌Œ�
        .Range("a4").Select
        ActiveWindow.FreezePanes = True
        '�r��
        With .Range(.Cells(3, 1), .Cells(addRow, UBound(Val) + ���i�_�� + 2))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        '�\�[�g
        With .Sort.SortFields
            .Clear
            .add key:=Cells(3, 11), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(3, 12), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(3, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(3, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange .Range(.Rows(4), Rows(addRow))
        With .Sort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
Exit Function

�i�[���s:
    �i�[temp = V(2) & "," & V(3) & "," & V(4) & "," & V(5) & "," & V(6) & "," & V(7) & "," & V(8) & "," & V(9) & "," & V(10) & "," & V(15) & "," & V(12) & "," & V(11)
    '��������������
    For cc = 1 To c
        If �i�[V(cc) = �i�[temp Then
            For nn = 1 To ���i�_��
                If ��������(nn, 1) = V(0) Then
                    �i�[L(nn, cc) = CLng(�i�[L(nn, cc)) + CLng(V(13))
                    Return
                End If
            Next nn
        End If
    Next cc
    '�V�K�o�^
    For nn = 1 To ���i�_��
        If ��������(nn, 1) = V(0) Then
            c = c + 1
            ReDim Preserve �i�[V(c)
            ReDim Preserve �i�[L(���i�_��, c)
            �i�[V(c) = �i�[temp
            �i�[L(nn, c) = V(13)
        End If
    Next nn
Return
        
End Function

Public Function PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬_2017()
    '�g�p����V�[�g�̐��i�i�Ԃ̕��т��}�b�`���Ă��邩�m�F���鏈��_�ǉ��v
    Call �œK��
'    Dim my���i�i�� As String
'    If �T�u�}���i�i�� = "" Then
'        my���i�i�� = "821113B380" '��n���̐F�t���Ŏg�p���鐻�i�i�ԁ��u�����N�̎��̓A���}�b�`�������Ă��֌W������n���ɂ���"
'    Else
'        my���i�i�� = �T�u�}���i�i��
'    End If
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name    '�\�[�X
    Dim outSheetName As String: outSheetName = "PVSW_RLTF" '�o�͐�
    Dim myRefrentName As String: myRefrentName = "�[���ꗗ" '�Q��
    
    'PVSW_RLTF����[�������擾
    With wb(0).Sheets("�ݒ�")
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

    '���i�i��
    'Call ���i�i��RAN_set2(���i�i��RAN, "", "", my���i�i��)
    '���i�ʒ[���ꗗ
    With wb(0).Sheets(myRefrentName)
        Dim ref���Row As Long: ref���Row = .Cells.Find("�[�����i��", , , 1).Row
        Dim ref���Col As Long: ref���Col = .Cells.Find("�[�����i��", , , 1).Column
        Dim ref�[��Col As Long: ref�[��Col = .Cells.Find("�[����", , , 1).Column
        Dim refLastRow As Long: refLastRow = .Cells(.Rows.count, ref���Col).End(xlUp).Row
        Dim refLastCol As Long: refLastCol = .UsedRange.Columns.count
        Dim ref�^�C�g��Ran As Range: Set ref�^�C�g��Ran = .Rows(ref���Row)
        Dim ref���i�ʒ[���ꗗRan As Range: Set ref���i�ʒ[���ꗗRan = .Range(.Cells(1, 1), .Cells(refLastRow, refLastCol))
        'Dim ref���[��Ran As Range: Set ref���[��Ran = .Range(.Cells(ref���Row, ref���Col), .Cells(ref���Row, ref�[��Col))
    End With
    'PVSW_RLTF
    With wb(0).Sheets(outSheetName)
        Dim out�^�C�g��Row As Long: out�^�C�g��Row = .Cells.Find("�i��_", , , 1).Row
        Dim out�^�C�g��Col As Long: out�^�C�g��Col = .Cells(out�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Dim out�^�C�g��Ran As Range: Set out�^�C�g��Ran = .Range(.Cells(out�^�C�g��Row, 1), .Cells(out�^�C�g��Row, out�^�C�g��Col))
        Dim out�d�����ʖ�Col As Long: out�d�����ʖ�Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim outJCDFcol As Long: outJCDFcol = .Cells.Find("JCDF_", , , 1).Column
        Dim out�i��Col As Long: out�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim out�ڑ�Gcol As Long: out�ڑ�Gcol = .Cells.Find("�ڑ�G_", , , 1).Column
        Dim out�T�C�YCol As Long: out�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim out�FCol As Long: out�FCol = .Cells.Find("�F_", , , 1).Column
        Dim outABCol As Long: outABCol = .Cells.Find("AB_", , , 1).Column
        Dim out�F��Col(1) As Long
        out�F��Col(0) = out�^�C�g��Ran.Cells.Find("�F��_", , , 1).Column
        out�F��Col(1) = out�^�C�g��Ran.Cells.Find("�d���F", , , 1).Column
        Dim out�����i��Col As Long: out�����i��Col = .Cells.Find("�����i��", , , 1).Column
        Dim out����Col As Long: out����Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim out����Col(1) As Long
        out����Col(0) = .Cells.Find("�n�_������_", , , 1).Column
        out����Col(1) = .Cells.Find("�I�_������_", , , 1).Column
        Dim out��Col(1) As Long
        out��Col(0) = .Cells.Find("�n�_����H����", , , 1).Column
        out��Col(1) = .Cells.Find("�I�_����H����", , , 1).Column
        Dim out�[��Col(1) As Long
        out�[��Col(0) = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        out�[��Col(1) = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        Dim out���Col(1) As Long
        out���Col(0) = .Cells.Find("�n�_���[�����i��", , , 1).Column
        out���Col(1) = .Cells.Find("�I�_���[�����i��", , , 1).Column
        Dim out�[�qCol(1) As Long
        out�[�qCol(0) = .Cells.Find("�n�_���[�q�i��", , , 1).Column
        out�[�qCol(1) = .Cells.Find("�I�_���[�q�i��", , , 1).Column
        Dim outCavCol(1) As Long
        outCavCol(0) = .Cells.Find("�n�_���L���r�e�B", , , 1).Column
        outCavCol(1) = .Cells.Find("�I�_���L���r�e�B", , , 1).Column
        Dim out�}Col(1) As Long
        out�}Col(0) = .Cells.Find("�n�_���}_", , , 1).Column
        out�}Col(1) = .Cells.Find("�I�_���}_", , , 1).Column
        Dim out�}�VCol(1) As Long
        out�}�VCol(0) = .Cells.Find("�n�_���}���}�F�P", , , 1).Column
        out�}�VCol(1) = .Cells.Find("�I�_���}���}�F�P", , , 1).Column
        '��n���p�f�[�^
        Dim out2�n��Col(1) As Long
        out2�n��Col(0) = .Cells.Find("�n�_���n��", , , 1).Column
        out2�n��Col(1) = .Cells.Find("�I�_���n��", , , 1).Column
        Dim out2�[�qCol(1) As Long
        out2�[�qCol(0) = .Cells.Find("�n�_���[�q_", , , 1).Column
        out2�[�qCol(1) = .Cells.Find("�I�_���[�q_", , , 1).Column
        Dim out2���i�i��Col As Long: out2���i�i��Col = .Cells.Find("���i�i��", , , 1).Column
        .Cells(out�^�C�g��Row - 1, out2���i�i��Col).ClearContents
        .Activate
        .Range(Cells(out�^�C�g��Row + 1, out2���i�i��Col), Cells(.UsedRange.Rows.count, out2���i�i��Col)).ClearContents
        Dim out2�T�uCol As Long: out2�T�uCol = .Cells.Find("�T�u", , , 1).Column
        Dim out2�ڑ�Gcol As Long: out2�ڑ�Gcol = .Cells.Find("�ڑ�G", , , 1).Column
        Dim out2�F��col As Long: out2�F��col = .Cells.Find("�F��", , , 1).Column
        Dim out2�[��Col(1) As Long
        out2�[��Col(0) = .Cells.Find("�n�_���[��", , , 1).Column
        out2�[��Col(1) = .Cells.Find("�I�_���[��", , , 1).Column
        Dim out2��Col(1) As Long
        out2��Col(0) = .Cells.Find("�n�_����H����_", , , 1).Column
        out2��Col(1) = .Cells.Find("�I�_����H����_", , , 1).Column
        Dim out2�}Col(1) As Long
        out2�}Col(0) = .Cells.Find("�n�_���}", , , 1).Column
        out2�}Col(1) = .Cells.Find("�I�_���}", , , 1).Column
        Dim out2����Col As Long: out2����Col = .Cells.Find("����_", , , 1).Column
        Dim out2�\��Col As Long: out2�\��Col = .Cells.Find("�\��", , , 1).Column: .Columns(out2�\��Col).NumberFormat = "@"
        Dim out2����Col As Long: out2����Col = .Cells.Find("����__", , , 1).Column
        Dim out���[�n��Col As Long: out���[�n��Col = .Cells.Find("���[�n��", , , 1).Column
        Dim out���[���[�qCol As Long: out���[���[�qCol = .Cells.Find("���[���[�q", , , 1).Column
        Dim outRLTFCol As Long: outRLTFCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim outLastRow As Long: outLastRow = .Cells(.Rows.count, out�d�����ʖ�Col).End(xlUp).Row
        Dim outLastCol As Long: outLastCol = .Cells(out�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Dim outPVSWcsvRAN As Range: Set outPVSWcsvRAN = .Range(.Cells(1, 1), .Cells(outLastRow, outLastCol))
        '.Range(.Columns(1), .Columns(out�d�����ʖ�Col)).Interior.Pattern = xlNone
    End With
    
    Dim myKey As Variant, myY As Long, myX As Long, findC As Long, findR As Long, refY As Long, outY As Long, gawaLong As Long, sLong As Long, myLineStyle As Long
    Dim �\�� As String, ���(1) As String, �[��(1) As String, my�T�u As String, ���i�i�� As String, ��(1) As Variant, �� As String
    Dim ref�T�u As String, �F�� As String, �}���}�F(1) As String, �[�q(1) As String, ����(1) As String, cav(1) As String, �ڑ�G As String
    Dim findResult As Boolean
    Dim my�F(1) As Long, my�Fb(1) As Long, �n��(1) As String, �n��n(1) As Variant

    With wb(0).Sheets(outSheetName)
        .Range(.Cells(out�^�C�g��Row + 1, out2���i�i��Col), .Cells(outLastRow, out���[���[�qCol)).ClearContents
        .Range(.Cells(out�^�C�g��Row + 1, out2���i�i��Col), .Cells(outLastRow, out���[���[�qCol)).Interior.Pattern = xlNone
        For r = 1 To ���i�i��RANc
            ���i�i�� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            Set myKey = out�^�C�g��Ran.Find(���i�i��, , , 1)
            If myKey Is Nothing Then GoTo NextR
            myX = myKey.Column
                For myY = out�^�C�g��Row + 1 To outLastRow
                    �\�� = .Cells(myY, out�d�����ʖ�Col): If �\�� = "" Then GoTo nextY
                    �ڑ�G = .Cells(myY, out�ڑ�Gcol)
                    �F�� = .Cells(myY, out�F��Col(0))
                    ���Y�敪 = .Cells(myY, out2����Col)
                    JCDF = .Cells(myY, outJCDFcol)
                    
                    RLTF = .Cells(myY, outRLTFCol): If RLTF <> "Found" Then GoTo nextY 'RLFT�ɏ�����������
                    my�T�u = .Cells(myY, myX): If my�T�u = "" Then GoTo nextY 'my�T�u���u�����N�̎�
                    If .Cells(myY, out���Col(0)) = "" And .Cells(myY, out���Col(1)) = "" Then GoTo nextY '���[����""�̎�
                    
                    For a = 0 To 1
                        ���(a) = .Cells(myY, out���Col(a))
                        �[��(a) = .Cells(myY, out�[��Col(a)): If �[��(a) = "" Then GoTo NextA
                        cav(a) = .Cells(myY, outCavCol(a))
                        Set ��(a) = .Cells(myY, out��Col(a))
                        �[�q(a) = .Cells(myY, out�[�qCol(a))
                        If a = 0 Then
                            ����(a) = .Cells(myY, out�[��Col(1)) & "_" & .Cells(myY, outCavCol(1)) & "_" & .Cells(myY, out��Col(1))
                        Else
                            ����(a) = .Cells(myY, out�[��Col(0)) & "_" & .Cells(myY, outCavCol(0)) & "_" & .Cells(myY, out��Col(0))
                        End If
                        �}���}�F(a) = Replace(.Cells(myY, out�}Col(a)), " ", "")
                        
                        If �F�Ŕ��f = True Or �n����ƕ\�� <> "" Then
                            my�F(a) = ��(a).Font.color
                            my�Fb(a) = ��(a).Interior.color
                            For i2 = 1 To UBound(�n���F�ݒ�, 2)
                                If my�F(a) = �n���F�ݒ�(1, i2) And my�Fb(a) = �n���F�ݒ�(3, i2) Then
                                    �n��(a) = �n���F�ݒ�(2, i2)
                                    �n��n(a) = �n���F�ݒ�(0, i2)
                                End If
                            Next i2
                        Else
                            If Left(�[�q(a), 4) = "7009" Then
                                my�F(a) = RGB(150, 150, 240)
                                �n��(a) = "Earth"
                            ElseIf Left(�[�q(a), 4) = "7409" Then
                                my�F(a) = RGB(150, 240, 150)
                                �n��(a) = "Bonda"
                            ElseIf JCDF <> "" And �[��(a) = "" Then
                                my�F(a) = RGB(200, 200, 200)
                                �n��(a) = "JOINT"
                                .Range(.Cells(myY, out2�[��Col(a)), .Cells(myY, out2�n��Col(a))).Interior.color = my�F(a)
                            Else
                                '�[���ꗗ����[���̃T�u�i���o�[������
                                findResult = False
                                findC = ref�^�C�g��Ran.Find(���i�i��, , , 1).Column
                                For refY = ref���Row + 1 To refLastRow
                                    If Replace(���(a), "-", "") = ref���i�ʒ[���ꗗRan(refY, ref���Col) Then
                                        If �[��(a) = ref���i�ʒ[���ꗗRan(refY, ref�[��Col) Then
                                            ref�T�u = ref���i�ʒ[���ꗗRan(refY, findC).Value
                                            If CStr(my�T�u) = CStr(ref�T�u) Then
                                                my�F(a) = RGB(240, 150, 150)
                                                �n��(a) = "��n��"
                                            Else
                                                my�F(a) = xlNone
                                                �n��(a) = "��"
                                            End If
                                            findResult = True
                                            Exit For
                                        End If
                                    End If
                                Next refY
                                If findResult = 0 Then Stop '�[���ꗗ�ɊY���������������
                            End If
                        End If
                        
                        '.Cells(myY, out��Col(a)).Interior.Color = my�F(a)
                        If �F�Ŕ��f = True Or �n����ƕ\�� <> "" Then
                            .Cells(myY, out2�n��Col(a)).Font.color = my�F(a)
                            .Cells(myY, out2�n��Col(a)).Font.Bold = True
                            .Cells(myY, out2�n��Col(a) + 1).Font.color = my�F(a)
                            .Cells(myY, out2�n��Col(a) + 1).Font.Bold = True
                        Else
                            .Cells(myY, out2�n��Col(a)).Interior.color = my�F(a)
                        End If
                        .Cells(myY, out2���i�i��Col) = ���i�i��
                        .Cells(myY, out2�T�uCol) = my�T�u
                        .Cells(myY, out2�\��Col) = Left(.Cells(myY, out�d�����ʖ�Col), 4)
                        .Cells(myY, out2�ڑ�Gcol) = �ڑ�G
                        .Cells(myY, out2����Col) = .Cells(myY, out����Col)
                        .Cells(myY, out2�F��col) = �F��
                        Select Case ���Y�敪
                            Case "#", "*", "=", "<" '�c�C�X�g
                                ��ƋL�� = "Tw"
                            Case "E"           '�V�[���h
                                ��ƋL�� = "S"
                            Case Else
                                ��ƋL�� = ""
                        End Select
                        .Cells(myY, out2�F��col + 1) = ��ƋL��
                        .Cells(myY, out2�[��Col(a)) = .Cells(myY, out�[��Col(a))
                        .Cells(myY, out2��Col(a)) = .Cells(myY, out��Col(a))
                        .Cells(myY, out2�n��Col(a)) = �n��(a)
                        If ��n����Ǝ� = True Then
                            If �n��(a) = "��n��" Then
                                For b = LBound(��n����Ǝ�RAN, 2) + 1 To UBound(��n����Ǝ�RAN, 2)
                                    If Left(�\��, 4) = ��n����Ǝ�RAN(0, b) Then
                                        .Cells(myY, out2�n��Col(a)) = "��n���F" & ��n����Ǝ�RAN(1, b)
                                        GoTo line10
                                    End If
                                Next b
                                Stop '��n����Ǝ҂�������Ȃ�
                            End If
line10:
                        End If
                        .Cells(myY, out2�n��Col(a) + 1) = �n��n(a)
                        .Cells(myY, out2�}Col(a)) = �}���}�F(a)
                        .Cells(myY, out����Col(a)) = ����(a)
                        If Left(�n��(a), 3) = "��n��" And �}���}�F(a) <> "" Then
                            .Cells(myY, out2�}Col(a)) = .Cells(myY, out2�}Col(a)) & "��"
                            Call �F�ϊ�(�}���}�F(a), clocode1, clocode2, clofont)
                            .Cells(myY, out2�}Col(a)).Characters(Len(�}���}�F(a)) + 1, 1).Font.color = clocode1
                        End If
NextA:
                    Next a
                    If �F�Ŕ��f = True Then
                        If �n��(0) = �n��(1) Then ���[��n��flg = "1" Else ���[��n��flg = "0"
                    Else
                        If �n��(0) = "��n��" And �n��(1) = "��n��" Then ���[��n��flg = "1" Else ���[��n��flg = "0"
                    End If
                    .Cells(myY, out���[�n��Col) = ���[��n��flg
                    If �[�q(0) = �[�q(1) Then ���[���[�qflg = "1" Else ���[���[�qflg = "0"
                    .Cells(myY, out���[���[�qCol) = ���[���[�qflg
                    Call �d���F�ŃZ����h��(myY, out2�F��col + 1, �F��)
nextY:
                Next myY
NextR:
        Next r
        If ���i�i��RANc = 1 Then
            .Cells(out�^�C�g��Row - 1, out2���i�i��Col) = my���i�i��
        Else
            .Cells(out�^�C�g��Row - 1, out2���i�i��Col) = ""
        End If
    End With
    
    Set out�^�C�g��Ran = Nothing
    
    Call �œK�����ǂ�
    
End Function

Public Function PVSWcsv����G�t����p�T�u�i���o�[txt�o��_Ver187()
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"

    Dim �ؒf(0) As String: Dim xx As Long
    �ؒf(0) = ""
    '�ؒf(1) = "SS"
    
    ���type = ""
    
    Call ���i�i��RAN_set2(���i�i��RAN, ���type, "", "")
        
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim lastRow As Long, KoseiCol As Long, firstRow As Long, keyRow As Long
        KoseiCol = .Cells.Find("�d�����ʖ�", , , xlWhole).Column
        keyRow = .Cells.Find("�d�����ʖ�", , , xlWhole).Row
        firstRow = keyRow + 1
        lastRow = .Cells(Rows.count, KoseiCol).End(xlUp).Row
    End With
    
    '�o�͐�e�L�X�g�t�@�C���ݒ�
    Dim outPutAddress As String: outPutAddress = ActiveWorkbook.Path & "\�T�u�i���o�[temp.txt"
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    
    Open outPutAddress For Output As #lntFlNo
    
    Dim �T�u�l As String, �\�� As String, ���i�i�� As String
    Dim ���� As Date: ���� = Now
    Dim X As Long, Y As Long, fndX As Long
    
    '�G�t�ւ̈�������ɐؒf�R�[�h���܂ވׁA�ؒf�R�[�h���ς���������ł��Ȃ��̂Œm���Ă�R�[�h�S�ďo�� ���P�v�΍�͐ؒf�R�[�h����������O���� ���V���V�X�e�����Ή��ς�2017/09/05��
    For xx = LBound(�ؒf) To UBound(�ؒf)
        For X = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            With Workbooks(myBookName).Sheets(mySheetName)
            ���i�i�� = ���i�i��RAN(1, X)
            fndX = .Rows(keyRow).Find(���i�i��, , , 1).Column
            ���i�i��v = Replace(���i�i��, " ", "")
                For Y = firstRow To lastRow
                        �T�u�l = .Cells(Y, fndX).Value
                        If �T�u�l = "" Then GoTo line20
                        �\�� = Left(.Cells(Y, KoseiCol), 4)
                        Print #lntFlNo, Chr(34) & �ؒf(xx) & Chr(34) & _
                                        Chr(44) & Chr(34) & ���i�i��v & Chr(34) & _
                                        Chr(44) & _
                                        Chr(44) & Chr(34) & �\�� & Chr(34) & _
                                        Chr(44) & Chr(34) & �T�u�l & Chr(34) & _
                                        Chr(44) & ����
                    
line20:
    
                Next Y
            End With
        Next X
    Next xx
    
    Close #lntFlNo
    
End Function

Public Function PVSWcsv����G�t����p�T�u�i���o�[txt�o��_Ver2012(ByVal myIP As String)
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheet As Worksheet: Set mySheet = myBook.Sheets("PVSW_RLTF")
    myIP = Mid(myIP, InStr(myIP, ".") + 1)
    myIP = Mid(myIP, InStr(myIP, ".") + 1)
    myIP = Left(myIP, InStrRev(myIP, ".") - 1)

    Dim kumitateList As Variant, myPosSP As Variant, mySQLon(1) As String
    If myIP = "120" Then '�����H��
        myPath = �A�h���X(0) & "\IP�ʐݒ�\" & myIP & "\kumitateCode.txt"
        kumitateList = readTextToArray(myPath)
        myPosSP = Split(",1,,2,3,4", ",") '�ؒf�A���i�i�ԁA�ݕρA�\���A�T�u�A�g���̏��@���̗񂪖����ꍇ�͋�
        mySQLon(0) = " ON a.F1 = b.F1 AND a.F2 = b.F2 " '���i�i�Ԃƍ\���i���o�[�̈ʒu�ԍ�
        mySQLon(1) = " ON a.F1 = b.F1 AND a.F2 = b.F2 WHERE b.F1 is null"
    ElseIf myIP = "140" Then
        myPath = �A�h���X(0) & "\IP�ʐݒ�\" & myIP & "\kumitateCode.txt"
        kumitateList = readTextToArray(myPath)
        If IsEmpty(kumitateList) Then ReDim kumitateList(0, 1)
        myPosSP = Split("1,2,3,4,5,,", ",")
        mySQLon(0) = " ON a.F2 = b.F2 AND a.F4 = b.F4 " '���i�i�Ԃƍ\���i���o�[�̈ʒu�ԍ�
        mySQLon(1) = " ON a.F2 = b.F2 AND a.F4 = b.F4 WHERE b.F2 is null"
    Else
        Stop '����IP�͓o�^����Ă��܂���
        kumitateList = ""
    End If
    
    Dim �ؒf(0) As String: Dim xx As Long
    �ؒf(0) = ""
    ���type = ""

    Call ���i�i��RAN_set2(���i�i��RAN, ���type, "", "")
    DoEvents

    With mySheet
        Dim lastRow As Long, KoseiCol As Long, firstRow As Long, keyRow As Long
        KoseiCol = .Cells.Find("�d�����ʖ�", , , xlWhole).Column
        keyRow = .Cells.Find("�d�����ʖ�", , , xlWhole).Row
        firstRow = keyRow + 1
        lastRow = .Cells(Rows.count, KoseiCol).End(xlUp).Row
    End With

    Call �A�h���X�Z�b�g(myBook)

    With myBook.Sheets("�ݒ�")
        temp�A�h���X = myBook.Path & "\efu_subNo_temp.txt"    '�G�t����̃T�u����f�[�^
        temp�A�h���X2 = myBook.Path & "\efu_subNo_temp2.txt"  '���̃t�@�C���̃T�u����f�[�^
        temp�A�h���X3 = myBook.Path & "\efu_subNo_temp3.txt"  '��L���������V��������f�[�^
    End With
    
    '1_�T�u�i���o�[����Ɏg���Ă���t�@�C�����J�����g�f�B���N�g���ɃR�s�[
    If Dir(�A�h���X(2)) = "" Then Stop ' �T�u����A�h���X�̃t�@�C���܂ł����Ȃ��A�V�[�g�ݒ�̃A�h���X�������Ă��鎖�̊m�F
    FileCopy �A�h���X(2), temp�A�h���X
    DoEvents
    '�v�ǉ�_�d���f�[�^������΍폜 ��1�ɑ΂��čs��
    
    '2_���̃t�@�C���̃T�u�i���o�[�f�[�^���쐬
    Call SQL_�T�u�i���o�[���_�f�[�^�쐬(���i�i��RAN, mySheet, temp�A�h���X2, myPosSP, kumitateList)
    DoEvents
    '3_1�ɑ΂�2�ōX�V�����t�@�C�����쐬
    Call SQL_�T�u�i���o�[���_�f�[�^�X�V(temp�A�h���X, temp�A�h���X2, temp�A�h���X3, mySQLon)
    DoEvents
    '4_�T�u�i���o�[����t�@�C���̃o�b�N�A�b�v���쐬
    �T�u����A�h���Xbak = Left(�A�h���X(2), InStrRev(�A�h���X(2), ".") - 1) & "_" & Replace(CStr(Date), "/", "") & "_0" & ".txt"
    Do
        If Dir(�T�u����A�h���Xbak) = "" Then Exit Do
        i = i + 1
        �T�u����A�h���Xbak = Left(�A�h���X(2), InStrRev(�A�h���X(2), ".") - 1) & "_" & Replace(CStr(Date), "/", "") & "_" & i & ".txt"
        If i > 50 Then Stop ' ��������H
    Loop
    FileCopy �A�h���X(2), �T�u����A�h���Xbak
    DoEvents
    '5_�X�V�����t�@�C�����T�u����p�ɂ���
    FileCopy temp�A�h���X3, �A�h���X(2)
    DoEvents
    '�o�͐�e�L�X�g�t�@�C���ݒ�
    'Dim outPutAddress As String: outPutAddress = ActiveWorkbook.path & "\�T�u�i���o�[temp.txt"
    
    PlaySound ("�����Ă�")
    MsgBox "�o�͂��������܂����B", , "���Y����+"

End Function


Public Function �T�u�}�쐬()
    '���O��[Ver181_PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬]�̎��s���K�v
    Dim my���i�i��(1) As String
    If �T�u�}���i�i�� = "" Then
        my���i�i��(0) = "821113B240"
    Else
        my���i�i��(0) = �T�u�}���i�i��
    End If
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newBookName As String: newBookName = Left(myBookName, InStr(myBookName, "_")) & "�T�u�}_" & my���i�i��(0)
    Dim baseBookName As String: baseBookName = "����_�T�u�}.xlsx"
    Dim �n���}sheetName As String: �n���}sheetName = ActiveSheet.Name
    
    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�͈�key As Range: Set ���i�͈�key = .Cells.Find("���C���i��", , , 1)
        Dim ���i�͈�Ran As Range: Set ���i�͈�Ran = .Range(.Cells(���i�͈�key.Row + 1, ���i�͈�key.Column), .Cells(.Cells(.Rows.count, ���i�͈�key.Column).End(xlUp).Row, ���i�͈�key.Column + 1))
    End With

    Dim i As Long
    '�G�A�o�b�N�i�Ԃ�T��
    For i = 1 To ���i�͈�Ran.count / 2
        If Replace(my���i�i��(0), " ", "") = Replace(���i�͈�Ran(i, 1), " ", "") Then
            my���i�i��(1) = Replace(���i�͈�Ran(i, 2), " ", "")
            Exit For
        End If
    Next
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(myBookpath & "\40_�T�u�}\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    '������ǂݎ���p�ŊJ��
    On Error Resume Next
    Workbooks.Open fileName:=Left(myBookpath, InStrRev(myBookpath, "\")) & "000_�V�X�e���p�[�c\" & baseBookName, ReadOnly:=True
    On Error GoTo 0
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBookpath & "\40_�T�u�}\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'PVSW_RLTF
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim in�^�C�g��Row As Long: in�^�C�g��Row = .Cells.Find("�i��_", , , 1).Row
        Dim in�^�C�g��Col As Long: in�^�C�g��Col = .Cells(in�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Dim in�^�C�g��Ran As Range: Set in�^�C�g��Ran = .Range(.Cells(in�^�C�g��Row, 1), .Cells(in�^�C�g��Row, in�^�C�g��Col))
        Dim in�d�����ʖ�Col As Long: in�d�����ʖ�Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim in�W���C���gGCol As Long: in�W���C���gGCol = .Cells.Find("�W���C���g�O���[�v", , , 1).Column
        Dim in�i��Col As Long: in�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim in�T�C�YCol As Long: in�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim in�FCol As Long: in�FCol = .Cells.Find("�F_", , , 1).Column
        Dim inABcol As Long: inABcol = .Cells.Find("AB_", , , 1).Column
        Dim in�F��Col(1) As Long
        in�F��Col(0) = in�^�C�g��Ran.Cells.Find("�F��_", , , 1).Column
        in�F��Col(1) = in�^�C�g��Ran.Cells.Find("�d���F", , , 1).Column
        Dim in�����i��col As Long: in�����i��col = .Cells.Find("�����i��", , , 1).Column
        Dim in����Col As Long: in����Col = .Cells.Find("����_", , , 1).Column
        Dim in��Col(1) As Long
        in��Col(0) = .Cells.Find("�n�_����H����", , , 1).Column
        in��Col(1) = .Cells.Find("�I�_����H����", , , 1).Column
        Dim in�[��Col(1) As Long
        in�[��Col(0) = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        in�[��Col(1) = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        Dim in���Col(1) As Long
        in���Col(0) = .Cells.Find("�n�_���[�����i��", , , 1).Column
        in���Col(1) = .Cells.Find("�I�_���[�����i��", , , 1).Column
        Dim in�[�qCol(1) As Long
        in�[�qCol(0) = .Cells.Find("�n�_���[�q�i��", , , 1).Column
        in�[�qCol(1) = .Cells.Find("�I�_���[�q�i��", , , 1).Column
        '��n���p�f�[�^
        Dim in2�n��Col(1) As Long
        in2�n��Col(0) = .Cells.Find("�n�_���n��", , , 1).Column
        in2�n��Col(1) = .Cells.Find("�I�_���n��", , , 1).Column
        Dim in2���i�i��Col As Long: in2���i�i��Col = .Cells.Find("���i�i��", , , 1).Column
        Dim in2�T�uCol As Long: in2�T�uCol = .Cells.Find("�T�u", , , 1).Column
        Dim in2�F��Col As Long: in2�F��Col = .Cells.Find("�F��", , , 1).Column
        Dim in2�[��Col(1) As Long
        in2�[��Col(0) = .Cells.Find("�n�_���[��", , , 1).Column
        in2�[��Col(1) = .Cells.Find("�I�_���[��", , , 1).Column
        Dim in2��Col(1) As Long
        in2��Col(0) = .Cells.Find("�n�_����H����_", , , 1).Column
        in2��Col(1) = .Cells.Find("�I�_����H����_", , , 1).Column
        Dim in2����Col As Long: in2����Col = .Cells.Find("����_", , , 1).Column
        Dim in2�\��Col As Long: in2�\��Col = .Cells.Find("�\��", , , 1).Column: .Columns(in2�\��Col).NumberFormat = "@"
        Dim inLastRow As Long: inLastRow = .Cells(.Rows.count, in�d�����ʖ�Col).End(xlUp).Row
        Dim inLastCol As Long: inLastCol = .Cells(in�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Dim inPVSWcsvRAN As Range: Set inPVSWcsvRAN = .Range(.Cells(1, 1), .Cells(inLastRow, inLastCol))
    End With
    
    Dim myVal As Range
    Dim Y As Long, addRow As Long
    
    'DB�ɓd�������o��
    addRow = 1
    For Y = in�^�C�g��Row To inLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            If Replace(.Cells(Y, in2���i�i��Col), " ", "") = my���i�i��(0) Or (my���i�i��(1) <> "" And Replace(.Cells(Y, in2���i�i��Col), " ", "") = my���i�i��(1)) Then
                addRow = addRow + 1
                Set myVal = .Range(.Cells(Y, in2���i�i��Col), .Cells(Y, in2�n��Col(1)))
                myVal.Copy Workbooks(newBookName).Sheets("DB").Cells(addRow, 2)
            End If
        End With
    Next Y
    'DB����בւ�
    With Workbooks(newBookName).Sheets("DB")
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 4).address), Order:=xlAscending
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    'DB�̃T�u�i���o�[���ɃV�[�g�쐬
    For Y = 2 To addRow
        With Workbooks(newBookName).Sheets("DB")
            If .Cells(Y, 3) <> .Cells(Y + 1, 3) Then
                Sheets("base").Copy after:=Sheets("DB")
                ActiveSheet.Name = .Cells(Y, 3)
                ActiveSheet.Cells(2, 12) = Replace(.Cells(Y, 2), " ", "")
                ActiveSheet.Cells(2, 15) = .Cells(Y, 3)
                ActiveSheet.PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
                ActiveSheet.PageSetup.RightHeader = "&R" & "&14 " & my���i�i��(0) & "&14 �T�u- " & .Cells(Y, 3) & "  " & "&P/&N"
            End If
        End With
    Next Y
    'DB�̃f�[�^���T�u�i���o�[�V�[�g�ɏo��
    Dim startRow As Long, �T�u As String
    Dim �[�� As String
    startRow = 2
    For Y = 2 To addRow
        With Workbooks(newBookName)
            With .Sheets("DB")
                �T�u = .Cells(Y, 3)
                If �T�u <> .Cells(Y + 1, 3) Then
                    Set myVal = .Cells(startRow, 4).Resize(Y - startRow + 1, 11)
                    myVal.Copy Workbooks(newBookName).Sheets(�T�u).Range("a5")
                    With Workbooks(newBookName).Sheets(�T�u).Range("a5").Resize(Y - startRow + 1, 11)
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlInsideVertical).LineStyle = xlContinuous
                        .Font.Size = 16
                    End With
                    startRow = Y + 1
                End If
            End With
        End With
    Next Y
    
    '���i�ʒ[���ꗗ�̃Z�b�g
    With Workbooks(myBookName).Sheets("���i�ʒ[���ꗗ")
        Dim ���i�ʒ[���ꗗ() As Variant: ReDim ���i�ʒ[���ꗗ(2, 0)
        Dim ref�[��key As Range: Set ref�[��key = .Cells.Find("�[����", , , 1)
        Dim ref�[��Col As Long: ref�[��Col = ref�[��key.Column
        Dim ref�T�uCol As Long: ref�T�uCol = .Cells.Find(my���i�i��(0) & String(15 - Len(my���i�i��(0)), " "), , , 1).Column
        Dim ref���Col As Long: ref���Col = .Cells.Find("�[�����i��", , , 1).Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ref�[��Col).End(xlUp).Row
        For i = ref�[��key.Row + 1 To lastRow
            If .Cells(i, ref�T�uCol) <> "" Then
                c = c + 1
                ReDim Preserve ���i�ʒ[���ꗗ(2, c)
                ���i�ʒ[���ꗗ(0, c) = .Cells(i, ref�[��Col)
                ���i�ʒ[���ꗗ(1, c) = .Cells(i, ref�T�uCol)
                ���i�ʒ[���ꗗ(2, c) = .Cells(i, ref���Col)
            End If
        Next i
    End With
    addrow40 = 5
    addrow50 = 6
    x40 = 1
    x50 = 1
    Dim �i�[V() As Variant: ReDim Preserve �i�[V(1)
    '���i���X�g�̍쐬
    With Workbooks(myBookName).Sheets("���i���X�g")
        Dim ���i���X�gkey As Range: Set ���i���X�gkey = .Cells.Find(my���i�i��(0), , , 1)
        lastRow = .Cells(.Rows.count, ���i���X�gkey.Column).End(xlUp).Row
        Dim ���i�i��Col As Long: ���i�i��Col = ���i���X�gkey.Column
        Dim ���i�i��Col As Long: ���i�i��Col = .Cells.Find("���i�i��", , , 1).Column
        Dim �T�C�Y1Col As Long: �T�C�Y1Col = .Cells.Find("����1", , , 1).Column
        Dim �T�C�Y2Col As Long: �T�C�Y2Col = .Cells.Find("����2", , , 1).Column
        Dim �ؒf��Col As Long: �ؒf��Col = .Cells.Find("�ؒf��", , , 1).Column
        Dim �H��Col(1) As Long: �H��Col(0) = .Cells.Find("�H��", , , 1).Column
        Dim �[��Col As Long: �[��Col = ���i���X�gkey.Column
        Dim ����Col As Long: ����Col = .Cells.Find("����", , , 1).Column
        Dim ���Col As Long: ���Col = .Cells.Find("���", , , 1).Column
        Dim �\��Col As Long: �\��Col = .Cells.Find("�\����", , , 1).Column
        Dim ���ޏڍ�Col As Long: ���ޏڍ�Col = .Cells.Find("���ޏڍ�", , , 1).Column
        For i = ���i���X�gkey.Row + 1 To lastRow
            ���i�i�� = .Cells(i, ���i�i��Col)
            �H�� = .Cells(i, �H��Col(0))
            �\�� = .Cells(i, �\��Col)
            ���i�i�� = .Cells(i, ���i�i��Col)
            ��� = .Cells(i, ���Col)
            ���� = 1
            �[�� = .Cells(i, �[��Col)
            ���ޏڍ� = .Cells(i, ���ޏڍ�Col)
            �T�u = ""
            If �[�� <> "" Then
                If �H�� = "40" Then
                    If �[�� <> "" Then
                        For cc = 1 To c
                            If CStr(���i�ʒ[���ꗗ(0, cc)) = �[�� Then
                                �T�u = ���i�ʒ[���ꗗ(1, cc)
                                Exit For
                            End If
                        Next cc
                    End If
                    With Workbooks(newBookName).Sheets("base2")
                        .Cells(addrow40, x40 + 1).Value = �A�h���X '���g���ĂȂ�
                        .Cells(addrow40, x40 + 2).Value = �T�u
                        .Cells(addrow40, x40 + 0) = �\��
                        .Cells(addrow40, x40 + 3).Value = ���i�i��
                        .Cells(addrow40, x40 + 4) = ����
                        .Cells(addrow40, x40 + 6) = ���ޏڍ�
                        addrow40 = addrow40 + 1
                    End With
                Else
                    Select Case �H��
                    Case "50"
                        �H��Col(1) = 2
                    Case "60"
                        �H��Col(1) = 3
                    Case "70"
                        �H��Col(1) = 4
                    Case "80"
                        �H��Col(1) = 5
                    Case Else
                        �H��Col(1) = 0
                    End Select
                    If �H��Col(1) <> 0 Then
                        With Workbooks(newBookName).Sheets("base3")
                            .Cells(addrow50, x50 + 0) = �\��
                            .Cells(addrow50, x50 + 1) = �A�h���X '���g���ĂȂ�
                            .Cells(addrow50, x50 + �H��Col(1)).Value = "��"
                            .Cells(addrow50, x50 + 6).Value = ���i�i��
                            .Cells(addrow50, x50 + 7) = ����
                            .Cells(addrow50, x50 + 9) = ���ޏڍ�
                            addrow50 = addrow50 + 1
                        End With
                    End If
                End If
            End If
        Next i
    End With
    
    With Workbooks(newBookName).Sheets("base2")
        For cc = 1 To c
            f = 0
            For i = 5 To addrow40
                If CStr(���i�ʒ[���ꗗ(2, cc)) = Replace(.Cells(i, 4), "-", "") Then
                    If .Cells(i, 3) = "" Then
                        .Cells(i, 3) = ���i�ʒ[���ꗗ(1, cc)
                        f = 1
                        Exit For
                    End If
                End If
            Next i
            If f = 0 Then
                If (���i�ʒ[���ꗗ(2, cc)) <> "74099913" Then
                    Debug.Print (���i�ʒ[���ꗗ(2, cc))
                    Stop '�������i�ʒ[���ꗗ���猩����Ȃ�
                End If
            End If
        Next cc
    End With
    
    '�\���܂Ƃ�
    With Workbooks(newBookName).Sheets("base2")
        .Range("g2") = my���i�i��(0)
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(1).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(2).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(3).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(4).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(8).LineStyle = 1
        .Columns(6).Borders(12).LineStyle = -4142
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Font.Size = 16
        .Columns(7).Font.Name = "�l�r �S�V�b�N"
        .PageSetup.PrintArea = .Range(.Cells(5, 1), .Cells(addrow40 - 1, 5))
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my���i�i��(0) & "&14 ���  " & "&P/&N"
        .Name = "���"
    End With
    With Workbooks(newBookName).Sheets("base3")
        .Range("h2") = my���i�i��(0)
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(1).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(2).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(3).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(4).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(8).LineStyle = 1
        .Columns(9).Borders(12).LineStyle = -4142
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Font.Size = 16
        .Columns(10).Font.Name = "�l�r �S�V�b�N"
        .PageSetup.PrintArea = ""
        '.PageSetup.PrintArea = .Range(.Cells(6, 1), .Cells(addRow50 - 1, 13))
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my���i�i��(0) & "&14 ��t  " & "&P/&N"
        .Name = "��t"
    End With
    Application.DisplayAlerts = False
    Worksheets("base").Delete
    Application.DisplayAlerts = True
    
    Call �œK��
    '�}�̔z�u
    '���i�ʒ[���ꗗ����בւ�
    With Workbooks(myBookName).Sheets("���i�ʒ[���ꗗ")
        Dim refKeyRow As Long: refKeyRow = .Cells.Find("�[�����i��", , , 1).Row
        Dim refKeyCol As Long: refKeyCol = .Cells.Find("�[�����i��", , , 1).Column
        Dim refKey2Col As Long: refKey2Col = .Cells.Find("�[����", , , 1).Column
        Dim refLastCol As Long: refLastCol = .Cells(refKeyRow, .Columns.count).End(xlToLeft).Column
        Dim refLastRow As Long: refLastRow = .Cells(.Rows.count, refKeyCol).End(xlUp).Row
        Dim X As Long, ref���i�i��col As Long
        For X = refKeyCol To refLastCol
            If Replace(.Cells(refKeyRow, X), " ", "") = Replace(my���i�i��(0), " ", "") Then
                ref���i�i��col = X
                Exit For
            End If
            If X = refLastCol Then Stop '���i�i�Ԃ�������Ȃ��ُ�
        Next X
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, ref���i�i��col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�}��z�z���Ă���
        Dim objShp As Shape
        Dim �T�ubak As String, addpoint As Long
        For Y = refKeyRow + 1 To refLastRow
            If .Cells(Y, refKey2Col) <> "" Then
                �T�u = .Cells(Y, ref���i�i��col)
                If �T�u <> "" Then
                    If �T�u <> �T�ubak Then
                        addpoint = Workbooks(newBookName).Sheets(�T�u).Cells(.Rows.count, 1).End(xlUp).Top + 32.25
                        �T�ubak = �T�u
                    End If
                    �[�� = .Cells(Y, refKey2Col)
                    With Workbooks(myBookName).Sheets(�n���}sheetName)
                        For Each objShp In Workbooks(myBookName).Sheets(�n���}sheetName).Shapes
                            'Debug.Print objShp.Name
                            If �[�� = Left(objShp.Name, InStr(objShp.Name, "_") - 1) Then
                                'Stop
                                objShp.Copy 'Workbooks(newBookName).Sheets(�T�u)
                                DoEvents
                                Sleep 5
                                DoEvents
                                Workbooks(newBookName).Sheets(�T�u).Paste
                                Workbooks(newBookName).Sheets(�T�u).Shapes(objShp.Name).Left = 3
                                Workbooks(newBookName).Sheets(�T�u).Shapes(objShp.Name).Top = addpoint
                                addpoint = addpoint + Workbooks(newBookName).Sheets(�T�u).Shapes(objShp.Name).Height + 13.5
                                Workbooks(newBookName).Sheets(�T�u).Activate
                                Workbooks(newBookName).Sheets(�T�u).Cells(1, 15).Select
                            End If
                        Next objShp
                    End With
                End If
            End If
        Next Y
    End With
    
    '���i�ʒ[���ꗗ�̕��т���ɖ߂�
    With Workbooks(myBookName).Sheets("���i�ʒ[���ꗗ")
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
    Call �œK�����ǂ�
    
    MsgBox "�쐬���܂���"
End Function

Public Function �T�u�}�쐬_Ver2023(my���i�i��) As String
    '���O��[Ver181_PVSWcsv�ɃT�u�i���o�[��n���ăT�u�}�f�[�^�쐬]�̎��s���K�v
    Call �œK��
    Set wb(0) = ActiveWorkbook
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newBookName As String: newBookName = Left(myBookName, InStr(myBookName, "_")) & "�T�u�}_" & Replace(my���i�i��, " ", "")
    Dim baseBookName As String: baseBookName = "����_�T�u�}_2.191.13.xlsx"
    
    Dim �n���}sheetName As String: �n���}sheetName = ActiveSheet.Name
'
'    With Workbooks(myBookName).Sheets("���i�i��")
'        �n���}�A�h���X = .Cells.Find("System+", , , 1).Offset(0, 1).Value
'        Dim ���i�͈�key As Range: Set ���i�͈�key = .Cells.Find("���C���i��", , , 1)
'        Dim ���i�͈�Ran As Range: Set ���i�͈�Ran = .Range(.Cells(���i�͈�key.Row + 1, ���i�͈�key.Column), .Cells(.Cells(.Rows.count, ���i�͈�key.Column).End(xlUp).Row, ���i�͈�key.Column + 1))
'    End With

    Dim i As Long
    '�G�A�o�b�N�i�Ԃ�T��
'    For i = 1 To ���i�͈�Ran.count / 2
'        If Replace(my���i�i��(0), " ", "") = Replace(���i�͈�Ran(i, 1), " ", "") Then
'            my���i�i��(1) = Replace(���i�͈�Ran(i, 2), " ", "")
'            Exit For
'        End If
'    Next
    '�o�͐�f�B���N�g����������΍쐬
    If Dir(myBookpath & "\40_�T�u�}", vbDirectory) = "" Then
        MkDir myBookpath & "\40_�T�u�}"
    End If
    
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(myBookpath & "\40_�T�u�}\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    '������ǂݎ���p�ŊJ��
    Workbooks.Open fileName:=�A�h���X(0) & "\genshi\" & baseBookName, ReadOnly:=True
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBookpath & "\40_�T�u�}\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
    '�v���O���X�o�[
    ProgressBar.Show vbModeless
    Dim step0T As Long, step0 As Long
    step0T = 1: step0 = step0 + 1
    Call ProgressBar_ref(�O���[�v��� & "_" & �O���[�v��, "�T�u�}�̍쐬��", step0T, step0, 100, 100)
    'PVSW_RLTF
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim in�^�C�g��Row As Long: in�^�C�g��Row = .Cells.Find("�i��_", , , 1).Row
        Dim in�^�C�g��Col As Long: in�^�C�g��Col = .Cells(in�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Dim in�^�C�g��Ran As Range: Set in�^�C�g��Ran = .Range(.Cells(in�^�C�g��Row, 1), .Cells(in�^�C�g��Row, in�^�C�g��Col))
        Dim in�d�����ʖ�Col As Long: in�d�����ʖ�Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim in�W���C���gGCol As Long: in�W���C���gGCol = .Cells.Find("�W���C���g�O���[�v", , , 1).Column
        Dim in�i��Col As Long: in�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim �ڑ�Gcol As Long: �ڑ�Gcol = .Cells.Find("�ڑ�G_", , , 1).Column
        
        Dim in�T�C�YCol As Long: in�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim in�FCol As Long: in�FCol = .Cells.Find("�F_", , , 1).Column
        Dim inABcol As Long: inABcol = .Cells.Find("AB_", , , 1).Column
        Dim in�F��Col(1) As Long
        in�F��Col(0) = in�^�C�g��Ran.Cells.Find("�F��_", , , 1).Column
        in�F��Col(1) = in�^�C�g��Ran.Cells.Find("�d���F", , , 1).Column
        Dim in�����i��col As Long: in�����i��col = .Cells.Find("�����i��", , , 1).Column
        Dim in����Col As Long: in����Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim in��Col(1) As Long
        in��Col(0) = .Cells.Find("�n�_����H����", , , 1).Column
        in��Col(1) = .Cells.Find("�I�_����H����", , , 1).Column
        Dim in�[��Col(1) As Long
        in�[��Col(0) = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        in�[��Col(1) = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        Dim in���Col(1) As Long
        in���Col(0) = .Cells.Find("�n�_���[�����i��", , , 1).Column
        in���Col(1) = .Cells.Find("�I�_���[�����i��", , , 1).Column
        Dim in�[�qCol(1) As Long
        in�[�qCol(0) = .Cells.Find("�n�_���[�q�i��", , , 1).Column
        in�[�qCol(1) = .Cells.Find("�I�_���[�q�i��", , , 1).Column
        '��n���p�f�[�^
        Dim in2�n��Col(1) As Long
        in2�n��Col(0) = .Cells.Find("�n�_���n��", , , 1).Column
        in2�n��Col(1) = .Cells.Find("�I�_���n��", , , 1).Column
        Dim in2���i�i��Col As Long: in2���i�i��Col = .Cells.Find("���i�i��", , , 1).Column
        Dim in2�T�uCol As Long: in2�T�uCol = .Cells.Find("�T�u", , , 1).Column
        Dim in2�F��Col As Long: in2�F��Col = .Cells.Find("�F��", , , 1).Column
        Dim in2����Col As Long: in2����Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim in3����Col As Long: in3����Col = .Cells.Find("����__", , , 1).Column
        Dim in2�[��Col(1) As Long
        in2�[��Col(0) = .Cells.Find("�n�_���[��", , , 1).Column
        in2�[��Col(1) = .Cells.Find("�I�_���[��", , , 1).Column
        Dim in2��Col(1) As Long
        in2��Col(0) = .Cells.Find("�n�_����H����_", , , 1).Column
        in2��Col(1) = .Cells.Find("�I�_����H����_", , , 1).Column
        Dim in2����Col As Long: in2����Col = .Cells.Find("����_", , , 1).Column
        Dim in2�\��Col As Long: in2�\��Col = .Cells.Find("�\��", , , 1).Column: .Columns(in2�\��Col).NumberFormat = "@"
        Dim inLastRow As Long: inLastRow = .Cells(.Rows.count, in�d�����ʖ�Col).End(xlUp).Row
        Dim inLastCol As Long: inLastCol = .Cells(in�^�C�g��Row, .Columns.count).End(xlToLeft).Column
        Dim inPVSWcsvRAN As Range: Set inPVSWcsvRAN = .Range(.Cells(1, 1), .Cells(inLastRow, inLastCol))
    End With
    
    Dim myVal As Range
    Dim Y As Long, addRow As Long
    
    'DB�ɓd�������o��
    addRow = 1
    For Y = in�^�C�g��Row To inLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            If .Cells(Y, in2���i�i��Col) = my���i�i�� Then
                addRow = addRow + 1
                Set myVal = .Range(.Cells(Y, in2���i�i��Col), .Cells(Y, in3����Col))
                myVal.Copy Workbooks(newBookName).Sheets("DB").Cells(addRow, 2)
            End If
        End With
    Next Y
    With Workbooks(newBookName).Sheets("DB")
        'Bonda���E�Ɉړ�
        Dim myRange(1) As Range
        Dim fff(1) As Long
        For Y = 2 To addRow
            If .Cells(Y, 11) = "Bonda" Then
                Set myRange(0) = .Range(.Cells(Y, 13), .Cells(Y, 17))
                Set myRange(1) = .Range(.Cells(Y, 8), .Cells(Y, 12))
                myRange(0).Cut
                .Cells(Y, 8).Insert Shift:=xlToRight
'                .Range(.Cells(y, 11), .Cells(y, 14)) = myRange(1).Value
'                .Range(.Cells(y, 7), .Cells(y, 10)) = myRange(0).Value
            End If
        Next Y
        'DB����ёւ�_1���
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 15).address), Order:=xlAscending
            .add key:=Range(Cells(2, 12).address), Order:=xlAscending
            .add key:=Range(Cells(2, 17).address), Order:=xlAscending
        End With
        .Sort.SetRange .Range(.Rows(2), .Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�s�v�ȗ���폜
        .Columns(12).Delete
        .Columns(17).Delete
        
        '�\�������ێ����Ȃ���ڑ�G���܂Ƃ߂���ёւ����o�u���\�[�g�ōs��
        Dim sw As Boolean, thisRow As Integer
        For Y = 2 To addRow
            �ڑ�gstr = .Cells(Y, 5)
            If �ڑ�gstr = "" Then GoTo line15
            '�����͂��߂̐ڑ�G�Ȃ�ŏI�s�܂Œ��Ԃ�T��
            thisRow = 0
            sw = sw + 1
            If sw Then .Cells(Y, 5).Interior.color = RGB(130, 130, 130)
            For Y2 = Y + 1 To addRow
                If �ڑ�gstr = .Cells(Y2, 5) Then
                    thisRow = thisRow + 1
                    If sw Then .Cells(Y2, 5).Interior.color = RGB(130, 130, 130)
                    '���̍s�łȂ��ꍇ�͈ړ�
                    If Y + thisRow <> Y2 Then
                        .Rows(Y2).Cut
                        .Rows(Y + thisRow).Insert Shift:=xlDown
                    End If
                End If
            Next Y2
line15:
            Y = Y + thisRow
        Next Y
    End With
    'DB�̃T�u�i���o�[���ɃV�[�g�쐬
    For Y = 2 To addRow
        With Workbooks(newBookName).Sheets("DB")
            If CStr(.Cells(Y, 3)) <> CStr(.Cells(Y + 1, 3)) Then
                Sheets("base").Copy before:=Sheets("DB")
                ActiveSheet.Name = .Cells(Y, 3)
                ActiveSheet.Cells(2, 13).NumberFormat = "@"
                ActiveSheet.Cells(2, 13) = Replace(.Cells(Y, 2), " ", "")
                ActiveSheet.Cells(2, 14).NumberFormat = "@"
                ActiveSheet.Cells(2, 14) = .Cells(Y, 3)
                ActiveSheet.PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 8)
                ActiveSheet.PageSetup.RightHeader = "&R" & "&14 " & Replace(my���i�i��, " ", "") & "&14 �T�u- " & .Cells(Y, 3) & "  " & "&P/&N"
            End If
        End With
    Next Y
    
    'DB�̃f�[�^���T�u�i���o�[�V�[�g�ɏo��
    Dim startRow As Long, �T�u As String
    Dim �[�� As String
    startRow = 2
    For Y = 2 To addRow
        With Workbooks(newBookName)
            With .Sheets("DB")
                �T�u = .Cells(Y, 3)
                If �T�u <> .Cells(Y + 1, 3) Then
                    Set myVal = .Cells(startRow, 4).Resize(Y - startRow + 1, 13)
                    myVal.Copy Workbooks(newBookName).Sheets(�T�u).Range("a5")
                    Sheets(�T�u).Activate

                    With Workbooks(newBookName).Sheets(�T�u).Range("a5").Resize(Y - startRow + 1, 13)
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlInsideVertical).LineStyle = xlContinuous
                        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        .Font.Size = 13
                        Workbooks(newBookName).Sheets(�T�u).Range("e5").Resize(Y - startRow + 1, 1).Borders(xlEdgeLeft).Weight = xlMedium
                        Workbooks(newBookName).Sheets(�T�u).Range("i5").Resize(Y - startRow + 1, 1).Borders(xlEdgeLeft).Weight = xlMedium
                        Workbooks(newBookName).Sheets(�T�u).Range("m5").Resize(Y - startRow + 1, 1).Borders(xlEdgeLeft).Weight = xlMedium
                    End With
                    If cbxQR = True Then
                        For i = 5 To Y - startRow + 5
                            myQR = "           " & Sheets(�T�u).Cells(i, 1).Value & "          " & my���i�i��
                            Call QR�R�[�h���N���b�v�{�[�h�Ɏ擾(myQR)
                            Workbooks(newBookName).Sheets(�T�u).PasteSpecial Format:="�} (JPEG)", Link:=False, DisplayAsIcon:=False
                            Selection.Height = Workbooks(newBookName).Sheets(�T�u).Cells(i, 1).Height
                            Selection.Top = Workbooks(newBookName).Sheets(�T�u).Cells(i, 1).Top + 0.5
                            Selection.Left = Workbooks(newBookName).Sheets(�T�u).Cells(i, 2).Left - Selection.Width
                        Next i
                    End If
                    '�{���_�[�̒[�����Ƀo�[�O���t
                    �[��r = 5
                    For i = 5 To Y - startRow + 5
                        �[�� = Sheets(�T�u).Cells(i, 8)
                        �敪 = Sheets(�T�u).Cells(i, 11)
                        If �敪 <> "Bonda" Then Exit For
                        If �[�� <> Sheets(�T�u).Cells(i + 1, 8) And �敪 = "Bonda" Then
                            �Fbf = Cells(i, 2)
                            If InStr(�Fbf, "/") > 0 Then
                                �Fb = Left(�Fbf, InStr(�Fbf, "/") - 1)
                            Else
                                �Fb = �Fbf
                            End If
                            Call �F�ϊ�(�Fb, clocode1, clocode2, clofont)
                            �F�R�[�h = clocode1
                            Sheets(�T�u).Range(Cells(i, 1), Cells(i, 13)).Borders(xlEdgeBottom).Weight = xlMedium
                            Sheets(�T�u).Range(Cells(�[��r, 12), Cells(i, 12)).FormatConditions.AddDatabar
                            
                            With Sheets(�T�u).Range(Cells(�[��r, 12), Cells(i, 12)).FormatConditions(1)
                                .BarColor.color = �F�R�[�h
                                .BarBorder.color.TintAndShade = 0
                                .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
                                .BarBorder.Type = xlDataBarBorderSolid
                            End With
                                
                            If �Fb = "W" Then
                                Sheets(�T�u).Range(Cells(�[��r, 12), Cells(i, 12)).Interior.color = RGB(200, 200, 200)
                            End If
                            If �Fb = "B" Then
                                Sheets(�T�u).Range(Cells(�[��r, 12), Cells(i, 12)).Font.color = RGB(255, 255, 255)
                            End If

                            For yyy = �[��r To i
                                If Sheets(�T�u).Cells(yyy, 7) = "��n��" Then
                                    Sheets(�T�u).Cells(yyy, 13) = yyy - �[��r + 1
                                End If
                            Next yyy
                            �[��r = i + 1
                        End If
                    Next i
                    startRow = Y + 1
                End If
            End With
        End With
    Next Y
    
    '���i�ʒ[���ꗗ�̃Z�b�g
    With Workbooks(myBookName).Sheets("�[���ꗗ")
        Dim ���i�ʒ[���ꗗ() As Variant: ReDim ���i�ʒ[���ꗗ(2, 0)
        Dim ref�[��key As Range: Set ref�[��key = .Cells.Find("�[����", , , 1)
        Dim ref�[��Col As Long: ref�[��Col = ref�[��key.Column
        Dim ref�T�uCol As Long: ref�T�uCol = .Cells.Find(my���i�i��, , , 1).Column
        Dim ref���Col As Long: ref���Col = .Cells.Find("�[�����i��", , , 1).Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ref�[��Col).End(xlUp).Row
        For i = ref�[��key.Row + 1 To lastRow
            If .Cells(i, ref�T�uCol) <> "" Then
                c = c + 1
                ReDim Preserve ���i�ʒ[���ꗗ(2, c)
                ���i�ʒ[���ꗗ(0, c) = .Cells(i, ref�[��Col)
                ���i�ʒ[���ꗗ(1, c) = .Cells(i, ref�T�uCol)
                ���i�ʒ[���ꗗ(2, c) = .Cells(i, ref���Col)
            End If
        Next i
    End With
    addrow40 = 5
    addrow50 = 6
    x40 = 1
    x50 = 1
    
    Dim �i�[V() As Variant: ReDim Preserve �i�[V(1)
    '���i���X�g�̍쐬
    With Workbooks(myBookName).Sheets("���i���X�g")
        Dim ���i���X�gkey As Range: Set ���i���X�gkey = .Cells.Find(my���i�i��, , , 1)
        lastRow = .Cells(.Rows.count, ���i���X�gkey.Column).End(xlUp).Row
        Dim ���i�i��Col As Long: ���i�i��Col = ���i���X�gkey.Column
        Dim ���i�i��Col As Long: ���i�i��Col = .Cells.Find("���i�i��", , , 1).Column
        Dim �T�C�Y1Col As Long: �T�C�Y1Col = .Cells.Find("����1", , , 1).Column
        Dim �T�C�Y2Col As Long: �T�C�Y2Col = .Cells.Find("����2", , , 1).Column
        Dim �ؒf��Col As Long: �ؒf��Col = .Cells.Find("�ؒf��", , , 1).Column
        Dim �H��Col(1) As Long: �H��Col(0) = .Cells.Find("�H��", , , 1).Column
        Dim �H��aCol As Long: �H��aCol = .Cells.Find("�H��a", , , 1).Column
        Dim �[��Col As Long: �[��Col = ���i���X�gkey.Column
        
        Dim ���Col As Long: ���Col = .Cells.Find("���", , , 1).Column
        Dim �\��Col As Long: �\��Col = .Cells.Find("�\����", , , 1).Column
        Dim ���ޏڍ�Col As Long: ���ޏڍ�Col = .Cells.Find("���ޏڍ�", , , 1).Column
        For i = ���i���X�gkey.Row + 1 To lastRow
            ���i�i�� = .Cells(i, ���i�i��Col)
            �H�� = .Cells(i, �H��Col(0))
            �H��a = .Cells(i, �H��aCol)
            If �H��a <> "" Then �H�� = �H��a
            �\�� = .Cells(i, �\��Col)
            ���i�i�� = .Cells(i, ���i�i��Col)
            ��� = .Cells(i, ���Col)
            ���� = 1
            �[�� = .Cells(i, �[��Col)
            ���ޏڍ� = .Cells(i, ���ޏڍ�Col)
            �T�u = ""
            If �[�� <> "" Then
                If �H�� = "40" Then
                    For cc = 1 To c
                        If CStr(���i�ʒ[���ꗗ(0, cc)) = �[�� Then
                            �T�u = ���i�ʒ[���ꗗ(1, cc)
                            Exit For
                        End If
                    Next cc
                    With Workbooks(newBookName).Sheets("base2")
                        '.Cells(addrow40, x40 + 1).Value = �A�h���X '���g���ĂȂ�
                        .Cells(addrow40, x40 + 2).Value = �T�u
                        .Cells(addrow40, x40 + 0) = �\��
                        .Cells(addrow40, x40 + 3).Value = ���i�i��
                        .Cells(addrow40, x40 + 4) = ����
                        .Cells(addrow40, x40 + 6) = ���ޏڍ�
                        If ��� = "B" Then
                            .Cells(addrow40, x40 + 10) = "1"
                        ElseIf ��� = "T" Then
                            .Cells(addrow40, x40 + 10) = "2"
                        End If
                        .Cells(addrow40, x40 + 11) = ���
                        addrow40 = addrow40 + 1
                    End With
                Else
                    Select Case �H��
                    Case "45"
                        �H��Col(1) = 2
                    Case "50"
                        �H��Col(1) = 3
                    Case "60"
                        �H��Col(1) = 4
                    Case "70"
                        �H��Col(1) = 5
                    Case "80", "90"
                        �H��Col(1) = 6
                    Case Else
                        �H��Col(1) = 0
                    End Select
                    If �H��Col(1) <> 0 Then
                        With Workbooks(newBookName).Sheets("base3")
                            .Cells(addrow50, x50 + 0) = �\��
                            '.Cells(addrow50, x50 + 1) = �A�h���X '���g���ĂȂ�
                            .Cells(addrow50, x50 + �H��Col(1)).Value = "��"
                            .Cells(addrow50, x50 + 7).Value = ���i�i��
                            .Cells(addrow50, x50 + 8) = ����
                            .Cells(addrow50, x50 + 10) = ���ޏڍ�
                            If ��� = "B" Then
                                .Cells(addrow50, x50 + 13) = "1"
                            ElseIf ��� = "T" Then
                                .Cells(addrow50, x50 + 13) = "2"
                            End If
                            .Cells(addrow50, x50 + 14) = ���
                            addrow50 = addrow50 + 1
                        End With
                    End If
                End If
            End If
        Next i
    End With
    
    '���i�ʒ[���ꗗ�����Ƃߕ��i���X�g�ɏo��
    With Workbooks(newBookName).Sheets("base2")
        For cc = 1 To c
            f = 0
            For i = 5 To addrow40
                If CStr(���i�ʒ[���ꗗ(2, cc)) = Replace(.Cells(i, 4), "-", "") Then
                    If .Cells(i, 3) = "" Then
                        .Cells(i, 3) = ���i�ʒ[���ꗗ(1, cc)
                        .Cells(i, 13) = ���i�ʒ[���ꗗ(0, cc)
                        .Cells(i, 11) = 1
                        .Cells(i, 12) = "A"
                        f = 1
                        Exit For
                    End If
                End If
            Next i
            If f = 0 Then
                If Left(���i�ʒ[���ꗗ(2, cc), 4) <> "7409" Then
                    If Left(���i�ʒ[���ꗗ(2, cc), 4) <> "7009" Then
                        Debug.Print (���i�ʒ[���ꗗ(2, cc))
                        Stop '�������i�ʒ[���ꗗ����bese2���݂����A������Ȃ�����
                    End If
                End If
            End If
        Next cc
    End With
    
    Call SQL_�T�u�}_��Ƃߕ��i���X�g_���(���RAN, my���i�i��, myBookName)
    'PVSW���[���������Ƃߕ��i���X�g�ɏo��
    With Workbooks(newBookName).Sheets("base2")
        For e = LBound(���RAN, 2) + 1 To UBound(���RAN, 2)
            �[��e = ���RAN(0, e)
            ��� = ���RAN(1, e)
            f = 0
            For i = 5 To addrow40
                If ��� = .Cells(i, 4) Then
                    If .Cells(i, 3) = "" Then
                        '�T�u�i���o�[��T��
                        �T�uflg = False
                        For cc = 1 To c
                            If �[��e = CStr(���i�ʒ[���ꗗ(0, cc)) Then
                                .Cells(i, 3) = ���i�ʒ[���ꗗ(1, cc)
                                �T�uflg = True
                                GoTo line20
                            End If
                        Next cc
                    End If
                End If
            Next i
            If �T�uflg = False Then
                Debug.Print �[��e & "_" & ���
                Stop '�����̏�����������Ȃ�
                '����̏ꍇ�́A[CAV�ꗗ]�̋��i�Ԃ������Ă邩�m�F
            End If
line20:
        Next e
    End With
    
    '�\���܂Ƃ�
    With Workbooks(newBookName).Sheets("base2")
        .Activate
        .Name = "���"
        .Range("g2") = my���i�i��
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(1).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(2).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(3).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(4).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(8).LineStyle = 1
        .Columns(6).Borders(12).LineStyle = -4142
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Font.Size = 16
        .Columns(7).Font.Name = "�l�r �S�V�b�N"
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(addrow40 - 1, 10)).address
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my���i�i�� & "&14 ���  " & "&P/&N"
        '���ёւ�
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 11).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(5), Rows(addrow40 - 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        '�����܂Ƃ߂�
        For Y = 5 To addrow40 - 1
            If .Cells(Y, 4) = "" Then Exit For
            If .Cells(Y, 3) = .Cells(Y + 1, 3) And .Cells(Y, 4) = .Cells(Y + 1, 4) And .Cells(Y, 7) = .Cells(Y + 1, 7) Then
                'Stop
                .Cells(Y, 5) = .Cells(Y, 5) + .Cells(Y + 1, 5)
                If .Cells(Y, 13) <> "" Then
                    .Cells(Y, 13) = .Cells(Y, 13) & "_" & .Cells(Y + 1, 13)
                End If
                .Rows(Y + 1).Delete
                addrow40 = addrow40 - 1
                Y = Y - 1
            End If
        Next Y
        Dim �T���v���^�ORAN() As String
        ReDim �T���v���^�ORAN(5, addrow40 - 6)
        '���ёւ�_�T���v���^�O�쐬�f�[�^�p
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 12).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(5), Rows(addrow40 - 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        For i = 5 To addrow40 - 1
            �T���v���^�ORAN(0, i - 5) = .Cells(i, 3)
            �T���v���^�ORAN(1, i - 5) = .Cells(i, 4)
            �T���v���^�ORAN(2, i - 5) = .Cells(i, 5)
            �T���v���^�ORAN(3, i - 5) = .Cells(i, 7)
            �T���v���^�ORAN(4, i - 5) = .Cells(i, 12)
            �T���v���^�ORAN(5, i - 5) = .Cells(i, 13)
        Next i
        
        '���ёւ�
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 11).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(5), Rows(addrow40 - 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        '���ёւ�_�`���[�u����
        Set myRow = .Columns(12).Find("T", , , 1)
        If Not (myRow Is Nothing) Then
            tRow = myRow.Row
            With .Sort.SortFields
                .Clear
                .add key:=Range(Cells(6, 7).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
                .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            End With
            .Sort.SetRange Range(Rows(tRow), Rows(addrow40 - 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
        End If
        
        
        '�X�g���C�v
        For Y = 5 To addrow40 - 1
            If Y Mod 2 = 0 Then .Rows(Y).Interior.color = RGB(220, 220, 220)
        Next Y
    End With
    
    With Workbooks(newBookName).Sheets("base3")
        .Name = "��t"
        .Range("i2") = my���i�i��
        '�����܂Ƃ߂�
        For Y = 5 To addrow50 - 1
            If .Cells(Y, 8) = "" Then Exit For
            If .Cells(Y, 3) = .Cells(Y + 1, 3) And .Cells(Y, 4) = .Cells(Y + 1, 4) And .Cells(Y, 5) = .Cells(Y + 1, 5) Then
                If .Cells(Y, 6) = .Cells(Y + 1, 6) And .Cells(Y, 7) = .Cells(Y + 1, 7) And .Cells(Y, 8) = .Cells(Y + 1, 8) Then
                    If .Cells(Y, 11) = .Cells(Y + 1, 11) Then
                    'Stop
                    .Cells(Y, 9) = .Cells(Y, 9) + .Cells(Y + 1, 9)
                    .Rows(Y + 1).Delete
                    addrow50 = addrow50 - 1
                    Y = Y - 1
                    End If
                End If
            End If
        Next Y
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(1).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(2).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(3).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(4).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(8).LineStyle = 1
        .Columns(10).Borders(12).LineStyle = -4142
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Font.Size = 16
        .Columns(10).Font.Name = "�l�r �S�V�b�N"
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(addrow50 - 1, 13)).address
        '.PageSetup.PrintArea = .Range(.Cells(6, 1), .Cells(addRow50 - 1, 13))
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my���i�i�� & "&14 ��t  " & "&P/&N"
        
        '�X�g���C�v
        For Y = 5 To addrow50 - 1
            If Y Mod 2 = 1 Then .Rows(Y).Interior.color = RGB(200, 200, 200)
        Next Y
    End With
    
    With Workbooks(newBookName).Sheets("base4")
        .Name = "�^�O"
        .Activate
        Dim �摜URL As String, partName As String, �摜�� As String, ���v As Long, �^�Orow As Long
        Dim yy As Long
        For i = LBound(�T���v���^�ORAN, 2) To UBound(�T���v���^�ORAN, 2)
            If �T�u�^�Obak <> �T���v���^�ORAN(0, i) Or i = LBound(�T���v���^�ORAN, 2) Then
                .Range(.Rows(1 + yy), Rows(44 + yy)).Copy .Range(.Rows(45 + yy), .Rows(88 + yy))
                .Range("e" & 4 + yy) = my���i�i��
                If �T���v���^�ORAN(0, i) <> "" Then
                    .Range("e" & 5 + yy) = �T���v���^�ORAN(0, i)
                Else
                    .Range("e" & 5 + yy) = "�ΏۊO"
                End If
                aRow = 9 + yy: aCou = 0
                tRow = 24 + yy: tCou = 0
                bRow = 36 + yy: bCou = 0
                yy = yy + 44
                ActiveSheet.HPageBreaks.add (.Cells(yy + 1, 21))
            End If
            
            Select Case �T���v���^�ORAN(4, i)
                Case "A"
                �^�Orow = aRow
                aRow = aRow + 1
                aCou = aCou + �T���v���^�ORAN(2, i)
                ���v = aCou
                ����x = 1
                ���i���� = ""
                Case "B"
                �^�Orow = bRow
                bRow = bRow + 1
                bCou = bCou + �T���v���^�ORAN(2, i)
                ���v = bCou
                ����x = 1
                ���i���� = "_" & �T���v���^�ORAN(3, i)
                Case "T"
                �^�Orow = tRow
                tRow = tRow + 1
                tCou = tCou + �T���v���^�ORAN(2, i)
                ���v = tCou
                ����x = 3
                ���i���� = ""
            End Select
            partName = �T���v���^�ORAN(����x, i)
            .Cells(�^�Orow, 5) = partName & ���i����
            
            �摜flg = 0
            '�ʐ^
            �摜URL = �A�h���X(1) & "\���ވꗗ+_�ʐ^\" & partName & "_1_" & Format(1, "000") & ".png"
            If Dir(�摜URL) = "" Then
                '���}
                �摜URL = �A�h���X(1) & "\���ވꗗ+_���}\" & partName & "_1_" & Format(1, "000") & ".emf"
                If Dir(�摜URL) = "" Then GoTo line18
            End If
            
            �摜�� = partName & "_" & �^�Orow
            
            Dim myHeight As Single, myWidth As Single, cellHeight As Single, myScale As Single
            cellHeight = .Cells(�^�Orow, 4).Height
            With .Pictures.Insert(�摜URL)
                .Name = �摜��
                .ShapeRange(�摜��).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft '�摜���傫���ƃT�C�Y������������邩���̃T�C�Y�ɖ߂�
                myHeight = .Height
                myWidth = .Width
                myScale = cellHeight / myHeight
                .ShapeRange(�摜��).ScaleHeight myScale, msoTrue, msoScaleFromTopLeft
                .CopyPicture
                .Delete
            End With
            DoEvents
            Sleep 10
            DoEvents
            .Paste
            Selection.Name = �摜��
            .Shapes(�摜��).Height = .Cells(�^�Orow, 4).Height
            .Shapes(�摜��).Left = .Cells(�^�Orow, 7).Left - .Shapes(�摜��).Width - 1
            .Shapes(�摜��).Top = .Cells(�^�Orow, 4).Top
line18:
            .Cells(�^�Orow, 7) = �T���v���^�ORAN(2, i)
            .Cells(�^�Orow, 8) = �T���v���^�ORAN(5, i)
            .Cells(�^�Orow + 1, 7) = ���v
            
            �T�u�^�Obak = �T���v���^�ORAN(0, i)
        Next i
        .Range(Rows(yy + 1), Rows(yy + 44)).Delete
    End With
    
    Application.DisplayAlerts = False
    Worksheets("base").Delete
    Application.DisplayAlerts = True
    
    Call �œK��
    '�}�̔z�u
    '���i�ʒ[���ꗗ����בւ�
    With Workbooks(myBookName).Sheets("�[���ꗗ")
        Dim refKeyRow As Long: refKeyRow = .Cells.Find("�[�����i��", , , 1).Row
        Dim refKeyCol As Long: refKeyCol = .Cells.Find("�[�����i��", , , 1).Column
        Dim refKey2Col As Long: refKey2Col = .Cells.Find("�[����", , , 1).Column
        Dim refLastCol As Long: refLastCol = .Cells(refKeyRow, .Columns.count).End(xlToLeft).Column
        Dim refLastRow As Long: refLastRow = .Cells(.Rows.count, refKeyCol).End(xlUp).Row
        Dim X As Long, ref���i�i��col As Long
        For X = refKeyCol To refLastCol
            If Replace(.Cells(refKeyRow, X), " ", "") = Replace(my���i�i��, " ", "") Then
                ref���i�i��col = X
                Exit For
            End If
            If X = refLastCol Then Stop '���i�i�Ԃ�������Ȃ��ُ�
        Next X
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, ref���i�i��col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '�}��z�z���Ă���
        Dim objShp As Shape
        Dim �T�ubak As String, addRowPoint As Long, addRowPoint2 As Long
        For Y = refKeyRow + 1 To refLastRow
            If .Cells(Y, refKey2Col) <> "" Then
                �T�u = .Cells(Y, ref���i�i��col)
'                If �T�u = "17" Then Stop
                If �T�u <> "" Then
                    If �T�u <> �T�ubak Then
                        On Error Resume Next
                        Workbooks(newBookName).Sheets(�T�u).Activate
                        If Err = 9 Then
                            If InStr(�����T�u, vbCrLf & �T�u & vbCrLf) = 0 Then
                            �����T�u = �����T�u & vbCrLf & �T�u & vbCrLf
                            End If
                            On Error GoTo 0
                            GoTo nextY
                        End If
                        On Error GoTo 0
                        addRowPoint = Workbooks(newBookName).Sheets(�T�u).Cells(.Rows.count, 1).End(xlUp).Top + 32.25
                        nextrowpoint = addRowPoint
                        addcolpoint = 3
                        maxcolpoint = addcolpoint
                        �T�ubak = �T�u
                        c = 0
                        zure = 0
                    End If
                    
                    �[�� = .Cells(Y, refKey2Col)
                    With Workbooks(myBookName).Sheets(�n���}sheetName)
                        'For Each objShp In Workbooks(myBookName).Sheets(�n���}sheetName).Shapes
                            'Debug.Print objShp.Name
                            .Shapes(�[�� & "_1").Copy
                            'If �[�� = left(objShp.Name, InStr(objShp.Name, "_") - 1) Then
                            'Stop
                            'objShp.Copy 'Workbooks(newBookName).Sheets(�T�u)
                            'Sleep 5
                            DoEvents
                            Sleep 10
                            DoEvents
                            Workbooks(newBookName).Sheets(�T�u).Paste
                            
                            'Workbooks(newBookName).Sheets(�T�u).Shapes(objShp.Name).left = 3
                            'Workbooks(newBookName).Sheets(�T�u).Shapes(objShp.Name).Top = addRowPoint
                            'addRowPoint = addRowPoint + Workbooks(newBookName).Sheets(�T�u).Shapes(objShp.Name).Height + 13.5
                            
                            '�z�u��̃A�h���X���v�Z
                            addRowPoint2 = addRowPoint + Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Height
                            If (addRowPoint - zure) \ 597 <> (addRowPoint2 - zure) \ 597 Then 'Y����������͈͂��o�鎞
                                maxcolpoint2 = maxcolpoint + Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Width
                                If maxcolpoint2 < 878 Then ' X�����Ɏ��܂鎞
                                    If c = 0 Then '1���ڂ̃n���}����ʂɎ��܂�Ȃ���
                                        Workbooks(newBookName).Sheets(�T�u).HPageBreaks.add Workbooks(newBookName).Sheets(�T�u).Cells(.Rows.count, 1).End(xlUp).Offset(1, 0)
                                        addRowPoint = Workbooks(newBookName).Sheets(�T�u).Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Top + 3
                                        nextrowpoint = addRowPoint
                                        zure = nextrowpoint
                                    Else
                                        addRowPoint = nextrowpoint
                                        addcolpoint = maxcolpoint
                                    End If
                                Else                       ' X�����Ɏ��܂�Ȃ���
                                    addRowPoint = ((addRowPoint \ 597) + 1) * 597
                                    addcolpoint = 3
                                    nextrowpoint = addRowPoint
                                End If
                            End If
                            '�z�u
                            Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Left = addcolpoint
                            Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Top = addRowPoint
                            If Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Left + Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Width > maxcolpoint Then
                                maxcolpoint = Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Left + Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Width + 3
                            End If
                            
                            addRowPoint = addRowPoint + Workbooks(newBookName).Sheets(�T�u).Shapes(�[�� & "_1").Height + 3
                            
                            Workbooks(newBookName).Sheets(�T�u).Cells(1, 15).Select
                            c = c + 1
                            'End If
                        'Next objShp
                    End With
                    With Workbooks(newBookName).Sheets(�T�u)
                        .PageSetup.PrintArea = False
'                        For Each n In ActiveWorkbook.Names
'                            If n.Name = "'" & �T�u & "'!Print_Area" Then
'                                PrintLastRow = Mid(n.Value, InStrRev(n.Value, "$") + 1)
'                                Exit For
'                            End If
'                        Next
'                        ActiveSheet.PageSetup.PrintArea = Range(Cells(1, 1), Cells(Val(PrintLastRow), 15)).Address
                    End With
                End If
            End If
nextY:
        Next Y
    End With
    
    '���i�ʒ[���ꗗ�̕��т���ɖ߂�
    With Workbooks(myBookName).Sheets("�[���ꗗ")
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
    Call �œK�����ǂ�
    Unload ProgressBar
    DoEvents
    
    Application.DisplayAlerts = False
        'Workbooks(newBookName).Save
    Application.DisplayAlerts = True
    
    �T�u�}�쐬_Ver2023 = �����T�u
    
End Function



Function ���i���X�g�̍쐬()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "���i���X�g"
    
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    
    '���i�i�Ԃ̃��C���i�Ԃ�RLTF��Ǎ���
    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�i��key As Range: Set ���i�i��key = .Cells.Find("���C���i��", , , 1)
        Dim RLTFkey As Range: Set RLTFkey = .Cells.Find("RLTF", , , 1)
        Dim ���i�i��lastRow As Long: ���i�i��lastRow = .Cells(.Rows.count, ���i�i��key.Column).End(xlUp).Row
        Dim ��������() As String: ReDim ��������(���i�i��lastRow - ���i�i��key.Row, 2)
        Dim ���i�_�� As Long: ���i�_�� = ���i�i��lastRow - ���i�i��key.Row
        Dim n As Long
        For n = 1 To ���i�_��
            ��������(n, 1) = .Cells(���i�i��key.Row + n, ���i�i��key.Column)
            ��������(n, 2) = .Cells(RLTFkey.Row + n, RLTFkey.Column)
        Next n
        Set ���i�i��key = Nothing
        Set RLTFkey = Nothing
    End With
    
    '���ޏڍ�txt�̓Ǎ���
    Dim ���ޏڍ�() As String
    Dim TargetFile As String: TargetFile = Left(myBookpath, InStrRev(myBookpath, "\")) & "\000_�V�X�e���p�[�c\���ޏڍ�" & ".txt"
    Dim intFino As Integer
    Dim aRow As String, aCel As Variant, ���ޏڍ�c As Long: ���ޏڍ�c = -1
    Dim ���ޏڍ�v As String
    intFino = FreeFile
    Open TargetFile For Input As #intFino
    Do Until EOF(intFino)
        Line Input #intFino, aRow
        aCel = Split(aRow, ",")
        ���ޏڍ�c = ���ޏڍ�c + 1
        For a = LBound(aCel) To UBound(aCel)
            ReDim Preserve ���ޏڍ�(UBound(aCel), ���ޏڍ�c)
            ���ޏڍ�(a, ���ޏڍ�c) = aCel(a)
        Next a
    Loop
    Close #intFino
    
    Dim �i�[V() As Variant: ReDim �i�[V(0)
    Dim V(15) As String
    Dim c As Long
    '�^�C�g���s
    �i�[V(c) = "���i�i��,�ݕ�,�\����,���i�i��,�ď�,����1,����2,�F,�ؒf��,,,�H��,���,����,�[����,���ޏڍ�"
    '���i�i�Ԗ���RLTF����ǂݍ���
    For n = 1 To ���i�_��
        
        '���͂̐ݒ�(�C���|�[�g�t�@�C��)
        TargetFile = myBookpath & "\05_RLTF_A\" & ��������(n, 2) & ".txt"
        
        intFino = FreeFile
        Open TargetFile For Input As #intFino
        Do Until EOF(intFino)
            Line Input #intFino, aRow
            If Replace(��������(n, 1), " ", "") = Replace(Mid(aRow, 1, 15), " ", "") Then
                If Mid(aRow, 27, 1) = "T" Then '�`���[�u
                    V(0) = Replace(Mid(aRow, 1, 15), " ", "") '���i�i��
                    V(1) = Mid(aRow, 19, 3)   '�ݕ�
                    V(2) = Mid(aRow, 27, 4)   'T�\����
                    V(3) = Replace(Mid(aRow, 375, 8), " ", "") '���i�i��
                    Select Case Len(V(3))
                        Case 8
                            V(3) = Left(V(3), 3) & "-" & Mid(V(3), 4, 3) & "-" & Mid(V(3), 7, 3)
                        Case Else
                            Stop
                    End Select
                    V(4) = Mid(aRow, 383, 6)  'T�ď�
                    V(5) = Mid(aRow, 389, 4)  'T����1
                    V(6) = Mid(aRow, 393, 4)  'T����2
                    V(7) = Replace(Mid(aRow, 397, 6), " ", "") 'T�F
                    V(8) = CLng(Mid(aRow, 403, 5))  'T�ؒf��
                    V(9) = Mid(aRow, 544, 1) '�Ȃ�1
                    V(10) = Mid(aRow, 544, 4) '�Ȃ�2
                    V(11) = Mid(aRow, 153, 2)  '�H��
                    V(12) = "T"
                    V(13) = 1 '����
                If V(5) <> "    " And V(6) <> "    " Then 'VO
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "�~" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & " L=" & String(4 - Len(Format(Mid(aRow, 403, 5), 0)), " ") & Format(Mid(aRow, 403, 5), 0)
                ElseIf V(5) <> "    " Then 'COT
                    V(15) = Left(V(3), 3) & "-D" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "�~" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                ElseIf V(6) <> "    " Then 'VS
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & "�~" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                End If
                    GoSub �i�[���s
                ElseIf Mid(aRow, 27, 1) = "B" Then '40�H���ȍ~�̕��i
                    For X = 0 To 9
                        If Mid(aRow, 175 + (X * 20) + 10, 3) = "ATO" Then
                            V(0) = Replace(Mid(aRow, 1, 15), " ", "") '���i�i��
                            V(1) = Mid(aRow, 19, 3)   '�ݕ�
                            V(2) = ""                 'T�\����
                            V(3) = Replace(Mid(aRow, 175 + (X * 20), 10), " ", "") '���i�i��
                            Select Case Len(V(3))
                                Case 8
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4)
                                Case 9, 10
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4) & "-" & Mid(V(3), 9, 2)
                                Case Else
                                    Stop
                            End Select
                            '���ޏڍׂ̎擾
                            ���ޏڍ�v = ""
                            For a = 0 To ���ޏڍ�c
                                If ���ޏڍ�(0, a) = V(3) Then
                                    If Left(���ޏڍ�(2, a), 2) = "F1" Then '�N���b�v�̎�
                                        ���ޏڍ�v = Mid(���ޏڍ�(4, a), 6)
                                    Else
                                        ���ޏڍ�v = Mid(���ޏڍ�(3, a), 7)
                                    End If
                                    Exit For
                                End If
                            Next a
                            V(4) = ""  'T�ď�
                            V(5) = ""  'T����1
                            V(6) = ""  'T����2
                            V(7) = ""  'T�F
                            V(8) = ""  'T�ؒf��
                            V(9) = "" '�Ȃ�1
                            V(10) = "" '�Ȃ�2
                            V(11) = Mid(aRow, 558 + (X * 2), 2) '�H��
                            V(12) = "B"
                            V(13) = CLng(Mid(aRow, 189 + (X * 20), 4)) '����
                            V(15) = ���ޏڍ�v
                            GoSub �i�[���s
                        End If
                    Next X
                End If
            End If
        Loop
        Close #intFino
    Next n
    
    '�V�[�g�ǉ�
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    
    Dim Val As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        .Columns("A:P").NumberFormat = "@"
        .Columns("I").NumberFormat = 0
        For a = LBound(�i�[V) To UBound(�i�[V)
            Val = Split(�i�[V(a), ",")
            For b = LBound(Val) To UBound(Val)
                .Cells(a + 1, b + 1) = Val(b)
            Next b
        Next a
        'T�ď̂̃t�H���g�ݒ�
        .Columns("P").Font.Name = "�l�r �S�V�b�N"
        '�H��a�̒ǉ�
        .Columns("P").Insert
        .Range("p1") = "�H��a"
        '�t�B�b�g
        .Columns("A:q").AutoFit
        '�E�B���h�E�g�̌Œ�
        .Range("a2").Select
        ActiveWindow.FreezePanes = True
        '�r��
        With .Range(.Cells(1, 1), .Cells(UBound(�i�[V) + 1, UBound(Val) + 2))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        '�\�[�g
        With .Sort.SortFields
            .Clear
            .add key:=Cells(1, 1), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 12), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 13), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 4), Order:=xlAscending, DataOption:=0
            .add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange .Range(.Rows(2), Rows(UBound(�i�[V) + 1))
        With .Sort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With

Exit Function
�i�[���s:
    �i�[temp = V(0) & "," & V(1) & "," & V(2) & "," & V(3) & "," & V(4) & "," & V(5) & "," & V(6) & "," & V(7) & "," & V(8) & "," & V(9) & "," & V(10) & "," & V(11) & "," & V(12)
    If V(11) = "40" Or V(15) = "�X���[�N���b�v" Then
        For a = 1 To V(13)
            c = c + 1
            ReDim Preserve �i�[V(c)
            �i�[V(c) = �i�[temp & "," & 1 & ",," & V(15)
        Next a
    Else
        c = c + 1
        ReDim Preserve �i�[V(c)
        �i�[V(c) = �i�[temp & "," & V(13) & ",," & V(15)
    End If
Return

End Function

Public Function PVSWcsv���[�̃V�[�g�쐬_Ver2001()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "PVSW_RLTF���["
        
    Dim my����() As String, my����c As Long
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inKey As Range: Set inKey = .Cells.Find("�d�����ʖ�", , , 1)
        Dim lastInRow As Long: lastInRow = .Cells(.Rows.count, inKey.Column).End(xlUp).Row
        Dim lastINcol As Long: lastINcol = .Cells(inKey.Row, .Columns.count).End(xlToLeft).Column
        
        'PVSW_RLTF��Column���擾
        For X = inKey.Column To lastINcol
            If Left(.Cells(inKey.Row, X), 3) = "�I�_��" Then
                For c = 1 To my����c
                    If Mid(my����(0, c), 4) = Mid(.Cells(inKey.Row, X), 4) Then
                        my����(2, c) = .Cells(inKey.Row, X).Column
                        Exit For
                    End If
                Next c
            Else
                my����c = my����c + 1
                ReDim Preserve my����(3, my����c)
                my����(a + 0, my����c) = .Cells(inKey.Row, X)
                my����(a + 1, my����c) = .Cells(inKey.Row, X).Column
            End If
        Next X
    End With
    
    '���[�N�V�[�g�̒ǉ�
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Tab.color = False
        
    '�o�͂��鐻�i�i�Ԃ̑I��
    Dim ���i�g����c As Long, addCol As Long, addRow As Long: addRow = 1
    Dim ���i�g����() As String: ReDim Preserve ���i�g����(���i�i��RANc, 3)
    
    For Y = inKey.Row To lastInRow
        '�o��_���i�g����
        With Workbooks(myBookName).Sheets(mySheetName)
            If Y = inKey.Row Then
                For X = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                    Set f = .Rows(inKey.Row).Find(���i�i��RAN(1, X), , , 1)
                    ���i�g����(X, 0) = f.Value
                    ���i�g����(X, 1) = ""
                    ���i�g����(X, 2) = f.Column
                Next X
            Else
                ���i�g����str = ""
                For X = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                    ���i�g����(X, 1) = .Cells(Y, Val(���i�g����(X, 2)))
                    ���i�g����str = ���i�g����str & ���i�g����(X, 1)
                Next X
                If ���i�g����str = "" Then GoTo line20
            End If
        End With
        '�o��
        With Workbooks(myBookName).Sheets(newSheetName)
            If Y = inKey.Row Then
                For X = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                    .Cells.NumberFormat = "@"
                    'If ���i�o��(x) = 1 Then
                    addCol = addCol + 1
                    .Cells(1, addCol) = ���i�g����(X, 0)
                    ���i�g����(X, 3) = addCol
                    .Columns(addCol).NumberFormat = "@"
                    'End If
                Next X
                For c = 1 To my����c
                    .Cells(1, addCol + c) = Replace(my����(0, c), "�n�_��", "")
                    If InStr("�ؒf��_,�d�㐡�@_", my����(0, c)) > 0 Then
                        .Columns(addCol + c).NumberFormat = 0
                    End If
                Next c
                    .Cells(1, addCol + my����c + 1) = "��_"
                    .Cells(1, addCol + my����c + 2) = "LED_"
                    .Cells(1, addCol + my����c + 3) = "�|�C���g1_"
                    .Cells(1, addCol + my����c + 4) = "�|�C���g2_"
                    .Cells(1, addCol + my����c + 5) = "FUSE_"
                    .Cells(1, addCol + my����c + 6) = "��d�W�~_"
                    .Cells(1, addCol + my����c + 7) = "PVSWtoPOINT_"
                    .Cells(1, addCol + my����c + 8) = "�F��SI_"
            Else
                addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row + 1
                For X = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                    'If ���i�o��(x) = 1 Then
                        .Cells(addRow, CLng(���i�g����(X, 3))) = ���i�g����(X, 1)
                        .Cells(addRow + 1, CLng(���i�g����(X, 3))) = ���i�g����(X, 1)
                    'End If
                Next X
                For c = 1 To my����c
                    If my����(2, c) = "" Then
                        .Cells(addRow + 0, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my����(1, c)))
                        .Cells(addRow + 1, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my����(1, c)))
                    Else
                        .Cells(addRow + 0, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my����(1, c)))
                        .Cells(addRow + 1, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my����(2, c)))
                    End If
                    .Cells(addRow + 0, addCol + my����c + 1) = "�n"
                    .Cells(addRow + 1, addCol + my����c + 1) = "�I"
                Next c
            End If
        End With
line20:
    Next Y
    
    '�[�����i�Ԃ�74099913(bonda)�̎��A���i(���C�P���Ȃ�)�ɒu��������
    Dim ���Col As Long, ���iCol(5) As Long
    With Workbooks(myBookName).Sheets(newSheetName)
        ���Col = .Rows(1).Find("�[�����i��", , , 1).Column
        ���iCol(1) = .Rows(1).Find("���i_", , , 1).Column
        ���iCol(2) = .Rows(1).Find("���i2_", , , 1).Column
        ���iCol(3) = .Rows(1).Find("���i3_", , , 1).Column
        ���iCol(4) = .Rows(1).Find("���i4_", , , 1).Column
        ���iCol(5) = .Rows(1).Find("���i5_", , , 1).Column
        addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row
        For i = 2 To addRow
            If .Cells(i, ���Col) = "74099913" Then
                For k = 1 To 5
                    ���istr = Replace(.Cells(i, ���iCol(k)), " ", "")
                    If ���istr <> "" Then
                        .Cells(i, ���Col) = ���istr
                        GoTo line25
                    End If
                Next k
            End If
line25:
        Next i
    End With
 
    '�F��SI(�V�[���h�h����)�̎��A�d���F���`���[�u�F�ɕϊ�����
    Dim �F��Col As Long, ���i2Col As Long
    With Workbooks(myBookName).Sheets(newSheetName)
        �F��Col = .Rows(1).Find("�F��_", , , 1).Column
        ���i2Col = .Rows(1).Find("���i2_", , , 1).Column
        addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row
        For i = 2 To addRow
            If .Cells(i, �F��Col) = "SI" Then
                .Cells(i, ���i2Col).Select
                If Left(.Cells(i, ���i2Col), 4) = "7139" Then
                    'Call SQL_���ޏڍׂ̐F�擾("���ޏڍ�.txt", .Cells(i, ���i2Col), �F��SI)
                    �F��SI = ���ޏڍׂ̓ǂݍ���(�[�����i�ԕϊ�(.Cells(i, ���i2Col)), "�F_")
                    .Cells(i, addCol + my����c + 8) = �F��SI
                End If
            End If
        Next i
    End With
    
    '���בւ�
    �D��1 = addCol + 3
    �D��2 = addCol + 17
    �D��3 = addCol + 4
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow + 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
End Function

Function PVSWcsv�̋��ʉ�_Ver1944()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim i As Long, ii As Long
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        'PVSW���Ƃ��Ƃ̃f�[�^
        Dim PVSW����Row As Long: PVSW����Row = .Cells.Find("�d�����ʖ�", , , 1).Row
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim PVSW���i�i��sCol As Long: PVSW���i�i��sCol = .Cells.Find("���i�i��s", , , 1).Column
        Dim PVSW���i�i��eCol As Long
        Set ���i�i��ekey = .Cells.Find("���i�i��e", , , 1)
        If ���i�i��ekey Is Nothing Then
            PVSW���i�i��eCol = PVSW���i�i��sCol
        Else
            PVSW���i�i��eCol = ���i�i��ekey.Column
        End If
        
        Dim PVSW���i�i��RAN As Range: Set PVSW���i�i��RAN = .Range(.Cells(PVSW����Row, PVSW���i�i��sCol), .Cells(PVSW����Row, PVSW���i�i��eCol))
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW����Col).End(xlUp).Row
        Dim PVSW�d��sCol As Long: PVSW�d��sCol = .Cells.Find("�d�������擾s", , , 1).Column
        Dim PVSW�d��eCol As Long: PVSW�d��eCol = .Cells.Find("�d�������擾e", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW�n����Col As Long: PVSW�n����Col = .Cells.Find("�n�_������_", , , 1).Column
        Dim PVSW�I����Col As Long: PVSW�I����Col = .Cells.Find("�I�_������_", , , 1).Column

        Dim PVSW�n��HCol As Long: PVSW�n��HCol = .Cells.Find("�n�_����H����", , , 1).Column
        Dim PVSW�n�[��Col As Long: PVSW�n�[��Col = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        Dim PVSW�nCavCol As Long: PVSW�nCavCol = .Cells.Find("�n�_���L���r�e�B", , , 1).Column
        Dim PVSW�n���Col As Long: PVSW�n���Col = .Cells.Find("�n�_����햼��", , , 1).Column
        Dim PVSW�n���Ӑ�Col As Long: PVSW�n���Ӑ�Col = .Cells.Find("�n�_���[�����Ӑ�i��", , , 1).Column
        Dim PVSW�n���Col As Long: PVSW�n���Col = .Cells.Find("�n�_���[�����i��", , , 1).Column
        Dim PVSW�I��HCol As Long: PVSW�I��HCol = .Cells.Find("�I�_����H����", , , 1).Column
        Dim PVSW�I�[��Col As Long: PVSW�I�[��Col = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        Dim PVSW�ICavCol As Long: PVSW�ICavCol = .Cells.Find("�I�_���L���r�e�B", , , 1).Column
        Dim PVSW�I���Col As Long: PVSW�I���Col = .Cells.Find("�I�_����햼��", , , 1).Column
        Dim PVSW�I���Ӑ�Col As Long: PVSW�I���Ӑ�Col = .Cells.Find("�I�_���[�����Ӑ�i��", , , 1).Column
        Dim PVSW�I���Col As Long: PVSW�I���Col = .Cells.Find("�I�_���[�����i��", , , 1).Column
        'RLTF����擾�����f�[�^
        Dim PVSW�\��Col As Long: PVSW�\��Col = .Cells.Find("�\��_", , , 1).Column
        Dim PVSW�i��Col As Long: PVSW�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim PVSW�T�C�YCol As Long: PVSW�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim PVSW�T�C�Y�ď�Col As Long: PVSW�T�C�Y�ď�Col = .Cells.Find("�T��_", , , 1).Column
        Dim PVSW�FCol As Long: PVSW�FCol = .Cells.Find("�F_", , , 1).Column
        Dim PVSW�F��Col As Long: PVSW�F��Col = .Cells.Find("�F��_", , , 1).Column
        Dim PVSW��IDcol As Long: PVSW��IDcol = .Cells.Find("��ID_", , , 1).Column
        Dim PVSW�ڑ�Col As Long: PVSW�ڑ�Col = .Cells.Find("��ID_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        Dim PVSW�d�㐡�@Col As Long: PVSW�d�㐡�@Col = .Cells.Find("�d�㐡�@_", , , 1).Column
        Dim PVSW�ؒf��Col As Long: PVSW�ؒf��Col = .Cells.Find("�ؒf��_", , , 1).Column
        Dim PVSW�n�[Col As Long: PVSW�n�[Col = .Cells.Find("�n�_���[�q_", , , 1).Column
        Dim PVSW�n�}Col As Long: PVSW�n�}Col = .Cells.Find("�n�_���}_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_���ڑ��\��_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_�����i_", , , 1).Column
        Dim PVSW�I�[Col As Long: PVSW�I�[Col = .Cells.Find("�I�_���[�q_", , , 1).Column
        Dim PVSW�I�}Col As Long: PVSW�I�}Col = .Cells.Find("�I�_���}_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_���ڑ��\��_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_�����i_", , , 1).Column
        Dim PVSW�T�u0Col As Long: PVSW�T�u0Col = .Cells.Find("���0_", , , 1).Column
        
        '��r����
        Dim PVSW��rCol(25) As Long
        'RLTF����̃f�[�^
        PVSW��rCol(0) = PVSW�i��Col
        PVSW��rCol(1) = PVSW�T�C�YCol
        PVSW��rCol(2) = PVSW�T�C�Y�ď�Col
        PVSW��rCol(3) = PVSW�FCol
        PVSW��rCol(4) = PVSW�F��Col
        PVSW��rCol(5) = PVSW����Col
        PVSW��rCol(6) = PVSW����Col
        PVSW��rCol(7) = PVSWJCDFCol
        PVSW��rCol(8) = PVSW�d�㐡�@Col
        PVSW��rCol(9) = PVSW�ؒf��Col
        PVSW��rCol(10) = PVSW�n�[Col
        PVSW��rCol(11) = PVSW�n�}Col
        PVSW��rCol(12) = PVSW�n��Col
        PVSW��rCol(13) = PVSW�I�[Col
        PVSW��rCol(14) = PVSW�I�}Col
        PVSW��rCol(15) = PVSW�I��Col
        
        'PVSW����̃f�[�^
        PVSW��rCol(16) = PVSW�\��Col
        PVSW��rCol(17) = PVSWRLTFtoPVSWCol
        PVSW��rCol(18) = PVSW�n��HCol
        PVSW��rCol(19) = PVSW�n�[��Col
        PVSW��rCol(20) = PVSW�nCavCol
        'PVSW��rCol(20) = PVSW�n���Col
        'PVSW��rCol(19) = PVSW�n���Ӑ�Col
        PVSW��rCol(21) = PVSW�n���Col
        PVSW��rCol(22) = PVSW�I��HCol
        PVSW��rCol(23) = PVSW�I�[��Col
        PVSW��rCol(24) = PVSW�ICavCol
        'PVSW��rCol(26) = PVSW�I���Col
        'PVSW��rCol(25) = PVSW�I���Ӑ�Col
        PVSW��rCol(25) = PVSW�I���Col
        'PVSW��rCol(26) = PVSW�T�u0Col
        
        '���������ł���Γ����s�ɂ܂Ƃ߂�
        Dim ��rA() As String, ��rB() As String
        For i = PVSW����Row + 1 To PVSWlastRow
            '�����Z�b�gA
            ReDim ��rA(���i�i��c)
            For X = LBound(PVSW��rCol) To UBound(PVSW��rCol)
                ��rA(0) = ��rA(0) & .Cells(i, PVSW��rCol(X)) & "_"
            Next X
            For ii = i + 1 To PVSWlastRow
                ReDim ��rB(���i�i��c)
                '�����Z�b�gB
                For X = LBound(PVSW��rCol) To UBound(PVSW��rCol)
                    ��rB(0) = ��rB(0) & .Cells(ii, PVSW��rCol(X)) & "_"
                Next X
                'A��B�̔�r
                If ��rA(0) = ��rB(0) Then
                    '���i�i�ԃZ�b�gB
                    .Cells(i, 1).Select
                    For c = PVSW���i�i��sCol To PVSW���i�i��eCol
                        If .Cells(i, c) = "" And .Cells(ii, c) <> "" Then
                            .Cells(i, c) = .Cells(ii, c)
                            .Cells(i, c).Interior.color = .Cells(ii, c).Interior.color
                            '.Cells(ii, c).Interior.Color = xlNone
                        ElseIf .Cells(i, c) <> "" And .Cells(ii, c) <> "" Then
                            Stop '���肦��H�v�m�F
                        End If
                    Next c
                    .Cells(ii, 1).Select
                    
                    Sleep 5
                    DoEvents
                    .Rows(ii).Delete
                    ii = ii - 1: PVSWlastRow = PVSWlastRow - 1
                End If
            Next ii
        Next i
        
        '�������Ă�̂ɕ�ID���������̂ɕ�ID��^����_F�R�A�Ȃ�
        'JCDF���󗓂ł͂Ȃ��A�ڑ��\�����󗓁A��ID���Ȃ�
        Dim myJCDF As String, ��idA As Long: ��idA = 1
        For i = PVSW����Row + 1 To PVSWlastRow
            If .Cells(i, PVSW��IDcol) = "" Then
                myJCDF = .Cells(i, PVSWJCDFCol)
                If myJCDF <> "" Then
                    For i2 = i To PVSWlastRow
                        If myJCDF = .Cells(i2, PVSWJCDFCol) Then
                            If .Cells(i2, PVSW�n��Col) = "" And .Cells(i2, PVSW�I��Col) = "" Then
                                .Cells(i2, PVSW��IDcol) = "A" & ��idA
                            End If
                        End If
                    Next i2
                    ��idA = ��idA + 1
                End If
            End If
        Next i
        
        '�q�����Ă����H��ID��A�Ԃŗ^����_�ڑ�ID
        Dim ��idA As Long: ��idA = 1
        For i = PVSW����Row + 1 To PVSWlastRow
            If .Cells(i, PVSW�ڑ�Col) = "" Then
                myJCDF = .Cells(i, PVSWJCDFCol)
                If myJCDF <> "" Then
                    If .Cells(i, PVSW�n��Col) <> "" Or .Cells(i, PVSW�I��Col) <> "" Then
                        For i2 = i To PVSWlastRow
                            If myJCDF = .Cells(i2, PVSWJCDFCol) Then
                                If .Cells(i2, PVSW�n��Col) <> "" Or .Cells(i2, PVSW�I��Col) <> "" Then
                                    .Cells(i2, PVSW�ڑ�Col) = Format(��idA, "00")
                                End If
                            End If
                        Next i2
                        ��idA = ��idA + 1
                    End If
                End If
            End If
        Next i
        
        '�t�B�[���h����GY�͉B��
        For X = PVSW����Col To .Cells(PVSW����Row, .Columns.count).End(xlToLeft).Column
            If .Cells(1, X) = "PVSW" Or .Cells(1, X) = "RLTFA" Then
                If .Cells(PVSW����Row, X).Interior.color = 12566463 Then
                    .Columns(X).Hidden = True
                End If
            End If
        Next X
        
        '�R�����g�̐���
        For iii = PVSW����Row + 1 To PVSWlastRow
            If Not .Cells(iii, PVSW�n��HCol).Comment Is Nothing Then
                .Cells(iii, PVSW�n��HCol).Comment.Shape.Top = .Cells(iii - 1, PVSW�n��HCol).Top
                .Cells(iii, PVSW�n��HCol).Comment.Shape.Left = .Cells(iii - 1, PVSW�n��HCol).Left
            End If
            If Not .Cells(iii, PVSW�n�[��Col).Comment Is Nothing Then
                .Cells(iii, PVSW�n�[��Col).Comment.Shape.Top = .Cells(iii - 1, PVSW�n�[��Col + 1).Top
                .Cells(iii, PVSW�n�[��Col).Comment.Shape.Left = .Cells(iii - 1, PVSW�n�[��Col + 1).Left
            End If
            If Not .Cells(iii, PVSW�I��HCol).Comment Is Nothing Then
                .Cells(iii, PVSW�I��HCol).Comment.Shape.Top = .Cells(iii - 1, PVSW�I��HCol).Top
                .Cells(iii, PVSW�I��HCol).Comment.Shape.Left = .Cells(iii - 1, PVSW�I��HCol).Left
            End If
            If Not .Cells(iii, PVSW�I�[��Col).Comment Is Nothing Then
                .Cells(iii, PVSW�I�[��Col).Comment.Shape.Top = .Cells(iii - 1, PVSW�I�[��Col + 1).Top
                .Cells(iii, PVSW�I�[��Col).Comment.Shape.Left = .Cells(iii - 1, PVSW�I�[��Col + 1).Left
            End If
        Next iii

        '�󗓂Ȃ̂�isEmpty��false���Ԃ�Z����������SQL�ŊJ�����G���[�ɂȂ鎖�ւ̑΍�_�b��
        For iii = PVSW����Row + 1 To PVSWlastRow
            For X = PVSW���i�i��sCol To PVSW���i�i��eCol
                If IsEmpty(.Cells(iii, X)) = False Then
                    If .Cells(iii, X) = "" Then
                        .Cells(iii, X) = Empty
                    End If
                End If
            Next X
        Next iii
        
    End With
            
End Function
Function PVSWcsv�̋��ʉ�_Ver1944_�����ύX() '�����ύX�p

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF_temp"
    Dim i As Long, ii As Long

    With Workbooks(myBookName).Sheets("���i�i��")
        Dim ���i�i��() As String
        Set ���i�i��key = .Cells.Find("���C���i��", , , 1)
        Dim ���i�i��lastRow: ���i�i��lastRow = .Cells(.Rows.count, ���i�i��key.Column).End(xlUp).Row
        Dim ���i�i�ԍ��ڐ� As Long: ���i�i�ԍ��ڐ� = 8
        ReDim ���i�i��(���i�i�ԍ��ڐ�, ���i�i��lastRow - ���i�i��key.Row)
        Dim ���i�i��c As Long: ���i�i��c = 0
        For i = ���i�i��key.Row + 1 To ���i�i��lastRow
            ���i�i��c = ���i�i��c + 1
            For ii = 0 To ���i�i�ԍ��ڐ�
                ���i�i��(ii, ���i�i��c) = .Cells(i, ���i�i��key.Column + ii)
            Next ii
        Next i
    End With
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF_temp")
        'PVSW���Ƃ��Ƃ̃f�[�^
        Dim PVSW����Row As Long: PVSW����Row = .Cells.Find("�d�����ʖ�", , , 1).Row
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("�d�����ʖ�", , , 1).Column
        Dim PVSW���i�i��sCol As Long: PVSW���i�i��sCol = .Cells.Find("���i�i��s", , , 1).Column
        Dim PVSW���i�i��eCol As Long: PVSW���i�i��eCol = .Cells.Find("���i�i��e", , , 1).Column
        Dim PVSW���i�i��RAN As Range: Set PVSW���i�i��RAN = .Range(.Cells(PVSW����Row, PVSW���i�i��sCol), .Cells(PVSW����Row, PVSW���i�i��eCol))
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW����Col).End(xlUp).Row
        Dim PVSW�d��sCol As Long: PVSW�d��sCol = .Cells.Find("�d�������擾s", , , 1).Column
        Dim PVSW�d��eCol As Long: PVSW�d��eCol = .Cells.Find("�d�������擾e", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW�n����Col As Long: PVSW�n����Col = .Cells.Find("�n�_������_", , , 1).Column
        Dim PVSW�I����Col As Long: PVSW�I����Col = .Cells.Find("�I�_������_", , , 1).Column

        Dim PVSW�n��HCol As Long: PVSW�n��HCol = .Cells.Find("�n�_����H����", , , 1).Column
        Dim PVSW�n�[��Col As Long: PVSW�n�[��Col = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        Dim PVSW�nCavCol As Long: PVSW�nCavCol = .Cells.Find("�n�_���L���r�e�BNo.", , , 1).Column
        Dim PVSW�n���Col As Long: PVSW�n���Col = .Cells.Find("�n�_����햼��", , , 1).Column
        Dim PVSW�n���Ӑ�Col As Long: PVSW�n���Ӑ�Col = .Cells.Find("�n�_���[�����Ӑ�i��", , , 1).Column
        Dim PVSW�n���Col As Long: PVSW�n���Col = .Cells.Find("�n�_���[�����i��", , , 1).Column
        Dim PVSW�I��HCol As Long: PVSW�I��HCol = .Cells.Find("�I�_����H����", , , 1).Column
        Dim PVSW�I�[��Col As Long: PVSW�I�[��Col = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        Dim PVSW�ICavCol As Long: PVSW�ICavCol = .Cells.Find("�I�_���L���r�e�BNo.", , , 1).Column
        Dim PVSW�I���Col As Long: PVSW�I���Col = .Cells.Find("�I�_����햼��", , , 1).Column
        Dim PVSW�I���Ӑ�Col As Long: PVSW�I���Ӑ�Col = .Cells.Find("�I�_���[�����Ӑ�i��", , , 1).Column
        Dim PVSW�I���Col As Long: PVSW�I���Col = .Cells.Find("�I�_���[�����i��", , , 1).Column
        'RLTF����擾�����f�[�^
        Dim PVSW�\��Col As Long: PVSW�\��Col = .Cells.Find("�\��_", , , 1).Column
        Dim PVSW�i��Col As Long: PVSW�i��Col = .Cells.Find("�i��_", , , 1).Column
        Dim PVSW�T�C�YCol As Long: PVSW�T�C�YCol = .Cells.Find("�T�C�Y_", , , 1).Column
        Dim PVSW�T�C�Y�ď�Col As Long: PVSW�T�C�Y�ď�Col = .Cells.Find("�T��_", , , 1).Column
        Dim PVSW�FCol As Long: PVSW�FCol = .Cells.Find("�F_", , , 1).Column
        Dim PVSW�F��Col As Long: PVSW�F��Col = .Cells.Find("�F��_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        Dim PVSW����Col As Long: PVSW����Col = .Cells.Find("����_", , , 1).Column
        Dim PVSW�n�[Col As Long: PVSW�n�[Col = .Cells.Find("�n�_���[�q_", , , 1).Column
        Dim PVSW�n�}Col As Long: PVSW�n�}Col = .Cells.Find("�n�_���}_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_����_", , , 1).Column
        Dim PVSW�n��Col As Long: PVSW�n��Col = .Cells.Find("�n�_�����i_", , , 1).Column
        Dim PVSW�I�[Col As Long: PVSW�I�[Col = .Cells.Find("�I�_���[�q_", , , 1).Column
        Dim PVSW�I�}Col As Long: PVSW�I�}Col = .Cells.Find("�I�_���}_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_����_", , , 1).Column
        Dim PVSW�I��Col As Long: PVSW�I��Col = .Cells.Find("�I�_�����i_", , , 1).Column
        
        '��r����
        Dim PVSW��rCol(7) As Long
        'RLTF����̃f�[�^
        'PVSW��rCol(0) = PVSW�i��Col
        'PVSW��rCol(1) = PVSW�T�C�YCol
        'PVSW��rCol(2) = PVSW�T�C�Y�ď�Col
        'PVSW��rCol(3) = PVSW�FCol
        'PVSW��rCol(4) = PVSW�F��Col
        'PVSW��rCol(5) = PVSW����Col
        'PVSW��rCol(6) = PVSW����Col
        'PVSW��rCol(7) = PVSWJCDFCol
        PVSW��rCol(0) = PVSW����Col
        'PVSW��rCol(1) = PVSW�n�[Col
        'PVSW��rCol(10) = PVSW�n�}Col
        'PVSW��rCol(11) = PVSW�n��Col
        'PVSW��rCol(2) = PVSW�I�[Col
        'PVSW��rCol(13) = PVSW�I�}Col
        'PVSW��rCol(14) = PVSW�I��Col
        'PVSW����̃f�[�^
        PVSW��rCol(1) = PVSW�\��Col
        'PVSW��rCol(16) = PVSWRLTFtoPVSWCol
        PVSW��rCol(2) = PVSW�n��HCol
        PVSW��rCol(3) = PVSW�n�[��Col
        PVSW��rCol(4) = PVSW�nCavCol
        'PVSW��rCol(20) = PVSW�n���Col
        'PVSW��rCol(21) = PVSW�n���Ӑ�Col
        'PVSW��rCol(22) = PVSW�n���Col
        PVSW��rCol(5) = PVSW�I��HCol
        PVSW��rCol(6) = PVSW�I�[��Col
        PVSW��rCol(7) = PVSW�ICavCol
        'PVSW��rCol(26) = PVSW�I���Col
        'PVSW��rCol(27) = PVSW�I���Ӑ�Col
        'PVSW��rCol(28) = PVSW�I���Col
        
        Dim ��rA() As String, ��rB() As String
        For i = PVSW����Row + 1 To PVSWlastRow
            '�����Z�b�gA
            ReDim ��rA(���i�i��c)
            For X = LBound(PVSW��rCol) To UBound(PVSW��rCol)
                ��rA(0) = ��rA(0) & .Cells(i, PVSW��rCol(X)) & "_"
            Next X
            For ii = i + 1 To PVSWlastRow
                ReDim ��rB(���i�i��c)
                '�����Z�b�gB
                For X = LBound(PVSW��rCol) To UBound(PVSW��rCol)
                    ��rB(0) = ��rB(0) & .Cells(ii, PVSW��rCol(X)) & "_"
                Next X
                'A��B�̔�r
                If ��rA(0) = ��rB(0) Then
                    '���i�i�ԃZ�b�gB
                    .Cells(i, 1).Select
                    For c = 1 To ���i�i��c
                        If .Cells(i, c) = "" And .Cells(ii, c) <> "" Then
                            .Cells(i, c) = .Cells(ii, c)
                            .Cells(i, c).Interior.color = .Cells(ii, c).Interior.color
                        ElseIf .Cells(i, c) <> "" And .Cells(ii, c) <> "" Then
                            Stop '���肦��H�v�m�F
                        End If
                    Next c
                    .Cells(ii, 1).Select
                    .Rows(ii).Delete
                    ii = ii - 1: PVSWlastRow = PVSWlastRow - 1
                End If
            Next ii
        Next i
    End With
    
End Function

Public Function ���i�i��RAN_set()
    With ActiveWorkbook.Sheets("���i�i��")
        Dim ���C���i�� As Range: Set ���C���i�� = .Cells.Find("���C���i��", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ���C���i��.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(���C���i��.Row, .Columns.count).End(xlToLeft).Column
        
        ���i�i��RANc = lastRow - ���C���i��.Row
        ReDim ���i�i��RAN(lastCol - ���C���i��.Column + 1, ���i�i��RANc)
        For Y = 0 To ���i�i��RANc
            For X = 0 To lastCol - ���C���i��.Column + 1
                Set ���i�i��RAN(X, Y) = .Cells(Y + ���C���i��.Row, X + ���C���i��.Column - 1)
            Next X
        Next Y
    End With
End Function

Public Function ���i�i��RAN_set2(���i�i��RAN, Optional ���type, Optional �����, Optional ��n�����i�i��)
    
    Call �A�h���X�Z�b�g(myBook)
    
    With ThisWorkbook.Sheets("PVSW_RLTF")
        Dim PVSW_RLTF_fieldName As Range
        Set sikibetu = .Cells.Find("�d�����ʖ�", , , 1)
        Set PVSW_RLTF_fieldName = .Rows(sikibetu.Row)
    End With
    
    With ActiveWorkbook.Sheets("���i�i��")
        Dim ���C���i�� As Range: Set ���C���i�� = .Cells.Find("���C���i��", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ���C���i��.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(���C���i��.Row, .Columns.count).End(xlToLeft).Column
        Dim �����Col As Long: �����Col = .Rows(���C���i��.Row).Find(���type, , , 1).Column
        '��n���}�\�� = .Cells.Find("��n���}�\��", , , 1).Offset(1, 0).Value
        Dim flg As Range
        ReDim ���i�i��RAN(lastCol - ���C���i��.Column + 2, 0)
        ���i�i��RANc = 0
        Dim �o�^c As Long: �o�^c = 0

        For Y = ���C���i��.Row To lastRow
            If Y = ���C���i��.Row Then
                '�t�B�[���h����ǉ�
                For X = ���C���i��.Column - 1 To lastCol
                    Set ���i�i��RAN(X - ���C���i��.Column + 1, �o�^c) = .Cells(Y, X)
                Next X
                ���i�i��RAN(lastCol - ���C���i��.Column + 2, �o�^c) = "��ԍ�"
            Else
                '����ނ���������Ȃ���Ύ��̃��R�[�h�Ɉړ�
                If CStr(.Cells(Y, �����Col)) <> CStr(�����) And ����� <> "" Then GoTo nextY
                '���C���i�Ԃ�[PVSW_RLTF]�ɖ�����ΐ��i�i��RAN�ɒǉ����Ȃ�
                Set ���i�i��v = .Cells(Y, ���C���i��.Column)
                Set flg = PVSW_RLTF_fieldName.Find(���i�i��v, , , 1)
                If flg Is Nothing Then GoTo nextY
                
                '���i�i��RAN�ɒǉ�
                ReDim Preserve ���i�i��RAN(lastCol - ���C���i��.Column + 2, �o�^c)
                For X = ���C���i��.Column - 1 To lastCol
                    Set ���i�i��RAN(X - ���C���i��.Column + 1, �o�^c) = .Cells(Y, X)
                    '���̂��u�����N�Ȃ痪�̂�t����
                    If .Cells(���C���i��.Row, X) = "����" Then
                        If .Cells(Y, X) = "" Then
                            ���� = Replace(.Cells(Y, ���C���i��.Column), " ", "")
                            If Len(����) = 10 Then
                                ���� = Mid(����, 8)
                            Else
                                ���� = Mid(����, 5)
                            End If
                            .Cells(Y, X).NumberFormat = "@"
                            .Cells(Y, X) = ����
                        End If
                    End If
                Next X
                '���̐��i�i�Ԃ�[PVSW_RLTF]�ł̗�ԍ����Z�b�g
                ���i�i��RAN(lastCol - ���C���i��.Column + 2, �o�^c) = flg.Column
                ���i�i��RANc = ���i�i��RANc + 1
            End If
            �o�^c = �o�^c + 1
nextY:
        Next Y
    End With
'
'    With Sheets("Sheet2")
'        For i = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
'            For x = LBound(���i�i��RAN, 1) To UBound(���i�i��RAN, 1)
'                .Cells(i + 1, x + 1) = ���i�i��RAN(x, i)
'            Next x
'        Next i
'    End With
    
    Set ���C���i�� = Nothing
    'If ���i�i��RANc = 0 Then Stop '�Ώۂ̐��i�i�Ԃ�0
    
End Function


Public Function SQL_JUNK(mySQL0, mySheetName, sqlRow, sqlCol, �\��Col)
    '���̃v���V�[�W����2��ނ�SQL�����s���Ă��邩�番����Â炭�Ȃ��Ă�
    Dim sqlSheetName As String: sqlSheetName = "SQLtemp0"

    '�c�[���@���@�Q�Ɛݒ�@��
    ' Microsoft ActiveX Data Objects 2.8 Library
    '�`�F�b�N
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String

    xl_file = ThisWorkbook.FullName '���̃u�b�N���w�肵�Ă��ǂ�

'    Set cn = New ADODB.Connection
'    cn.Provider = "MSDASQL"
'    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
'    cn.Open
'    Set rs = New ADODB.Recordset

    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    rs.Open mySQL0, cn, adOpenStatic
    
    '���[�N�V�[�g�̒ǉ�
    For Each ws(0) In Worksheets
        If ws(0).Name = sqlSheetName Then
            Application.DisplayAlerts = False
            ws(0).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = sqlSheetName
    Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName).Cells.NumberFormat = "@"
    
    With Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName)
        '�t�B�[���hNAME�\��
        For i = 0 To rs.Fields.count - 1
            .Cells(1, i + 1).Value = rs(i).Name
        Next
        
        j = 2
        Do Until rs.EOF
          '1 ���R�[�h���̏���
            If rs(�\��Col).Value <> "" Or rs(�\��Col - 1).Value <> "" Then
                For i = 0 To rs.Fields.count - 1
                    .Cells(j, i + 1).Value = rs(i).Value
                Next
                j = j + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
        
        '�f�[�^��SQL�ɃZ�b�g�o����`�ɕύX
        ���i�i��Rc = .Cells.Find("�[�����i��", , , 1).Column - 1
        ReDim ���i�i��R(10, ���i�i��Rc)
        For X = 1 To ���i�i��Rc
            If 3 < Len(.Cells(2, X)) Then
                ���i�i��h = .Cells(2, X)
            End If
            .Cells(3, X) = ���i�i��h & .Cells(3, X)
            ���i�i��R(1, X) = .Cells(3, X)
        Next X
        Call ���i�i��RAN_seek
        
        .Range(.Rows(1), .Rows(2)).Delete
        �[�����i��Col = .Cells.Find("�[�����i��", , , 1).Column
        .Columns(�[�����i��Col).Insert
        .Cells(1, �[�����i��Col) = "Products"
        For Y = 2 To j - 1
            Products = ""
            For X = 1 To �[�����i��Col - 1
                If .Cells(Y, X) <> "" Then
                    Products = Products & "1"
                Else
                    Products = Products & "0"
                End If
            Next X
            .Cells(Y, �[�����i��Col) = Products
        Next Y
    End With
    
    'SQL1�ŊJ��
    '���[�N�V�[�g�̒ǉ�
    sqlSheetName = "SQLtemp1"
    sqlsheetname0 = ActiveSheet.Name
    For Each ws(0) In Worksheets
        If ws(0).Name = sqlSheetName Then
            Application.DisplayAlerts = False
            ws(0).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = sqlSheetName
    Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName).Cells.NumberFormat = "@"

    mySQL1 = " SELECT Products,�\��,�T�C�Y,�F�ď�,�[����,Cav,��,�},�}1,�� from [" & sqlsheetname0 & "$] where �} <> �}1"
    On Error Resume Next
        rs.Open mySQL1, cn, adOpenStatic
    
        myErrFlg = False
        If Err.Number = -2147467259 Then 'RS��OPEN�ŃG���[�o��B�Ȃ񂩂����悭�������A�����s���ăG���[�Œ�~�����čēx���s��������G���[�o�Ȃ�����G���[�o����ŏ�������s�����Ƃ��A�Ȃ񂩂��߂�
            myErrFlg = True
            
            Exit Function
        End If
    On Error GoTo 0
    
    Call DeleteDefinedNames
    With Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName)
        '�t�B�[���hNAME�\��
        For i = 0 To rs.Fields.count - 1
            .Cells(1, i + 1).Value = rs(i).Name
        Next
        j = 2
        Do Until rs.EOF
          '1 ���R�[�h���̏���
            For i = 0 To rs.Fields.count - 1
                If i > 3 And rs(9) = "�I" Then
                    .Cells(j, i + 1 + 6).Value = rs(i).Value
                Else
                    .Cells(j, i + 1 + 0).Value = rs(i).Value
                End If
            Next
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
        .Cells(1, 11) = "�[����_"
        .Cells(1, 12) = "CAV_"
        .Cells(1, 13) = "��_"
        .Cells(1, 14) = "�}_"
        .Cells(1, 15) = "�}1_"
        .Cells(1, 16) = "��_"
        
        '�n�_�ƏI�_���܂Ƃ߂�
        For i = 2 To j - 1
            For i2 = i + 1 To j - 1
                If i = i2 Then Stop '���肦���
                    If .Cells(i, 2) = .Cells(i2, 2) Then
                        If .Cells(i, 1) = .Cells(i2, 1) Then
                            If .Cells(i, 10) = "" Then
                                .Range(.Cells(i, 5), .Cells(i, 10)).Value = .Range(.Cells(i2, 5), .Cells(i2, 10)).Value
                                .Rows(i2).Delete
                                j = j - 1
                            ElseIf .Cells(i, 16) = "" Then
                                .Range(.Cells(i, 11), .Cells(i, 16)).Value = .Range(.Cells(i2, 11), .Cells(i2, 16)).Value
                                .Rows(i2).Delete
                                j = j - 1
                            Else
                                Stop '���肦���
                            End If
                        End If
                    End If
            Next i2
        Next i
    End With

    'cs.Close
    
End Function


Public Function SQL_�}���}�ύX�˗�(mysql)

    '�c�[���@���@�Q�Ɛݒ�@��
    ' Microsoft ActiveX Data Objects 2.8 Library
    '�`�F�b�N
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String

    xl_file = ThisWorkbook.FullName '���̃u�b�N���w�肵�Ă��ǂ�

'    Set cn = New ADODB.Connection
'    cn.Provider = "MSDASQL"
'    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
'    cn.Open
'    Set rs = New ADODB.Recordset


    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    rs.Open mysql, cn, adOpenStatic
        
    With Workbooks(ActiveWorkbook.Name).Sheets("��A��_�}���}")
        '�t�B�[���hNAME�\��
'        For i = 0 To rs.Fields.Count - 1
'            .Cells(1, i + 1).Value = rs(i).Name
'        Next
        Dim out�\��r As Long: out�\��r = .Cells.Find("�\��" & Chr(10) & "W-No.", , , xlWhole).Row
        out�N����r = .Cells.Find("�N����", , , 1).Row
        out�N����c = .Cells.Find("�N����", , , 1).Column
        out�^��r = .Cells.Find("�^��", , , 1).Row
        Dim out�\��c As Long: out�\��c = .Cells.Find("�\��" & Chr(10) & "W-No.", , , xlWhole).Column
        Dim out������c As Long: out������c = .Cells.Find("������_", , , 1).Column
        Dim out�T�C�Yc As Long: out�T�C�Yc = .Cells.Find("�T�C�Y" & Chr(10) & "Size", , , xlWhole).Column
        Dim out�Fc As Long: out�Fc = .Cells.Find("�F" & Chr(10) & "Color", , , xlWhole).Column
        Dim out�n�_��c As Long: out�n�_��c = .Cells.Find("�n�_��", , , 1).Column
        Dim out�n�_�[��c As Long: out�n�_�[��c = .Cells.Find("�[��" & Chr(10) & "Tno", , , xlWhole).Column
        Dim out�n�_��c As Long: out�n�_��c = .Cells.Find("��" & Chr(10) & "Cno", , , xlWhole).Column
        Dim out�n�_��c As Long: out�n�_��c = .Cells.Find("��H����" & Chr(10) & "Circuit", , , xlWhole).Column
        Dim out�n�_�}���}�Oc As Long: out�n�_�}���}�Oc = .Cells.Find("�}���}" & Chr(10) & "�ύX�O", , , xlWhole).Column
        Dim out�n�_����c As Long: out�n�_����c = .Cells.Find("����", , , xlWhole).Column
        Dim out�n�_�}���}��c As Long: out�n�_�}���}��c = .Cells.Find("�}���}" & Chr(10) & "�ύX��", , , xlWhole).Column
        Dim out�I�_��c As Long: out�I�_��c = .Cells.Find("�I�_��", , , 1).Column
        Dim out�I�_�[��c As Long: out�I�_�[��c = .Cells.Find("�[��" & Chr(10) & "Tno_", , , xlWhole).Column
        Dim out�I�_��c As Long: out�I�_��c = .Cells.Find("��" & Chr(10) & "Cno_", , , xlWhole).Column
        Dim out�I�_��c As Long: out�I�_��c = .Cells.Find("��H����" & Chr(10) & "Circuit_", , , xlWhole).Column
        Dim out�I�_�}���}�Oc As Long: out�I�_�}���}�Oc = .Cells.Find("�}���}" & Chr(10) & "�ύX�O_", , , xlWhole).Column
        Dim out�I�_����c As Long: out�I�_����c = .Cells.Find("����_", , , xlWhole).Column
        Dim out�I�_�}���}��c As Long: out�I�_�}���}��c = .Cells.Find("�}���}" & Chr(10) & "�ύX��_", , , xlWhole).Column
        Dim outKeyc As Long: outKeyc = .Cells.Find("key_", , , xlWhole).Column
        .Range(.Columns(out�N����c + 1), .Columns(.Columns.count)).ClearContents
        addRow = .Cells(.Rows.count, out�\��c).End(xlUp).Row + 1
        .Range(.Rows(out�\��r + 1), .Rows(addRow)).Delete
        addRow = .Cells(.Rows.count, out�\��c).End(xlUp).Row + 1
        Do Until rs.EOF
            For s = 1 To Len(rs(0))
                If Mid(rs(0), s, 1) = "1" Then
                    Set xx = .Rows(out�\��r).Find(�}���}���i�i��(s - 1, 0), , , 1)
                    If xx Is Nothing Then
                        xxx = .Cells(out�\��r, .Columns.count).End(xlToLeft).Column + 1
                        .Cells(out�\��r, xxx) = �}���}���i�i��(s - 1, 0)
                        .Cells(out�\��r, xxx).Orientation = -90
                        .Cells(out�\��r - 1, xxx) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), �}���}���i�i��(s - 1, 1))
                        .Cells(out�\��r - 1, xxx).ShrinkToFit = True
                        Set xx = .Rows(out�\��r).Find(�}���}���i�i��(s - 1, 0), , , 1)
                    End If
                    .Cells(addRow + j, xx.Column) = "1"
                    If .Cells(out�N����r, xx.Column).Value = "" Then
                        .Cells(out�N����r, xx.Column).NumberFormat = "mm/dd"
                        .Cells(out�N����r, xx.Column) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�N����"), �}���}���i�i��(s - 1, 1))
                        .Cells(out�N����r, xx.Column).ShrinkToFit = True
                        .Cells(out�^��r, xx.Column) = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�^��"), �}���}���i�i��(s - 1, 1))
                        .Cells(out�^��r, xx.Column).ShrinkToFit = True
                    End If
                End If
            Next s
            .Cells(addRow + j, out�\��c) = rs(1)
            .Cells(addRow + j, out�T�C�Yc) = rs(2)
            .Cells(addRow + j, out�Fc) = rs(3)
            .Cells(addRow + j, out�n�_�[��c) = rs(4)
            .Cells(addRow + j, out�n�_��c) = rs(5)
            .Cells(addRow + j, out�n�_��c) = rs(6)
            .Cells(addRow + j, out�n�_�}���}�Oc) = rs(7)
            .Cells(addRow + j, out�n�_�}���}��c) = rs(8)
            If rs(7) = "" Then
                ���� = "ADD"
            ElseIf rs(8) = "" Then
                ���� = "DEL"
            ElseIf rs(7) <> "" And rs(8) <> "" Then
                ���� = "CH"
            Else
                ���� = ""
            End If
            .Cells(addRow + j, out�n�_����c) = ����
            If ���� <> "" Then .Cells(addRow + j, out�n�_�}���}��c).Interior.color = RGB(255, 100, 100)
            .Cells(addRow + j, out�I�_�[��c) = rs(9)
            .Cells(addRow + j, out�I�_��c) = rs(10)
            .Cells(addRow + j, out�I�_��c) = rs(11)
            .Cells(addRow + j, out�I�_�}���}�Oc) = rs(12)
            .Cells(addRow + j, out�I�_�}���}��c) = rs(13)
            If rs(12) = "" Then
                ���� = "ADD"
            ElseIf rs(13) = "" Then
                ���� = "DEL"
            ElseIf rs(12) <> "" And rs(13) <> "" Then
                ���� = "CH"
            Else
                ���� = ""
            End If
            .Cells(addRow + j, out�I�_����c) = ����
            If ���� <> "" Then .Cells(addRow + j, out�I�_�}���}��c).Interior.color = RGB(255, 100, 100)
            .Cells(addRow + j, out������c) = Date
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
        
        addRow = .Cells(.Rows.count, out�\��c).End(xlUp).Row
        '�r��
        maxCol = .Cells(out�N����r, .Columns.count).End(xlToLeft).Column
        With .Range(.Cells(out�\��r, 1), .Cells(addRow, maxCol))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        .Range(.Cells(out�\��r - 1, out�n�_��c), .Cells(addRow, out�n�_��c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out�\��r - 1, out�I�_��c), .Cells(addRow, out�I�_��c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out�\��r - 1, out�N����c + 1), .Cells(addRow, out�N����c + 1)).Borders(1).Weight = xlMedium
        '�\�[�g
        With .Sort.SortFields
            .Clear
            .add key:=Cells(out�\��r, out������c), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(out�\��r, out�\��c), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '       .Add key:=Cells(out�\��r, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 4), Order:=xlAscending, DataOption:=0
    '            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Activate
        .Sort.SetRange .Range(.Rows(out�\��r), Rows(addRow))
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        .Activate
    End With
    
End Function

Public Function �z���}�쐬(Optional ���i�i��str = "", Optional �T�ustr, Optional ���}�̂�, Optional ���type, Optional ��n���摜Sheet)

    'Application.OnKey "%{ENTER}", "�I�[�g�V�F�C�v�폜"
    'Application.OnKey�ŌĂяo�������̏���
    Dim key As Range
    If IsError(�T�ustr) Then  '���W�o�^�p
        Dim U�i���o�[�\�����[�h As Boolean
        If ���i�i��str = "U�i���o�[" Then U�i���o�[�\�����[�h = True
        PlaySound "��������2"
        ���i�i��str = ""
        �T�ustr = ""
        ���}�̂� = "1"
        ���type = Mid(ActiveSheet.Name, 4)
        ��n���摜Sheet = ""
        Call ���i�i��RAN_set2(���i�i��RAN, "����", ���type, "")
        '���W���͎x��
        With ActiveSheet
            lastRow = .UsedRange.Rows.count
            Set myKey = .Cells.Find("Size_", , , 1)
            For i = myKey.Row + 1 To lastRow
                If .Cells(i, myKey.Column) = "" Then
                    myLastCol = .Cells(i, Columns.count).End(xlToLeft).Column
                    If myLastCol Mod 2 = 1 Then GoTo line05
                    For X = 1 To myLastCol
                        If .Cells(i, X) <> "" Then GoTo line05
                        .Cells(i, X) = .Cells(i - 1, X)
                    Next X
                End If
line05:
            Next i
            .Columns.AutoFit
        End With
    End If
    
'    If IsError(���i�i��str) Then
'        ���i�i��str = "8501K006"
'        �T�ustr = "2"
'        ���}�̂� = "0"
'        ���type = "F"
'        ��n���摜Sheet = "�n���}_���C���i��8501K006"
'        Call ���i�i��RAN_set2(���i�i��RAN, "����", "F", "8501K006")
'    End If
    
    Call �œK��
    '���i�i��str = ""
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
        
    For Each ws(0) In wb.Sheets
        If ws(0).Name = "���_" Then
            Stop
        End If
    Next ws
    
    If IsError(���type) Or ���type = "" Then
        ���type = Mid(ActiveSheet.Name, 4)
    End If
    
    On Error Resume Next
    wb.Sheets("���_" & ���type).Activate
    If Err = 9 Then
        Call �œK�����ǂ�
        End
    End If
    On Error GoTo 0
    
    With wb.Sheets("���_" & ���type)
        Set key = .Cells.Find("Size_", , , 1)
        '���̃T�C�Y
        �T�C�Y = .Cells(key.Row, key.Column).Offset(, 1)
        If InStr(�T�C�Y, "_") = 0 Then
            �z���}�쐬temp = 1
            Exit Function
        Else
            �z���}�쐬temp = 0
            �T�C�Ys = Split(�T�C�Y, "_")
            �T�C�Yx = �T�C�Ys(0)
            �T�C�Yy = �T�C�Ys(1)
        End If
        �{�� = 1220 / �T�C�Yx '�T�C�Yx / 1220
        �{��y = 480 / �T�C�Yy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "�l�r �S�V�b�N"
        '�I�[�g�V�F�C�v���폜
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            objShp.Delete
        Next objShp
        
        Dim ���Oc As Long
        '���}�̍쐬
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim �����͈� As Range: Set �����͈� = .Range(.Rows(key.Row + 1), .Rows(lastRow))
        For Y = 2 To lastRow
            '�[�����^�C�g��
            �[�� = .Cells(Y, X)
            ���Ws = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(���Ws) < 1 Then ���WErr = 1 Else ���WErr = 0
            
            If ���WErr = 0 Then
                ���Wx = ���Ws(0) * �{��
                ���Wy = ���Ws(1) * �{��y
                
                ���Od = 0
                On Error Resume Next
                ���Od = wb.ActiveSheet.Shapes.Range(�[��).count
                If Err = 1004 Then ���Od = 0
                On Error GoTo 0
                
                If ���Od = 0 Then
                    Select Case Left(�[��, 1)
                    Case "U"
                        With wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            If U�i���o�[�\�����[�h = True Then
                                .TextFrame2.TextRange.Characters.Text = Mid(�[��, 2)
                                .TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent = 0
                                .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                                .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
                                .TextFrame2.TextRange.Characters.ParagraphFormat.Alignment = msoAlignLeft
                                .TextFrame2.MarginLeft = 0
                                .TextFrame2.WordWrap = msoFalse
                                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
                            End If
                            .Name = �[��
                            .Left = ���Wx - 4
                            .Top = ���Wy - 4
                            If ���}�̂� = "1" Then
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
'                                .TextFrame.Characters.Font.Size = 4
'                                .TextFrame.Characters.Text = Replace(�[��, "U", "")
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    Case Else
                        With wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = �[��
                            .OnAction = "���}_�[���o�H�\��"
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = �[��
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.ForeColor.RGB = RGB(250, 250, 250)
                            If ���}�̂� = "1" Then
                                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                            Else
                                .TextFrame.Characters.Font.color = RGB(200, 200, 200)
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                            
                            .Left = ���Wx - 15
                            .Top = ���Wy - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If ���Wxbak <> "" Then
                    
                    On Error Resume Next
                    ���Oc1 = wb.Sheets("���_" & ���type).Shapes.Range(�[��bak & " to " & �[��).count
                    If Err = 1004 Then ���Oc1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    ���Oc2 = wb.Sheets("���_" & ���type).Shapes.Range(�[�� & " to " & �[��bak).count
                    If Err = 1004 Then ���Oc2 = 0
                    On Error GoTo 0
                        
                    If ���Oc1 = 0 And ���Oc2 = 0 And �[�� <> �[��bak Then
                        With wb.Sheets("���_" & ���type).Shapes.AddLine(���Wxbak, ���Wybak, ���Wx, ���Wy)
                            .Name = �[��bak & " to " & �[��
                            .Line.Weight = 3.2
                            If ���}�̂� = "1" Then
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    End If
                End If
                ���Wxbak = ���Wx
                ���Wybak = ���Wy
                �[��bak = �[��
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                ���Wsbak = Split(.Cells(Y, 2), "_")
                ���Wxbak = ���Wsbak(0) * �{��
                ���Wybak = ���Wsbak(1) * �{��y
                �[��bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        wb.Sheets("���_" & ���type).Shapes.SelectAll
        If wb.Sheets("���_" & ���type).Shapes.count = 0 Then GoTo line30
        
        Selection.Left = 5
        Selection.Top = 10
    
    If ���}�̂� = "1" Then GoTo line99
        �摜add = �T�C�Yy * �{��y
        '���z������[���̐F�t��
        Call SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            If �z���[��RAN(0, i) = "" Then GoTo nextI
            Set �z�� = �����͈�.Cells.Find(�z���[��RAN(0, i), , , 1)
            If �z�� Is Nothing Then GoTo nextI
            ��F = �z���[��RAN(1, i)
            If ��F = "" Then
                With wb.Sheets("���_" & ���type).Shapes(�z��.Value)
                    .Select
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = RGB(255, 100, 100)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    .Copy
                    DoEvents
                    Sleep 5
                    DoEvents
                    ActiveSheet.Paste
                    Selection.Name = �z��.Value & "!"
                    Selection.Left = myLeft
                    Selection.Top = �摜add
                    �摜add = �摜add + Selection.Height
                End With
                '��n���}�̎擾�Ɣz�z
                With wb.Sheets(��n���摜Sheet)
                    .Activate
                    n = 0
                    For Each obj In .Shapes(�z��.Value & "_1").GroupItems
                        If obj.Name Like �z��.Value & "_1*" Then
                            If obj.Name <> �z��.Value & "_1_t" Then
                                If obj.Name <> �z��.Value & "_1_b" Then
                                    If n = 0 Then
                                        obj.Select True
                                    Else
                                        obj.Select False
                                    End If
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next obj
                    Selection.Copy
                    .Cells(1, 1).Select
                End With
                
                .Activate
                ActiveSheet.Pictures.Paste.Select
                'Sheets(��n���摜Sheet).Shapes(�z��.Value & "_1").Copy
                'Selection.Top = (�T�C�Yy * �{��y) + �摜add + myHeight
                Selection.Left = myLeft
                �{��a = (myWidth / Selection.Width) * 3
                If �{��a > 0.7 Then �{��a = 0.7
                Selection.ShapeRange.ScaleHeight �{��a, msoFalse, msoScaleFromTopLeft
                Selection.Top = �摜add
                ActiveSheet.Shapes(�z��.Value & "!").Select False
                Selection.Group.Select
                Selection.Name = �z��.Value & "!"
                �摜add = �摜add + Selection.Height
            End If
            Set �z��bak = �z��
            ��Fbak = ��F
nextI:
        Next i
            
        '���z������[���Ԃ̃��C���ɐF�t��
        Dim myStep As Long
        
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            For i2 = i + 1 To UBound(�z���[��RAN, 2)
                Set �[��from = �����͈�.Cells.Find(�z���[��RAN(0, i), , , 1)
                Set �[��to = �����͈�.Cells.Find(�z���[��RAN(0, i2), , , 1)
                If �[��from Is Nothing Or �[��to Is Nothing Then GoTo line31
                If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                    
                Set �[��1 = �[��from
                �㉺�ɐi��flg = 0
                For Y = �[��from.Row To �[��to.Row Step myStep
                    'from���獶�ɐi��
                    If �[��1.Row = �[��from.Row Or �㉺�ɐi��flg = 0 Then
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                            If �[��1 = �[��2.Offset(myStep, 0) Then
                                �㉺�ɐi��flg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'to�̍s�܂ŏ�܂��͉��ɐi��
                    If (�[��1.Column = 1 Or �㉺�ɐi��flg = 1) And �[��1.Row <> �[��to.Row Then
line15:
                        Set �[��2 = �[��1.Offset(myStep, 0)
                        If �[��1 <> �[��2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                        If �[��1 <> �[��2.Offset(myStep, 0) Then
                            �㉺�ɐi��flg = 0
                        End If
                        'If �㉺�ɐi��flg = 1 Then GoTo line15
                    End If
                    
                    'to�̍s���E�ɐi��
                    If �[��1.Row = �[��to.Row Then
                        Do Until �[��1.Column = �[��to.Column
                            Set �[��2 = �[��1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set �[��2 = Nothing
            Next i2
line31:
        Next i

        '���z�������n���d����\��
        Call SQL_�z����n���擾(�z����n��RAN, ���i�i��str, �T�ustr)
        Dim �Fv As String, �Tv As String, �[��v As String, �}v As String, �n��v As String
        For i = LBound(�z����n��RAN, 2) To UBound(�z����n��RAN, 2)
            �Fv = �z����n��RAN(0, i)
            If �Fv = "" Then Exit For
            �Tv = �z����n��RAN(1, i)
            �[��v = �z����n��RAN(2, i)
            �}v = �z����n��RAN(3, i)
            �n��v = �z����n��RAN(4, i)
            
            ���Oc = 0
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name = �[��v & "_" Then
                    ���Oc = ���Oc + 1
                End If
            Next objShp
                
            With ActiveSheet.Shapes(�[��v)
                .Select
                .Line.ForeColor.RGB = RGB(255, 100, 100)
                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                .ZOrder msoBringToFront
                ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (���Oc * 15), Selection.Top, 15, 15).Select
                Call �F�ϊ�(�Fv, clocode1, clocode2, clofont)
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(�Tv, "F", ""), 3)
                Selection.ShapeRange.Adjustments.Item(1) = 0.15
                'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Name = �[��v & "_"
                If InStr(�Fv, "/") > 0 Then
                    �x�[�X�F = Left(�Fv, InStr(�Fv, "/") - 1)
                Else
                    �x�[�X�F = �Fv
                End If
                
                myFontColor = clofont '�t�H���g�F���x�[�X�F�Ō��߂�
                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 6
                Selection.Font.Name = myFont
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                Selection.ShapeRange.TextFrame2.MarginLeft = 0
                Selection.ShapeRange.TextFrame2.MarginRight = 0
                Selection.ShapeRange.TextFrame2.MarginTop = 0
                Selection.ShapeRange.TextFrame2.MarginBottom = 0
                '�X�g���C�v�͌��ʂ��g��
                If clocode1 <> clocode2 Then
                    With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                        .color = clocode1
                        .color.TintAndShade = 0
                        .color.Brightness = 0
                        .Transparency = 0#
                        .Radius = 8
                    End With
                End If
                '�}���}
                If �}v <> "" Then
                    myLeft = Selection.Left
                    myTop = Selection.Top
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    For Each objShp In Selection.ShapeRange
                        Set objShpTemp = objShp
                    Next objShp
                    ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                    Call �F�ϊ�(�}v, clocode1, clocode2, clofont)
                    myFontColor = clofont
                    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                    objShpTemp.Select False
                    Selection.Group.Select
                    Selection.Name = �[��v & "_"
                End If
            End With
        Next i

        '��n���d���������ɕ\��
'        With Sheets("���_" & ���type)
'            For Each objShp In ActiveSheet.Shapes
'                If objShp.Line.ForeColor.RGB = RGB(255, 100, 100) Then
'                    If objShp.Type = 1 Then '���C�����}�b�`���鎖�̉��
'                        If Right(objShp.Name, 1) <> "!" Then
'                            ��n���[�� = objShp.Name
'                            myLeft = objShp.Left
'                            ActiveSheet.Shapes(��n���[��).Select True
'                            For Each objShp2 In ActiveSheet.Shapes
'                                If ��n���[�� & "_" = objShp2.Name Then
'                                    objShp2.Select False
'                                End If
'                            Next objShp2
'                            Selection.Copy
'                            Sleep 5
'                            .Paste
'                            Selection.Group.Select
'                            Selection.Name = ��n���[�� & "!"
'                            Selection.Left = myLeft
'                            Selection.Top = (�T�C�Yy * �{��y) + �摜add
'                            �摜add = �摜add + Selection.Height
'                        End If
'                    End If
'                End If
'            Next objShp
'        End With
        
        
'        '�[�����őO�ʂɈړ�
'        For Each objShp In Wb.Sheets("Sheet1").Shapes
'            If objShp.Type = 1 Then
'              objShp.ZOrder msoBringToFront
'            End If
'        Next objShp
'
'        '��n���d�����őO�ʂɈړ�
'        For Each objShp In Wb.Sheets("Sheet1").Shapes
'            If InStr(objShp.Name, "_") > 0 Then
'              objShp.ZOrder msoBringToFront
'            End If
'        Next objShp
'
'        '�D�F�̒[�����Ŕw�ʂɈړ�
'        For Each objShp In Wb.Sheets("Sheet1").Shapes
'            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
'              objShp.ZOrder msoSendToBack
'            End If
'        Next objShp
                
        '�[�����őO�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '��n���d�����őO�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '�D�F�̒[�����Ŕw�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
              objShp.ZOrder msoSendToBack
            End If
        Next objShp
line99:
        
        '�D�F�̃��C�����Ŕw�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = RGB(150, 150, 150) Or objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
               
        Dim SyTop As Long
        Dim flg As Long, �摜flg As Long, Sx As Long, Sy As Long
        '�}����̋󂢂Ă���X�y�[�X�Ɉړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            �摜flg = 0: SyTop = (�T�C�Yy * �{��y) + 5
line20:
            flg = 0
            For Each objShp2 In wb.Sheets("���_" & ���type).Shapes
                'If objShp.Name = "501!" And objShp2.Name = "843!" Then Stop
                If Right(objShp.Name, 1) = "!" And Right(objShp2.Name, 1) = "!" Then
                    If objShp.Name <> objShp2.Name Then
                        �摜flg = 1
                        For Sx = objShp.Left To objShp.Left + objShp.Width Step 1
                            If objShp2.Left <= Sx And objShp2.Left + objShp2.Width >= Sx Then
                                If objShp2.Top <= SyTop And objShp2.Top + objShp2.Height >= SyTop Then
                                    flg = 1
                                    SyTop = SyTop + 10
                                    GoTo line20
                                End If
                            End If
                        Next Sx
                    End If
                End If
            Next objShp2
            
            If flg = 1 Then GoTo line20
            
            If �摜flg = 1 Then
                objShp.Top = SyTop
            End If
        Next objShp
                
        wb.Sheets("���_" & ���type).Shapes.SelectAll
        If wb.Sheets("���_" & ���type).Shapes.count > 1 Then Selection.Group.Select
        Selection.Name = "���"
        Selection.Top = 10
        Selection.Left = 10
       
line30:
        wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, �T�C�Yx * �{��, �T�C�Yy * �{��y).Select
        Selection.Name = "��"
        wb.Sheets("���_" & ���type).Shapes("��").Adjustments.Item(1) = 0.02
        wb.Sheets("���_" & ���type).Shapes("��").ZOrder msoSendToBack
        'WB.Sheets("���_" & ���type).Shapes("��").Fill.PresetTextured 23
        wb.Sheets("���_" & ���type).Shapes("��").Fill.Patterned msoPatternDashedHorizontal
        wb.Sheets("���_" & ���type).Shapes("��").Fill.ForeColor.RGB = RGB(120, 120, 120)
        wb.Sheets("���_" & ���type).Shapes("��").Fill.BackColor.RGB = RGB(0, 0, 0)
        wb.Sheets("���_" & ���type).Shapes("��").Fill.Transparency = 0.8
        '�؂�ڂ̕\��
        Set kk = wb.Sheets("���_" & ���type).Cells.Find("k_", , , 1)
        If kk Is Nothing Then
            key.Offset(0, 2).Value = "k_"
            key.Offset(0, 3).Value = 42.2
        End If
        Dim k As String
        k = wb.Sheets("���_" & ���type).Cells.Find("k_", , , 1).Offset(0, 1)
        
        If IsNumeric(k) Then
            With wb.Sheets("���_" & ���type).Shapes.AddLine(k * �{��, 0, k * �{��, �T�C�Yy * �{��y)
                .Line.Weight = 1
                .Name = "k"
                .Line.ForeColor.RGB = RGB(150, 150, 150)
                .ZOrder msoSendToBack
                .Select False
            End With
        End If
        If wb.Sheets("���_" & ���type).Shapes.count > 2 Then
            wb.Sheets("���_" & ���type).Shapes("���").Select False
            Selection.Group.Select
            Selection.Name = "�z��"
        End If
        
        '.Cells(1, 1).Select
    End With
    If ���}�̂� = "1" Then
        �����[�� = SQL_�z���}_�[���ꗗ(wb.Name, ���type)
        If �����[��(0) <> Empty Then
            Dim myMsg As String: myMsg = "���̒[�����s�����Ă��܂��B�c��=" & UBound(�����[��) & vbCrLf
            For u = LBound(�����[��) To UBound(�����[��)
                myMsg = myMsg & vbCrLf & �����[��(u)
            Next u
        End If
        
        With wb.Sheets("���_" & ���type).Shapes("��")
            If myMsg = "" Then
                myMsg = "�s���[���͂���܂���"
            End If
            '�Ώۂ̐��i�i��
            myMsg = myMsg & vbCrLf & vbCrLf & "�Ώۂ̐��i�i��"
            For r = 1 To ���i�i��RANc
                myMsg = myMsg & vbCrLf & ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            Next r
            .TextFrame.Characters.Text = myMsg
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 255
        End With
        PlaySound "��������2"
    End If

    Call �œK�����ǂ�

End Function
Public Function �z���}�쐬3(Optional ���i�i��str, Optional ��zstr, Optional �T�ustr, Optional ���}�̂�, Optional ���type, Optional ��n���摜Sheet)

    temp = False
    'Application.OnKey "%{ENTER}", "�I�[�g�V�F�C�v�폜"
    'Application.OnKey�ŌĂяo�������̏���
    If IsError(���i�i��str) Then
        PlaySound "��������2"
        ���i�i��str = "8211158560"
        �T�ustr = "Base"
        ���}�̂� = "0"
        ���type = Mid(ActiveSheet.Name, 4)
        ��n���摜Sheet = ""
        Call ���i�i��RAN_set2(���i�i��RAN, "����", ���type, "")
    End If
    
    'Dim rootColor As Long: rootColor = RGB(50, 250, 50)
    Dim rootColor As Long: rootColor = RGB(0, 255, 102)
    Dim elseColor As Long: elseColor = RGB(160, 160, 160)
       
    Call �œK��
                    
    For Each WS_ In wb(0).Sheets
        If WS_.Name = "���_" Then
            Stop
        End If
    Next WS_
    
    If IsError(���type) Or ���type = "" Then
        ���type = Mid(ActiveSheet.Name, 4)
    End If
    
    With wb(0).Sheets("���_" & ���type)
        Dim key As Range
        'k_���[�����Əd������ꍇ�̏���_���߂�
        Set key = .Cells.Find("k_", , , 1).Offset(0, 1)
        If InStr(key, ".") = 0 Then key.Value = key.Value & ".1"
        
        Set key = .Cells.Find("Size_", , , 1)
        '���̃T�C�Y
        �T�C�Y = .Cells(key.Row, key.Column).Offset(, 1)
        �T�C�Ys = Split(�T�C�Y, "_")
        �T�C�Yx = �T�C�Ys(0)
        �T�C�Yy = �T�C�Ys(1)
                
        �{�� = 1220 / �T�C�Yx '�T�C�Yx / 1220
        �{��y = 480 / �T�C�Yy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "�l�r �S�V�b�N"
        '�I�[�g�V�F�C�v���폜
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb(0).Sheets("���_" & ���type).Shapes
            objShp.Delete
        Next objShp
        
        Dim ���Oc As Long
        '���}�̍쐬
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For Y = 2 To lastRow
            '�[�����^�C�g��
            �[�� = .Cells(Y, X)
            ���Ws = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(���Ws) < 1 Then ���WErr = 1 Else ���WErr = 0
            
            If ���WErr = 0 Then
                ���Wx = ���Ws(0) * �{��
                ���Wy = ���Ws(1) * �{��y
                
                ���Od = 0
                On Error Resume Next
                ���Od = wb(0).ActiveSheet.Shapes.Range(�[��).count
                If Err = 1004 Then ���Od = 0
                On Error GoTo 0
                
                If ���Od = 0 Then
                    Select Case Left(�[��, 1)
                    Case "U"
                        With wb(0).Sheets("���_" & ���type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            .Name = �[��
                            .Left = ���Wx - 4
                            .Top = ���Wy - 4
                            .Line.ForeColor.RGB = RGB(0, 10, 21)
                            .Fill.ForeColor.RGB = elseColor
                        End With
                    Case Else
                        With wb(0).Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = �[��
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = �[��
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 10, 21) '�[���̐F
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 10, 21)
                            .Fill.ForeColor.RGB = elseColor
                            .Left = ���Wx - 15
                            .Top = ���Wy - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If ���Wxbak <> "" Then
                    
                    On Error Resume Next
                    ���Oc1 = wb(0).Sheets("���_" & ���type).Shapes.Range(�[��bak & " to " & �[��).count
                    If Err = 1004 Then ���Oc1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    ���Oc2 = wb(0).Sheets("���_" & ���type).Shapes.Range(�[�� & " to " & �[��bak).count
                    If Err = 1004 Then ���Oc2 = 0
                    On Error GoTo 0
                        
                    If ���Oc1 = 0 And ���Oc2 = 0 And �[�� <> �[��bak Then
                        With wb(0).Sheets("���_" & ���type).Shapes.AddLine(���Wxbak, ���Wybak, ���Wx, ���Wy)
                            .Name = �[��bak & " to " & �[��
                            .Line.Weight = 3.2
                            .Line.ForeColor.RGB = elseColor '�[���ԃ��C��
                        End With
                    End If
                End If
                ���Wxbak = ���Wx
                ���Wybak = ���Wy
                �[��bak = �[��
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                ���Wsbak = Split(.Cells(Y, 2), "_")
                ���Wxbak = ���Wsbak(0) * �{��
                ���Wybak = ���Wsbak(1) * �{��y
                �[��bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        wb(0).Sheets("���_" & ���type).Activate
        wb(0).Sheets("���_" & ���type).Shapes.SelectAll
        Selection.Group.Name = "temp00"
        wb(0).Sheets("���_" & ���type).Shapes("temp00").Select
        Selection.Left = 5
        Selection.Top = 5
        Selection.Ungroup
        If wb(0).Sheets("���_" & ���type).Shapes.count = 0 Then GoTo line30
        wb(0).Sheets("���_" & ���type).Activate
        'Selection.Left = 5
        'Selection.Top = 5
        If �T�ustr = "Base" Then GoTo line99
        �[��count = 0
        '���z������[���̐F�t��
        Call SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            If �z���[��RAN(0, i) = "" Then GoTo nextI
            Set �z�� = .Cells.Find(�z���[��RAN(0, i), , , 1)
            If �z�� Is Nothing Then GoTo nextI
            ��F = �z���[��RAN(1, i)
            If ��F = "" Then
                �z��str = CStr(�z��.Value)
                ActiveSheet.Shapes(�z��str).Select
                With Selection.ShapeRange
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = rootColor
                    .Line.ForeColor.RGB = RGB(0, 10, 21)
                    .TextFrame.Characters.Font.color = RGB(0, 10, 21)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    Sleep 5
                End With
                                
                If Not (temp) Then
                    '��n���}�̎擾�Ɣz�z
                    With wb(0).Sheets(��n���摜Sheet)
                        .Activate
                        n = 0
                        For Each obj In .Shapes(�z��.Value & "_1").GroupItems
                            If obj.Name Like �z��.Value & "_1*" Then
                                If obj.Name <> �z��.Value & "_1_t" Then
                                    If obj.Name <> �z��.Value & "_1_b" Then
                                        If n = 0 Then
                                            obj.Select True
                                        Else
                                            obj.Select False
                                        End If
                                        n = n + 1
                                    End If
                                End If
                            End If
                        Next obj
                        Selection.Copy
                        .Cells(1, 1).Select
                    End With
                    
                    .Activate
                    DoEvents
                    Sleep 5
                    DoEvents
                    ActiveSheet.Pictures.Paste.Select
                    'Sheets(��n���摜Sheet).Shapes(�z��.Value & "_1").Copy
                    'Selection.Top = (�T�C�Yy * �{��y) + �摜add + myHeight
                    Selection.Left = myLeft
                    �{��a = (myWidth / Selection.Width) * 3
                    If �{��a > 0.7 Then �{��a = 0.7
                    Selection.ShapeRange.ScaleHeight �{��a, msoFalse, msoScaleFromTopLeft
                    Selection.ShapeRange.Glow.color.RGB = 16777215
                    Selection.ShapeRange.Glow.Radius = 5
                    Selection.ShapeRange.Glow.Transparency = 0.2
                    Selection.Top = myTop + myHeight
                    '���̌�n���}�Əd�Ȃ�Ȃ����m�F
                    top�ړ�flg = False
line12:
                    myleft2 = Selection.Left
                    mytop2 = Selection.Top
                    mytop3 = mytop2
                    myright2 = Selection.Width + myleft2
                    mybottom2 = Selection.Height + mytop2
                    If myright2 > 1220 Then
                        Selection.Left = myleft2 - (myright2 - 1220)
                        myleft2 = Selection.Left
                    End If
                    usednodesp = Split(usednode, ",")
                    For ii = LBound(usednodesp) + 1 To UBound(usednodesp)
        
                        node = Split(usednodesp(ii), "_")
                        �d�Ȃ�flg = False
                        nodeleft = CLng(node(0))
                        nodetop = CLng(node(1))
                        noderight = CLng(node(2))
                        nodebottom = CLng(node(3))
                        
                        For xx = nodeleft To noderight Step 2
                            For yy = nodetop To nodebottom Step 2
'                                If xx > myleft2 And yy > mytop2 Then Stop
                                �d�Ȃ�flg = xx >= myleft2 And xx <= myright2 And yy >= mytop2 And yy <= mybottom2
                                If �d�Ȃ�flg = True Then Exit For
                            Next yy
                            If �d�Ȃ�flg = True Then Exit For
                        Next xx
                        If �d�Ȃ�flg = True And nodebottom + 2 <> mytop3 Then
                            mybottom2 = mybottom2 + nodebottom + 2 - mytop3
                            mytop3 = nodebottom + 2
                            Selection.Top = mytop3
                            top�ړ�flg = True
                            GoTo line12
                        End If
                    Next ii
                    usednode = usednode & "," & Selection.Left & "_" & Selection.Top & "_" & Selection.Width + Selection.Left & "_" & Selection.Height + Selection.Top
                    If top�ړ�flg = True Then
                        With ActiveSheet.Shapes.AddLine(myLeft + myWidth / 2, myTop + myHeight, myleft2 + ((myright2 - myleft2) / 2), Selection.Top)
                            .Line.ForeColor.RGB = RGB(255, 0, 0)
                            .Line.Weight = 3
                            .Line.Transparency = 0.4
                            .Select False
                            Selection.Group.Select
                        End With
                    End If
                    Selection.Name = �z��.Value & "!"
                    �[��count = �[��count + 1
                End If
            End If
            Set �z��bak = �z��
            ��Fbak = ��F
nextI:
        Next i
           
        '���z�����郉�C���ɐF�t��
        Dim myStep As Long
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            For i2 = i + 1 To UBound(�z���[��RAN, 2)
                Set �[��from = .Cells.Find(�z���[��RAN(0, i), , , 1)
                Set �[��to = .Cells.Find(�z���[��RAN(0, i2), , , 1)
                If �[��from Is Nothing Or �[��to Is Nothing Then GoTo line31
                If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                wb(0).Activate
                On Error Resume Next
                ActiveSheet.Shapes(�[��from).ZOrder msoBringToFront
                ActiveSheet.Shapes(�[��to).ZOrder msoBringToFront
                On Error GoTo 0
                Set �[��1 = �[��from
                �㉺�ɐi��flg = 0
                For Y = �[��from.Row To �[��to.Row Step myStep
                    'from���獶�ɐi��
                    If �[��1.Row = �[��from.Row Or �㉺�ɐi��flg = 0 Then
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.ShapeStyle = msoLineStylePreset17
                            Selection.ShapeRange.Line.ForeColor.RGB = rootColor
                            Selection.ShapeRange.Line.DashStyle = 11 '�_��
                            Selection.ShapeRange.Line.Weight = 3
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = rootColor
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 10, 21)
                            End If
                            If �[��1 = �[��2.Offset(myStep, 0) Then
                                �㉺�ɐi��flg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'to�̍s�܂ŏ�܂��͉��ɐi��
                    If (�[��1.Column = 1 Or �㉺�ɐi��flg = 1) And �[��1.Row <> �[��to.Row Then
line15:
                        Set �[��2 = �[��1.Offset(myStep, 0)
                        If �[��1 <> �[��2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = rootColor
                            Selection.ShapeRange.Line.Weight = 3
                            Selection.ShapeRange.Line.DashStyle = 11 '�_��
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = rootColor
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 10, 21)
                        End If
                        If �[��1 <> �[��2.Offset(myStep, 0) Then
                            �㉺�ɐi��flg = 0
                        End If
                        'If �㉺�ɐi��flg = 1 Then GoTo line15
                    End If
                    
                    'to�̍s���E�ɐi��
                    If �[��1.Row = �[��to.Row Then
                        Do Until �[��1.Column = �[��to.Column
                            Set �[��2 = �[��1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = rootColor
                            Selection.ShapeRange.Line.Weight = 3
                            Selection.ShapeRange.Line.DashStyle = 11 '�_��
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = rootColor
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 10, 21)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set �[��2 = Nothing
            Next i2
line31:
        Next i

        '���z�������n���d����\��
        Call SQL_�z����n���擾(�z����n��RAN, ���i�i��str, �T�ustr)
        Dim �Fv As String, �Tv As String, �[��v As String, �}v As String, �n��v As String
        For i = LBound(�z����n��RAN, 2) To UBound(�z����n��RAN, 2)
            �Fv = �z����n��RAN(0, i)
            If �Fv = "" Then Exit For
            �Tv = �z����n��RAN(1, i)
            �[��v = �z����n��RAN(2, i)
            If IsNull(�z����n��RAN(3, i)) Then �z����n��RAN(3, i) = ""
            �}v = �z����n��RAN(3, i)
            �n��v = �z����n��RAN(4, i)
            ��v = �z����n��RAN(5, i)
            If ��v <> "" Then
                If ��v = "#" Or ��v = "*" Or ��v = "=" Or ��v = "<" Then
                    �Tv = "Tw"
                ElseIf ��v = "E" Then
                    �Tv = "S"
                Else
                    �Tv = ��v
                End If
            End If
            ���Oc = 0
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name = �[��v & "_" Then
                    ���Oc = ���Oc + 1
                End If
            Next objShp
                
            With ActiveSheet.Shapes(�[��v)
                .Select
                .Line.ForeColor.RGB = rootColor
                .Line.Weight = 3
                .TextFrame.Characters.Font.color = RGB(0, 10, 21)
                .ZOrder msoBringToFront
                ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (���Oc * 15), Selection.Top, 15, 15).Select
                Call �F�ϊ�(�Fv, clocode1, clocode2, clofont)
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(�Tv, "F", ""), 3)
                Selection.ShapeRange.Adjustments.Item(1) = 0.15
                'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Name = �[��v & "_"
                If InStr(�Fv, "/") > 0 Then
                    �x�[�X�F = Left(�Fv, InStr(�Fv, "/") - 1)
                Else
                    �x�[�X�F = �Fv
                End If
                myFontColor = clofont '�t�H���g�F���x�[�X�F�Ō��߂�
                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
                Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 8.5
                Selection.Font.Name = myFont
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                Selection.ShapeRange.TextFrame2.MarginLeft = 0
                Selection.ShapeRange.TextFrame2.MarginRight = 0
                Selection.ShapeRange.TextFrame2.MarginTop = 0
                Selection.ShapeRange.TextFrame2.MarginBottom = 0
                '�X�g���C�v�͌��ʂ��g��
                If clocode1 <> clocode2 Then
                    With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                        .color = clocode1
                        .color.TintAndShade = 0
                        .color.Brightness = 0
                        .Transparency = 0#
                        .Radius = 8
                    End With
                End If
                '���F�̎��̓��C������
                If clocode1 = 1315860 Then
                    Selection.ShapeRange.Line.ForeColor.RGB = RGB(250, 250, 250)
                End If
                '�}���}
                If �}v <> "" Then
                    myLeft = Selection.Left
                    myTop = Selection.Top
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    For Each objShp In Selection.ShapeRange
                        Set objShpTemp = objShp
                    Next objShp
                    ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                    Call �F�ϊ�(�}v, clocode1, clocode2, clofont)
                    myFontColor = clofont
                    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                    objShpTemp.Select False
                    Selection.Group.Select
                    Selection.Name = �[��v & "_"
                End If
                top�ړ�flg = False
line90:
                myleft2 = Selection.Left
                mytop2 = Selection.Top
                mytop3 = mytop2
                myright2 = Selection.Width + myleft2
                mybottom2 = Selection.Height + mytop2
                If myright2 > 1220 Then
                    Selection.Left = myleft2 - (myright2 - 1220)
                    myleft2 = Selection.Left
                End If
                usednodesp2 = Split(usednode2, ",")
                For ii = LBound(usednodesp2) + 1 To UBound(usednodesp2)
                    node = Split(usednodesp2(ii), "_")
                    �d�Ȃ�flg = False
                    nodeleft = CLng(node(0))
                    nodetop = CLng(node(1))
                    noderight = CLng(node(2))
                    nodebottom = CLng(node(3))
                    
                    For xx = nodeleft To noderight Step 2
                        For yy = nodetop To nodebottom Step 2
'                                If xx > myleft2 And yy > mytop2 Then Stop
                            �d�Ȃ�flg = xx > myleft2 And xx < myright2 And yy > mytop2 And yy < mybottom2
                            If �d�Ȃ�flg = True Then Exit For
                        Next yy
                        If �d�Ȃ�flg = True Then Exit For
                    Next xx
                    If (�d�Ȃ�flg = True And nodebottom + 2 <> myTop) Or myright2 > 1220 Then
                        mybottom2 = mybottom2 + nodebottom - myTop
                        myTop = nodebottom
                        Selection.Top = myTop
                        top�ړ�flg = True
                        GoTo line90
                    End If
                Next ii
                usednode2 = usednode2 & "," & Selection.Left & "_" & Selection.Top & "_" & Selection.Width + Selection.Left & "_" & Selection.Height + Selection.Top
                If top�ړ�flg = True Then
                    myTop = ActiveSheet.Shapes(�[��v).Top
                    myLeft = ActiveSheet.Shapes(�[��v).Left
                    myWidth = ActiveSheet.Shapes(�[��v).Width
                    With ActiveSheet.Shapes.AddLine(myLeft + myWidth / 2, myTop + myHeight, myleft2 + ((myright2 - myleft2) / 2), Selection.Top)
                        .Line.ForeColor.RGB = RGB(255, 0, 0)
                        .Line.Weight = 3
                        .Line.Transparency = 0.4
                        .Select False
                        Selection.Group.Select
                    End With
                End If
            End With
        Next i
        
        '���̐��i�i�Ԃ̒[���ꗗ�̍쐬
        Call SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, wb(0).Name)

        '���̐��i�i�ԂŎg�p����[�����őO�ʂɈړ�
        For Each objShp In wb(0).Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 Then
                If InStr(objShp.Name, "U") = 0 Then
                    For i = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
                        If �[���ꗗran(1, i) = objShp.Name Then
                            objShp.ZOrder msoBringToFront
                            Exit For
                        End If
                    Next i
                End If
            End If
        Next objShp
        
        '�D�F�̒[�����őO�ʂɈړ�
        For Each objShp In wb(0).Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 And objShp.Fill.ForeColor.RGB = elseColor Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '��n���d�����őO�ʂɈړ�
        For Each objShp In wb(0).Sheets("���_" & ���type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
line99:
        
        '�D�F�̃��C�����Ŕw�ʂɈړ�
        For Each objShp In wb(0).Sheets("���_" & ���type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = elseColor Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
               
        Dim SyTop As Long
        Dim flg As Long, �摜flg As Long, Sx As Long, Sy As Long
        
        '�[���摜���o�͂���ׂɃO���[�v�ɂ���
        wb(0).Sheets("���_" & ���type).Activate
        wb(0).Sheets("���_" & ���type).Cells(1, 1).Select
        myc = 0
        For Each objShp In wb(0).Sheets("���_" & ���type).Shapes
            If Right(objShp.Name, 1) = "!" Then
                objShp.Select False
                myc = myc + 1
            End If
        Next objShp
        
        If myc = 1 Then
            Selection.Name = "temp�[���摜"
        ElseIf myc > 1 Then
            Selection.Group.Name = "temp�[���摜"
        End If
        If myc > 0 Then
            wb(0).Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, �T�C�Yx * �{��, �T�C�Yy * �{��y).Select
            Selection.Name = "��f"
            wb(0).Sheets("���_" & ���type).Shapes("��f").Adjustments.Item(1) = 0
            wb(0).Sheets("���_" & ���type).Shapes("��f").Fill.Transparency = 1
            wb(0).Sheets("���_" & ���type).Shapes("��f").Line.Visible = msoFalse
            wb(0).Sheets("���_" & ���type).Shapes("temp�[���摜").Select False
            Selection.Group.Name = "temp�[���摜"
            wb(0).Sheets("���_" & ���type).Shapes("temp�[���摜").Select
            myfootwidth = Selection.Width
            myfootleft = Selection.Left
            myfootheight = Selection.Height
            '�o��
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '�摜�\��t���p�̖��ߍ��݃O���t���쐬
            Set cht = ActiveSheet.ChartObjects.add(0, 0, �T�C�Yx * �{��, myfootheight).Chart
             '���ߍ��݃O���t�ɓ\��t����
             DoEvents
             Sleep 10
             DoEvents
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            '�T�C�Y����
            ActiveWindow.Zoom = 100
            '��l = 1000
            '�{�� = 1
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth 1, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight 1, False, msoScaleFromTopLeft
            '
            cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\img\" & �T�ustr & "_foot.png", filtername:="PNG"
            cht.Parent.Delete
            wb(0).Sheets("���_" & ���type).Shapes("temp�[���摜").Delete
        End If
        
        wb(0).Sheets("���_" & ���type).Shapes.SelectAll
        Selection.Group.Select
        Selection.Name = "���"
        Selection.Top = 5
        Selection.Left = 5
line30:

        wb(0).Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, �T�C�Yx * �{��, �T�C�Yy * �{��y).Select
        Selection.Name = "��a"
        wb(0).Sheets("���_" & ���type).Shapes("��a").Adjustments.Item(1) = 0.02
        wb(0).Sheets("���_" & ���type).Shapes("��a").ZOrder msoSendToBack
'        WB(0).Sheets("���_" & ���type).Shapes("��a").Fill.PresetTextured 23
        wb(0).Sheets("���_" & ���type).Shapes("��a").Fill.Patterned msoPatternDashedHorizontal
        wb(0).Sheets("���_" & ���type).Shapes("��a").Fill.ForeColor.RGB = RGB(0, 10, 21) '���w�i�F
        wb(0).Sheets("���_" & ���type).Shapes("��a").Fill.BackColor.RGB = RGB(0, 10, 21)
        wb(0).Sheets("���_" & ���type).Shapes("��a").Fill.Transparency = 1
        '�؂�ڂ̕\��
        Dim k As String
        k = wb(0).Sheets("���_" & ���type).Cells.Find("k_", , , 1).Offset(0, 1)
        If IsNumeric(k) Then
            With wb(0).Sheets("���_" & ���type).Shapes.AddLine(k * �{��, 0, k * �{��, �T�C�Yy * �{��y)
                .Line.Weight = 1
                .Line.ForeColor.RGB = elseColor
                .Name = "k"
                .Select False
                Selection.Group.Select
                Selection.Name = "��"
                wb(0).Sheets("���_" & ���type).Shapes("��").ZOrder msoSendToBack
            End With
        End If
'        If WB(0).Sheets("���_" & ���type).Shapes.Count > 1 Then
'            WB(0).Sheets("���_" & ���type).Shapes("���").Select False
'            Selection.Group.Select
'            Selection.Name = "�z��"
'        End If
        
        '.Cells(1, 1).Select
    End With
    If ���}�̂� = "1" Then
        �����[�� = SQL_�z���}_�[���ꗗ(wb(0).Name, ���type)
        If �����[��(0) <> Empty Then
            Dim myMsg As String: myMsg = "���̒[���̍��W���s�����Ă��܂��B" & vbCrLf
            For u = LBound(�����[��) To UBound(�����[��)
                myMsg = myMsg & vbCrLf & �����[��(u)
            Next u
        End If
        
        With wb(0).Sheets("���_" & ���type).Shapes("��")
            If myMsg = "" Then
                myMsg = "�s���[���͂���܂���"
            End If
            '�Ώۂ̐��i�i��
            myMsg = myMsg & vbCrLf & vbCrLf & "�Ώۂ̐��i�i��"
            For r = 1 To ���i�i��RANc
                myMsg = myMsg & vbCrLf & ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            Next r
            .TextFrame.Characters.Text = myMsg
        End With
        PlaySound "��������2"
    End If
    
    �z���}�쐬3 = Split(myfootleft & "_" & myfootwidth & "_" & myfootheight, "_")

    Call �œK�����ǂ�

End Function


Public Function �z���}�쐬_�o�H����(���i�i��str, ���type)

    
    If IsError(���i�i��str) Then
        PlaySound "��������2"
        ���i�i��str = "8211158560"
        �T�ustr = ""
        ���}�̂� = "1"
        ���type = "�⋋"
        ��n���摜Sheet = ""
        Call ���i�i��RAN_set2(���i�i��RAN, "����", ���type, "")
    End If
    
    '�f�B���N�g���쐬
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��"
    End If
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str, vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img"
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\css", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\css"
    End If
    
'    If IsError(���i�i��str) Then
'        ���i�i��str = "8501K006"
'        �T�ustr = "2"
'        ���}�̂� = "0"
'        ���type = "F"
'        ��n���摜Sheet = "�n���}_���C���i��8501K006"
'        Call ���i�i��RAN_set2(���i�i��RAN, "����", "F", "8501K006")
'    End If
    
    Call �œK��
        
    '���i�i��str = ""
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
        
    For Each ws(0) In wb.Sheets
        If ws(0).Name = "���_" Then
            Stop
        End If
    Next ws
    
    If IsError(���type) Or ���type = "" Then
        ���type = Mid(ActiveSheet.Name, 4)
    End If
    
    On Error Resume Next
    wb.Sheets("���_" & ���type).Activate
    If Err = 9 Then
        Call �œK�����ǂ�
        End
    End If
    On Error GoTo 0
    
    With wb.Sheets("���_" & ���type)
        Dim key As Range: Set key = .Cells.Find("Size_", , , 1)
    
        '���̃T�C�Y
        �T�C�Y = .Cells(key.Row, key.Column).Offset(, 1)
        �T�C�Ys = Split(�T�C�Y, "_")
        �T�C�Yx = �T�C�Ys(0)
        �T�C�Yy = �T�C�Ys(1)
                
        �{�� = 1220 / �T�C�Yx '�T�C�Yx / 1220
        �{��y = 480 / �T�C�Yy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "�l�r �S�V�b�N"
        '�I�[�g�V�F�C�v���폜
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            objShp.Delete
        Next objShp
        
        Dim ���Oc As Long
        '���}�̍쐬
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For Y = 2 To lastRow
            '�[�����^�C�g��
            �[�� = .Cells(Y, X)
            ���Ws = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(���Ws) < 1 Then ���WErr = 1 Else ���WErr = 0
            
            If ���WErr = 0 Then
                ���Wx = ���Ws(0) * �{��
                ���Wy = ���Ws(1) * �{��y
                
                ���Od = 0
                On Error Resume Next
                ���Od = wb.ActiveSheet.Shapes.Range(�[��).count
                If Err = 1004 Then ���Od = 0
                On Error GoTo 0
                
                If ���Od = 0 Then
                    Select Case Left(�[��, 1)
                    Case "U"
                        With wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            .Name = �[��
                            .Left = ���Wx - 4
                            .Top = ���Wy - 4
                            If ���}�̂� = "1" Then
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
'                                .TextFrame.Characters.Font.Size = 4
'                                .TextFrame.Characters.Text = Replace(�[��, "U", "")
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    Case Else
                        With wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = �[��
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = �[��
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.ForeColor.RGB = RGB(250, 250, 250)
                            If ���}�̂� = "1" Then
                                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                            Else
                                .TextFrame.Characters.Font.color = RGB(200, 200, 200)
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                            
                            .Left = ���Wx - 15
                            .Top = ���Wy - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If ���Wxbak <> "" Then
                    
                    On Error Resume Next
                    ���Oc1 = wb.Sheets("���_" & ���type).Shapes.Range(�[��bak & " to " & �[��).count
                    If Err = 1004 Then ���Oc1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    ���Oc2 = wb.Sheets("���_" & ���type).Shapes.Range(�[�� & " to " & �[��bak).count
                    If Err = 1004 Then ���Oc2 = 0
                    On Error GoTo 0
                        
                    If ���Oc1 = 0 And ���Oc2 = 0 And �[�� <> �[��bak Then
                        With wb.Sheets("���_" & ���type).Shapes.AddLine(���Wxbak, ���Wybak, ���Wx, ���Wy)
                            .Name = �[��bak & " to " & �[��
                            .Line.Weight = 3.2
                            If ���}�̂� = "1" Then
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    End If
                End If
                ���Wxbak = ���Wx
                ���Wybak = ���Wy
                �[��bak = �[��
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                ���Wsbak = Split(.Cells(Y, 2), "_")
                ���Wxbak = ���Wsbak(0) * �{��
                ���Wybak = ���Wsbak(1) * �{��y
                �[��bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        '�[�����őO�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '��n���d�����őO�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '�D�F�̒[�����Ŕw�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
              objShp.ZOrder msoSendToBack
            End If
        Next objShp
line99:
        
        '�D�F�̃��C�����Ŕw�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = RGB(150, 150, 150) Or objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
        
        wb.Sheets("���_" & ���type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("���_" & ���type).Shapes("temp").Select
        Selection.Left = 5
        Selection.Top = 5
        Selection.Ungroup
        
        wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, �T�C�Yx * �{��, �T�C�Yy * �{��y).Select
        Selection.Name = "��"
        wb.Sheets("���_" & ���type).Shapes("��").Adjustments.Item(1) = 0.02
        wb.Sheets("���_" & ���type).Shapes("��").ZOrder msoSendToBack
        wb.Sheets("���_" & ���type).Shapes("��").Fill.PresetTextured msoTextureBlueTissuePaper
        wb.Sheets("���_" & ���type).Shapes("��").Fill.Transparency = 0.62

        wb.Sheets("���_" & ���type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("���_" & ���type).Shapes("temp").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height
        Stop
        '�o��
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '�摜�\��t���p�̖��ߍ��݃O���t���쐬
        Set cht = ActiveSheet.ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
         '���ߍ��݃O���t�ɓ\��t����
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        
        '�T�C�Y����
        ActiveWindow.Zoom = 100
        ��l = 1000
        myW = Selection.Width
        myH = Selection.Height
        If myW > myH Then
            �{�� = ��l / myW
        Else
            �{�� = ��l / myH
        End If
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft

        cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img\" & "Base.png", filtername:="PNG"

         '���ߍ��݃O���t���폜
        cht.Parent.Delete
        wb.Sheets("���_" & ���type).Shapes("temp").Select
        Selection.Ungroup
        
'       ���o�H
        Call SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
        Stop
        Set ws(2) = wb.Sheets("���_" & ���type)
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            For i2 = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
                '���[����
                ws(2).Shapes(�z���[��RAN(0, i)).Select
                ws(2).Shapes(�z���[��RAN(0, i2)).Select False
                Set �[��from = .Cells.Find(�z���[��RAN(0, i), , , 1)
                Set �[��to = .Cells.Find(�z���[��RAN(0, i2), , , 1)
                    
                If �z���[��RAN(0, i) <> �z���[��RAN(0, i2) Then
                    '���z������[���Ԃ̃��C���ɐF�t��
                    If �[��from Is Nothing Or �[��to Is Nothing Then GoTo nextI
                    If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                        
                    Set �[��1 = �[��from
                    Set �[��2 = Nothing
                    For Y = �[��from.Row To �[��to.Row Step myStep
                        'from���獶�ɐi��
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                            On Error GoTo 0
                           
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Select False
                            End If
                        Loop
                        'to�̍s�܂ŏ�܂��͉��ɐi��
                        Do Until �[��1.Row = �[��to.Row
                            Set �[��2 = �[��1.Offset(myStep, 0)
                            If �[��1 <> �[��2 Then
                                On Error Resume Next
                                    ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                    ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                On Error GoTo 0
                            End If
                            Set �[��1 = �[��2
                        Loop
                        'to�̍s���E�ɐi��
                        Do Until �[��1.Column = �[��to.Column
                            Set �[��2 = �[��1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                            On Error GoTo 0
                            Set �[��1 = �[��2
                        Loop
                    Next Y
                End If
                '�o�H�̍��W���擾����ׂɃO���[�v��
                If Selection.ShapeRange.count > 1 Then
                    Selection.Group.Name = "temp"
                    ws(2).Shapes("temp").Select
                End If
                myLeft = Selection.Left
                myTop = Selection.Top
                myWidth = Selection.Width
                myHeight = Selection.Height
                Selection.Copy
                
                If Selection.ShapeRange.Type = msoGroup Then
                    ws(2).Shapes("temp").Select
                    Selection.Ungroup
                End If
                ws(2).Paste
                
                If Selection.ShapeRange.Type = msoGroup Then
                    For Each ob In Selection.ShapeRange.GroupItems
                        If InStr(ob.Name, "to") > 0 Then
                            ob.Line.ForeColor.RGB = RGB(255, 100, 100)
                        Else
                            ob.Fill.ForeColor.RGB = RGB(255, 100, 100)
                        End If
                    Next
                    ws(2).Shapes("temp").Select
                Else
                    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 100, 100)
                End If
                
                Selection.Name = "temp"
                
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '�摜�\��t���p�̖��ߍ��݃O���t���쐬
                Set cht = ActiveSheet.ChartObjects.add(0, 0, myWidth, myHeight).Chart
                 '���ߍ��݃O���t�ɓ\��t����
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                
                '�T�C�Y����
                ActiveWindow.Zoom = 100
                ��l = 1000
                If myWidth > myHeight Then
                    �{�� = ��l / myWidth
                Else
                    �{�� = ��l / myHeight
                End If
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
        
                cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img\" & �z���[��RAN(0, i) & "to" & �z���[��RAN(0, i2) & ".png", filtername:="PNG"

                myPath = ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\" & �z���[��RAN(0, i) & "to" & �z���[��RAN(0, i2) & ".html"
                Stop
                'Call TEXT�o��_�z���o�Hhtml(mypath, �[��from.Value, �[��to.Value)
                Stop
                myPath = ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\css\" & �z���[��RAN(0, i) & "to" & �z���[��RAN(0, i2) & ".css"
'                Call TEXT�o��_�z���o�Hcss(mypath, myLeft / mybasewidth, myTop / mybaseheight, myWidth / mybasewidth, myHeight / mybaseheight, 255)
                cht.Parent.Delete
            
                ws(2).Shapes("temp").Delete
            Next i2
            
nextI:
        Next i
        
        Stop '�����܂�
If ���}�̂� = "1" Then GoTo line99
    
        �摜add = �T�C�Yy * �{��y
        '���z������[���̐F�t��
        Call SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            If �z���[��RAN(0, i) = "" Then GoTo nextI
            Set �z�� = .Cells.Find(�z���[��RAN(0, i), , , 1)
            If �z�� Is Nothing Then GoTo nextI
            ��F = �z���[��RAN(1, i)
            If ��F = "" Then
                With wb.Sheets("���_" & ���type).Shapes(�z��.Value)
                    .Select
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = RGB(255, 100, 100)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    .Copy
                    Sleep 5
                    ActiveSheet.Paste
                    Selection.Name = �z��.Value & "!"
                    Selection.Left = myLeft
                    Selection.Top = �摜add
                    �摜add = �摜add + Selection.Height
                End With
                
                '��n���}�̎擾�Ɣz�z
                With wb.Sheets(��n���摜Sheet)
                    .Activate
                    n = 0
                    For Each obj In .Shapes(�z��.Value & "_1").GroupItems
                        If obj.Name Like �z��.Value & "_1*" Then
                            If obj.Name <> �z��.Value & "_1_t" Then
                                If obj.Name <> �z��.Value & "_1_b" Then
                                    If n = 0 Then
                                        obj.Select True
                                    Else
                                        obj.Select False
                                    End If
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next obj
                    Selection.Copy
                    .Cells(1, 1).Select
                End With
                
                .Activate
                ActiveSheet.Pictures.Paste.Select
                'Sheets(��n���摜Sheet).Shapes(�z��.Value & "_1").Copy
                'Selection.Top = (�T�C�Yy * �{��y) + �摜add + myHeight
                Selection.Left = myLeft
                �{��a = (myWidth / Selection.Width) * 3
                If �{��a > 0.7 Then �{��a = 0.7
                Selection.ShapeRange.ScaleHeight �{��a, msoFalse, msoScaleFromTopLeft
                Selection.Top = �摜add
                ActiveSheet.Shapes(�z��.Value & "!").Select False

                Selection.Name = �z��.Value & "!"
                �摜add = �摜add + Selection.Height
            End If
            Set �z��bak = �z��
            ��Fbak = ��F

        Next i
        
        '���z������[���Ԃ̃��C���ɐF�t��
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            For i2 = i + 1 To UBound(�z���[��RAN, 2)
                Set �[��from = .Cells.Find(�z���[��RAN(0, i), , , 1)
                Set �[��to = .Cells.Find(�z���[��RAN(0, i2), , , 1)
'                If �[��from Is Nothing Or �[��to Is Nothing Then GoTo line31
                If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                    
                Set �[��1 = �[��from
                �㉺�ɐi��flg = 0
                For Y = �[��from.Row To �[��to.Row Step myStep
                    'from���獶�ɐi��
                    If �[��1.Row = �[��from.Row Or �㉺�ɐi��flg = 0 Then
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            �o�H = �o�H & "," & Selection
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                            If �[��1 = �[��2.Offset(myStep, 0) Then
                                �㉺�ɐi��flg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'to�̍s�܂ŏ�܂��͉��ɐi��
                    If (�[��1.Column = 1 Or �㉺�ɐi��flg = 1) And �[��1.Row <> �[��to.Row Then

                        Set �[��2 = �[��1.Offset(myStep, 0)
                        If �[��1 <> �[��2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                        If �[��1 <> �[��2.Offset(myStep, 0) Then
                            �㉺�ɐi��flg = 0
                        End If
                        'If �㉺�ɐi��flg = 1 Then GoTo line15
                    End If
                    
                    'to�̍s���E�ɐi��
                    If �[��1.Row = �[��to.Row Then
                        Do Until �[��1.Column = �[��to.Column
                            Set �[��2 = �[��1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set �[��2 = Nothing
            Next i2
        Next i
                

               
        Dim SyTop As Long
        Dim flg As Long, �摜flg As Long, Sx As Long, Sy As Long
'        '�}����̋󂢂Ă���X�y�[�X�Ɉړ�
'        For Each objShp In WB.Sheets("���_" & ���type).Shapes
'            �摜flg = 0: SyTop = (�T�C�Yy * �{��y) + 5
'line20:
'            flg = 0
'            For Each objShp2 In WB.Sheets("���_" & ���type).Shapes
'                'If objShp.Name = "501!" And objShp2.Name = "843!" Then Stop
'                If Right(objShp.Name, 1) = "!" And Right(objShp2.Name, 1) = "!" Then
'                    If objShp.Name <> objShp2.Name Then
'                        �摜flg = 1
'                        For Sx = objShp.Left To objShp.Left + objShp.width Step 1
'                            If objShp2.Left <= Sx And objShp2.Left + objShp2.width >= Sx Then
'                                If objShp2.Top <= SyTop And objShp2.Top + objShp2.height >= SyTop Then
'                                    flg = 1
'                                    SyTop = SyTop + 10
'                                    GoTo line20
'                                End If
'                            End If
'                        Next Sx
'                    End If
'                End If
'            Next objShp2
'
'            If flg = 1 Then GoTo line20
'
'            If �摜flg = 1 Then
'                objShp.Top = SyTop
'            End If
'        Next objShp
                
       
line30:

        
        '.Cells(1, 1).Select
    End With
    If ���}�̂� = "1" Then

    End If

    Call �œK�����ǂ�

End Function

Public Function �z���}�쐬_�o�H����2(Optional ���i�i��str, Optional ���type)

    
    If IsError(���i�i��str) Then
        PlaySound "��������2"
        ���i�i��str = "8211158560"
        �T�ustr = ""
        ���type = "�⋋"
        ��n���摜Sheet = ""
        Call ���i�i��RAN_set2(���i�i��RAN, "����", ���type, "")
    End If
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Set ws(2) = wb.Sheets("���_" & ���type)
    
    '�f�B���N�g���쐬
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��"
    End If
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str, vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img"
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\css", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\css"
    End If
       
    Call �œK��
    
        
    For Each ws(0) In wb.Sheets
        If ws(0).Name = "���_" Then
            Stop
        End If
    Next ws
    
    If IsError(���type) Or ���type = "" Then
        ���type = Mid(ActiveSheet.Name, 4)
    End If
    
    On Error Resume Next
    wb.Sheets("���_" & ���type).Activate
    If Err = 9 Then
        Call �œK�����ǂ�
        End
    End If
    On Error GoTo 0
    
    With wb.Sheets("���_" & ���type)
        Dim key As Range: Set key = .Cells.Find("Size_", , , 1)
    
        '���̃T�C�Y
        �T�C�Y = .Cells(key.Row, key.Column).Offset(, 1)
        �T�C�Ys = Split(�T�C�Y, "_")
        �T�C�Yx = �T�C�Ys(0)
        �T�C�Yy = �T�C�Ys(1)
                
        �{�� = 1220 / �T�C�Yx '�T�C�Yx / 1220
        �{��y = 480 / �T�C�Yy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "�l�r �S�V�b�N"
        '�I�[�g�V�F�C�v���폜
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            objShp.Delete
        Next objShp
        
        Dim ���Oc As Long
        '���}�̍쐬
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For Y = 2 To lastRow
            '�[�����^�C�g��
            �[�� = .Cells(Y, X)
            ���Ws = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(���Ws) < 1 Then ���WErr = 1 Else ���WErr = 0
            
            If ���WErr = 0 Then
                ���Wx = ���Ws(0) * �{��
                ���Wy = ���Ws(1) * �{��y
                
                ���Od = 0
                On Error Resume Next
                ���Od = wb.ActiveSheet.Shapes.Range(�[��).count
                If Err = 1004 Then ���Od = 0
                On Error GoTo 0
                
                If ���Od = 0 Then
                    Select Case Left(�[��, 1)
                    Case "U"
                        With wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            .Name = �[��
                            .Left = ���Wx - 4
                            .Top = ���Wy - 4
                            If ���}�̂� = "1" Then
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
'                                .TextFrame.Characters.Font.Size = 4
'                                .TextFrame.Characters.Text = Replace(�[��, "U", "")
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    Case Else
                        With wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = �[��
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = �[��
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.ForeColor.RGB = RGB(250, 250, 250)
                            If ���}�̂� = "1" Then
                                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                            Else
                                .TextFrame.Characters.Font.color = RGB(200, 200, 200)
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                            
                            .Left = ���Wx - 15
                            .Top = ���Wy - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If ���Wxbak <> "" Then
                    
                    On Error Resume Next
                    ���Oc1 = wb.Sheets("���_" & ���type).Shapes.Range(�[��bak & " to " & �[��).count
                    If Err = 1004 Then ���Oc1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    ���Oc2 = wb.Sheets("���_" & ���type).Shapes.Range(�[�� & " to " & �[��bak).count
                    If Err = 1004 Then ���Oc2 = 0
                    On Error GoTo 0
                        
                    If ���Oc1 = 0 And ���Oc2 = 0 And �[�� <> �[��bak Then
                        With wb.Sheets("���_" & ���type).Shapes.AddLine(���Wxbak, ���Wybak, ���Wx, ���Wy)
                            .Name = �[��bak & " to " & �[��
                            .Line.Weight = 3.2
                            If ���}�̂� = "1" Then
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    End If
                End If
                ���Wxbak = ���Wx
                ���Wybak = ���Wy
                �[��bak = �[��
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                ���Wsbak = Split(.Cells(Y, 2), "_")
                ���Wxbak = ���Wsbak(0) * �{��
                ���Wybak = ���Wsbak(1) * �{��y
                �[��bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        '�[�����őO�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '��n���d�����őO�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '�D�F�̒[�����Ŕw�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
              objShp.ZOrder msoSendToBack
            End If
        Next objShp
line99:
        
        '�D�F�̃��C�����Ŕw�ʂɈړ�
        For Each objShp In wb.Sheets("���_" & ���type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = RGB(150, 150, 150) Or objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
        
        wb.Sheets("���_" & ���type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("���_" & ���type).Shapes("temp").Select
        Selection.Left = 5
        Selection.Top = 5
        Selection.Ungroup
        
        wb.Sheets("���_" & ���type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, �T�C�Yx * �{��, �T�C�Yy * �{��y).Select
        Selection.Name = "��"
        wb.Sheets("���_" & ���type).Shapes("��").Adjustments.Item(1) = 0.02
        wb.Sheets("���_" & ���type).Shapes("��").ZOrder msoSendToBack
        wb.Sheets("���_" & ���type).Shapes("��").Fill.PresetTextured msoTextureBlueTissuePaper
        wb.Sheets("���_" & ���type).Shapes("��").Fill.Transparency = 0.62

        wb.Sheets("���_" & ���type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("���_" & ���type).Shapes("temp").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height

        '�o��
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '�摜�\��t���p�̖��ߍ��݃O���t���쐬
        Set cht = ActiveSheet.ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
         '���ߍ��݃O���t�ɓ\��t����
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        
        '�T�C�Y����
        ActiveWindow.Zoom = 100
        ��l = 1000
        myW = Selection.Width
        myH = Selection.Height
        If myW > myH Then
            �{�� = ��l / myW
        Else
            �{�� = ��l / myH
        End If
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
        '��Base�̏o��
        cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img\" & "Base.png", filtername:="PNG"

         '���ߍ��݃O���t���폜
        cht.Parent.Delete
        wb.Sheets("���_" & ���type).Shapes("temp").Select
        Selection.Ungroup
'�@�@�@ ���T�u���̔z���}
        Call SQL_�z���T�u�擾(�z���T�uRAN, ���i�i��str)
        For i = LBound(�z���T�uRAN, 2) + 1 To UBound(�z���T�uRAN, 2) '�T�u�ꗗ
            �T�ustr = �z���T�uRAN(0, i)
            Call SQL_�z���[���擾2(�z���[��RAN, ���i�i��str, �T�ustr)
            Call SQL_�z����n���擾(�z����n��RAN, ���i�i��str, �T�ustr)
            For i2 = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2) '�T�u�̒[���ꗗ
                For i3 = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2) '�T�u�̒[���ꗗ
                    If i2 <> i3 Then
                        On Error Resume Next
                        ws(2).Shapes(�z���[��RAN(0, i2)).Select False
                        ws(2).Shapes(�z���[��RAN(0, i3)).Select False
                        errNumber = Err.Number
                        On Error GoTo 0
                        If errNumber = -2147024809 Then GoTo nextI3
                        Set �[��from = .Cells.Find(�z���[��RAN(0, i2), , , 1)
                        Set �[��to = .Cells.Find(�z���[��RAN(0, i3), , , 1)
                            
                        If �z���[��RAN(0, i2) <> �z���[��RAN(0, i3) Then
                            '���z������[���Ԃ̃��C���ɐF�t��
                            If �[��from Is Nothing Or �[��to Is Nothing Then GoTo nextI
                            If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                                
                            Set �[��1 = �[��from
                            Set �[��2 = Nothing
                            For Y = �[��from.Row To �[��to.Row Step myStep
                                'from���獶�ɐi��
                                Do Until �[��1.Column = 1
                                    Set �[��2 = �[��1.Offset(0, -2)
                                    On Error Resume Next
                                        ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                        ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                    On Error GoTo 0
                                   
                                    Set �[��1 = �[��2
                                    If Left(�[��1.Value, 1) = "U" Then
                                        ActiveSheet.Shapes(�[��1.Value).Select False
                                    End If
                                Loop
                                'to�̍s�܂ŏ�܂��͉��ɐi��
                                Do Until �[��1.Row = �[��to.Row
                                    Set �[��2 = �[��1.Offset(myStep, 0)
                                    If �[��1 <> �[��2 Then
                                        On Error Resume Next
                                            ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                            ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                        On Error GoTo 0
                                    End If
                                    Set �[��1 = �[��2
                                Loop
                                'to�̍s���E�ɐi��
                                Do Until �[��1.Column = �[��to.Column
                                    Set �[��2 = �[��1.Offset(0, 2)
                                    On Error Resume Next
                                        ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                        ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                    On Error GoTo 0
                                    Set �[��1 = �[��2
                                Loop
                            Next Y
                        End If
                        
                    End If
nextI3:
                Next i3
            Next i2
                    
            '�o�͏���
            Selection.Group.Name = "temp"
            ws(2).Shapes("temp").Select
            myLeft = Selection.Left
            myTop = Selection.Top
            Selection.Copy
            ws(2).Paste
            Selection.Left = myLeft
            Selection.Top = myTop
            ws(2).Shapes("temp").Ungroup
            For Each ob In Selection.ShapeRange.GroupItems
                ob.Name = ob.Name & "!"
            Next ob
            Selection.Name = "temp2"
            If Selection.ShapeRange.Type = msoGroup Then
                For Each ob In Selection.ShapeRange.GroupItems
                    If InStr(ob.Name, "to") > 0 Then
                        ob.Line.ForeColor.RGB = RGB(255, 100, 100)
                    Else
                        '���z�������n���d����\��
                        Dim �Fv As String, �Tv As String, �[��v As String, �}v As String, �n��v As String
                        For i4 = LBound(�z����n��RAN, 2) To UBound(�z����n��RAN, 2)
                            Debug.Print �z����n��RAN(2, i4)
                            If ob.Name = �z����n��RAN(2, i4) & "!" Then
                                �Fv = �z����n��RAN(0, i4)
                                If �Fv = "" Then Exit For
                                �Tv = �z����n��RAN(1, i4)
                                �[��v = �z����n��RAN(2, i4) & "!"
                                �}v = �z����n��RAN(3, i4)
                                �n��v = �z����n��RAN(4, i4)
                                
                                ���Oc = 0
                                For Each objShp In ActiveSheet.Shapes
                                    If objShp.Name = �[��v & "_" Then
                                        ���Oc = ���Oc + 1
                                    End If
                                Next objShp
                                    
                                With ActiveSheet.Shapes(�[��v)
                                    .Select
                                    .Line.ForeColor.RGB = RGB(255, 100, 100)
                                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                                    .ZOrder msoBringToFront
                                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (���Oc * 15), Selection.Top, 15, 15).Select
                                    Call �F�ϊ�(�Fv, clocode1, clocode2, clofont)
                                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(�Tv, "F", ""), 3)
                                    Selection.ShapeRange.Adjustments.Item(1) = 0.15
                                    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                                    Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                                    Selection.ShapeRange.Name = �[��v & "_"
                                    If InStr(�Fv, "/") > 0 Then
                                        �x�[�X�F = Left(�Fv, InStr(�Fv, "/") - 1)
                                    Else
                                        �x�[�X�F = �Fv
                                    End If
                                    
                                    myFontColor = clofont '�t�H���g�F���x�[�X�F�Ō��߂�
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 6
                                    Selection.Font.Name = myFont
                                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                                    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                                    Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                    '�X�g���C�v�͌��ʂ��g��
                                    If clocode1 <> clocode2 Then
                                        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                                            .color = clocode1
                                            .color.TintAndShade = 0
                                            .color.Brightness = 0
                                            .Transparency = 0#
                                            .Radius = 8
                                        End With
                                    End If
                                    '�}���}
                                    If �}v <> "" Then
                                        myLeft = Selection.Left
                                        myTop = Selection.Top
                                        myHeight = Selection.Height
                                        myWidth = Selection.Width
                                        For Each objShp In Selection.ShapeRange
                                            Set objShpTemp = objShp
                                        Next objShp
                                        ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                                        Call �F�ϊ�(�}v, clocode1, clocode2, clofont)
                                        myFontColor = clofont
                                        Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                                        Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                                        objShpTemp.Select False
                                        Selection.Group.Select
                                        Selection.Name = �[��v & "_"
                                    End If
                                End With
                            Else
                                ob.Fill.ForeColor.RGB = RGB(255, 100, 100)
                            End If
                        Next i4
                        
                        
                    End If
                Next ob

            Else
                Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 100, 100)
            End If
     
            Dim choseiFlg As Boolean: choseiFlg = False
            If mybasewidth < Selection.Left + 15 Then choseiFlg = True
            
            Dim flgCount As Long: flgCount = 0
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name Like "*!_" Then
                    If flgCount = 0 Then objShp.Select Else objShp.Select False
                    flgCount = flgCount + 1
                End If
            Next objShp
            If choseiFlg = True Then Selection.Left = mybasewidth - (���Oc + 1) * 15
                
            ws(2).Shapes("temp2").Select False
            On Error Resume Next
                Selection.Group.Name = "temp2"
            On Error GoTo 0
            ws(2).Shapes.SelectAll
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '�摜�\��t���p�̖��ߍ��݃O���t���쐬
            Set cht = ActiveSheet.ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
             '���ߍ��݃O���t�ɓ\��t����
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            
            '�T�C�Y����
            ActiveWindow.Zoom = 100
            ��l = 1000
            myW = Selection.Width
            myH = Selection.Height
            If myW > myH Then
                �{�� = ��l / myW
            Else
                �{�� = ��l / myH
            End If
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
    
            cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img\" & �T�ustr & ".png", filtername:="PNG"
            cht.Parent.Delete
            ws(2).Shapes("temp2").Delete
        Next i
'       ���o�H
        Call SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
        Stop
        
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            For i2 = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
                '���[����
                ws(2).Shapes(�z���[��RAN(0, i)).Select
                ws(2).Shapes(�z���[��RAN(0, i2)).Select False
                Set �[��from = .Cells.Find(�z���[��RAN(0, i), , , 1)
                Set �[��to = .Cells.Find(�z���[��RAN(0, i2), , , 1)
                    
                If �z���[��RAN(0, i) <> �z���[��RAN(0, i2) Then
                    '���z������[���Ԃ̃��C���ɐF�t��
                    If �[��from Is Nothing Or �[��to Is Nothing Then GoTo nextI
                    If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                        
                    Set �[��1 = �[��from
                    Set �[��2 = Nothing
                    For Y = �[��from.Row To �[��to.Row Step myStep
                        'from���獶�ɐi��
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                            On Error GoTo 0
                           
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Select False
                            End If
                        Loop
                        'to�̍s�܂ŏ�܂��͉��ɐi��
                        Do Until �[��1.Row = �[��to.Row
                            Set �[��2 = �[��1.Offset(myStep, 0)
                            If �[��1 <> �[��2 Then
                                On Error Resume Next
                                    ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                    ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                                On Error GoTo 0
                            End If
                            Set �[��1 = �[��2
                        Loop
                        'to�̍s���E�ɐi��
                        Do Until �[��1.Column = �[��to.Column
                            Set �[��2 = �[��1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select False
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select False
                            On Error GoTo 0
                            Set �[��1 = �[��2
                        Loop
                    Next Y
                End If
                '�o�H�̍��W���擾����ׂɃO���[�v��
                If Selection.ShapeRange.count > 1 Then
                    Selection.Group.Name = "temp"
                    ws(2).Shapes("temp").Select
                End If
                myLeft = Selection.Left
                myTop = Selection.Top
                myWidth = Selection.Width
                myHeight = Selection.Height
                Selection.Copy
                
                If Selection.ShapeRange.Type = msoGroup Then
                    ws(2).Shapes("temp").Select
                    Selection.Ungroup
                End If
                ws(2).Paste
                
                If Selection.ShapeRange.Type = msoGroup Then
                    For Each ob In Selection.ShapeRange.GroupItems
                        If InStr(ob.Name, "to") > 0 Then
                            ob.Line.ForeColor.RGB = RGB(255, 100, 100)
                        Else
                            ob.Fill.ForeColor.RGB = RGB(255, 100, 100)
                        End If
                    Next
                    ws(2).Shapes("temp").Select
                Else
                    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 100, 100)
                End If
                
                Selection.Name = "temp"
                
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '�摜�\��t���p�̖��ߍ��݃O���t���쐬
                Set cht = ActiveSheet.ChartObjects.add(0, 0, myWidth, myHeight).Chart
                 '���ߍ��݃O���t�ɓ\��t����
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                '�T�C�Y����
                ActiveWindow.Zoom = 100
                ��l = 1000
                If myWidth > myHeight Then
                    �{�� = ��l / myWidth
                Else
                    �{�� = ��l / myHeight
                End If
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
        
                cht.Export fileName:=ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\img\" & �z���[��RAN(0, i) & "to" & �z���[��RAN(0, i2) & ".png", filtername:="PNG"

                myPath = ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\" & �z���[��RAN(0, i) & "to" & �z���[��RAN(0, i2) & ".html"
                Stop
                'Call TEXT�o��_�z���o�Hhtml(mypath, �[��from.Value, �[��to.Value)
                Stop
                myPath = ActiveWorkbook.Path & "\56_�z���}_�U��\" & ���i�i��str & "\css\" & �z���[��RAN(0, i) & "to" & �z���[��RAN(0, i2) & ".css"
'                Call TEXT�o��_�z���o�Hcss(mypath, myLeft / mybasewidth, myTop / mybaseheight, myWidth / mybasewidth, myHeight / mybaseheight, 255)
                cht.Parent.Delete
            
                ws(2).Shapes("temp").Delete
            Next i2
            
nextI:
        Next i
        
        Stop '�����܂�
If ���}�̂� = "1" Then GoTo line99
    
        �摜add = �T�C�Yy * �{��y
        '���z������[���̐F�t��
        Call SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            If �z���[��RAN(0, i) = "" Then GoTo nextI
            Set �z�� = .Cells.Find(�z���[��RAN(0, i), , , 1)
            If �z�� Is Nothing Then GoTo nextI
            ��F = �z���[��RAN(1, i)
            If ��F = "" Then
                With wb.Sheets("���_" & ���type).Shapes(�z��.Value)
                    .Select
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = RGB(255, 100, 100)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    .Copy
                    Sleep 5
                    ActiveSheet.Paste
                    Selection.Name = �z��.Value & "!"
                    Selection.Left = myLeft
                    Selection.Top = �摜add
                    �摜add = �摜add + Selection.Height
                End With
                
                '��n���}�̎擾�Ɣz�z
                With wb.Sheets(��n���摜Sheet)
                    .Activate
                    n = 0
                    For Each obj In .Shapes(�z��.Value & "_1").GroupItems
                        If obj.Name Like �z��.Value & "_1*" Then
                            If obj.Name <> �z��.Value & "_1_t" Then
                                If obj.Name <> �z��.Value & "_1_b" Then
                                    If n = 0 Then
                                        obj.Select True
                                    Else
                                        obj.Select False
                                    End If
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next obj
                    Selection.Copy
                    .Cells(1, 1).Select
                End With
                
                .Activate
                ActiveSheet.Pictures.Paste.Select
                'Sheets(��n���摜Sheet).Shapes(�z��.Value & "_1").Copy
                'Selection.Top = (�T�C�Yy * �{��y) + �摜add + myHeight
                Selection.Left = myLeft
                �{��a = (myWidth / Selection.Width) * 3
                If �{��a > 0.7 Then �{��a = 0.7
                Selection.ShapeRange.ScaleHeight �{��a, msoFalse, msoScaleFromTopLeft
                Selection.Top = �摜add
                ActiveSheet.Shapes(�z��.Value & "!").Select False

                Selection.Name = �z��.Value & "!"
                �摜add = �摜add + Selection.Height
            End If
            Set �z��bak = �z��
            ��Fbak = ��F

        Next i
        
        '���z������[���Ԃ̃��C���ɐF�t��
        For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
            For i2 = i + 1 To UBound(�z���[��RAN, 2)
                Set �[��from = .Cells.Find(�z���[��RAN(0, i), , , 1)
                Set �[��to = .Cells.Find(�z���[��RAN(0, i2), , , 1)
'                If �[��from Is Nothing Or �[��to Is Nothing Then GoTo line31
                If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                    
                Set �[��1 = �[��from
                �㉺�ɐi��flg = 0
                For Y = �[��from.Row To �[��to.Row Step myStep
                    'from���獶�ɐi��
                    If �[��1.Row = �[��from.Row Or �㉺�ɐi��flg = 0 Then
                        Do Until �[��1.Column = 1
                            Set �[��2 = �[��1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            �o�H = �o�H & "," & Selection
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                            If �[��1 = �[��2.Offset(myStep, 0) Then
                                �㉺�ɐi��flg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'to�̍s�܂ŏ�܂��͉��ɐi��
                    If (�[��1.Column = 1 Or �㉺�ɐi��flg = 1) And �[��1.Row <> �[��to.Row Then

                        Set �[��2 = �[��1.Offset(myStep, 0)
                        If �[��1 <> �[��2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                        If �[��1 <> �[��2.Offset(myStep, 0) Then
                            �㉺�ɐi��flg = 0
                        End If
                        'If �㉺�ɐi��flg = 1 Then GoTo line15
                    End If
                    
                    'to�̍s���E�ɐi��
                    If �[��1.Row = �[��to.Row Then
                        Do Until �[��1.Column = �[��to.Column
                            Set �[��2 = �[��1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Select
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set �[��1 = �[��2
                            If Left(�[��1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set �[��2 = Nothing
            Next i2
        Next i
                

               
        Dim SyTop As Long
        Dim flg As Long, �摜flg As Long, Sx As Long, Sy As Long
'        '�}����̋󂢂Ă���X�y�[�X�Ɉړ�
'        For Each objShp In WB.Sheets("���_" & ���type).Shapes
'            �摜flg = 0: SyTop = (�T�C�Yy * �{��y) + 5
'line20:
'            flg = 0
'            For Each objShp2 In WB.Sheets("���_" & ���type).Shapes
'                'If objShp.Name = "501!" And objShp2.Name = "843!" Then Stop
'                If Right(objShp.Name, 1) = "!" And Right(objShp2.Name, 1) = "!" Then
'                    If objShp.Name <> objShp2.Name Then
'                        �摜flg = 1
'                        For Sx = objShp.Left To objShp.Left + objShp.width Step 1
'                            If objShp2.Left <= Sx And objShp2.Left + objShp2.width >= Sx Then
'                                If objShp2.Top <= SyTop And objShp2.Top + objShp2.height >= SyTop Then
'                                    flg = 1
'                                    SyTop = SyTop + 10
'                                    GoTo line20
'                                End If
'                            End If
'                        Next Sx
'                    End If
'                End If
'            Next objShp2
'
'            If flg = 1 Then GoTo line20
'
'            If �摜flg = 1 Then
'                objShp.Top = SyTop
'            End If
'        Next objShp
                
       
line30:

        
        '.Cells(1, 1).Select
    End With
    If ���}�̂� = "1" Then

    End If

    Call �œK�����ǂ�

End Function



Public Function �݊����Z�o()
    
    Dim ���type As String: ���type = "C"
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    Dim newBookName As String
    
    Call �œK��
    Call ���i�i��RAN_set2(���i�i��RAN, ���type, "����", "")
    
    newBookName = Left(myBookName, InStrRev(myBookName, ".") - 1) & "_�݊���"
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(myBookpath & "\45_�݊���\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    
    '������ǂݎ���p�ŊJ��
    On Error Resume Next
    Workbooks.Open fileName:=Left(myBookpath, InStrRev(myBookpath, "\")) & "000_�V�X�e���p�[�c\����_�݊���.xlsx", ReadOnly:=True
    On Error GoTo 0
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks("����_�݊���.xlsx").SaveAs fileName:=myBookpath & "\45_�݊���\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
        
    For e1 = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2) - 1 '���i�i�Ԗ�
        '���i�i�Ԃ̃V�[�g��ǉ�
        Workbooks(newBookName).Sheets("Sheet1").Copy after:=Workbooks(newBookName).Sheets("Sheet1")
        With ActiveSheet
            .Name = Replace(���i�i��RAN(1, e1 + 1), " ", "")
            .Cells(1, 1) = newBookName
            .Cells(2, 1) = ���i�i��RAN(1, e1 + 1)
            For e3 = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2) - 1
                .Cells(4, e3 + 3) = ���i�i��RAN(1, e3 + 1)
                .Cells(5, e3 + 3) = Right(���i�i��RAN(1, e3 + 1), 3)
            Next e3
        End With
        
        With Workbooks(newBookName).Sheets(Replace(���i�i��RAN(1, e1 + 1).Value, " ", ""))
            For e2 = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2) - 1 '�Ώۂ̐��i�i�Ԗ�
    
                ���i�i��str0 = ���i�i��RAN(1, e1 + 1) & String(15 - Len(���i�i��RAN(1, e1 + 1)), " ") '���i�i��A
                ���i�i��str1 = ���i�i��RAN(1, e2 + 1) & String(15 - Len(���i�i��RAN(1, e2 + 1)), " ") '���i�i��B
                
                Call SQL_�݊��[��(�݊��[��0ran, ���i�i��str0, myBookName, ���type)          '���i�i��A�̒[�����Ǝ�����W��z��ɓ����
                Call SQL_�݊��[��cav_1998(�݊��[��cav0ran, �݊��[��0ran, ���i�i��str0, myBookName)
                
                Call SQL_�݊��[��(�݊��[��1RAN, ���i�i��str1, myBookName, ���type)
                Call SQL_�݊��[��cav_1998(�݊��[��cav1RAN, �݊��[��1RAN, ���i�i��str1, myBookName)
                
                '���W��cav��������
                Dim �n�_�}�b�`flg As Boolean, �I�_�}�b�`flg As Boolean
                For i = LBound(�݊��[��cav0ran, 2) To UBound(�݊��[��cav0ran, 2)     '�[����cav_���i�i��A
                    �n�_�}�b�`flg = False: �I�_�}�b�`flg = False
                    ���Wcav0 = �݊��[��cav0ran(4, i) & "_" & �݊��[��cav0ran(5, i)   '�����W
                    For p = LBound(�݊��[��cav1RAN, 2) To UBound(�݊��[��cav1RAN, 2)   '�[����cav_���i�i��B
                        ���Wcav1 = �݊��[��cav1RAN(4, p) & "_" & �݊��[��cav1RAN(5, p) '�����W
                        '��r
                        If ���Wcav0 = ���Wcav1 Then '�[����cav��������
                            For pp = LBound(�݊��[��0ran, 2) To UBound(�݊��[��0ran, 2)
                                If �݊��[��cav0ran(0, i) = "" Then Stop
                                If �݊��[��cav0ran(2, i) = Null Then Stop
                                '�n�_��
                                If �n�_�}�b�`flg = False Then
                                    If �݊��[��cav0ran(4, i) = �݊��[��0ran(1, pp) Then
                                        �݊��[��0ran(3, pp) = �݊��[��0ran(3, pp) + 1 '�[��cav�}�b�`�̃J�E���g
                                        �n�_�}�b�`flg = True
                                    End If
                                End If
                                '�I�_��
                                If �I�_�}�b�`flg = False Then
                                    If �݊��[��cav0ran(5, i) = �݊��[��0ran(1, pp) Then
                                        �݊��[��0ran(3, pp) = �݊��[��0ran(3, pp) + 1 '�[��cav�}�b�`�̃J�E���g
                                        �I�_�}�b�`flg = True
                                    End If
                                End If
                                If �n�_�}�b�`flg = True And �I�_�}�b�`flg = True Then GoTo line20
                            Next pp
                        End If
                    Next p
line20:
                Next i
            
                '�������W�͒[���𕡐��܂Ƃ߂�
                Dim cnt As Long
                For pp = LBound(�݊��[��0ran, 2) To UBound(�݊��[��0ran, 2)
                    For ppp = LBound(�݊��[��0ran, 2) To UBound(�݊��[��0ran, 2)
                        If �݊��[��0ran(1, pp) = �݊��[��0ran(1, ppp) Then
                            If �݊��[��0ran(1, pp) <> "" Then
                                If pp <> ppp Then
                                    �݊��[��0ran(0, pp) = �݊��[��0ran(0, pp) & "&" & �݊��[��0ran(0, ppp) '�����W
                                    �݊��[��0ran(2, pp) = (�݊��[��0ran(2, pp) + �݊��[��0ran(2, ppp)) '��cav��
                                    �݊��[��0ran(3, pp) = �݊��[��0ran(3, pp) + �݊��[��0ran(3, ppp) '�}�b�`��
                                    �݊��[��0ran(0, ppp) = ""
                                    �݊��[��0ran(1, ppp) = ""
                                    �݊��[��0ran(2, ppp) = ""
                                    �݊��[��0ran(3, ppp) = ""
                                End If
                            End If
                        End If
                    Next ppp
                Next pp
                
                '�V�[�g�ɏo��
                ��cav��total = 0: ���}�b�`��total = 0
                For pp = LBound(�݊��[��0ran, 2) To UBound(�݊��[��0ran, 2)
                    If �݊��[��0ran(0, pp) <> "" Then
                        cnt = 1: ��cav�� = 0: ���}�b�`�� = 0
                        For n = 1 To Len(�݊��[��0ran(0, pp))
                            If InStr(Mid(�݊��[��0ran(0, pp), n, 1), "&") > 0 Then cnt = cnt + 1
                        Next n
                        
                        Set myfind = .Columns(1).Find(�݊��[��0ran(0, pp), , , 1)
                        If myfind Is Nothing Then
                            addRow = .Cells(.Rows.count, 1).End(xlUp).Row + 1
                        Else
                            addRow = myfind.Row
                        End If
                        
                        ��cav�� = RoundUp(��cav�� + �݊��[��0ran(2, pp) / cnt, 0)
                        ���}�b�`�� = RoundUp(���}�b�`�� + �݊��[��0ran(3, pp) / cnt, 0)
                        keyCol = e2 + 3
                        .Cells(addRow, 1) = �݊��[��0ran(0, pp)
                        .Cells(addRow, 2) = �݊��[��0ran(1, pp)
                        .Cells(addRow, keyCol).NumberFormat = 0
                        .Cells(addRow, keyCol).Value = ���}�b�`��
                        Set Rng = .Cells(addRow, keyCol)
                        Rng.FormatConditions.Delete
                        Dim dBar As Databar
                        Set dBar = Rng.FormatConditions.AddDatabar
                        ' Set the endpoints for the data bars:
                        dBar.MinPoint.Modify xlConditionValueNumber, 0
                        dBar.MaxPoint.Modify xlConditionValueNumber, ��cav��
                        dBar.BarFillType = xlDataBarFillSolid
                        If ���}�b�`�� = ��cav�� Then
                            dBar.BarColor.color = RGB(200, 200, 255)
                        Else
                            dBar.BarColor.color = RGB(200, 200, 200)
                        End If
                        If e1 = e2 Then
                            dBar.BarColor.color = RGB(177, 160, 199)
                        End If
                        ���}�b�`��total = ���}�b�`��total + ���}�b�`��
                        ��cav��total = ��cav��total + ��cav��
                    End If
                Next pp
                .Cells(addRow + 1, keyCol).NumberFormat = 0
                .Cells(addRow + 1, keyCol) = ���}�b�`��total
                Set Rng = .Cells(addRow + 1, keyCol)
                Rng.FormatConditions.Delete
                Set dBar = Rng.FormatConditions.AddDatabar
                ' Set the endpoints for the data bars:
                dBar.MinPoint.Modify xlConditionValueNumber, 0
                dBar.MaxPoint.Modify xlConditionValueNumber, ��cav��total
                dBar.BarFillType = xlDataBarFillSolid
                If ���}�b�`��total = ��cav��total Then
                    dBar.BarColor.color = RGB(200, 200, 255)
                Else
                    dBar.BarColor.color = RGB(200, 200, 200)
                End If
                If e1 = e2 Then
                    dBar.BarColor.color = RGB(177, 160, 199)
                End If

            Next e2
            .Range(.Columns(3), .Columns(UBound(���i�i��RAN, 2) + 2)).ColumnWidth = 5
        End With
    Next e1
    
    Call �œK�����ǂ�
End Function

Sub ����΂�쐬()
    'Call ���i�i��RAN_set2
    
    For c = 1 To ���i�i��RANc
        With Sheets("�d�|������΂�")
            ���i�i�� = Replace(���i�i��RAN(1, c).Value, " ", "")
            If Len(���i�i��) = 8 Then
                barcode = "*" & ���i�i�� & "*"
                ���i�i�� = Left(���i�i��, 4) & "-" & Right(���i�i��, 4)
            Else
                barcode = "*" & Mid(���i�i��, 2, 8) & "*"
                ���i�i�� = Left(���i�i��, 5) & "-" & Right(���i�i��, 5)
            End If
            .Cells(3, 4) = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, "_") - 1)
            .Cells(9, 4) = Left(���i�i��, 5)
            .Cells(7, 9) = Mid(���i�i��, 6)
            .Cells(14, 4) = barcode
            .Cells(14, 8) = barcode
        End With
    Next c
End Sub

Sub �T�u�\���̐ڑ����m�F()

    Dim ���i�i��str As String: ���i�i��str = "821113B340"
    ���i�i��str = ���i�i��str + String(15 - Len(���i�i��str), " ")
    myBookName = ActiveWorkbook.Name
    
    Call SQL_�T�u�m�F_�d���ꗗ(�d��RAN, ���i�i��str, myBookName)
    
    Call SQL_�T�u�[����(�T�u�[����RAN, ���i�i��str, myBookName)
    
    Call SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, myBookName)
    
    Dim �T�u�ڑ��[��RAN()
    Dim �n�_flg As Boolean, �I�_flg As Boolean
    Dim j As Long: j = 0
    For i = LBound(�d��RAN, 2) To UBound(�d��RAN, 2)
        �n�_flg = False: �I�_flg = False
        For ii = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
            '�����̒[������null�̎�
            If IsNull(�d��RAN(3, i)) And IsNull(�d��RAN(5, i)) Then GoTo line20
            
            'PVSW�n�_��
            If �n�_flg = False Then
                If �d��RAN(2, i) & "_" & �d��RAN(3, i) = �[���ꗗran(0, ii) & "_" & �[���ꗗran(1, ii) Then
                    If �d��RAN(0, i) = �[���ꗗran(2, ii) Then '�����T�u���m�F
                        �n�_flg = True
                    End If
                End If
            End If
            'PVSW�I�_��
            If �I�_flg = False Then
                If �d��RAN(4, i) & "_" & �d��RAN(5, i) = �[���ꗗran(0, ii) & "_" & �[���ꗗran(1, ii) Then
                    If �d��RAN(0, i) = �[���ꗗran(2, ii) Then '�����T�u���m�F
                        �I�_flg = True
                    End If
                End If
            End If
            If �n�_flg = True And �I�_flg = True Then Exit For
        Next ii
        
        If �n�_flg = True And �I�_flg = True Then
            '�n�_
            ReDim Preserve �T�u�ڑ��[��RAN(1, j)
            For p = LBound(�T�u�ڑ��[��RAN, 2) To UBound(�T�u�ڑ��[��RAN, 2)
                If �T�u�ڑ��[��RAN(0, p) = �d��RAN(0, i) Then
                    If �T�u�ڑ��[��RAN(1, p) = �d��RAN(3, i) Then
                        GoTo line10
                    End If
                End If
            Next p
            '�����̂Œǉ�
            �T�u�ڑ��[��RAN(0, j) = �d��RAN(0, i)
            �T�u�ڑ��[��RAN(1, j) = �d��RAN(3, i)
            j = j + 1
line10:
            '�I�_
            ReDim Preserve �T�u�ڑ��[��RAN(1, j)
            For p = LBound(�T�u�ڑ��[��RAN, 2) To UBound(�T�u�ڑ��[��RAN, 2)
                If �T�u�ڑ��[��RAN(0, p) = �d��RAN(0, i) Then
                    If �T�u�ڑ��[��RAN(1, p) = �d��RAN(5, i) Then
                        GoTo line15
                    End If
                End If
            Next p
            '�����̂Œǉ�
            �T�u�ڑ��[��RAN(0, j) = �d��RAN(0, i)
            �T�u�ڑ��[��RAN(1, j) = �d��RAN(5, i)
            j = j + 1
line15:
        End If
        
        If �n�_flg = False And �I�_flg = False Then
            �������JCDF = �d��RAN(6, i) & �d��RAN(7, i) & �d��RAN(8, i)
            �q����Ȃ��d�� = �q����Ȃ��d�� & �d��RAN(1, i) & "  " & �������JCDF & vbCrLf
        End If
line20:
    Next i
    
    '�[���ꗗ����T�u�ڑ��[�����Q�Ƃ��Ė�����Όq����Ȃ��[���Ƃ���
    If j > 0 Then
        For ii = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
            For iii = LBound(�T�u�ڑ��[��RAN, 2) To UBound(�T�u�ڑ��[��RAN, 2)
                If �[���ꗗran(2, ii) = �T�u�ڑ��[��RAN(0, iii) Then
                    If �[���ꗗran(1, ii) = �T�u�ڑ��[��RAN(1, iii) Then
                    GoTo line30
                    End If
                End If
            Next iii
            '�����̂Œǉ�
            '�T�u�̒[�����𒲂ׂ�
            For b = LBound(�T�u�[����RAN, 2) To UBound(�T�u�[����RAN, 2)
                If �[���ꗗran(2, ii) = �T�u�[����RAN(0, b) Then
                    �q����Ȃ��[�� = �q����Ȃ��[�� & �[���ꗗran(2, ii) & String(5 - Len(�[���ꗗran(2, ii)), " ") & _
                                     �[���ꗗran(1, ii) & String(5 - Len(�[���ꗗran(1, ii)), " ") & _
                                     String(3 - Len(�T�u�[����RAN(1, b)), " ") & �T�u�[����RAN(1, b) & _
                                     "  " & �[���ꗗran(0, ii) & _
                                     vbCrLf
                End If
            Next b
line30:
        Next ii
    End If
    
    Debug.Print vbCrLf & ���i�i��str
    Debug.Print "�q����Ȃ��[��_��" & "�T�u,�[��,�T�u�̒[����" & vbCrLf & �q����Ȃ��[��
    Debug.Print "�q����Ȃ��d��_ |" & vbCrLf & �q����Ȃ��d��

End Sub

Public Function �n���}�쐬_�w��()
    Dim sTime As Single: sTime = Timer
    Debug.Print "0= " & Round(Timer - sTime, 2)

    Call �œK��
    
    ��n�����i�i�� = "" '�w�肵���琻�i�g�������쐬���Ȃ�_���̐��i�i�Ԃ̒l���g�p���Ă��Ȃ�
    
    �[�� = ""
    
    �n���}�^�C�v = "0" '0:�쐬���Ȃ� or ��� or �`�F�b�J�[�p or ��H���� or �\��
    Dim ����� As String: ����� = "����" '3:�����A4:����
    Dim ����G As String: ����G = "A~"
    �n���\�� = "1" '0:�����A1:��n���}�A2:��n���}
    �������i = "0" '0:�\�����Ȃ��A40:��n�����i�A50:��n�����i
    Dim ��ƕ\���ϊ� As String: ��ƕ\���ϊ� = "1" '0:�ϊ����Ȃ��A1:�T�C�Y����ƕ\���L���ɕϊ�����
    Dim �n���}�D�� As String: �n���}�D�� = "�ʐ^" '�ʐ^ or ���}_�ʐ^���������͗��}��T��
    Dim �{�����[�h As Long: �{�����[�h = "1" '0:�����{,1:��{
    
    myFont = "�l�r �S�V�b�N"
    Dim minW�w�� As Long
    Dim minH As Single: minH = -1
    Dim X, Y, w, h, minW As Single: minW = -1
    Select Case �n���}�^�C�v
    Case "�`�F�b�J�[�p"
        minW�w�� = 24
    Case "��H����", "�\��"
        minW�w�� = 28
    Case Else
        minW�w�� = 18
    End Select
    
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    
    Call ���i�i��RAN_set2(���i�i��RAN, ����G, �����, ��n�����i�i��)
    
    Dim ws As Worksheet
    Dim i As Long
    
    Debug.Print "1= " & Round(Timer - sTime, 2): sTime = Timer
    
    Call SQL_�n���}�쐬_1(���i�i��RAN, �n���}�쐬RAN, �[��, myBook, newSheet)
    
    With newSheet
        Dim �\��Col As Long: �\��Col = .Rows(1).Find("�\��_", , , 1).Column
        Dim �D��1 As Long: �D��1 = .Rows(1).Find("�[�����ʎq", , , 1).Column
        Dim �D��2 As Long: �D��2 = .Rows(1).Find("�[�����i��", , , 1).Column
        Dim �D��3 As Long: �D��3 = .Rows(1).Find("�L���r�e�B", , , 1).Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �\��Col).End(xlUp).Row
        Call �\�[�g0(newSheet, 2, lastRow, �D��1, �D��2, �D��3)
    End With
    
    '��Cav��ǉ�
    With newSheet
        Dim �[��Col As Long: �[��Col = .Rows(1).Find("�[�����ʎq", , , 1).Column
        Dim cavCol As Long: cavCol = .Rows(1).Find("�L���r�e�B", , , 1).Column
        Dim ���Col As Long: ���Col = .Rows(1).Find("�[�����i��", , , 1).Column
        Dim �ɐ�Col As Long: �ɐ�Col = .Rows(1).Find("�R�l�N�^�ɐ�_", , , 1).Column
        Dim aRow As Long: aRow = 2
        lastRow = .Cells(.Rows.count, �\��Col).End(xlUp).Row
        Dim addRow As Long: addRow = lastRow
        For i = 2 To lastRow
            If .Cells(i, �[��Col) & "_" & .Cells(i, ���Col) <> .Cells(i + 1, �[��Col) & "_" & .Cells(i + 1, ���Col) Then
                addrows = ""
                �ɐ� = .Cells(i, �ɐ�Col)
                If �ɐ� = "" Then �ɐ� = 1
                For p = 1 To �ɐ�
                    For i2 = aRow To i
                        If CStr(p) = .Cells(i2, cavCol) Then GoTo line10
                    Next i2
                    addRow = addRow + 1
                    .Cells(addRow, �[��Col) = .Cells(i, �[��Col)
                    .Cells(addRow, ���Col) = .Cells(i, ���Col)
                    .Cells(addRow, cavCol) = p
                    If addrows = "" Then addrows = addRow
line10:
                Next p
                    '���i�i�ԂŎg�p������΃T�u��0��t����
                    If addrows <> "" Then
                        Dim jj As Range
                        For X = 1 To ���i�i��RANc
                            Set jj = .Range(.Cells(aRow, X), .Cells(i, X))
                            If WorksheetFunction.CountA(jj) > 0 Then
                                .Range(.Cells(addrows, X), .Cells(addRow, X)) = "0"
                            End If
                        Next X
                    End If
                aRow = i + 1
            End If
        Next i
    End With
    
    With newSheet
        �D��1 = .Rows(1).Find("�[�����ʎq", , , 1).Column
        �D��2 = .Rows(1).Find("�[�����i��", , , 1).Column
        �D��3 = .Rows(1).Find("�L���r�e�B", , , 1).Column
        Call �\�[�g0(newSheet, 2, addRow, �D��1, �D��2, �D��3)
    End With
    
    Stop
    
    Debug.Print "2= " & Round(Timer - sTime, 2): sTime = Timer
    
    '���W�f�[�^�̎擾
    Call SQL_CAV���W�擾(���i�i��RAN, myBook, newSheet)
    
    'Call SQL_�n���}�쐬_2(���i�i��RAN, myBook, newSheet)
    
    Debug.Print "���W�f�[�^�̎擾 " & Round(Timer - sTime, 2): sTime = Timer
    
    '���[�N�V�[�g�̒ǉ�
    For Each ws In Worksheets
        If ws.Name = "�n���}_" & ����� & "_" & ����G Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet2 = Worksheets.add(after:=newSheet)
    newSheet2.Name = "�n���}_" & ����� & "_" & ����G
    newSheet2.Cells.NumberFormat = "@"
        
    With newSheet
        '����0:�d���f�[�^
        Dim ����0 As String: Dim ����0Col() As Long
        ����0 = "�[�����i��,�\��_,�i��_,�T��_,�F��_,�[�����ʎq,�L���r�e�B,��H����,��_,�}_,����_,��"
        ����0s = Split(����0, ",")
        ReDim ����0Col(UBound(����0s))
        For i = LBound(����0s) To UBound(����0s)
            ����0Col(i) = .Rows(1).Find(����0s(i), , , 1).Column
        Next i
        
        '����1:�n���}�����ʉ����������
        Dim ����1 As String: Dim ����1Col() As Long
        ����1 = "�[�����i��,�\��_,�T��_,�F��_,�[�����ʎq,�L���r�e�B,��H����,��_,�}_"
        ����1s = Split(����1, ",")
        ReDim ����1Col(UBound(����1s))
        For i = LBound(����1s) To UBound(����1s)
            ����1Col(i) = .Rows(1).Find(����1s(i), , , 1).Column
        Next i
        
        '�i�[����z��
        Dim �f�[�^() As String: ReDim �f�[�^(1, 1, 0) '���i�g����,����,�l
        Dim j As Long:
        
        '�n���}temp���Q�Ƃ��鍀��
        �\��Col = .Rows(1).Find("�\��_", , , 1).Column
        �[��Col = .Rows(1).Find("�[�����ʎq", , , 1).Column
        ���Col = .Rows(1).Find("�[�����i��", , , 1).Column
        h = .Rows(1).Find("Height", , , 1).Column
        w = .Rows(1).Find("Width", , , 1).Column
        If h <> 0 Then If w < minH Or minH = -1 Then minH = h
        If w <> 0 Then If w < minW Or minW = -1 Then minW = w
        lastRow = .Cells(.Rows.count, �\��Col).End(xlUp).Row
        addRow = 3
        For i = 2 To lastRow
            ReDim Preserve �f�[�^(1, 1, j)
            For D = 1 To ���i�i��RANc
                �f�[�^(1, 0, j) = �f�[�^(1, 0, j) & "," & .Cells(i, D)
                �f�[�^(1, 1, j) = �f�[�^(1, 1, j) & "," & .Cells(i, D)
            Next D
            �f�[�^(1, 0, j) = Right(�f�[�^(1, 0, j), Len(�f�[�^(1, 0, j)) - 1)
            �f�[�^(1, 1, j) = Right(�f�[�^(1, 1, j), Len(�f�[�^(1, 1, j)) - 1)
            
            For D = LBound(����0s) To UBound(����0s)
                �f�[�^(0, 0, j) = �f�[�^(0, 0, j) & "," & .Cells(i, ����0Col(D))
            Next D
            �f�[�^(0, 0, j) = Right(�f�[�^(0, 0, j), Len(�f�[�^(0, 0, j)) - 1)
            
            For D = LBound(����1s) To UBound(����1s)
                �f�[�^(0, 1, j) = �f�[�^(0, 1, j) & "," & .Cells(i, ����1Col(D))
            Next D
            �f�[�^(0, 1, j) = Right(�f�[�^(0, 1, j), Len(�f�[�^(0, 1, j)) - 1)
            j = j + 1
            
            '�[��,�[�����i�Ԃ����s�ňقȂ鎞�A�d���f�[�^���o�͂��ăn���}�쐬
            If .Cells(i, �[��Col) & "_" & .Cells(i, ���Col) <> .Cells(i + 1, �[��Col) & "_" & .Cells(i + 1, ���Col) Then
                '�������Ȃ琻�i�g����������,�l��""
                For D = LBound(�f�[�^, 3) To UBound(�f�[�^, 3)
                    For d2 = D To UBound(�f�[�^, 3)
                        If D <> d2 Then
                            '�d���f�[�^
                            If �f�[�^(0, 0, D) = �f�[�^(0, 0, d2) Then
                                �f�[�^(1, 0, D) = ���i�g��������(�f�[�^(1, 0, D), �f�[�^(1, 0, d2))
                                �f�[�^(1, 0, d2) = ""
                                �f�[�^(0, 0, d2) = ""
                            End If
                            '�n���}�쐬�f�[�^
                            If �f�[�^(0, 1, D) = �f�[�^(0, 1, d2) Then
                                �f�[�^(1, 1, D) = ���i�g��������(�f�[�^(1, 1, D), �f�[�^(1, 1, d2))
                                �f�[�^(1, 1, d2) = ""
                                �f�[�^(0, 1, d2) = ""
                            End If
                        End If
                    Next d2
                Next D
                '�d���f�[�^�̏o��
                With newSheet2
                    '�t�B�[���h��
                    If addRow = 3 Then
                        For p = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
                            If Left(���i�i��RAN(1, p), 7) <> strbak Then
                                .Cells(addRow - 1, p + 1) = Left(���i�i��RAN(1, p), 7)
                            End If
                            .Cells(addRow - 0, p + 1) = Mid(���i�i��RAN(1, p), 8, 3)
                            .Columns(p + 1).ColumnWidth = 3.2
                            strbak = Left(���i�i��RAN(1, p), 7)
                        Next p
                        .Range(.Cells(addRow, ���i�i��RANc + 1), .Cells(addRow, ���i�i��RANc + UBound(����0s))) = Split(����0, ",")
                        .Range(.Cells(addRow, ���i�i��RANc + 1), .Cells(addRow, ���i�i��RANc + UBound(����0s))).Columns.AutoFit
                        addRow = addRow + 1
                    End If
                    '�d���f�[�^
                    For D = LBound(�f�[�^, 3) To UBound(�f�[�^, 3)
                        If �f�[�^(0, 0, D) <> "" Then
                            .Range(.Cells(addRow, 1), .Cells(addRow, ���i�i��RANc)) = Split(�f�[�^(1, 0, D), ",")
                            .Range(.Cells(addRow, ���i�i��RANc + 1), .Cells(addRow, ���i�i��RANc + UBound(����0s))) = Split(�f�[�^(0, 0, D), ",")
                            addRow = addRow + 1
                        End If
                    Next D
                    '�n���}�̐��i�g�����p�^�[����g�����ɓ����
                    �g���� = �z������ւ���(�f�[�^)
                    For D = LBound(�g����, 2) To UBound(�g����, 2)
                        If �g����(1, D) <> "" Then

                            �g����s = Split(�g����(0, D), ",")
                            Stop
                            For p = LBound(�g����s) To UBound(�g����s)
                                Stop
                                '�摜�̔z�u�Ƃ�
                                If p = 0 Then
                                    �f�[�^s = Split(�f�[�^(0, 1, 0), ",")
                                    �[�����i�� = �f�[�^s(0)
                                    If Len(�[�����i��) = 8 Then
                                        �[�����i�� = Left(�[�����i��, 4) & "-" & Mid(�[�����i��, 5, 4)
                                    ElseIf Len(�[�����i��) = 10 Then
                                        �[�����i�� = Left(�[�����i��, 4) & "-" & Mid(�[�����i��, 5, 4) & "-" & Mid(�[�����i��, 9, 2)
                                    Else
                                        Stop
                                    End If
                                    �[�� = �f�[�^s(4)
                                    '�ʐ^��T��
                                    �n���}URL = �n���}�A�h���X & "\" & �[�����i�� & "_1_001.png"
                                    If Dir(�n���}URL) = "" Then
                                        '���}��T��
                                        �n���}URL = Left(�n���}�A�h���X, InStrRev(�n���}�A�h���X, "_") - 1) & "_���}\" & �[�����i�� & "_0_001.emf"
                                        If Dir(�n���}URL) = "" Then GoTo line20
                                    End If
                                    '���i�g������2�i���ɂ���
                                    myBIN = ""
                                    For e = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
                                        If InStr(�g����(1, D), ���i�i��RAN(1, e)) > 0 Then
                                            myBIN = myBIN & "1"
                                        Else
                                            myBIN = myBIN & "0"
                                        End If
                                    Next e
                                    '�g������16�i���ɕϊ�
                                    myHEX = BIN2HEX(myBIN)
                                    �[���} = �[�� & "_" & myHEX
                                    '�摜�̔z�u
                                    With .Pictures.Insert(�n���}URL)
                                        .Name = �[���}
                                        .ShapeRange(�[���}).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft
                                        If �{�����[�h = "1" Then
                                            If minW < minH Then
                                                my�� = (minW�w�� / minW)
                                            Else
                                                my�� = (minW�w�� / minH)
                                            End If
                                            If �`�� = "Cir" Then my�� = my�� * 1.2
                                        Else
                                            my�� = .Width / (.Width / 3.08) * ��
                                            my�� = my�� / .Width * �{��
                                        End If
                                        .ShapeRange(�[���}).ScaleHeight my��, msorue, msoScaleFromTopLeft
                                        .CopyPicture
                                        .Delete
                                    End With
                                    Sleep 1
                                    .Paste
                                    Selection.Name = �[���}
                                    Stop
                                End If
                            Next p
                        End If
line20:
                    Next D
                End With
                ReDim �f�[�^(1, 1, 0)
                j = 0
                addRow = addRow + 1
            End If
        Next i
    End With
    
End Function

Sub �T�u�ꗗ�\�̍쐬()
    
    ����� = ""
    ���type = ""
    
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    a = InStr(myBook.Name, "_")
    b = InStrRev(myBook.Name, ".")
    Dim newBookName As String: newBookName = "�T�u�ꗗ�\_" & Replace(Mid(myBook.Name, a + 1, b - a - 1), ���type, "") & ����� & ���type
    Dim newDir As String: newDir = "51_�T�u�ꗗ�\"
    Dim newPath As String
    
    Call ���i�i��RAN_set2(���i�i��RAN, �����, ���type, "")
    
    '�o�̓t�H���_�m�F
    If Dir(myBook.Path & "\" & newDir, vbDirectory) = "" Then
        MkDir (myBook.Path & "\" & newDir)
    End If
    
    '�o�̓t�@�C���A�Ԋm�F
    Dim �A�� As Long: �A�� = 0
    Do
        newPath = myBook.Path & "\" & newDir & "\" & newBookName & "_" & Format(�A��, "000") & Mid(myBook.Name, InStrRev(myBook.Name, "."))
        If Dir(newPath) = "" Then
            Exit Do
        End If
        If �A�� = 999 Then Stop '���߂�
        �A�� = �A�� + 1
    Loop
    
    '�o�̓t�@�C���쐬
    With Workbooks.add
        Set newBook = ActiveWorkbook
        Application.DisplayAlerts = False
        .SaveAs newPath, xlOpenXMLWorkbookMacroEnabled 'xlsm
        Application.DisplayAlerts = True
    End With
    
    '���i�ʒ[���ꗗ����擾
    Dim �T�uRAN() As String, �T�uRANc As Long
    ReDim �T�uRAN(�T�uRANc)
    Dim ���i�T�uRAN() As String, ���i�T�uRANc As Long
    ReDim ���i�T�uRAN(1, ���i�T�uRANc)
    With myBook.Sheets("�[���ꗗ")
        Dim key As Range: Set key = .Cells.Find("�[����", , , 1)
        Dim lastRow As Long: lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = key.End(xlToRight).Column
        
        Dim fndCol As Long, flg As Boolean
        For p = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            ���i�i��v = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), p)
            fndCol = .Rows(key.Row).Find(���i�i��v, , , 1).Column
            For Y = key.Row + 1 To lastRow
                �T�u = CStr(.Cells(Y, fndCol))
                If �T�u <> "" Then
                    '�o�^�����邩�m�F_�T�uRAN
                    flg = False
                    For r = LBound(�T�uRAN) To UBound(�T�uRAN)
                        If �T�uRAN(r) = �T�u Then flg = True: Exit For
                    Next r
                    If flg = False Then
                        ReDim Preserve �T�uRAN(�T�uRANc)
                        �T�uRAN(�T�uRANc) = �T�u
                        �T�uRANc = �T�uRANc + 1
                    End If
                    '�o�^�����邩�m�F_���i�T�uRAN
                    flg = False
                    For r = LBound(���i�T�uRAN, 2) To UBound(���i�T�uRAN, 2)
                        If ���i�T�uRAN(0, r) = ���i�i��v And ���i�T�uRAN(1, r) = �T�u Then flg = True: Exit For
                    Next r
                    If flg = False Then
                        ReDim Preserve ���i�T�uRAN(1, ���i�T�uRANc)
                        ���i�T�uRAN(0, ���i�T�uRANc) = ���i�i��v
                        ���i�T�uRAN(1, ���i�T�uRANc) = �T�u
                        ���i�T�uRANc = ���i�T�uRANc + 1
                    End If
                End If
            Next Y
        Next p
    End With
    
    '�o��
    With newBook.Sheets(1)
        .Range("a1") = newBookName
        .Range("a2") = "�o�͌��t�@�C��= " & myBook.Name
        .Range("a2").Font.Size = 10
        .Range("a5") = "�T�u��"
        .Columns(1).ColumnWidth = 5.5
        .Cells.NumberFormat = "@"
        Set �T�u�͈� = .Range(Rows(6), Rows(UBound(�T�uRAN) + 6))
        '�T�uRAN�̏o��
        For Y = LBound(�T�uRAN) To UBound(�T�uRAN)
            .Cells(Y + 6, 1) = CStr(�T�uRAN(Y))
        Next Y
        .Cells(UBound(�T�uRAN) + 7, 1) = "total"
        '�T�u���̕��ёւ�
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange �T�u�͈�
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '���i�T�uRAN�̏o��
        Dim xFnd As Object
        X = 1
        For s = LBound(���i�T�uRAN, 2) To UBound(���i�T�uRAN, 2)
            ���i�i�� = ���i�T�uRAN(0, s)
            Set xFnd = .Rows(4).Find(���i�i��, , , 1)
            '�V�������i�i��
            If xFnd Is Nothing Then
                X = X + 1
                .Cells(4, X) = ���i�i��
                .Cells(5, X) = Mid(���i�i��, 8, 3)
                .Columns(X).ColumnWidth = 3.6
                .Range(.Cells(6, X), .Cells(UBound(�T�uRAN) + 7, X)).Interior.color = RGB(200, 200, 200)
                .Range(.Cells(5, X), .Cells(UBound(�T�uRAN) + 7, X)).HorizontalAlignment = xlCenter
                �T�uc = 0
            End If
            �T�u = ���i�T�uRAN(1, s)
            yfnd = �T�u�͈�.Find(�T�u, , , 1).Row
            .Cells(yfnd, X) = �T�u
            �T�uc = �T�uc + 1
            .Cells(UBound(�T�uRAN) + 7, X) = �T�uc
            .Cells(yfnd, X).Interior.Pattern = xlNone
        Next s
        '�r��
        .Range(.Cells(5, 1), .Cells(UBound(�T�uRAN) + 6, X)).Borders.LineStyle = True
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0.8)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .Zoom = 100
'            .PaperSize = �v�����g�T�C�Y
'            .Orientation = �v�����g�z�E�R�E
            .LeftFooter = "&L" & "&11 " & ActiveWorkbook.FullName
'            .PageSetup.RightHeader = "&R" & "&14 " & my���i�i��(0) & "&14 ���  " & "&P/&N"
        End With
            newBook.Save
    End With
End Sub

Public Sub PVSW_RLTF�̃T�u0�ɑ����i�̃T�u�����蓖�Ă�_2047()

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    
    With myBook.Sheets(mySheetName)
        Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For X = 1 To key.Column - 1
            For Y = key.Row + 1 To lastRow
                If .Cells(Y, X) = "0" Then
                    .Cells(Y, X).Select
                    ��� = ""
                    For x2 = 1 To key.Column - 1 '���̍s�̑��T�u�i���o�[���S�ē����Ȃ炻�̃T�u�i���o�[�����蓖�Ă�
                        If .Cells(Y, x2) <> "" And .Cells(Y, x2) <> "0" Then
                            If X <> x2 Then
                                ���A = .Cells(Y, x2)
                                If ��� = "" Or ��� = .Cells(Y, x2) Then
                                    ��� = ���A
                                Else
                                    ��� = ""
                                    Exit For
                                End If
                            End If
                        End If
                    Next x2
                    If ��� <> "" Then .Cells(Y, X) = ���
                End If
            Next Y
        Next X
    End With
    
    Set myBook = Nothing
End Sub
Public Sub PVSW_RLTF�̃T�u0�ɑ����i�̃T�u�����蓖�Ă�_2048()

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    
    With myBook.Sheets(mySheetName)
        Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim ��r(5) As Long, ��� As String, ��rr As Boolean
        ��r(0) = .Cells.Find("�n�_����H����", , , 1).Column
        ��r(1) = .Cells.Find("�I�_����H����", , , 1).Column
        ��r(2) = .Cells.Find("�n�_���[�����ʎq", , , 1).Column
        ��r(3) = .Cells.Find("�I�_���[�����ʎq", , , 1).Column
        ��r(4) = .Cells.Find("�n�_���L���r�e�B", , , 1).Column
        ��r(5) = .Cells.Find("�I�_���L���r�e�B", , , 1).Column
        
        For X = 1 To key.Column - 1
            For Y = key.Row + 1 To lastRow
                If .Cells(Y, X) = "0" Then
                    .Cells(Y, X).Select
                    ��� = ""
                        For Y2 = key.Row + 1 To lastRow
                            ��rr = False
                            For h = LBound(��r) To UBound(��r)
                                If CStr(.Cells(Y, ��r(h))) <> CStr(.Cells(Y2, ��r(h))) Then
                                    ��rr = True
                                End If
                            Next h
                            If ��rr = True Then GoTo Next_y2

                            For x2 = 1 To key.Column - 1
                                If X = x2 Then GoTo Next_x2
                                ���A = .Cells(Y2, x2)
                                If ���A = "" Or ���A = "0" Then GoTo Next_x2
                                If ��� = "" Or ��� = ���A Then
                                    ��� = ���A
                                Else
                                    ��� = ""
                                    GoTo result
                                End If
Next_x2:
                            Next x2
Next_y2:
                        Next Y2
result:
                    If ��� <> "" Then .Cells(Y, X) = ���
                End If
Next_y:
            Next Y
        Next X
    End With
    
    Set myBook = Nothing
End Sub

Public Function ���V�[�g�̍쐬()
    Call �A�h���X�Z�b�g(myBook)
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim myCount As Long
    Dim myMessage As String
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    
    For i = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
        newSheetName = "���_" & ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), i)
        '�������O�̃t�@�C�������邩�m�F
        Dim ws As Worksheet
        flg = False
        For Each ws In Worksheets
            If ws.Name = newSheetName Then
                flg = True
                Exit For
            End If
        Next ws
        
        If flg = True Then GoTo next_I
            
        Dim newSheet As Worksheet
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Tab.color = 14470546
        newSheet.Cells.NumberFormat = "@"
        newSheet.Cells(1, 1).Value = "Size_"
        newSheet.Cells(1, 1).AddComment
        newSheet.Cells(1, 1).Comment.Text "Ctrl+ENTER�Ŗ��}�̍쐬"
        newSheet.Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
        newSheet.Cells(1, 1).Interior.color = RGB(255, 255, 0)
        newSheet.Cells(1, 2).Value = "1000_300"
        newSheet.Cells(1, 3).Value = "k_"
        newSheet.Cells(1, 3).AddComment
        newSheet.Cells(1, 3).Comment.Text "����̂Ȃ��ڂ̃��C��"
        newSheet.Cells(1, 3).Comment.Shape.TextFrame.AutoSize = True
        newSheet.Cells(1, 3).Interior.color = RGB(255, 255, 0)
        newSheet.Cells(1, 4).Value = "100.1"
        newSheet.Cells(1, 5).Value = "Width_"
        newSheet.Cells(1, 5).AddComment
        newSheet.Cells(1, 5).Comment.Text "����̉���mm"
        newSheet.Cells(1, 5).Comment.Shape.TextFrame.AutoSize = True
        newSheet.Cells(1, 5).Interior.color = RGB(255, 255, 0)
        newSheet.Cells(1, 6).Value = "1800"
        myCount = myCount + 1
        '�C�x���g�̒ǉ�
        
line11:
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile �A�h���X(0) & "\onKey\000_CodeModule_���.txt"
        If Err.Number <> 0 Then GoTo line11
        On Error GoTo 0
        
        Application.OnKey "^{ENTER}", "�z���}�쐬"
        
        PlaySound "��������"
        Sleep 500
next_I:
    Next i
        If myCount > 0 Then
            myMessage = "�V�[�g��ǉ����܂���"
        Else
            myMessage = "�ǉ��V�[�g�͂���܂���ł����B"
        End If
        ���V�[�g�̍쐬 = myMessage
    
End Function

Public Sub ��H�}�g���N�X�쐬_������()
    
    �Ώ� = "�V"
    �V����r = True
    
    Call �œK��

    Set wb(0) = ThisWorkbook
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    
    
    Call ���i�i��RAN_set2(���i�i��RAN, "", "", "")
    
    '�������J��
    Set wb(1) = �����̐ݒ�(wb(0), "genshi\��H��ظ�.xlsx", "A1_��H�}�g���N�X", "��H��ظ�" & "_" & �Ώ�)
    Set ws(1) = wb(1).Worksheets("Sheet1")
    
    With ws(0)
        �Ώۗ� = "�\��_,SubNo,SubNo2,SubNo3,�����@,SSC,�i���_,�T��_,�F��_,��ID_,��ID_,����_,����_," & _
                 "�n�_����H����,�n�_���[�����ʎq,�n�_���}_," & _
                 "�I�_����H����,�I�_���[�����ʎq,�I�_���}_," & _
                 "�n�_���[�q_,�n�_�����i_," & _
                 "�I�_���[�q_,�I�_�����i_," & _
                 "�d�㐡�@_"
        �Ώۗ�sp = Split(�Ώۗ�, ",")
        
        Dim �Ώۗ�col() As Long
        ReDim �Ώۗ�col(UBound(�Ώۗ�sp))
        For X = LBound(�Ώۗ�sp) To UBound(�Ώۗ�sp)
            �Ώۗ�col(X) = .Rows(sikibetu.Row).Find(�Ώۗ�sp(X), , , 1).Column
        Next X
        Dim �V��n As Long, ��ԍ�n As Long, �V���A��n As Long, �Ԏ�n As Long, ����n As Long, ���C��n As Long
        Dim �N����An As Long, ��H��n As Long, ��H��An As Long, ��H��ABn As Long, ��H��Bn As Long
        ���C��n = ���i�i��RAN_read(���i�i��RAN, "���C���i��")
        �V��n = ���i�i��RAN_read(���i�i��RAN, "�V��")
        ��ԍ�n = ���i�i��RAN_read(���i�i��RAN, "��ԍ�")
        �V���A��n = ���i�i��RAN_read(���i�i��RAN, "�A��")
        �Ԏ�n = ���i�i��RAN_read(���i�i��RAN, "�Ԏ�")
        ����n = ���i�i��RAN_read(���i�i��RAN, "����")
        �N����An = ���i�i��RAN_read(���i�i��RAN, "�N����")
        ��H��An = ���i�i��RAN_read(���i�i��RAN, "��H��")
        ��H��ABn = ���i�i��RAN_read(���i�i��RAN, "��H��AB")
        ��H��Bn = ���i�i��RAN_read(���i�i��RAN, "��H��_")
        
        Set outkey = ws(1).Cells.Find("�@�@�@��H��(A/B�܂�)", , , 1)
        
        '�Ώۂ̐��i�i�Ԃ��J�E���g�A�Ԏ���擾
        Set addkey = ws(1).Cells.Find("CONP No", , , 1)
        typeCol = ws(1).Cells.Find("TYPE", , , 1).Column
        �d�㐡�@col = ws(1).Cells.Find("�d�㐡�@", , , 1).Column
        Dim �Ԏ�(1) As String, �Ԏ�str As String
        Dim ����(1) As String, ���i�i��bak(1) As String
        For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
            If ���i�i��RAN(�V��n, r) = �Ώ� Then
                �Ώ�count = �Ώ�count + 1
                If ���i�i��header = "" Or ���i�i��header <> Left(���i�i��RAN(���C��n, r), Len(���i�i��header)) Then
                    ���i�i��header = Replace(Replace(���i�i��RAN(���C��n, r), ���i�i��RAN(����n, r), ""), " ", "")
                    ws(1).Cells(addkey.Row - 1, �d�㐡�@col + �Ώ�count) = ���i�i��header
                    If �Ώ�count <> 1 Then ws(1).Cells(addkey.Row - 1, �d�㐡�@col + �Ώ�count).Borders(xlEdgeLeft).Weight = xlThin
                End If
                ws(1).Cells(addkey.Row, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(����n, r)
                ws(1).Cells(outkey.Row + 0, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(��H��An, r)
                ws(1).Cells(outkey.Row + 1, �d�㐡�@col + �Ώ�count) = ws(1).Cells(outkey.Row, �d�㐡�@col + �Ώ�count) - ws(1).Cells(outkey.Row + 2, �d�㐡�@col + �Ώ�count)
                ws(1).Cells(outkey.Row + 2, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(��H��ABn, r)
                ws(1).Cells(outkey.Row + 3, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(��H��Bn, r)
                �Ԏ�str = ���i�i��RAN(�Ԏ�n, r)
                If InStr(�Ԏ�(0), �Ԏ�str) = 0 Then �Ԏ�(0) = �Ԏ�(0) & "," & �Ԏ�str
                '��r�Ώۂ̏����o��
                If �V����r = True Then
                    ����(1) = ""
                    For rr = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                        If r <> rr Then
                            If ���i�i��RAN(�V���A��n, r) = ���i�i��RAN(�V���A��n, rr) Then
                                If �Ώ� = "�V" Then ��r�L�� = "��" Else ��r�L�� = "��"
                                �Ԏ�str = ���i�i��RAN(�Ԏ�n, rr)
                                If InStr(�Ԏ�(1), �Ԏ�str) = 0 Then �Ԏ�(1) = �Ԏ�(1) & "," & �Ԏ�str
                                ws(1).Cells(addkey.Row - 3, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(����n, rr)
                                ws(1).Cells(outkey.Row + 8, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(����n, rr)
                                ws(1).Cells(outkey.Row + 9, �d�㐡�@col + �Ώ�count) = ���i�i��RAN(��H��An, rr)
                            End If
                        End If
                    Next rr
                End If
                ���i�i��bak(0) = ���i�i��RAN(���C��n, r)
            End If
        Next r
        '�w�b�_�[�����o��
        �Ԏ�(0) = Mid(�Ԏ�(0), 2): �Ԏ�(1) = Mid(�Ԏ�(1), 2)
        ws(1).Cells.Find("�Ԏ� / CAR STYLE", , , 1).Offset(1, 0).Value = �Ԏ�(0)
        ws(1).Cells(addkey.Row - 3, �d�㐡�@col) = �Ԏ�(1) & "��"
            
        Dim ���i�i��array() As String, sub_bak(2) As String
        '[PVSW_RLTF]���o��
        lastRow = .Cells(.UsedRange.Rows.count, sikibetu.Column).Row
        For i = sikibetu.Row + 1 To lastRow
            '�ΏۂɎg�p����������z���1���i�[
            Dim flg As Boolean: flg = False: Dim c As Long: c = 0
            ReDim ���i�i��array(�Ώ�count - 1)
            For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
                If ���i�i��RAN(�V��n, r) = �Ώ� Then
                    If .Cells(i, ���i�i��RAN(��ԍ�n, r)) <> "" Then
                        flg = True
                        ���i�i��array(c) = "1"
                    End If
                    c = c + 1
                End If
            Next r
            '�g�p������Ώo�͂���
            If flg = True Then
                Dim addFlg As Boolean: addFlg = False
                addRow = ws(1).Cells(.Rows.count, typeCol).End(xlUp).Row + 1
                If addkey.Row + 1 = addRow Then addFlg = True '�擪��1�s�󂯂�
                For p = 0 To 2  '�e�T�u�i���o�[���قȂ�Ȃ�1�s�󂯂�
                    If sub_bak(p) <> "" And sub_bak(p) <> ws(0).Cells(i, �Ώۗ�col(p + 1)) Then addFlg = True
                    If addFlg = True Then Exit For
                Next p
                If addFlg = True Then addRow = addRow + 1
                '�d�����̏o��
                For X = LBound(�Ώۗ�sp) To UBound(�Ώۗ�sp)
                    ws(0).Cells(i, �Ώۗ�col(X)).Copy Destination:=ws(1).Cells(addRow, X + 1)
                    ws(1).Cells(addRow, X + 1) = ws(0).Cells(i, �Ώۗ�col(X)).Value '�o�͌������̏ꍇ���l��
                    'WS(1).Cells(addRow + 1, x + 1) = �Ώۗ�sp(x)
                Next X
                ws(1).Rows(addRow).ShrinkToFit = True '�k�����đS�̂�\��
                '���i�g�������o��
                ws(1).Range(Cells(addRow, �d�㐡�@col + 1), Cells(addRow, �d�㐡�@col + �Ώ�count)) = ���i�i��array
                sub_bak(0) = ws(0).Cells(i, �Ώۗ�col(1))
                sub_bak(1) = ws(0).Cells(i, �Ώۗ�col(2))
                sub_bak(2) = ws(0).Cells(i, �Ώۗ�col(3))
            End If
        Next i
    End With
    
    '�r��
    With ws(1)
        
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, UBound(�Ώۗ�sp) + c + 1)).Borders.Weight = xlThin
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 6)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 19)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 24)).Borders(xlEdgeRight).Weight = xlMedium
        '�s�v�ȍs���폜
        If addRow + 1 < outkey.Row - 1 Then
            .Range(.Rows(addRow + 1), .Rows(outkey.Row - 1)).Delete
        Else
            Stop '�s��������
        End If
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, UBound(�Ώۗ�sp) + c + 1)).Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    Call �œK�����ǂ�
    
End Sub

Public Function openMenu()
    UI_Menu.Show
End Function

Public Function ���[�J���T�u�i���o�[�̎擾()

    Unload UI_07

    Call �A�h���X�Z�b�g(myBook)
    Call ���i�i��RAN_set2(���i�i��RAN, "�^��", "", "")
    
    For r = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2) - 1
        Dim ���i�i��str As String
        ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r + 1)
        '�d���T�u�i���o�[
        Call SQL_���[�J���d���T�u�i���o�[�擾(RAN, ���i�i��RAN(1, r + 1))
        With myBook.Sheets("PVSW_RLTF")
            .Activate
            Dim myCol As Long, myRow As Long, myKey, lastRow As Long
            Set myKey = .Cells.Find(���i�i��str, , , 1)
            myCol = .Cells.Find("�d�����ʖ�", , , 1).Column
            lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
            For i = myKey.Row + 1 To lastRow
                If .Cells(i, myKey.Column) <> "" Then
                    �\��str = Left(.Cells(i, myCol), 4)
                    For Y = LBound(RAN, 2) To UBound(RAN, 2)
                        If �\��str = RAN(1, Y) Then
                            ActiveWindow.ScrollColumn = myKey.Column
                            ActiveWindow.ScrollRow = i
                            .Cells(i, myKey.Column) = RAN(2, Y)
                            DoEvents
                            Sleep 20
                            DoEvents
                            Exit For
                        End If
                    Next Y
                End If
            Next i
        End With
        '�[���T�u�i���o�[
        aa = SQL_���[�J���[���T�u�i���o�[�擾(RAN, ���i�i��RAN(1, r + 1))
        With myBook.Sheets("�[���ꗗ")
            .Activate
            Set myKey = .Cells.Find(���i�i��str, , , 1)
            myCol = .Cells.Find("�[�����i��", , , 1).Column
            Dim myCol2 As Long
            myCol2 = .Cells.Find("�[����", , , 1).Column
            lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
            For i = myKey.Row + 1 To lastRow
                If .Cells(i, myKey.Column) <> "" Then
                    findFlg = False
                    ���i�i��str = �[�����i�ԕϊ�(.Cells(i, myCol))
                    �[��str = .Cells(i, myCol2)
                    For Y = LBound(RAN, 2) To UBound(RAN, 2)
                        If ���i�i��str = RAN(3, Y) And �[��str = RAN(2, Y) Then
                            ActiveWindow.ScrollColumn = myKey.Column
                            ActiveWindow.ScrollRow = i
                            .Cells(i, myKey.Column) = RAN(4, Y)
                            DoEvents
                            Sleep 20
                            findFlg = True
                            Exit For
                        End If
                    Next Y
                    '��H�}�g���N�X�̓A�[�X�ƃ{���_�[�͕��i�i�Ԃ�������Ă��Ȃ��ׂ̎b�菈��_������Ȃ�������[���������ŒT��
                    If findFlg = False Then
                        For Y = LBound(RAN, 2) To UBound(RAN, 2)
                            If �[��str = RAN(2, Y) Then
                                ActiveWindow.ScrollColumn = myKey.Column
                                ActiveWindow.ScrollRow = i
                                .Cells(i, myKey.Column) = RAN(4, Y)
                                DoEvents
                                Sleep 10
                                Exit For
                            End If
                        Next Y
                    End If
                End If
            Next i
        End With
        
    Next r
    
    MsgBox "���[�J������T�u�i���o�[���擾���܂����B"
    
End Function
