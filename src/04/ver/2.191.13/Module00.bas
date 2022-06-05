Attribute VB_Name = "Module00"
Public Function BIN2HEX(myBIN)
    If Len(myBIN) Mod 4 > 0 Then
        myBIN = String(((Len(myBIN) \ 4) + 1) * 4 - Len(myBIN), "0") & myBIN
    End If
    
    For u = 1 To Len(myBIN) Step 4
        Select Case Mid(myBIN, u, 4)
            Case "0000"
            myHEX = myHEX & "0"
            Case "0001"
            myHEX = myHEX & "1"
            Case "0010"
            myHEX = myHEX & "2"
            Case "0011"
            myHEX = myHEX & "3"
            Case "0100"
            myHEX = myHEX & "4"
            Case "0101"
            myHEX = myHEX & "5"
            Case "0110"
            myHEX = myHEX & "6"
            Case "0111"
            myHEX = myHEX & "7"
            Case "1000"
            myHEX = myHEX & "8"
            Case "1001"
            myHEX = myHEX & "9"
            Case "1010"
            myHEX = myHEX & "A"
            Case "1011"
            myHEX = myHEX & "B"
            Case "1100"
            myHEX = myHEX & "C"
            Case "1110"
            myHEX = myHEX & "D"
            Case "1111"
            myHEX = myHEX & "F"
        End Select
    Next u
    BIN2HEX = myHEX
End Function

Public Function �����̐ݒ�(myBook, ����, �ۑ��t�H���_��, newBookName) As Workbook

    �g���q = Mid(����, InStrRev(����, "."))
    newBookName = Left(myBook.Name, InStrRev(myBook.Name, ".") - 1) & "_" & newBookName
    
    '�d�����Ȃ��t�@�C�����Ɍ��߂�
    For i = 0 To 999
        If Dir(wb(0).Path & "\" & �ۑ��t�H���_�� & "\" & newBookName & "_" & Format(i, "000") & �g���q) = "" Then
            newBookName = newBookName & "_" & Format(i, "000")
            Exit For
        End If
        If i = 999 Then Stop '�z�肵�Ă��Ȃ���
    Next i
    
    '������ǂݎ���p�ŊJ��
    On Error Resume Next
    Workbooks.Open fileName:=�A�h���X(0) & "\" & ����, ReadOnly:=True
    On Error GoTo 0
    
    '�������T�u�}�̃t�@�C�����ɕύX���ĕۑ�
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=wb(0).Path & "\" & �ۑ��t�H���_�� & "\" & newBookName & �g���q
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set �����̐ݒ� = ActiveWorkbook
End Function

Public Function �I�[�g�V�F�C�v�폜()
    '�I�[�g�V�F�C�v���폜
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        objShp.Delete
    Next objShp
End Function

Public Function ���i�g��������(�����O1, �����O2)
    �����O1s = Split(�����O1, ",")
    �����O2s = Split(�����O2, ",")
    
    For i = LBound(�����O1s) To UBound(�����O1s)
        If �����O1s(i) <> "" Then
            ������ = ������ & "," & �����O1s(i)
        Else
            ������ = ������ & "," & �����O2s(i)
        End If
    Next i
    
    ���i�g�������� = Right(������, Len(������) - 1)
    
End Function

Public Function ���}_�[���o�H�\��()
    Call �œK��
    Set myBook = ActiveWorkbook
    Dim �[��str As String
    �[��str = Application.Caller
    On Error Resume Next
    ActiveSheet.Shapes("�z��").Ungroup
    ActiveSheet.Shapes("���").Ungroup
    On Error GoTo 0
    Call SQL_�z���[���擾_�[���p�[��(�z���[��RAN, �[��str)

    For Each ob In ActiveSheet.Shapes
        If InStr(ob.Name, "!") Then
            ob.Delete
        Else
            If ob.Type = 1 Then
                ob.Line.ForeColor.RGB = RGB(0, 0, 0)
                ob.Fill.ForeColor.RGB = RGB(255, 255, 255)
            ElseIf ob.Type = 9 Then
                ob.Line.ForeColor.RGB = RGB(150, 150, 150)
            End If
        End If
    Next ob
    Dim �z��toStr As String
    With ActiveSheet
        '���I�������[���̐F�t��
        With .Shapes(�[��str)
            .Select
            .ZOrder msoBringToFront
            .Fill.ForeColor.RGB = RGB(255, 100, 100)
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame.Characters.Font.color = RGB(0, 0, 0)
            '.Line.Weight = 2
            myTop = Selection.Top
            myLeft = Selection.Left
            myHeight = Selection.Height
            myWidth = Selection.Width
        End With

        '���z������[���Ԃ̃��C���ɐF�t��
        Set �[��from = .Cells.Find(�[��str, , , 1)
        For i = LBound(�z���[��RAN) To UBound(�z���[��RAN)
            Dim myStep As Long
            �[��to = �z���[��RAN(i)
            If �[��to = "" Then GoTo nextI
            Set �z�� = .Cells.Find(�[��str, , , 1)
            If �z�� Is Nothing Then GoTo nextI
                Set �[��to = .Cells.Find(�z���[��RAN(i), , , 1)
                If �[��to Is Nothing Then GoTo nextI
                If �[��from.Row < �[��to.Row Then myStep = 1 Else myStep = -1
                ActiveSheet.Shapes(�[��to.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                ActiveSheet.Shapes(�[��to.Value).ZOrder msoBringToFront
                �z��toStr = �z��toStr & "," & �[��to.Value
                Set �[��1 = �[��from
                For Y = �[��from.Row To �[��to.Row Step myStep
                    'from���獶�ɐi��
                    Do Until �[��1.Column = 1
                        Set �[��2 = �[��1.Offset(0, -2)
                        On Error Resume Next
                            ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).ZOrder msoBringToFront
                            ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).ZOrder msoBringToFront
                        On Error GoTo 0
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                    Loop
                    
line15:
                    'to�̍s�܂ŏ�܂��͉��ɐi��
                    Do Until �[��1.Row = �[��to.Row
                        Set �[��2 = �[��1.Offset(myStep, 0)
                        If �[��1 <> �[��2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).ZOrder msoBringToFront
                                ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).ZOrder msoBringToFront
                            On Error GoTo 0
                        End If
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                    Loop
                    
                    'to�̍s���E�ɐi��
                    Do Until �[��1.Column = �[��to.Column
                        Set �[��2 = �[��1.Offset(0, 2)
                        On Error Resume Next
                            ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value & " to " & �[��2.Value).ZOrder msoBringToFront
                            ActiveSheet.Shapes(�[��2.Value & " to " & �[��1.Value).ZOrder msoBringToFront
                        On Error GoTo 0
                        Set �[��1 = �[��2
                        If Left(�[��1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(�[��1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(�[��1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                    Loop
                Next Y
                Set �[��2 = Nothing
nextI:
        Next i

        For Each ob In ActiveSheet.Shapes
            If ob.Type = 1 And ob.Name <> "��" Then
                ob.ZOrder msoBringToFront
            Else
                
            End If
        Next ob
        '���z�������n���d����\��
        Dim �z��toStrSp
        �z��toStrSp = Split(�z��toStr, ",")
        Dim �Fv As String, �Tv As String, �[��v As String, �}v As String, �n��v As String
        For ii = LBound(�z��toStrSp) + 1 To UBound(�z��toStrSp)
            �[��v = �z��toStrSp(ii) '�[��v=�s����
            Call SQL_�z���[���擾_�[���p��H(�z���[��RAN, �[��v, �[��str)
            For i = LBound(�z���[��RAN, 2) To UBound(�z���[��RAN, 2)
                �Fv = �z���[��RAN(2, i)
                If �Fv = "" Then Exit For
                �}v = �z���[��RAN(6, i)
                �Tv = �z���[��RAN(4, i)
                �n��v = �z���[��RAN(4, i)
                �\��v = �z���[��RAN(3, i)
                ���Oc = 0
                For Each objShp In ActiveSheet.Shapes
                    If objShp.Name = �[��v & "!" Then
                        ���Oc = ���Oc + 1
                    End If
                Next objShp
                    
                With .Shapes(�[��v)
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
                    Selection.ShapeRange.Name = �[��v & "!"
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
                        Selection.Name = �[��v & "!"
                    End If
                End With
            Next i
        Next ii
    End With
    Call �œK�����ǂ�
End Function

Public Function �z������ւ���(�f�[�^)
    '���i�i�Ԗ��̐��i�g�����ɒu��������_�T�u����1�ɒu��������
    Dim �z��() As String
    ReDim �z��(1, ���i�i��RANc - 1) '0:�d���g����,1:���i�g����
    
    For i = LBound(�f�[�^, 3) To UBound(�f�[�^, 3)
        �f�[�^s = Split(�f�[�^(1, 1, i), ",")
        For a = LBound(�f�[�^s) To UBound(�f�[�^s)
            If �f�[�^s(a) <> "" Then
                �z��(0, a) = �z��(0, a) & ",1"
            Else
                �z��(0, a) = �z��(0, a) & ",0"
            End If
        Next a
    Next i
    '�]����","���폜
    For i = LBound(�z��, 2) To UBound(�z��, 2)
        �z��(0, i) = Right(�z��(0, i), Len(�z��(0, i)) - 1)
    Next i
    '�d��������ΐ��i�i�Ԃ��Z�b�g����
    For i = LBound(�z��, 2) To UBound(�z��, 2)
        If InStr(�z��(0, i), "1") > 0 Then �z��(1, i) = ���i�i��RAN(1, i)
    Next i
    '�d���g�������������́A�Е����폜����
    For i = LBound(�z��, 2) To UBound(�z��, 2)
        If �z��(0, i) <> "0" Then
            For i2 = i To UBound(�z��, 2)
                If i <> i2 Then
                        If �z��(0, i) = �z��(0, i2) Then
                            �z��(0, i2) = ""
                            �z��(1, i) = �z��(1, i) & "," & �z��(1, i2)
                            �z��(1, i2) = ""
                        End If
                End If
            Next i2
        End If
    Next i
    �z������ւ��� = �z��
End Function

Public Function �\�[�g0(newSheet, startRow, lastRow, �D��1, �D��2, �D��3)
    '�\�[�g
    With newSheet
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, �D��1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, �D��3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange .Range(.Rows(startRow), .Rows(lastRow))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
End Function

Sub DeleteDefinedNames()
 
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        If n.MacroType = -4142 Then
            n.Delete
        End If
    Next
 
End Sub

Public Function ���i�i��RAN_read(���i�i��RAN, ���i�i��FIE)

    For i = LBound(���i�i��RAN, 1) To UBound(���i�i��RAN, 1)
        If ���i�i��RAN(i, 0) = ���i�i��FIE Then
            ���i�i��RAN_read = i
            Exit Function
        End If
    Next i

End Function
Public Function ���i�i��RAN_seek()
    For X = 1 To ���i�i��Rc
        If ���i�i��RAN(1, X) = "" Then Stop '���i�i�Ԃ��Z�b�g����ĂȂ��ƒT���Ȃ�
        For xx = 1 To ���i�i��RANc
            If ���i�i��RAN(1, X) = ���i�i��RAN(1, xx) Then
                For a = 1 To 10
                    ���i�i��RAN(a, X) = ���i�i��RAN(a, xx)
                Next a
                GoTo line10
            End If
        Next xx
        Stop '���i�i�Ԃ�������Ȃ�
line10:
    Next X
End Function
Public Function ProgressBar_ref(������ As String, �������e As String, step0T As Long, step0 As Long, Step1T As Long, Step1 As Long)
    With ProgressBar
        .Caption = "������ " & ������
        
        .ProgBar0.Max = step0T
        .ProgBar0.Value = step0
        .msg0.Caption = step0 & "/" & step0T & "  " & �������e
        
        .ProgBar1.Max = Step1T
        .ProgBar1.Value = Step1
        .msg1.Caption = Step1 & "/" & Step1T
        '.Repaint
        DoEvents
        'If .StopBtn.Value = True Then Stop
        
    End With
End Function
Public Function �R�����g�\���ؑ�()
    Dim �R�����g�\�� As Boolean
    With Sheets("�ݒ�")
        �R�����g�\�� = .Cells.Find("�R�����g�\���ؑ�", , , 1).Offset(0, 1).Value
    End With
    
    �R�����g�\�� = �R�����g�\�� + 1
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
        
    For Each cmt In ws.Comments
        cmt.Visible = �R�����g�\��
    Next cmt
    
    With Sheets("�ݒ�")
        .Cells.Find("�R�����g�\���ؑ�", , , 1).Offset(0, 1) = �R�����g�\��
    End With
End Function

Public Function �������܂���(Optional myBook)
    myBook.Activate
    'Set aa = ActiveSheet.Shapes.AddPicture("H:\�쐬���܂���.png", False, True, 0, 0, 164, 128)
    Set aa = ActiveSheet.Pictures.Insert(�A�h���X(0) & "\picture\�쐬���܂���.png")
    winW = Application.Width
    winH = Application.Height
    aa.Left = (winW - aa.Width) / 2
    aa.Top = (winH - aa.Height) / 2
    aa.OnAction = "����"
    PlaySound ("���񂹂�")
End Function

Public Function ����()
    Set myBook = ActiveWorkbook
    myme = Application.Caller
    ActiveSheet.Shapes(myme).Delete
    PlaySound ("�Ƃ���2")
    ActiveWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile �A�h���X(0) & "\OnKey\002_��A���쐬_�}���}.txt"
    Application.OnKey "^{ENTER}", "���_�A����_�}���}_Ver2002"
    Application.OnKey "^~", "���_�A����_�}���}_Ver2002"
End Function

Public Function �A�h���X�Z�b�g(myBook)
    If �A�h���X(0) = "" Or myBook Is Nothing Then Set myBook = ActiveWorkbook
    With myBook.Sheets("�ݒ�")
        �A�h���X(0) = .Cells.Find("�V�X�e���p�[�c_", , , 1).Offset(0, 1).Value
        �A�h���X(1) = .Cells.Find("���ވꗗ+_", , , 1).Offset(0, 1).Value
        �A�h���X(2) = .Cells.Find("subNo.txt", , , 1).Offset(0, 1).Value
    End With
    
End Function

Public Function �Q�ƕs������΂��̃t�H���_���쐬����()

    Call �A�h���X�Z�b�g(ActiveWorkbook)

    Dim Ref, buf As String, bufS, myCount As Long
    Dim myProject(8) As String
    myProject(0) = ""            'VBE�̃o�[�W�����ɂ��̂Ŏg�p���Ȃ�_VBE7.DLL
    myProject(1) = ""            'EXCEL.EXE�̃o�[�W�����ɂ��̂Ŏg�p���Ȃ�_Office15
    myProject(2) = "stdole2.tlb"
    myProject(3) = "MSO.DLL"
    myProject(4) = "scrrun.dll"
    myProject(5) = "FM20.DLL"
    myProject(6) = "msado15.dll"
    myProject(7) = "REFEDIT.DLL"
    myProject(8) = "MSCOMCTL.OCX"
    
    '�Q�ƕs������ꍇbuf�ɃZ�b�g����
    For Each Ref In ActiveWorkbook.VBProject.References
        If Ref.isbroken = True Then
            buf = buf & myCount & vbTab & Ref.Name & vbTab & Ref.Description & vbTab & Ref.FullPath & vbCrLf
        End If
        myCount = myCount + 1
    Next Ref
    
    Debug.Print buf
    '�Q�ƕs������ꍇ
    If buf <> "" Then
        bufS = Split(buf, vbCrLf)
        For i = LBound(bufS) To UBound(bufS) - 1
            bufss = Split(bufS(i), vbTab)
            '�t�H���_��������΍쐬
            dirsp = Split(bufss(3), "\")
            dirstr = ""
            For i2 = LBound(dirsp) To UBound(dirsp) - 1
                dirstr = dirstr & "\" & dirsp(i2)
                If Dir(Mid(dirstr, 2), vbDirectory) = "" Then
                    MkDir Mid(dirstr, 2)
                End If
            Next i2
        Next i
        '���C�u�����t�@�C���̃R�s�[
        If Dir(bufss(3)) = "" Then
            FileCopy �A�h���X(0) & "\DLL\" & myProject(bufss(0)), bufss(3)
        End If
    End If
    
End Function

Public Function �n���F�ύX()
    Dim keyRow As Long, keyCol As Long
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim �� As String
    Dim �t�B�[���h��(2) As String
    �t�B�[���h��(0) = "����H����"
    �t�B�[���h��(1) = "���[�����ʎq"
    �t�B�[���h��(2) = "���L���r�e�B"
    
    Dim �F As Variant: �F = RGB(0, 102, 0)
    
    With ActiveSheet
        keyRow = .Cells.Find("�d�����ʖ�", , , 1).Row
        X = ActiveCell.Column
        'y = ActiveCell.Row
        �� = Left(.Cells(keyRow, X).Value, 2)
        If �� = "�n�_" Or �� = "�I�_" Then
            For Y = Selection(1).Row To Selection(Selection.count).Row
                If Y <= keyRow Then GoTo line10
                For i = LBound(�t�B�[���h��) To UBound(�t�B�[���h��)
                    keyCol = .Cells.Find(�� & �t�B�[���h��(i), , , 1).Column
                    .Cells(Y, keyCol).Font.color = �F
                    .Cells(Y, keyCol).Font.Bold = True
                Next i
line10:
            Next Y
        Else
            Exit Function
        End If
    End With

End Function

Public Function �f�B���N�g���쐬()
    Dim myDir As String, myDirS As Variant
    
    myDir = "\01_PVSW_csv,\05_RLTF_A,\06_RLTF_B,\07_SUB,\08_MD,\08_MD,\08_hsf�f�[�^�ϊ�,\08_hsf�f�[�^�ϊ�\log,\A0_���ވꗗ+"
    
    myDirS = Split(myDir, ",")
    For i = LBound(myDirS) To UBound(myDirS)
        If Dir(ActiveWorkbook.Path & "\" & myDirS(i), vbDirectory) = "" Then
            MkDir ActiveWorkbook.Path & myDirS(i)
        End If
    Next i
End Function
Public Function �K�v�t�@�C���̎擾()
    'exe
    Dim myDir As String, myDirS As Variant
    myDir = "\08_hsf�f�[�^�ϊ�\WH_DataConvert.exe"
    myDirS = Split(myDir, ",")
    For i = LBound(myDirS) To UBound(myDirS)
        If Dir(ActiveWorkbook.Path & "\" & myDirS(i)) = "" Then
            FileCopy �A�h���X(0) & "\hsf�f�[�^�ϊ�\WH_DataConvert.exe", ActiveWorkbook.Path & "\" & myDirS(i)
        End If
    Next i
    'ini�t�@�C���𖈉�쐬���Ȃ���
    Open ActiveWorkbook.Path & "\08_hsf�f�[�^�ϊ�\HsfDataConvert.ini" For Output As #1
        Print #1, "[Data]"
        Print #1, "HsfDataPath=" & ActiveWorkbook.Path & "\08_hsf�f�[�^�ϊ�"
        Print #1, "GuideDataPath=" & ActiveWorkbook.Path & "\08_MD"
        Print #1, "HsfSearchCnt=200"
        Print #1, "HsfReadState=0"
        Print #1, "[Time]"
        Print #1, "StartHour=0"
        Print #1, "StartMin=0"
        Print #1, "StartSec=0"
        Print #1, "EndHour=23"
        Print #1, "EndMin=0"
        Print #1, "EndSec=0"
    Close #1
    '���ވꗗ+�����邩�`�F�b�N
    Dim buf As String, cnt As Long
    Dim Path1 As String: Path1 = ActiveWorkbook.Path & "\" & "A0_���ވꗗ+\���ވꗗ+"
    buf = Dir(Path1 & "*.xlsm")
    Do While buf <> ""
        cnt = cnt + 1
        buf = Dir()
    Loop
    '���ވꗗ+�������ꍇ�͍ŐV�ł��擾
    If cnt = 0 Then
        Dim Path2 As String: Path2 = �A�h���X(1) & "\down\���ވꗗ+"
        buf = Dir(Path2 & "*.xlsm")
        Dim thisVer As String, newVer As String, fileName As String
        Do While buf <> ""
            thisVer = Mid(buf, InStr(buf, "+") + 1, InStr(buf, "_") - InStr(buf, "+") - 1)
            If newVer = "" Then
                newVer = thisVer
            Else
                If thisVer > newVer Then
                    newVer = thisVer
                End If
            End If
            buf = Dir()
        Loop
        FileCopy Path2 & newVer & "_.xlsm", Path1 & newVer & "_.xlsm"
    End If
End Function
Public Sub ���ޏڍ�_�[�q�t�@�~���[(strFilePath, �[�q�t�@�~���[)
    Dim intCount As Integer
    Dim intNo As Integer
    Dim strFileName As String
    Dim strBuff As String, getFlg As Boolean
    
    ' �t�@�C���I�[�v��
    intNo = FreeFile()                      ' �t���[�t�@�C��No���擾
    Open strFilePath For Input As #intNo    ' �t�@�C�����I�[�v��
    
    ' �t�@�C���̓ǂݍ���
    intCount = 0
    Do Until EOF(intNo)                     ' �t�@�C���̍Ō�܂Ń��[�v
        getFlg = False
        Line Input #intNo, strBuff          ' �t�@�C�������s�ǂݍ���
        For k = LBound(�[�q�t�@�~���[, 2) To UBound(�[�q�t�@�~���[, 2)
            If InStr(strBuff, "," & �[�q�t�@�~���[(0, k)) > 0 Then
                getFlg = True
                Exit For
            End If
        Next k
        
        If intCount = 0 Or getFlg = True Then
            ReDim Preserve strArray(intCount)   ' �z�񒷂�ύX
            strArray(intCount) = strBuff        ' �z��̍ŏI�v�f�ɓǂݍ��񂾒l����
            intCount = intCount + 1             ' �z��̗v�f�������Z
        End If
    Loop
    
    ' �t�@�C���N���[�Y
    Close #intNo
    
    ' �ǂݍ��񂾒l���m�F
'    Dim i As Integer
'    For i = 0 To UBound(strArray)
'        Debug.Print strArray(i)
'    Next i
    
End Sub

Public Sub SUB�f�[�^�擾(SUB�f�[�^RAN, strFilePath)
    Dim intCount As Integer
    Dim intNo As Integer
    Dim strFileName As String
    Dim strBuff As String, getFlg As Boolean
    
    ' �t�@�C���I�[�v��
    intNo = FreeFile()                      ' �t���[�t�@�C��No���擾
    Open strFilePath For Input As #intNo    ' �t�@�C�����I�[�v��
    ReDim SUB�f�[�^RAN(0)
    ' �t�@�C���̓ǂݍ���
    intCount = 0
    Do Until EOF(intNo)                     ' �t�@�C���̍Ō�܂Ń��[�v
        getFlg = False
        Line Input #intNo, strBuff          ' �t�@�C�������s�ǂݍ���
        ReDim Preserve SUB�f�[�^RAN(UBound(SUB�f�[�^RAN) + 1)
        SUB�f�[�^RAN(UBound(SUB�f�[�^RAN)) = strBuff
    Loop
    
    ' �t�@�C���N���[�Y
    Close #intNo
    
    ' �ǂݍ��񂾒l���m�F
'    Dim i As Integer
'    For i = 0 To UBound(strArray)
'        Debug.Print strArray(i)
'    Next i
    
End Sub


Public Sub �[�q�t�@�~���[����(myCell, �[�q�t�@�~���[)
    For i = LBound(strArray) To UBound(strArray)
        strArrayS = Split(strArray(i), ",")
        '���i�i�Ԃ̃}�b�`�m�F
        If myCell = Replace(strArrayS(0), "-", "") Then
            '�t�@�~���[�ԍ��̃}�b�`�m�F
            For ii = LBound(�[�q�t�@�~���[, 2) To UBound(�[�q�t�@�~���[, 2)
                If Left(strArrayS(13), 3) = �[�q�t�@�~���[(0, ii) Then
                    If strArrayS(14) = �[�q�t�@�~���[(1, ii) Or "*" = �[�q�t�@�~���[(1, ii) Then
                        myCell.Interior.color = �[�q�t�@�~���[(3, ii)
                        '���Ltemp�ɓo�^�����邩�m�F
                        Set fnd = Range("�[�q�t�@�~���[�͈�").Find(�[�q�t�@�~���[(0, ii) & �[�q�t�@�~���[(1, ii), , , 1)
                        If fnd Is Nothing Then
                            For Each f In Range("�[�q�t�@�~���[�͈�")
                                If f.Value = "" Then
                                    Sheets("�ݒ�").Hyperlinks.add anchor:=f, address:=�[�q�t�@�~���[(2, ii), ScreenTip:="", TextToDisplay:=�[�q�t�@�~���[(0, ii) & �[�q�t�@�~���[(1, ii)
                                    f.Interior.color = �[�q�t�@�~���[(3, ii)
                                    f.Font.color = 0
                                    f.Font.Underline = False
                                    f.AddComment
                                    f.Comment.Text �[�q�t�@�~���[(5, ii)
                                    f.Comment.Shape.TextFrame.AutoSize = True
                                    Exit Sub
                                End If
                            Next f
                        End If
                    End If
                End If
            Next ii
            Exit Sub
        End If
    Next i
    '������Ȃ�����
    'Stop  '���ވꗗ�̏���������?
End Sub

Public Sub �d���i�팟��(myCell, �d���i��)
    '�d���i��̃}�b�`�m�F
    For ii = LBound(�d���i��, 2) To UBound(�d���i��, 2)
        If myCell = �d���i��(1, ii) Then
                myCell.Interior.color = �d���i��(3, ii)
                '�d���i��temp�ɓo�^�����邩�m�F
                Set fnd = Range("�d���i��͈�").Find(�d���i��(0, ii), , , 1)
                If fnd Is Nothing Then
                    For Each f In Range("�d���i��͈�")
                        If f.Value = "" Then
                            Sheets("�ݒ�").Hyperlinks.add anchor:=f, address:=�d���i��(2, ii), ScreenTip:="", TextToDisplay:=�d���i��(0, ii)
                            f.Interior.color = �d���i��(3, ii)
                            f.Font.color = 0
                            f.Font.Underline = False
                            If �d���i��(5, ii) <> "" Then
                                f.AddComment
                                f.Comment.Text �d���i��(5, ii)
                                f.Comment.Shape.TextFrame.AutoSize = True
                            End If
                            Exit Sub
                        End If
                    Next f
                End If
        End If
    Next ii
End Sub

Public Function ���ޏڍ�_set(strFilePath, filterWord, u, myX)
    Dim intCount As Integer
    Dim intNo As Integer
    Dim strFileName As String
    Dim strBuff As String, getFlg As Boolean
    
    ' �t�@�C���I�[�v��
    intNo = FreeFile()                      ' �t���[�t�@�C��No���擾
    Open strFilePath For Input As #intNo    ' �t�@�C�����I�[�v��
    
    ' �t�@�C���̓ǂݍ���
    intCount = 0
    Do Until EOF(intNo)                     ' �t�@�C���̍Ō�܂Ń��[�v
        getFlg = False
        Line Input #intNo, strBuff          ' �t�@�C�������s�ǂݍ���
        '�t�B�[���h�����w��
        If intCount = 0 Then
            strbuffsp = Split(strBuff, ",")
            For i = LBound(strbuffsp) To UBound(strbuffsp)
                If strbuffsp(i) = filterWord Then
                    myX = i
                    Exit For
                End If
            Next i
        End If
        '�o�^�������
        
        strbuffsp = Split(strBuff, ",")
        If strbuffsp(myX) <> "" Then
            '�o�^
            ReDim Preserve strArray(intCount)   ' �z�񒷂�ύX
            strArray(intCount) = strBuff        ' �z��̍ŏI�v�f�ɓǂݍ��񂾒l����
            intCount = intCount + 1             ' �z��̗v�f�������Z
        End If
    Loop
    
    ' �t�@�C���N���[�Y
    Close #intNo

End Function

Public Function TEXT�o��_�ėp���������V�X�e��(myDir, �\��, �F��, �T�u, point, �[��, ��ƍH��)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 14) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir & "\" & Format(point, "0000") & ".html"

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift_JIS" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=8" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./img/wh.css" & Chr(34) & ">"
        outdats(6) = "<title>" & �\�� & "</title>"
        outdats(7) = "</head>"
        outdats(8) = "<body>"
        outdats(9) = "<table>"
        outdats(10) = "<tr><td class=" & Chr(34) & "title" & Chr(34) & "> �\��:" & �\�� & " " & �F�� & " �H��:" & �T�u & " " & ��ƍH�� & "</td></tr>"
        outdats(11) = "<tr><td><img src=" & Chr(34) & "./img/" & Format(point, "0000") & ".jpg" & Chr(34) & "></td></tr>"
        outdats(12) = "</table>"
        outdats(13) = "</body>"
        outdats(14) = "</html>"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber
    
End Function
Public Function TEXT�o��_�ėp���������V�X�e��html(myDir, �\��, �F��, �T�u, point, �[��, ��ƍH��, cav)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 17) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir & "\" & Format(point, "0000") & ".html"

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift_JIS" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=8" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/wh" & Format(point, "0000") & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & point & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink()" & Chr(34) & " >"
        
        outdats(9) = "<table>"
        
        outdats(10) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>" & �\�� & "</td><td>" & �F�� & "</td>"
        If ��n����Ǝ� = True Then outdats(10) = outdats(10) & "<td>" & myVer & " " & ��n����Ǝ҃V�[�g�� & "</td>"
        outdats(10) = outdats(10) & "<td>" & �T�u & "</td><td>" & ��ƍH�� & "</td></tr>"
        outdats(11) = "</table>"
                
        outdats(12) = "<div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[�� & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " ></div>"
        outdats(13) = "<div id=" & Chr(34) & "box2" & Chr(34) & " ><img src=" & Chr(34) & "./img/" & �[�� & "_1_" & cav & ".png" & Chr(34) & "></div>"
        outdats(14) = ""
        
        outdats(15) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(16) = "</body>"
        outdats(17) = "</html>"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber
    
    TEXT�o��_�ėp���������V�X�e��html = myPath

End Function

Public Function TEXT�o��_�z���o�Hhtml(myDir, �[��from, �[��to, ���i�i��str, �T�u, �T�u2, �\��, �F��, �n�_�n��, �n�_cav, �I�_�n��, �I�_cav, �[��leftRAN)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 38) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift-jis" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & �\�� & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & �\�� & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();myBlink2();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>�\��:" & �\�� & " " & �F�� & "</td><td>" & �[��from & " to " & �[��to & "</td><td>SUB:" & �T�u & "</td><td>Ver:" & myVer & "</td>" & _
                               "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
        '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        Dim �[��fromleft As Single, �[��toleft As Single, �[��from1 As String, �[��from2 As String, �[��to1 As String, �[��to2 As String
        �[��fromleft = 0: �[��toleft = 0
        For i = LBound(�[��leftRAN, 2) + 1 To UBound(�[��leftRAN, 2)
            If �[��from = �[��leftRAN(0, i) Then �[��fromleft = �[��leftRAN(1, i)
            If �[��to = �[��leftRAN(0, i) Then �[��toleft = �[��leftRAN(1, i)
        Next i
        '�E�ɂ�������E�ɕ\��������box6��7���ƉE�ɂȂ�
        If Val(�[��fromleft) >= Val(�[��toleft) Then
            �[��from1 = "box6"
            �[��from2 = "box7"
            �[��to1 = "box4"
            �[��to2 = "box5"
        Else
            �[��from1 = "box4"
            �[��from2 = "box5"
            �[��to1 = "box6"
            �[��to2 = "box7"
        End If
        
        If Left(�n�_�n��, 1) = "��" Then
            outdats(14) = "  <div  id=" & Chr(34) & �[��from1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_1.png" & Chr(34) & " ></div>"
            outdats(15) = "  <div  id=" & Chr(34) & �[��from2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_1_" & �n�_cav & ".png" & Chr(34) & " ></div>"
        End If
        
        If Left(�I�_�n��, 1) = "��" Then
            outdats(16) = "  <div id=" & Chr(34) & �[��to1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��to & "_1.png" & Chr(34) & " ></div>"
            outdats(17) = "  <div id=" & Chr(34) & �[��to2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��to & "_1_" & �I�_cav & ".png" & Chr(34) & " ></div>"
        End If
        outdats(18) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �T�u2 & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(19) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "to" & �[��to & "_" & �F�� & ".png" & Chr(34) & " ></div>"
        outdats(20) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �T�u & "_foot.png" & Chr(34) & " ></div>"
        outdats(21) = "</body>"
        
        outdats(22) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(23) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink2.js" & Chr(34) & "></script>"
        outdats(24) = "<script>"
        outdats(25) = "function checkText(){"
        outdats(26) = "  var str1=document.myform.txtb.value;"
        outdats(27) = "  var seihin,kosei;"
        outdats(28) = "  var myLen=str1.length;"
        outdats(29) = "  if (myLen <=10){"
        outdats(30) = "    kosei=str1;"
        outdats(31) = "  }else{"
        outdats(32) = "    seihin=str1.substr(25,10);"
        outdats(33) = "    kosei=str1.substr(11,4);"
        outdats(34) = "  }"
        outdats(35) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(36) = "}"
        outdats(37) = "</script>"
        outdats(38) = "</html>"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�Hhtml_UTF8(myDir, �[��from, �[��to, ���i�i��str, �T�u, �T�u2, �\��, �F��, �n�_�n��, �n�_cav, �I�_�n��, �I�_cav, �[��leftRAN)
        
        Dim i As Long
        Dim outdats(1 To 38) As Variant

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & �\�� & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & �\�� & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();myBlink2();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>�\��:" & �\�� & " " & �F�� & "</td><td>" & �[��from & " to " & �[��to & "</td><td>SUB:" & �T�u & "</td><td>Ver:" & myVer & "</td>" & _
                               "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
        '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        Dim �[��fromleft As Single, �[��toleft As Single, �[��from1 As String, �[��from2 As String, �[��to1 As String, �[��to2 As String
        �[��fromleft = 0: �[��toleft = 0
        For i = LBound(�[��leftRAN, 2) + 1 To UBound(�[��leftRAN, 2)
            If �[��from = �[��leftRAN(0, i) Then �[��fromleft = �[��leftRAN(1, i)
            If �[��to = �[��leftRAN(0, i) Then �[��toleft = �[��leftRAN(1, i)
        Next i
        '�E�ɂ�������E�ɕ\��������box6��7���ƉE�ɂȂ�
        If Val(�[��fromleft) >= Val(�[��toleft) Then
            �[��from1 = "box6"
            �[��from2 = "box7"
            �[��to1 = "box4"
            �[��to2 = "box5"
        Else
            �[��from1 = "box4"
            �[��from2 = "box5"
            �[��to1 = "box6"
            �[��to2 = "box7"
        End If
        
        '2.191.01
        If Left(�n�_�n��, 1) = "��" Or ��n���_�� = True Then
            outdats(14) = "  <div  id=" & Chr(34) & �[��from1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_1.png" & Chr(34) & " ></div>"
            outdats(15) = "  <div  id=" & Chr(34) & �[��from2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_1_" & �n�_cav & ".png" & Chr(34) & " ></div>"
        End If
        If Left(�I�_�n��, 1) = "��" Or ��n���_�� = True Then
            outdats(16) = "  <div id=" & Chr(34) & �[��to1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��to & "_1.png" & Chr(34) & " ></div>"
            outdats(17) = "  <div id=" & Chr(34) & �[��to2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��to & "_1_" & �I�_cav & ".png" & Chr(34) & " ></div>"
        End If
        outdats(18) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �T�u2 & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(19) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "to" & �[��to & "_" & �F�� & ".png" & Chr(34) & " ></div>"
        outdats(20) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �T�u & "_foot.png" & Chr(34) & " ></div>"
        outdats(21) = "</body>"
        
        outdats(22) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(23) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink2.js" & Chr(34) & "></script>"
        outdats(24) = "<script>"
        outdats(25) = "function checkText(){"
        outdats(26) = "  var str1=document.myform.txtb.value;"
        outdats(27) = "  var seihin,kosei;"
        outdats(28) = "  var myLen=str1.length;"
        outdats(29) = "  if (myLen <=10){"
        outdats(30) = "    kosei=str1;"
        outdats(31) = "  }else{"
        outdats(32) = "    seihin=str1.substr(25,10);"
        outdats(33) = "    kosei=str1.substr(11,4);"
        outdats(34) = "  }"
        outdats(35) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(36) = "}"
        outdats(37) = "</script>"
        outdats(38) = "</html>"

        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            
            '��������BOM�����ɂ��鏈��
            .Position = 0
            .Type = adTypeBinary
            .Position = 3 'BOM�f�[�^��3�o�C�g�ڂ܂�
            Dim byteData() As Byte '�ꎞ�i�[
            byteData = .Read  '�ꎞ�i�[�p�ϐ��ɕۑ�
            .Close '�X�g���[�������_���Z�b�g
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With
End Function


Public Function TEXT�o��_�z���o�H_�[���o�Hhtml(myDir, �[��from, �[��to, ���i�i��str, �T�u, �\��, �F��)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 34) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift-jis" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/tanmatukeiro.css" & Chr(34) & ">"
        outdats(6) = "<title>" & �[��from & "-</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>�[��: " & �[��from & "-</td><td>" & �F�� & "</td><td>" & �T�u & "</td><td>Ver:" & myVer & "</td>" & _
                                "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                                "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �T�u & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(15) = "  <div id=" & Chr(34) & "box4" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_2" & "_foot.png" & Chr(34) & " ></div>"
        outdats(16) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_2" & ".png" & Chr(34) & " ></div>"
        outdats(17) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_2" & "_tansen.png" & Chr(34) & " ></div>"
        outdats(18) = "</body>"
        
        outdats(19) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(20) = "<script>"
        outdats(21) = "function checkText(){"
        outdats(22) = "  var str1=document.myform.txtb.value;"
        outdats(23) = "  var seihin,kosei;"
        outdats(24) = "  var myLen=str1.length;"
        outdats(25) = "  if (myLen <=10){"
        outdats(26) = "    kosei=str1;"
        outdats(27) = "  }else{"
        outdats(28) = "    seihin=str1.substr(25,10);"
        outdats(29) = "    kosei=str1.substr(11,4);"
        outdats(30) = "  }"
        outdats(31) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(32) = "}"
        outdats(33) = "</script>"
        outdats(34) = "</html>"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�H_�[���o�Hhtml_UTF8(myDir, �[��from, �[��to, ���i�i��str, �T�u, �\��, �F��)
    
    Dim i As Integer
    Dim outdats(1 To 34) As Variant

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/tanmatukeiro.css" & Chr(34) & ">"
        outdats(6) = "<title>" & �[��from & "-</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>�[��: " & �[��from & "-</td><td>" & �F�� & "</td><td>" & �T�u & "</td><td>Ver:" & myVer & "</td>" & _
                                "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                                "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �T�u & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(15) = "  <div id=" & Chr(34) & "box4" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_2" & "_foot.png" & Chr(34) & " ></div>"
        outdats(16) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_2" & ".png" & Chr(34) & " ></div>"
        outdats(17) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��from & "_2" & "_tansen.png" & Chr(34) & " ></div>"
        outdats(18) = "</body>"
        
        outdats(19) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(20) = "<script>"
        outdats(21) = "function checkText(){"
        outdats(22) = "  var str1=document.myform.txtb.value;"
        outdats(23) = "  var seihin,kosei;"
        outdats(24) = "  var myLen=str1.length;"
        outdats(25) = "  if (myLen <=10){"
        outdats(26) = "    kosei=str1;"
        outdats(27) = "  }else{"
        outdats(28) = "    seihin=str1.substr(25,10);"
        outdats(29) = "    kosei=str1.substr(11,4);"
        outdats(30) = "  }"
        outdats(31) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(32) = "}"
        outdats(33) = "</script>"
        outdats(34) = "</html>"
        
        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            '��������BOM�����ɂ��鏈��
            .Position = 0
            .Type = adTypeBinary
            .Position = 3 'BOM�f�[�^��3�o�C�g�ڂ܂�
            Dim byteData() As Byte '�ꎞ�i�[
            byteData = .Read  '�ꎞ�i�[�p�ϐ��ɕۑ�
            .Close '�X�g���[�������_���Z�b�g
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With

End Function


Public Function TEXT�o��_�z���o�H_�[��html(myDir, �[��str, �[��0, ���i�i��str)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 31) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir
    �[��0 = "�[��:" & �[��0
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift_JIS" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & "atohame" & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & point & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>" & �[��0 & "</td><td>" & ���i�i��str & "</td><td>" & "" & "</td>" _
                               & "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��str & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"

        outdats(15) = "</body>"
        
        outdats(16) = "<script>"
        outdats(17) = "function checkText(){"
        outdats(18) = "  var str1=document.myform.txtb.value;"
        outdats(19) = "  var seihin,kosei;"
        outdats(20) = "  var myLen=str1.length;"
        outdats(21) = "  if (myLen <=10){"
        outdats(22) = "    kosei=str1;"
        outdats(23) = "  }else{"
        outdats(24) = "    seihin=str1.substr(25,10);"
        outdats(25) = "    kosei=str1.substr(11,4);"
        outdats(26) = "  }"
        outdats(27) = "  "
        outdats(28) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(29) = "}"
        outdats(30) = "</script>"
        outdats(31) = "</html>"
        
        '�z��̗v�f���������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�H_�[��html_UTF8(myDir, �[��str, �[��0, ���i�i��str)
    
    Dim i As Integer
    Dim outdats(1 To 31) As Variant

    �[��0 = "�[��:" & �[��0

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & "atohame" & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & �[��0 & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>" & �[��0 & "</td><td>" & ���i�i��str & "</td><td>" & "" & "</td><td>" & myVer & "</td>" _
                               & "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & �[��str & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"

        outdats(15) = "</body>"
        
        outdats(16) = "<script>"
        outdats(17) = "function checkText(){"
        outdats(18) = "  var str1=document.myform.txtb.value;"
        outdats(19) = "  var seihin,kosei;"
        outdats(20) = "  var myLen=str1.length;"
        outdats(21) = "  if (myLen <=10){"
        outdats(22) = "    kosei=str1;"
        outdats(23) = "  }else{"
        outdats(24) = "    seihin=str1.substr(25,10);"
        outdats(25) = "    kosei=str1.substr(11,4);"
        outdats(26) = "  }"
        outdats(27) = "  "
        outdats(28) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(29) = "}"
        outdats(30) = "</script>"
        outdats(31) = "</html>"
        
        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            '��������BOM�����ɂ��鏈��
            .Position = 0
            .Type = adTypeBinary
            .Position = 3 'BOM�f�[�^��3�o�C�g�ڂ܂�
            Dim byteData() As Byte '�ꎞ�i�[
            byteData = .Read  '�ꎞ�i�[�p�ϐ��ɕۑ�
            .Close '�X�g���[�������_���Z�b�g
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With

End Function


Public Function TEXT�o��_�ݒ�_�ƃ��C�A�E�g�}(myDir)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 4) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    mess0 = "2�s�ڂ�Cav���ԍ��̕ϊ��Ɏg�p����t�@�C�����A3�s�ڂɕ��ވꗗ�̃f�B���N�g������͂��Ă��������B"
    mess1 = Left(myDir, InStr(myDir, "���Y����+") + 4) & "\010_����͏��\Exchange_CavToHole.xlsx"
    mess2 = �A�h���X(1)
    mess3 = Left(myDir, InStr(myDir, "���Y����+") + 4) & "\010_����͏��\�����@�ݒ�.xlsx"
    
    myPath = myDir
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
        
        outdats(1) = mess0
        outdats(2) = mess1
        outdats(3) = mess2
        outdats(4) = mess3
        
        '�z��̗v�f���������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function


Public Function TEXT�o��_�ėp���������V�X�e��js(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 7) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = " mSec = 300; //  �_�ő��x (1sec=1000)"
        outdats(2) = " function myBlink(){"
        outdats(3) = "     flag = document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility;"
        outdats(4) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & ") document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";"
        outdats(5) = "     else document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";"
        outdats(6) = "     setTimeout(" & Chr(34) & "myBlink()" & Chr(34) & ",mSec);"
        outdats(7) = " }"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�H_�[��js(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 20) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = " //mSec = 300; //  �_�ő��x (1sec=1000)"
        outdats(2) = " function myBlink(){"
        outdats(3) = "     flag = document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility;"
        outdats(4) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & "){"
        outdats(5) = "         document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";"
        outdats(6) = "         mSec = 600;"
        outdats(7) = "     }else {"
        outdats(8) = "         document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";"
        outdats(9) = "         mSec = 300;"
        outdats(10) = "     }"
        outdats(11) = "     flag = document.getElementById(" & Chr(34) & "box3" & Chr(34) & ").style.visibility;"
        outdats(12) = "     if (flag == " & Chr(34) & "hidden" & Chr(34) & "){"
        outdats(13) = "         document.getElementById(" & Chr(34) & "box3" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";"
        outdats(14) = "         mSec = 600;"
        outdats(15) = "     }else {"
        outdats(16) = "         document.getElementById(" & Chr(34) & "box3" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";"
        outdats(17) = "         mSec = 300;"
        outdats(18) = "     }"
        outdats(19) = "     setTimeout(" & Chr(34) & "myBlink()" & Chr(34) & ",mSec);"
        outdats(20) = " }"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function
Public Function TEXT�o��_�z���o�H_�[��js2(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 17) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = " mSec = 300; //  �_�ő��x (1sec=1000)"
        outdats(2) = " function myBlink2(){"
        outdats(3) = " mSec = 175;"
        outdats(4) = "     try{flag = document.getElementById(" & Chr(34) & "box5" & Chr(34) & ").style.visibility;} catch(e){}"
        outdats(5) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & "){"
        outdats(6) = "         try{document.getElementById(" & Chr(34) & "box5" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";} catch(e){}"
        outdats(7) = "     }else {"
        outdats(8) = "         try{document.getElementById(" & Chr(34) & "box5" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";} catch(e){}"
        outdats(9) = "     }"
        outdats(10) = "     try{flag = document.getElementById(" & Chr(34) & "box7" & Chr(34) & ").style.visibility;} catch(e){}"
        outdats(11) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & "){"
        outdats(12) = "         try{document.getElementById(" & Chr(34) & "box7" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";} catch(e){}"
        outdats(13) = "     }else {"
        outdats(14) = "         try{document.getElementById(" & Chr(34) & "box7" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";} catch(e){}"
        outdats(15) = "     }"
        outdats(16) = "     setTimeout(" & Chr(34) & "myBlink2()" & Chr(34) & ",mSec);"
        outdats(17) = " }"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�H_ver(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 3) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
    outdats(1) = "���t:" & Date
    outdats(2) = "ver:" & myVer
    outdats(3) = "��n���̂�:" & �z���}�쐬temp
    
    Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�Hcss(myPath, box2l, box2t, box2w, box2h, clocode1, clofont)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 66) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#" & clocode1 & ";"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #" & clofont & ";"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #000000;"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #" & clofont & "; /* �\�����̐��F1px,����,�ΐF */"
        outdats(15) = "    border-left:2px solid #" & clofont & ";"
        outdats(16) = "    border-right:2px solid  #" & clofont & ";"
        outdats(17) = "    padding: 1px;            /* �Z�������̗]���F3�s�N�Z�� */"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:99%;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:99%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=70); /* IE 6,7*/"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:99%;"
        outdats(33) = "    opacity:0.7;"
        outdats(34) = "    zoom:1;"
        outdats(35) = "    display:inline-block;"
        outdats(36) = "}"
        outdats(37) = "#box3 img{"
        outdats(38) = "    position:absolute;"
        outdats(39) = "    width:99%;"
        outdats(40) = "}"
        outdats(41) = "#box4 img{"
        outdats(42) = "    position:absolute;"
        outdats(43) = "    bottom:0%;"
        outdats(44) = "    height:30%;"
        outdats(45) = "}"
        outdats(46) = "#box5 img{"
        outdats(47) = "    position:absolute;"
        outdats(48) = "    bottom:0%;"
        outdats(49) = "    height:30%;"
        outdats(50) = "    filter:alpha(opacity=70);"
        outdats(51) = "    opacity:0.7;"
        outdats(52) = "}"
        outdats(53) = "#box6 img{"
        outdats(54) = "    position:absolute;right:0%;"
        outdats(55) = "    bottom:0%;"
        outdats(56) = "    height:30%;"
        outdats(57) = "}"
        outdats(58) = "#box7 img{"
        outdats(59) = "    position:absolute;right:0%;"
        outdats(60) = "    bottom:0%;"
        outdats(61) = "    height:30%;"
        outdats(62) = "    filter:alpha(opacity=70);"
        outdats(63) = "    opacity:0.7;"
        outdats(64) = "}"
        outdats(65) = "body{background-color:#111111;}"
        outdats(66) = ".myB{color:#" & clofont & ";background-color:#" & clocode1 & ";}"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function

Public Function TEXT�o��_�z���o�H_�[��css(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 47) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    clocode1 = "FFFFFF"
    clofont = "000000"
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#" & clocode1 & ";"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #" & clofont & ";"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #000000;"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #" & clofont & "; /* �\�����̐��F1px,����,�ΐF */"
        outdats(15) = "    border-left:2px solid #" & clofont & ";"
        outdats(16) = "    border-right:2px solid  #" & clofont & ";"
        outdats(17) = "    padding: 1px;            /* �Z�������̗]���F3�s�N�Z�� */"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:auto;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:100%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=60); /* IE 6,7*/"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:100%;"
        outdats(33) = "    opacity:0.8;"
        outdats(34) = "    zoom:1;"
        outdats(35) = "    display:inline-block;"
        outdats(36) = "}"
        outdats(37) = "#box3 img{"
        outdats(38) = "    position:absolute;"
        outdats(39) = "    width:100%;"
        outdats(40) = "}"
        outdats(41) = "#box4 img{"
        outdats(42) = "    position:absolute;"
        outdats(43) = "    bottom:0%;"
        outdats(44) = "    width:100%;"
        outdats(45) = "}"
        outdats(46) = "body{background-color:#111111;}"
        outdats(47) = ".myB{color:#000000;background-color:#FFFFFF;}"
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function


Public Function TEXT�o��_�ėp���������V�X�e��css(myPath, clocode1, clofont)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 47) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#" & clocode1 & ";"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #" & clofont & ";"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #" & clofont & ";"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #" & clofont & "; /* �\�����̐��F1px,����,�ΐF */"
        outdats(15) = "    border-left:2px solid #" & clofont & ";"
        outdats(16) = "    border-right:2px solid  #" & clofont & ";"
        outdats(17) = "    padding: 1px;            /* �Z�������̗]���F3�s�N�Z�� */"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:auto;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:98%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=60); /* IE 6,7*/"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:auto;"
        outdats(33) = "    height:auto;"
        outdats(34) = "    max-width:98%;"
        outdats(35) = "    max-height:95%;"
        outdats(36) = "    opacity:0.6;"
        outdats(37) = "    display:inline-block;"
        outdats(38) = "}"
        outdats(39) = "#box3 img{"
        outdats(40) = "    position:absolute;"
        outdats(41) = "    width:100%;"
        outdats(42) = "}"
        outdats(43) = "#box4 img{"
        outdats(44) = "    position:absolute;"
        outdats(45) = "    bottom:0%;"
        outdats(46) = "    width:100%;"
        outdats(47) = "}"
        
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function



Public Function �n�_�I�_����ւ�()
    '�t�B�[���h��.row�𒴂��ĂȂ���Ώ������o��
    Dim keyRow As Long: keyRow = Cells.Find("�d�����ʖ�", , , 1).Row
    If Selection.Row < keyRow Then Exit Function
    
    Call �œK��
    '�n�_�����܂ރt�B�[���h�����擾
    Dim changeTitle As String: Dim lastCol As Long
    lastCol = Cells(keyRow, Columns.count).End(xlToLeft).Column
    For X = 1 To lastCol
        If Left(Cells(keyRow, X), 3) = "�n�_��" Then
            changeTitle = changeTitle & "," & Mid(Cells(keyRow, X), 4)
        End If
    Next X
    '�n�_��/�I�_���̗���擾
    Dim gawa(1) As String: gawa(0) = "�n�_��": gawa(1) = "�I�_��"
    Dim retsu() As Long
    Dim changeTitleSP As Variant: changeTitleSP = Split(changeTitle, ",")
    ReDim retsu(1, UBound(changeTitleSP))
    For g = 0 To 1
        For u = 1 To UBound(changeTitleSP) '0�Ԗڂ��]���ȃf�[�^
            retsu(g, u) = Rows(keyRow).Find(gawa(g) & changeTitleSP(u), , , 1).Column
        Next u
    Next g
    '�n�_��/�I�_�������ւ���
    Dim tempKey As Variant, tempCol As Long
    Set tempKey = Rows(keyRow).Find("�I�_�n�_����ւ�temp", , , 1)
    If tempKey Is Nothing Then
        tempCol = lastCol + 1
        Cells(keyRow, tempCol) = "�n�_�I�_����ւ�temp"
    Else
        tempCol = tempKey.Column
    End If
    'temp�̗�ɃR�s�[���Ċe�񖈂Ɏn�_/�I�_�̓���ւ�
    Dim startRow As Long: startRow = Selection.Row
    Dim endRow As Long: endRow = Selection.Row + Selection.Rows.count - 1
    For u = 1 To UBound(changeTitleSP)
        Range(Cells(startRow, retsu(0, u)), Cells(endRow, retsu(0, u))).Copy Destination:=Range(Cells(startRow, tempCol), Cells(endRow, tempCol))
        Range(Cells(startRow, retsu(1, u)), Cells(endRow, retsu(1, u))).Copy Destination:=Range(Cells(startRow, retsu(0, u)), Cells(endRow, retsu(0, u)))
        Range(Cells(startRow, tempCol), Cells(endRow, tempCol)).Copy Destination:=Range(Cells(startRow, retsu(1, u)), Cells(endRow, retsu(1, u)))
    Next u
    '����ւ����s�͗�:�n�I�ւ�1�ɂ���
    Dim changeFlgCol As Long: changeFlgCol = Rows(keyRow).Find("�n�I��", , , 1).Column
    For Y = startRow To endRow
        If Cells(Y, changeFlgCol) = "1" Then
            Cells(Y, changeFlgCol) = Empty
        Else
            Cells(Y, changeFlgCol) = "1"
        End If
    Next Y
    '���
    Columns(tempCol).Delete
    Set tempKey = Nothing
    
    Call �œK�����ǂ�
    
    Call PlaySound("�����Ă�")
    
End Function

Public Function ��ƐF�ɒ��F(myNum)
    '�t�B�[���h��.row�𒴂��ĂȂ���ΏI��
    Dim keyRow As Long: keyRow = Cells.Find("�d�����ʖ�", , , 1).Row
    If Selection.Row < keyRow Then Exit Function

    '[�ݒ�]����F���擾
    With Sheets("�ݒ�")
        Dim myKey As Range, myRange As Range, myNumF As Range
        Dim myFontColor As Variant, myInteriorColor As Long, myBold As Boolean
        Set myKey = .Cells.Find("�n���F_", , , 1).Offset(0, 1)
        Set myRange = .Range(myKey, myKey.End(xlDown))
        If myNum = "-" Then
            myFontColor = 0
            myInteriorColor = 16777215
            myBold = False
        Else
            Set myNumF = myRange.Find(myNum, , , 1)
            If myNumF Is Nothing Then Exit Function '�Ăяo���ꂽmyNum��������ΏI��
            myFontColor = myNumF.Font.color
            myInteriorColor = myNumF.Interior.color
            myBold = True
        End If
    End With

    '�n�_���܂��͏I�_����I�����Ă��Ȃ���ΏI��
    Dim selectGawa As String, retsu(2) As Long, �Ώۖ� As String
    selectGawa = Left(Cells(keyRow, Selection.Column), 3)
    If Not selectGawa = "�n�_��" And Not selectGawa = "�I�_��" Then Exit Function
    
    '���F�������擾
    Dim myTitle As String: myTitle = "��H����,�[�����ʎq,�L���r�e�B"
    Dim myTitleSP As Variant: myTitleSP = Split(myTitle, ",")
    For X = LBound(myTitleSP) To UBound(myTitleSP)
        retsu(X) = Rows(keyRow).Find(selectGawa & myTitleSP(X), , , 1).Column
    Next X

    'temp�̗�ɃR�s�[���Ċe�񖈂Ɏn�_/�I�_�̓���ւ�
    Dim startRow As Long: startRow = Selection.Row
    Dim endRow As Long: endRow = Selection.Row + Selection.Rows.count - 1
    For u = LBound(myTitleSP) To UBound(myTitleSP)
        Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Font.color = myFontColor '�t�H���g�F��ƐF�ɒ��F
        Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Font.Bold = myBold '�t�H���g�𑾎��ɂ���

        If myInteriorColor <> 16777215 Then
            Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Interior.color = myInteriorColor
        Else
            Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Interior.ColorIndex = xlNone '�w�i���h��Ԃ������̎�
        End If
    Next u

    '���
    Set myKey = Nothing
    Set myRange = Nothing
    Set myNumF = Nothing
    
    Call �œK�����ǂ�
    
    Call PlaySound("�����Ă�")
    
End Function

Public Function QR�R�[�h���N���b�v�{�[�h�Ɏ擾(Optional myString)
'    If IsMissing(myString) Then myString = "            0607         8211158560"
'    Dim MiBar As Mibarcd.Auto
'    Set MiBar = New Mibarcd.Auto
'    MiBar.CodeType = 12 '12=QR
'    MiBar.BarScale = 1
'    MiBar.QRVersion = 3 '�傫��������傫���Ȃ�
'    MiBar.QRErrLevel = 1
'    MiBar.Code = myString
'    MiBar.Execute
End Function

Public Function �t�B�[���h���̒ǉ�(wsTemp, myKey, myArea, LR)
    retsu = myArea.count / 2
    With wsTemp
        For i = 1 To retsu
            myLR = myArea(i)
            If LR = "" Or myLR = "l" Then
                �t�B�[���h�� = myArea(retsu + i)
                Set mykey2 = .Cells.Find(�t�B�[���h��, , , 1)
                '�t�B�[���h�������ꍇ
                If mykey2 Is Nothing Then
                    .Columns(myKey.Column + i - 1).Insert
                    .Columns(myKey.Column + i - 1).Interior.Pattern = xlNone
                    .Cells(myKey.Row, myKey.Column + i - 1) = myArea(retsu + i)
                    .Columns(myKey.Column + i - 1).AutoFit
                    .Cells(myKey.Row, myKey.Column + i - 1).Interior.color = myArea(retsu + i).Interior.color
                    '�R�����g������ꍇ�̓R�����g�ǉ�
                    If TypeName(myArea(retsu + i).Comment) <> "Nothing" Then
                        .Cells(myKey.Row, myKey.Column + i - 1).ClearComments
                        .Cells(myKey.Row, myKey.Column + i - 1).AddComment myArea(retsu + i).Comment.Text
                    End If
                '�R�����g������ꍇ�̓R�����g�폜���Ă���R�����g�ǉ�
                ElseIf TypeName(myArea(retsu + i).Comment) <> "Nothing" Then
                    .Cells(myKey.Row, myKey.Column + i - 1).ClearComments
                    .Cells(myKey.Row, myKey.Column + i - 1).AddComment myArea(retsu + i).Comment.Text
                End If
            End If
        Next i
    End With
End Function

Public Function �[�����i�ԕϊ�(�[�����i��)
    '-���܂ޏꍇ�͍폜�A�܂܂Ȃ��ꍇ��-��t�^
    If InStr(�[�����i��, "-") = 0 Then
        Select Case Len(�[�����i��)
        Case 8
            �[�����i�ԕϊ� = Left(�[�����i��, 4) & "-" & Mid(�[�����i��, 5, 4)
        Case 10
            �[�����i�ԕϊ� = Left(�[�����i��, 4) & "-" & Mid(�[�����i��, 5, 4) & "-" & Mid(�[�����i��, 9, 2)
        End Select
    Else
        �[�����i�ԕϊ� = Replace(�[�����i��, "-", "")
    End If
End Function

Public Function �|�C���g�i���o�[�}�쐬(Optional ���i�i��, Optional �[��, Optional �z��)
    Call �œK��
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim minW�w�� As Long
    Dim myKey, actKey
    Dim cavCol As Long, �|�C���g1Col As Long, ��d�W�~Col As Long
    myFont = "�l�r �S�V�b�N"
    minW�w�� = 30
    '�V�[�g����Ăяo������
    If IsMissing(���i�i��) Then
        With Workbooks(myBookName).Sheets(mySheetName)
            �n���}�^�C�v = "�`�F�b�J�[�p"
            Set myKey = .Cells.Find("�[�����i��", , , 1)
            Set actKey = ActiveCell
            If actKey.Row <= myKey.Row Then Exit Function
            If .Cells(actKey.Row, myKey.Column) = "" Then Exit Function
            cavCol = .Cells.Find("Cav", , , 1).Column
            �|�C���g1Col = .Cells.Find("�|�C���g1", , , 1).Column
            ��d�W�~Col = .Cells.Find("��d�W�~", , , 1).Column
            Dim �[�����Col As Integer: �[�����Col = .Cells.Find("�[�����i��", , , 1).Column
            Dim �[��Col As Integer: �[��Col = .Cells.Find("�[����", , , 1).Column
            Dim ���}col As Integer: ���}col = .Cells.Find("���}_�\�ʎ�", , , 1).Column
            ReDim �z��(7, 0)
            ���i�i�� = .Cells(actKey.Row, �[�����Col).Value
            �[�� = .Cells(actKey.Row, �[��Col).Value
            Dim myCount1 As Long, myCount2 As Long
            Dim myTop As Long, myLeft As Long, myEnd As Long, myHeight As Long
            myCount1 = -1
            Do
                If ���i�i�� <> .Cells(actKey.Row + myCount1, �[�����Col) Or �[�� <> .Cells(actKey.Row + myCount1, �[��Col) Then
                    myTop = .Cells(actKey.Row + myCount1 + 1, 1).Top
                    myLeft = .Columns(���}col).Left
                    Exit Do
                End If
                myCount1 = myCount1 - 1
            Loop
            myCount2 = 1
            Do
                If ���i�i�� <> .Cells(actKey.Row + myCount2, �[�����Col) Or �[�� <> .Cells(actKey.Row + myCount2, �[��Col) Then
                    myEnd = .Cells(actKey.Row + myCount2, 1).Top
                    myHeight = myEnd - myTop
                    Exit Do
                End If
                myCount2 = myCount2 + 1
            Loop
            Dim Y As Long, addc As Long
            For Y = actKey.Row + myCount1 + 1 To actKey.Row + myCount2 - 1
                addc = UBound(�z��, 2) + 1
                ReDim Preserve �z��(7, addc)
                �z��(0, addc) = .Cells(Y, cavCol)
                �z��(1, addc) = .Cells(Y, �|�C���g1Col)
                �z��(2, addc) = .Cells(Y, ��d�W�~Col)
            Next Y
            ���i�i�� = �[�����i�ԕϊ�(���i�i��)
            '�摜������ꍇ�͍폜
            Dim objShp As Shape
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name = ���i�i�� & "_" & �[�� Then
                    objShp.Delete
                End If
            Next
        End With
    End If
    �[���} = ���i�i�� & "_" & �[��
    Call �A�h���X�Z�b�g(ActiveWorkbook)
    
    Dim �I���o�� As String
    Dim �{�����[�h As Long: �{�����[�h = 1 '0(�����{) or 1(Cav��{)
    Dim �{�� As Single
    Dim frameWidth As Long, frameWidth1 As Long, frameWidth2 As Long, frameHeight1 As Long, frameHeight2 As Long, cornerSize As Single
    Dim pp As Long

    Dim �n���}��� As String: �n���}��� = "�ʐ^" ' �ʐ^(�ʐ^���������͗��}) or ���}�B�g���q�̓n���}��ނɉ�����(�Œ�)PVSW_RLTF���[�Ƀn���}��ނ��o�͂��鎞�ɍs���B
    Dim �n���}�g���q As String
    Dim ex As Long
    Dim varBinary As Variant
    Dim colHValue As New Collection  '�A�z�z��ACollection�I�u�W�F�N�g�̍쐬
    Dim lngNu() As Long

    With Workbooks(myBookName).Sheets(mySheetName)
        
        '���W�f�[�^�̓Ǎ���(�C���|�[�g�t�@�C��)
        Dim Target As New FileSystemObject
        Dim TargetDir As String: TargetDir = �A�h���X(1) & "\200_CAV���W"
        If Dir(TargetDir, vbDirectory) = "" Then MsgBox "���L�̃t�@�C���������ׁA�e�L���r�e�B�̍��W��������܂���B" & vbCrLf & "���ވꗗ+�ō��W�̏o�͂��s���Ă�����s���ĉ������B" & vbCrLf & vbCrLf & �A�h���X(1) & "\CAV���W.txt"
        
        Dim lastgyo As Long: lastgyo = 1
        Dim fileCount As Long: fileCount = 0
        Dim �g�p���istr As String
        Dim �g�p���i_�[�� As String
        
        Dim aa As Variant, a As Variant
        Dim ���W����Flag As Boolean
        Dim �g�p���i_�[��s_count As Long
        '�g�p���iStr�ɁA����g�p������W�f�[�^������
        Dim intFino As Variant
        intFino = FreeFile
        Dim ���r(1) As String
        ���W����Flag = False
        ���r(0) = "png": ���r(1) = "emf"
        minW = 1000: minH = 1000
        For ss = 0 To 1
            '�ʐ^,���}�̏��ŒT��
            URL = �A�h���X(1) & "\200_CAV���W\" & ���i�i�� & "_1_001_" & ���r(ss) & ".txt"
            If Dir(URL) <> "" Then
                intFino = FreeFile
                Open URL For Input As #intFino
                Do Until EOF(intFino)
                    Line Input #intFino, aa
                    a = Split(aa, ",")
                    If a(0) <> "PartName" Then
                        For b = LBound(�z��, 2) + 1 To UBound(�z��, 2)
                            If CStr(�z��(0, b)) = a(1) Then
                                �z��(3, b) = a(2)
                                �z��(4, b) = a(3)
                                �z��(5, b) = a(4)
                                �z��(6, b) = a(5)
                                �z��(7, b) = a(7)
                                If minW > CLng(a(4)) Then minW = CLng(a(4))
                                If minH > CLng(a(5)) Then minH = CLng(a(5))
                                Exit For
                            End If
                        Next b
                    End If
                Loop
                Close #intFino
                Exit For
            End If
        Next ss
        Dim �g�p���i As Variant, �g�p���is As Variant, �g�p���ic As Variant
line15:
        ReDim �d���f�[�^(2, 1) As String
        '�摜�̔z�u
        ReDim ���\�L(2, 0): ���c = 0
        Dim �摜����flg As Boolean: �摜����flg = False
        '�ʐ^
        �摜URL = �A�h���X(1) & "\���ވꗗ+_�ʐ^\" & ���i�i�� & "_1_" & Format(1, "000") & ".png"
        If Dir(�摜URL) = "" Then
            '���}
            �摜URL = �A�h���X(1) & "\���ވꗗ+_���}\" & ���i�i�� & "_1_" & Format(1, "000") & ".emf"
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
        ElseIf Dir(URL) = "" Then
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
                .Shapes.Range(�[���} & "_1").TextFrame2.TextRange.Text = "���W.txt������"
                .Shapes.Range(�[���}).Select
                .Shapes.Range(�[���} & "_1").Select False
                Selection.Group.Select
                Selection.Name = �[���}
            End With
        Else
            With ActiveSheet.Pictures.Insert(�摜URL)
                .Name = �[���}
                If minW < minH Then
                    my�� = (minW�w�� / minW)
                Else
                    my�� = (minW�w�� / minH)
                End If
                .ShapeRange(�[���}).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft '�摜���傫���ƃT�C�Y������������邩���̃T�C�Y�ɖ߂�
                .ShapeRange(�[���}).ScaleHeight my��, msoTrue, msoScaleFromTopLeft
                .CopyPicture
                .Delete
            End With
            DoEvents
            Sleep 10
            DoEvents
            .Paste
            Selection.Name = �[���}
            
            .Shapes(�[���}).Left = 0
            .Shapes(�[���}).Top = 0
            For i = LBound(�z��, 2) + 1 To UBound(�z��, 2)
                cav = �z��(0, i)
                If �z��(7, i) = "Ter" Then �z��(7, i) = "Box"
                If �z��(2, i) = True Or �z��(2, i) = 1 Then ��d�W�~flg = True Else ��d�W�~flg = False
                Call ColorMark3(�[��, CStr(�z��(3, i)), CStr(�z��(4, i)), CStr(�z��(5, i)), CStr(�z��(6, i)), "", "", �z��(7, i), "", "", �z��(1, i), "", "", "", "", RowStr)
            Next i
            .Shapes.Range(�[���}).Select
            For i = LBound(�z��, 2) + 1 To UBound(�z��, 2)
                .Shapes.Range(�[���} & "_" & �z��(0, i)).Select False
            Next i
            Selection.Group.Select
            Selection.Name = �[���}
            Selection.ShapeRange.Flip msoFlipHorizontal
            Selection.Copy
            Selection.Delete
            ActiveSheet.PasteSpecial Format:="�} (PNG)", Link:=False, DisplayAsIcon:=False
            Selection.Name = �[���}
        End If
        '�V�[�g������s�������̏���
        If myTop <> 0 Then
            Selection.Left = myLeft
            Selection.Top = myTop
            Selection.Height = myHeight
            actKey.Select
        Else
            Set �|�C���g�i���o�[�}�쐬 = Selection
        End If
    End With
    Call �œK�����ǂ�
    
End Function
Public Function ��n���}�Ăяo���pQR����f�[�^�쐬(Optional ����str)
    If IsMissing(����str) Then
        ����str = "152"
    End If
    Set wb(0) = ActiveWorkbook
    Set ws(0) = wb(0).Worksheets("���_" & ����str)
    '���[�N�u�b�N�쐬
    myBookpath = wb(0).Path
    '�o�͐�f�B���N�g����������΍쐬
    If Dir(myBookpath & "\56_�z���}_�U��", vbDirectory) = "" Then
        MkDir myBookpath & "\56_�z���}_�U��"
    End If
    
    With ws(0)
        Set myKey = .Cells.Find("Size_", , , 1)
        Dim �[��ran As Variant
        ReDim �[��ran(0)
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            xx = 1
            Do Until .Cells(Y, xx) = ""
                �Z��str = .Cells(Y, xx).Value
                If Left(�Z��str, 1) <> "U" Then
                    ReDim Preserve �[��ran(UBound(�[��ran) + 1)
                    �[��ran(UBound(�[��ran)) = �Z��str
                End If
                xx = xx + 2
            Loop
        Next Y
    End With
    
    newBookName = "QR���_" & ����str & ".xlsx"
    Set wb(1) = Workbooks.add
    
    With wb(1).Sheets("Sheet1")
        .Cells.NumberFormat = "@"
        .Cells(1, 1) = "QR"
        .Cells(1, 2) = "�[��"
        .Cells(2, 2) = "����_" & ����str
        addRow = 3
        For Y = LBound(�[��ran) + 1 To UBound(�[��ran)
            .Cells(addRow, 1).Resize(1, 2) = �[��ran(Y)
            addRow = addRow + 1
        Next Y
    End With
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=myBookpath & "\56_�z���}_�U��\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    wb(1).Close
    
End Function

Public Function �U�����j�^�̈ړ��f�[�^�쐬_��n���}csv(���i�i��str, ��zstr, ���str)
    'temp
    Set myBook = ActiveWorkbook
    Dim �[���ꗗran()
    Call SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, myBook.Name)

    With myBook.Sheets("���_" & ���str)
        .Activate
        Dim moveX As Long, moveXpt As Single
        Dim ���Wmm As Single: ���Wmm = .Cells.Find("Width_", , , 1).Offset(0, 1)
        If ���Wmm = 0 Then Stop '����Wmm�����͂���Ă��܂���
        Dim ���Wpt As Single: ���Wpt = .Shapes.Range("��").Width
        Dim ��n���}dir As String
        ��n���}dir = ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\xMove"
        If Dir(��n���}dir, vbDirectory) = "" Then MkDir (��n���}dir)
        Dim ��n���}path As String
        ��n���}path = ��n���}dir & "\��n���}.csv"
        Open ��n���}path For Output As #1
        For i = LBound(�[���ꗗran, 2) To UBound(�[���ꗗran, 2)
            �[��str = �[���ꗗran(1, i)
            moveXpt = .Shapes.Range(�[��str).Left + (.Shapes.Range(�[��str).Width / 2)
            moveX = moveXpt / ���Wpt * ���Wmm
            Print #1, �[��str & "," & moveX & "," & �[���ꗗran(3, i)
        Next i
        Close #1
    End With
End Function
Public Function �U�����j�^�̈ړ��f�[�^�쐬_�\��_�\���̒��Scsv(���i�i��str, ��zstr, ���str)
    'temp
    Set myBook = ActiveWorkbook
    Call SQL_�z���}�p_��H(�z���[��RAN, ���i�i��str, myBook)
    'Call SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, myBook.Name)
    Dim �T�u���WRAN()
    ReDim �T�u���WRAN(2, 0)

    With myBook.Sheets("���_" & ���str)
        Dim moveX As Long, moveXpt As Single
        Dim ���Wmm As Single: ���Wmm = .Cells.Find("Width_", , , 1).Offset(0, 1)
        Dim ���Wpt As Single: ���Wpt = .Shapes.Range("��").Width
        '�T�u���̒��Spt�����߂ăT�u���Wran�Ɋi�[
        Dim minX As Single, maxX As Single, aveX As Single
        For i = LBound(�z���[��RAN, 2) + 1 To UBound(�z���[��RAN, 2)
            For X = 0 To 1 '�n�_�I�_�̒[��
                �[��str = �z���[��RAN(4 + X, i)
                If �[��str <> "" Then
                    �[��pt = .Shapes.Range(�[��str).Left + (.Shapes.Range(�[��str).Width / 2)
                    If �[��pt < minX Or minX = 0 Then minX = �[��pt
                    If �[��pt > maxX Then maxX = �[��pt
                End If
                If X = 1 Then
                    If minX = 0 Then minX = maxX
                    If maxX = 0 Then maxX = minX
                    aveX = minX + ((maxX - minX) / 2)
                    ReDim Preserve �T�u���WRAN(2, UBound(�T�u���WRAN, 2) + 1)
                    �T�u���WRAN(0, UBound(�T�u���WRAN, 2)) = aveX
                    �T�u���WRAN(1, UBound(�T�u���WRAN, 2)) = minX
                    �T�u���WRAN(2, UBound(�T�u���WRAN, 2)) = maxX
                    minX = 0: maxX = 0
                End If
            Next X
        Next i
        
        Dim ��n���}dir As String
        ��n���}dir = ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\xMove"
        If Dir(��n���}dir, vbDirectory) = "" Then MkDir (��n���}dir)
        Dim ��n���}path As String
        ��n���}path = ��n���}dir & "\�\��.csv"
        Open ��n���}path For Output As #1
        For i = LBound(�z���[��RAN, 2) + 1 To UBound(�z���[��RAN, 2)
            �\��str = �z���[��RAN(2, i)
            '�T�ustr = �z���[��RAN(1, i)
            aveX = �T�u���WRAN(0, i) / ���Wpt * ���Wmm
            minX = �T�u���WRAN(1, i) / ���Wpt * ���Wmm
            maxX = �T�u���WRAN(2, i) / ���Wpt * ���Wmm
            colorLong = �z���[��RAN(11, i)
            Print #1, �\��str & "," & aveX & "," & minX & "," & maxX & "," & colorLong
        Next i
        Close #1
    End With
End Function
Public Function �U�����j�^�̈ړ��f�[�^�쐬_�\��_�T�u�̒��Scsv(���i�i��str, ��zstr, ���str)
    'temp
    Set myBook = ActiveWorkbook
    Call SQL_�z���}�p_��H(�z���[��RAN, ���i�i��str, myBook)
    'Call SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, myBook.Name)
    Dim �T�u���WRAN()
    ReDim �T�u���WRAN(1, 0)
    
    With myBook.Sheets("���_" & ���str)
        Dim moveX As Long, moveXpt As Single
        Dim ���Wmm As Single: ���Wmm = .Cells.Find("Width_", , , 1).Offset(0, 1)
        Dim ���Wpt As Single: ���Wpt = .Shapes.Range("��").Width
        
        '�T�u���̒��Spt�����߂ăT�u���Wran�Ɋi�[
        Dim minX As Single, maxX As Single, aveX As Single
        For i = LBound(�z���[��RAN, 2) + 1 To UBound(�z���[��RAN, 2)
            �T�ustr = �z���[��RAN(1, i)
            For X = 0 To 1 '�n�_�I�_�̒[��
                �[��str = �z���[��RAN(4 + X, i)
                If �[��str <> "" Then
                    
                    �[��pt = .Shapes.Range(�[��str).Left + (.Shapes.Range(�[��str).Width / 2)
                    If �[��pt < minX Or minX = 0 Then minX = �[��pt
                    If �[��pt > maxX Then maxX = �[��pt
                End If
                
                If X = 1 Then
                    If i = UBound(�z���[��RAN, 2) Then
                        aveX = minX + ((maxX - minX) / 2)
                        ReDim Preserve �T�u���WRAN(1, UBound(�T�u���WRAN, 2) + 1)
                        �T�u���WRAN(0, UBound(�T�u���WRAN, 2)) = �T�ustr
                        �T�u���WRAN(1, UBound(�T�u���WRAN, 2)) = aveX
                        minX = 0: maxX = 0
                    Else
                        �T�unext = �z���[��RAN(1, i + 1)
                        If �T�ustr <> �T�unext Then
                            aveX = minX + ((maxX - minX) / 2)
                            ReDim Preserve �T�u���WRAN(1, UBound(�T�u���WRAN, 2) + 1)
                            �T�u���WRAN(0, UBound(�T�u���WRAN, 2)) = �T�ustr
                            �T�u���WRAN(1, UBound(�T�u���WRAN, 2)) = aveX
                            minX = 0: maxX = 0
                        End If
                    End If
                End If
            Next X
        Next i
        
        Dim ��n���}dir As String
        ��n���}dir = ActiveWorkbook.Path & "\56_�z���}_�U��\" & Replace(���i�i��str, " ", "") & "_" & ��zstr & "\xMove"
        If Dir(��n���}dir, vbDirectory) = "" Then MkDir (��n���}dir)
        Dim ��n���}path As String
        ��n���}path = ��n���}dir & "\�T�u.csv"
        Open ��n���}path For Output As #1
        For i = LBound(�z���[��RAN, 2) + 1 To UBound(�z���[��RAN, 2)
            �\��str = �z���[��RAN(2, i)
            �T�ustr = �z���[��RAN(1, i)
            For ii = LBound(�T�u���WRAN, 2) + 1 To UBound(�T�u���WRAN, 2)
                If �T�ustr = �T�u���WRAN(0, ii) Then
                    moveXpt = �T�u���WRAN(1, ii)
                    moveX = moveXpt / ���Wpt * ���Wmm
                    Print #1, �\��str & "," & moveX & "," & �T�ustr
                    Exit For
                End If
            Next ii
        Next i
        Close #1
    End With
End Function

Public Function checkSpace(address)
    If InStr(address, "\") = 1 Then '\\10.7.120.44�Ƃ�
        address = Left(address, InStr(Mid(address, 3), "\") + 1)
    Else
        address = Left(address, InStr(address, "\") - 2)
    End If
    Dim FSO As Object, DrvLetter As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    DrvLetter = address
    If DrvLetter = "" Then
        Set FSO = Nothing
        Exit Function
    End If
    If FSO.DriveExists(DrvLetter) Then
        Dim maxSize As Long, nowSize As Long
        maxSize = Format(FSO.GetDrive(DrvLetter).TotalSize / 1024 / 1024 / 1024, "0")
        nowSize = Format(FSO.GetDrive(DrvLetter).AvailableSpace / 1024 / 1024 / 1024, "0")
        checkSpace = "�e��:" & nowSize & "/" & maxSize & "GB (" & Format(nowSize / maxSize * 100, "0") & "%)"
    Else
        checkSpace = ""
    End If
    Set FSO = Nothing
End Function

Sub MakeShortcut(Path)
    Dim ShellObject
    Set ShellObject = CreateObject("WScript.Shell")
   
    Dim ShortcutObject
    Set ShortcutObject = ShellObject.CreateShortcut(Path & "\" & ActiveWorkbook.Name & ".lnk")
    
    With ShortcutObject
        .TargetPath = ActiveWorkbook.FullName
        .Save
    End With
End Sub

Sub ���O�o��test_temp()
    Call �A�h���X�Z�b�g(wb(0))
    Set wb(0) = ThisWorkbook
    Call ���O�o��("aaa", "bbb", "textttttttttttt")
End Sub

Public Function ���O�o��(�t�H���_, �t�@�C����, �e�L�X�g1)
    Dim myPath As String, myIP As String, myDir As String
    myPath = �A�h���X(0) & "\log\" & �t�H���_ & "\" & �t�@�C���� & ".txt"
    myDir = �A�h���X(0) & "\log\" & �t�H���_
    myIP = GetIPAddress
    '�t�H���_��������΍쐬
    If Dir(myDir, vbDirectory) = "" Then
        MkDir (myDir)
    End If
    '�e�L�X�g�t�@�C�����g�p����
    Dim tFso As FileSystemObject
    Dim tFile As TextStream
    '�t�@�C����������ΐV�K�쐬
    If Dir(myPath) = "" Then
        Set tFso = New FileSystemObject
        Set tFile = tFso.CreateTextFile(myPath, True)
        Set tFso = Nothing
        Set tFile = Nothing
    End If
                
    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Append���[�h�ŊJ���܂��B
    Open myPath For Append As #FileNumber
    Dim outdats(3)
    '�o�͗p�̔z��փf�[�^���Z�b�g���܂��B
    outdats(0) = Now
    outdats(1) = myIP
    outdats(2) = �e�L�X�g1
    outdats(3) = ThisWorkbook.FullName
    '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
    Print #FileNumber, Join(outdats, vbTab)

    '���̓t�@�C������܂��B
    Close #FileNumber
    
End Function

Public Function ���ޏڍׂ̓ǂݍ���(���i�i��str, �t�B�[���h��str)
        Dim Target As New FileSystemObject
        Dim Path As String: Path = �A�h���X(1) & "\300_���ޏڍ�\" & ���i�i��str & ".txt"
        If Dir(Path) = "" Then ���ޏڍׂ̓ǂݍ��� = False: Exit Function
        Dim intFino As Variant
        intFino = FreeFile
        Open Path For Input As #intFino
        myX = ""
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For X = LBound(temp) To UBound(temp)
                If Replace(temp(X), "-", "") = �t�B�[���h��str Then
                    Line Input #intFino, aa
                    temp = Split(aa, ",")
                    ���ޏڍׂ̓ǂݍ��� = temp(X)
                    Close #intFino
                    Exit Function
                End If
            Next X
        Loop
        Close #intFino
End Function

Public Function �Z���̒��g��S�ēn��(base As Range, aite As Range)
    base.Value = aite.Value
    If aite.Interior.ColorIndex <> xlNone Then base.Interior.color = aite.Interior.color
    If Not (aite.Comment Is Nothing) Then
        Set �R�����g = base.AddComment
        �R�����g.Text aite.Comment.Text
        �R�����g.Visible = False
        �R�����g.Shape.Fill.ForeColor.RGB = RGB(255, 192, 0)
        �R�����g.Shape.TextFrame.AutoSize = True
        �R�����g.Shape.TextFrame.Characters.Font.Size = 11
        �R�����g.Shape.Placement = xlMove
    End If
End Function

Public Function TEXT�o��_�z���o�H_�[���o�Hcss(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 47) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '�󂢂Ă���t�@�C���ԍ����擾���܂��B
    FileNumber = FreeFile
    '�t�@�C����Output���[�h�ŊJ���܂��B
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#232526;"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #FFFFFF;"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #000000;"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #FFFFFF;"
        outdats(15) = "    border-left:2px solid #FFFFFF;"
        outdats(16) = "    border-right:2px solid  #FFFFFF;"
        outdats(17) = "    padding: 1px;"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:99%;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:99%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=70);"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:99%;"
        outdats(33) = "    opacity:0.7;"
        outdats(34) = "    zoom:1;"
        outdats(35) = "    display:inline-block;"
        outdats(36) = "}"
        outdats(37) = "#box3 img{"
        outdats(38) = "    position:absolute;"
        outdats(39) = "    width:99%;"
        outdats(40) = "}"
        outdats(41) = "#box4 img{"
        outdats(42) = "    position:absolute;"
        outdats(43) = "    bottom:0%;"
        outdats(44) = "    width:99%;"
        outdats(45) = "}"
        outdats(46) = "body{background-color:#111111;}"
        outdats(47) = ".myB{color:#FFFFFF;background-color:#232526;}"
        
        '�z��̗v�f���J���}�Ō������ďo�͂��܂��B
        Print #FileNumber, Join(outdats, vbCrLf)

    '���̓t�@�C������܂��B
    Close #FileNumber

End Function


