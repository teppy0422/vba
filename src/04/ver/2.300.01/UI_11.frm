VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_11 
   Caption         =   "���i���\��_�쐬"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_11.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_11"
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
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "��������"
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "")
    
    Dim fileName As String: fileName = Replace(wb(0).Name, ".xlsm", "") & "_���i���\��_" & CB0.Value & "_" & CB1.Value & ".xlsx"
    Unload Me
    
    Dim i As Long, pNumbers As String
    For i = LBound(���i�i��Ran, 2) + 1 To UBound(���i�i��Ran, 2)
        pNumbers = pNumbers & "," & ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "���C���i��"), i)
    Next i
    
    Dim setWords As String, setWordsSP As Variant
    setWords = "���i�i��,�ď�,d,D,W,L,�F,���i����,���ޏڍ�,���,�H��,�H��a"
    setWordsSP = Split(setWords, ",")
    
    Set ws(1) = wb(0).Sheets("���i���X�g")
    msg = checkFieldName("���i�i��", ws(1), setWords)
    If msg <> "" Then
        msg = "[���i���X�g]�Ɏ��̃t�B�[���h��������܂���B" & msg & vbCrLf & vbCrLf & _
                   "���̋@�\���g�p����ɂ�Ver2.200.70�ȍ~�ō쐬����[���i���X�g]�ł���K�v������܂��B" & vbCrLf & _
                   "�쐬�𒆎~���܂��B"
        MsgBox msg, vbOKOnly, "PLUS+"
        End
    End If
    
    Dim Array_���i���X�g As Variant
    Array_���i���X�g = readSheetToRan2(ws(1), "���i�i��", setWords & pNumbers, "")
    
    '���i�i�Ԗ��Ɏg�p���������i���폜
    Dim x As Long, skipFlag As Boolean
    For i = LBound(Array_���i���X�g, 2) + 1 To UBound(Array_���i���X�g, 2)
        skipFlag = True
        For x = UBound(setWordsSP) + 1 To UBound(setWordsSP) + UBound(���i�i��Ran, 2)
            If Array_���i���X�g(x, i) <> "" Then
                skipFlag = False
                Exit For
            End If
        Next x
        If skipFlag = True Then
            Debug.Print i, Array_���i���X�g(0, i)
            Array_���i���X�g = delete_RanVer2(Array_���i���X�g, i)
            i = i - 1
        End If
        If i + 1 > UBound(Array_���i���X�g, 2) Then Exit For
    Next i
    
    '�o�͂���f�[�^�̂܂Ƃ�
    Dim addArray() As Variant
    ReDim addArray(2, UBound(Array_���i���X�g, 2))
    addArray(0, 0) = "A"
    addArray(1, 0) = "B"
    addArray(2, 0) = "C"
    
    For i = LBound(Array_���i���X�g, 2) + 1 To UBound(Array_���i���X�g, 2)
        If Array_���i���X�g(9, i) = "B" Then
            addArray(0, i) = Array_���i���X�g(0, i)
            If Array_���i���X�g(2, i) <> "" Then
                addArray(1, i) = Replace(Array_���i���X�g(2, i) & " L=" & Array_���i���X�g(5, i), ".0", "")
            End If
        ElseIf Array_���i���X�g(9, i) = "T" Then
            addArray(2, i) = Replace(Array_���i���X�g(0, i), "-", " ")
            addArray(1, i) = Replace(Array_���i���X�g(1, i), " ", "") & "-" & Array_���i���X�g(6, i)
            If Replace(Array_���i���X�g(3, i), " ", "") <> "" Then
                addArray(0, i) = "D" & Replace(Replace(Array_���i���X�g(2, i) & "�~" & Array_���i���X�g(3, i), ".0", ""), " ", "") & " L=" & Replace(Array_���i���X�g(5, i), " ", "")
            ElseIf Replace(Array_���i���X�g(2, i), " ", "") <> "" Then
                addArray(0, i) = "D" & Replace(Replace(Array_���i���X�g(2, i), ".0", ""), " ", "") & " L=" & Replace(Replace(Array_���i���X�g(5, i), ".0", ""), " ", "")
            ElseIf Replace(Array_���i���X�g(4, i), " ", "") <> "" Then
                addArray(0, i) = Replace(Replace("W" & Array_���i���X�g(4, i), ".0", ""), " ", "") & " L=" & Replace(Replace(Array_���i���X�g(5, i), ".0", ""), " ", "")
            End If
        Else
            
        End If
    Next i
    
    '����L��0����4���ɂ���
    Dim array_temp
    For i = LBound(Array_���i���X�g, 2) To UBound(Array_���i���X�g, 2)
        array_temp = Array_���i���X�g(5, i)
        If array_temp <> "" Then
            If IsNumeric(array_temp) Then
                array_temp = Int(array_temp)
                If (Len(array_temp) <= 4) Then
                    array_temp = String(4 - Len(array_temp), "0") & array_temp
                    Array_���i���X�g(5, i) = array_temp
                    
                End If
            End If
        End If
    Next
     
    Array_���i���X�g = merge_Array(addArray, Array_���i���X�g)

    export_ArrayToSheet Array_���i���X�g, "���i���\��", True
    
    Dim outputDirectory As String
    outputDirectory = wb(0).path & "\42_���i���\��"
    If Dir(outputDirectory, vbDirectory) = "" Then MkDir outputDirectory
    
    '�e�v����SPC100��csv��txt�́A�J���}����؂蕶���Ƃ��ĔF������ׁA�J���}��؂�ŏo�͂���ƃe�L�X�g���̃J���}�ŗ�Y������������ _
    �Ȃ̂�xlsx�ŏo��
    wb(0).Sheets("���i���\��").Move
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=outputDirectory & "\" & fileName, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Call �œK�����ǂ�
    PlaySound "���񂹂�"
    
    Dim myMsg As String: myMsg = "�쐬���܂���" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "���i���\��_�쐬")
End Sub

Private Sub UserForm_Initialize()
    Dim ����(1) As String
    With wb(0).Sheets("���i�i��")
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
            If ����s(i) = "����" Then myindex = i
        Next i
        .ListIndex = myindex
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub
