VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_09 
   Caption         =   "��n���U��_SSC����"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_09.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_09"
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
    '�폜����
    Set wb(0) = ThisWorkbook
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "��������"
    Call ���i�i��RAN_set2(���i�i��Ran, CB0.Value, CB1.Value, "")

    Unload Me
    
    '�n���}���쐬
    cb�I�� = "4,4,1,1,0,-1,1"
    �}���}�`�� = 160
    ���^�p�x����flag = True
    �[���i���o�[�\�� = False
    Call �n���}�쐬_Ver220098(cb�I��, "���C���i��", CB1.Value)
    
    ���i�i��str = Replace(���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "���C���i��"), 1), " ", "")
    �ݕ�str = ���i�i��Ran(���i�i��RAN_read(���i�i��Ran, "��z"), 1)
    
    '�����̃Z�b�g
    Dim myRan As Variant, myPath As String
    'myRan = setWorkRan(�[���T�uRAN)
    myPath = wb(0).path & dirString_09 & Replace(���i�i��str, " ", "") & "_wire.txt"
    myRan = readTextToArray(myPath)
    
    Call ��n���U��_SSC����(myRan, "�n���}_���C���i��_" & ���i�i��str, ���i�i��str & "_" & �ݕ�str, �[���T�uRAN)
    
    PlaySound "���񂹂�"
    
    Dim myMsg As String: myMsg = "�쐬���܂���" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "��n���U��_SSC����")
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
