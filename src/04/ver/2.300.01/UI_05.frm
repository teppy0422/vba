VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_05 
   Caption         =   "���̑�"
   ClientHeight    =   4650
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7530
   OleObjectBlob   =   "UI_05.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









































































































































































































Private Sub B0_Click()
    PlaySound "���񂽂�"
    CB0.ListIndex = 4
    CB1.ListIndex = 0
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = True
    PIC00.Picture = LoadPicture(myAddress(0, 1) & "\�n���}sample_" & "4511000000" & ".jpg")
End Sub

Private Sub CB0_Change()
    'Call CB�I��ύX
End Sub

Private Sub CB5_Change()
    If CB5.Value = "" Then Exit Sub
    With ActiveWorkbook.Sheets("���i�i��")
        Set key = .Cells.Find("�^��", , , 1)
        myCol = .Rows(key.Row).Find(CB5.Value, , , 1).Column
        lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
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
End Sub

Public Function CB�I��ύX()
    Call addressSet(wb(0))
    
    ��� = CB0.ListIndex & CB1.ListIndex & CB2.ListIndex & CB3.ListIndex
    ��� = ��� & "000000"
    
    ��� = Replace(���, "-1", "0")
    
    If Left(���, 1) = "0" Then ��� = "0000000000"
    
    PIC00.Picture = LoadPicture(myAddress(0, 1) & "\�n���}sample_" & ��� & ".jpg")
End Function

Private Sub CommandButton1_Click()
    Unload Me
    If CB6.ListIndex = -1 Then
        �R�����g.Visible = True
        �R�����g.Caption = "���i�i�Ԃ��I������Ă��܂���B"
        Beep
        Exit Sub
    End If
    
    PlaySound ("��������")
    cb�I�� = CB0.ListIndex
    
    Call ���i�i��RAN_set2(���i�i��Ran, CB5.Value, CB6.Value, "")
    
    If ���i�i��RANc = 0 Then
        �R�����g.Visible = True
        �R�����g.Caption = "�Y�����鐻�i�i�Ԃ�����܂���B" & vbCrLf _
                         & "�Ⴆ�ΑI�������������A" & vbCrLf & "[PVSW_RLTF]�ɍ݂�܂���B"
        Beep
        Exit Sub
    End If
    
    Unload UI_01
    
    Select Case CB0.ListIndex
    Case 0
        PlaySound ("�����Ă�")
        Call �T�u�ꗗ�\�̍쐬
        PlaySound ("���񂹂�")
    Case 1
        PlaySound ("�����Ă�")
        Call �ގ��R�l�N�^�ꗗb�쐬
        PlaySound ("���񂹂�")
    Case -1
        
    End Select
    
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("���ǂ�")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()

    Dim ����(6) As String
    ����(0) = "�T�u�ꗗ�\,�ގ��R�l�N�^�ꗗb"
    
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

    
    ����s = Split(����(5), ",")
    With CB5
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

