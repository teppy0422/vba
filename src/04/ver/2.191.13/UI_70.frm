VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_70 
   Caption         =   "70_��������"
   ClientHeight    =   3000
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4800
   OleObjectBlob   =   "UI_70.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False































































Private Sub cbx1_Click()
    If cbx1 = True Then

    Else
        Label0.Visible = True
        Label0.Caption = ""
    End If
End Sub

Private Sub CommandButton1_Click()

    Call ���i�i��RAN_set2(���i�i��RAN, "���C���i��", CB0.Value, "")
    �}���}�`�� = 160 'Tear
    If cbx0 = True Then
        cb�I�� = "2,0,0,1,0,-1"
        �F�Ŕ��f = True
    Else
        cb�I�� = "2,1,0,1,0,-1"
        �F�Ŕ��f = False
    End If
    Unload Me
    Call �n���}�쐬_Ver2001(cb�I��, "���C���i��", CB0.Value)
    Call ���������V�X�e���p�f�[�^�쐬v2182(CB0.Value)
    myBook.Activate
    MsgBox "�쐬���܂���" & vbLf & DateDiff("s", mytime, Time) & " s"
End Sub

Private Sub CommandButton4_Click()
    PlaySound "���ǂ�"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()
    Dim ����(0) As String
    Call �A�h���X�Z�b�g(myBook)
    With ActiveWorkbook.Sheets("���i�i��")
        Set myKey = .Cells.Find("���C���i��", , , 1)
        lastRow = .Cells(Rows.count, myKey.Column).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            ����(0) = ����(0) & "," & .Cells(Y, myKey.Column)
        Next Y
        ����(0) = Mid(����(0), 2)
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
End Sub

