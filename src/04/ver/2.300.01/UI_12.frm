VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_12 
   Caption         =   "�����w��������V�X�e���̑I��"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   10270
   OleObjectBlob   =   "UI_12.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













































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

Private Sub CommandButton5_Click()
    Dim Result As Variant, SjpSetting_decide(0, 0) As Variant
    If Me.CB0.ListIndex = -1 Then End
    SjpSetting_decide(0, 0) = Me.CB0.Value
    If SjpSetting_decide(0, 0) <> "" Then
        Result = SpjSetting_write(SjpSetting_decide)
    End If
    End
End Sub

Function SpjSetting_write(ByVal ary As Variant) As Boolean
    If Dir(SjpSetting_Path) = "" Then makeDir path
    export_Array_ShiftJis_ver2 ary, SjpSetting_Path, ","
    Unload Me
End Function

Private Sub UserForm_Initialize()
    '�O���[�o���ϐ�����l���󂯎��
    Dim sp As Variant, myDir As String
    sp = Split(SjpSetting_list, ",")
    With CB0
        .RowSource = ""
        For i = LBound(sp) To UBound(sp)
            .AddItem sp(i)
        Next i
        .ListIndex = 0
    End With
    myDir = Left(SjpSetting_Path, InStrRev(SjpSetting_Path, "\") - 1)
    If Dir(myDir, vbDirectory) = "" Then
        MkDir (myDir)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub
