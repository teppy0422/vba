VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_02 
   Caption         =   "���̓V�[�g�̍쐬"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6990
   OleObjectBlob   =   "UI_02.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




































Private Sub CommandButton1_Click()
    PlaySound "��������"
    RLTF�T�u = RLTF�T�ucbx.Value
    Unload UI_02
    Call ���i�ʒ[���ꗗ�̃V�[�g�쐬_2009
End Sub

Private Sub CommandButton2_Click()
    PlaySound "��������"
    Unload UI_02
    Call ���i���X�g�̍쐬_Ver2040
    MsgBox "�V�[�g[" & ActiveSheet.Name & "] ���쐬���܂����B"
End Sub

Private Sub CommandButton3_Click()
    PlaySound "��������"
    Unload UI_02
    mytime = �|�C���g�ꗗ�̃V�[�g�쐬_2190
    PlaySound "��������"
    Call MsgBox(mytime & "s �쐬���܂����B", vbOKOnly, "�|�C���g�ꗗ")
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("���ǂ�")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    PlaySound "��������"
    Unload Me
    MsgBox ���V�[�g�̍쐬
End Sub

Private Sub CommandButton6_Click()
    PlaySound "��������"
    Unload Me
    mytime = CAV�ꗗ�쐬2190
    PlaySound "��������"
    Call MsgBox(mytime & "s �쐬���܂����B", vbOKOnly, "CAV�ꗗ")

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub
