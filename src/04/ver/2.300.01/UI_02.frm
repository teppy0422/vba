VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_02 
   Caption         =   "���̓V�[�g�̍쐬"
   ClientHeight    =   8910
   ClientLeft      =   50
   ClientTop       =   410
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
    If Multiple = False Then Unload Me
    Call ���i�ʒ[���ꗗ�̃V�[�g�쐬_220081
End Sub

Private Sub CommandButton2_Click()
    PlaySound "��������"
    If Multiple = False Then Unload Me
    Call ���i���X�g�̍쐬_Ver220078(Me.isgetMDcbx)
    If Multiple = False Then MsgBox "�V�[�g[" & ActiveSheet.Name & "] ���쐬���܂����B"
End Sub

Private Sub CommandButton3_Click()
    PlaySound "��������"
    If Multiple = False Then Unload Me
    mytime = �|�C���g�ꗗ�̃V�[�g�쐬_2190
    PlaySound "��������"
    If Multiple = False Then MsgBox mytime & "s �쐬���܂����B", vbOKOnly, "�|�C���g�ꗗ"
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("���ǂ�")
    If Multiple = False Then Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    PlaySound "��������"
    Unload Me
    Call ���V�[�g�̍쐬
End Sub

Private Sub CommandButton6_Click()
    PlaySound "��������"
    If Multiple = False Then Unload Me
    mytime = CAV�ꗗ�쐬2190
    PlaySound "��������"
    If Multiple = False Then MsgBox mytime & "s �쐬���܂����B", vbOKOnly, "CAV�ꗗ"
End Sub

Private Sub CommandButton7_Click()
    Multiple = True
    CommandButton1_Click
    CommandButton2_Click
    CommandButton6_Click
    CommandButton3_Click
    CommandButton5_Click
    CommandButton8_Click
End Sub

Private Sub CommandButton8_Click()
    PlaySound "��������"
    Unload Me
    Dim wsTemp As Worksheet
    Set wsTemp = �ʒm���̍쐬_220060
    PlaySound "��������"
    If Multiple = False Then
        MsgBox wsTemp.Name & " ���쐬/�X�V���܂����B", vbOKOnly, "�ʒm��"
        wsTemp.Activate
    End If
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub
