VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_00 
   Caption         =   "1.�f�[�^�C���|�[�g"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   OleObjectBlob   =   "UI_00.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private Sub CommandButton1_Click()
    Unload UI_00
    PlaySound ("��������")
    Call PVSWcsv_csv�̃C���|�[�g_2029
    PlaySound ("��������")
    Call ���O�o��("test", "test", "PVSW�C���|�[�g���s")
    MsgBox "[PVSW_RLTF]�ւ�PVSW.csv�̃C���|�[�g���������܂����B"
    Sheets("���i�i��").Activate
End Sub

Private Sub CommandButton2_Click()
    PlaySound ("��������")
    RLTF�T�u = RLTF�T�ucbx.Value
    Unload UI_00
    Call PVSWcsv��RLTFA�����H�����擾_Ver2026
    Call PVSWcsv��RLTFB�����H�����擾
    PlaySound ("���񂹂�")
    Call ���O�o��("test", "test", "RLTF�C���|�[�g���s")
    MsgBox "�擾���������܂����B"
    Sheets("PVSW_RLTF").Activate
End Sub

Private Sub CommandButton3_Click()
    PlaySound ("��������")
    Unload UI_00
    Sheets("PVSW_RLTF").Activate
    Sleep 10
    'Call �œK��
    Call PVSWcsv�̋��ʉ�_Ver1944
    Call PVSW_RLTF�̃T�u0�ɑ����i�̃T�u�����蓖�Ă�_2048
    'Call �œK�����ǂ�
    PlaySound ("���񂹂�")
    Call ���O�o��("test", "test", "PVSW_RLTF�œK��")
    MsgBox "�������������܂����B"
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("���ǂ�")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "�Ƃ���"
End Sub
