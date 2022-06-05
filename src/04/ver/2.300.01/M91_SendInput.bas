Attribute VB_Name = "M91_SendInput"
Public Const KEY_DOWN = 0   '�L�[����
Public Const KEY_UP = 1     '�L�[�A�b�v

' �w�莞��Wait�i�~���b�j
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Type KEYBDINPUT
      wVk As Integer
      wScan As Integer
      dwFlags As Long
      time As Long
      dwExtraInfo As Long
      no_use1 As Long
      no_use2 As Long
End Type

Type INPUT_TYPE
      dwType As Long
      xi As KEYBDINPUT
End Type


'���z�L�[�R�[�h                    EXTES�̊��蓖��
Public Const �N���A = &H21         '33  PAGEUP
Public Const �^�u = &H22           '34  PAGEDOWN
Public Const �z�[�� = &H24         '36  HOME
Public Const �G���^�[ = &HD        '13  NUM_RETURN
Public Const �_�E�� = &H28         '40  �s�v
Public Const �R���g���[�� = &H11   '17  �s�v
Public Const �I���g = &H12         '18  �s�v
Public Const �R�s�[ = &H2D         '45  INSERT
Public Const �X�x�e�Z���^�N = &H6F '111 NUM_/
Public Const �V�t�g = &H10         'SHIFT

Public Const �G�[ = &H41         '65 �s�v
Public Const �V�[ = &H43         '67 �s�v
Public Const �C�[ = &H45         '69 �s�v

Private Const KEYEVENTF_KEYUP = &H2 '�L�[�A�b�v
Private Const KEYEVENTF_EXTENDEDKEY = &H1   '�X�L�����R�[�h�͊g���R�[�h
Private Const INPUT_KEYBOARD = 1    '���̓^�C�v�F�L�[�{�[�h

'���z�L�[�R�[�h�EASCII�l�E�X�L�����R�[�h�ԂŃR�[�h��ϊ�����
Declare Function MapVirtualKey Lib "user32" _
    Alias "MapVirtualKeyA" (ByVal wCode As Long, _
    ByVal wMapType As Long) As Long
'
' ���z�L�[�R�[�h���X�L�����R�[�h�A�܂��͕����̒l�iASCII �l�j�֕ϊ��B
' �܂��A�X�L�����R�[�h�����z�R�[�h�֕ϊ����B
'
'�m����
' �@wCode�F�L�[�̉��z�L�[�R�[�h�A�܂��̓X�L�����R�[�h���w��B
'�@�@�@�@�@���̒l�̉��ߕ��@�́AwMapType �p�����[�^�̒l�Ɉˑ��B
'
' �@uMapType:���s�������ϊ��̎�ނ��w��B
' �@���̃p�����[�^�̒l�Ɋ�Â��āAuCode �p�����[�^�̒l�͎��̂悤�ɉ��߁B
'
'�@�@�l �Ӗ�
' �@�@0 wCode �͉��z�L�[�R�[�h�ł���A�X�L�����R�[�h�֕ϊ��B
' �@�@�@���E�̃L�[����ʂ��Ȃ����z�L�[�R�[�h�̂Ƃ��́A�֐��͍����̃X�L�����R�[�h��ԋp�B
' �@�@1 wCode �̓X�L�����R�[�h�ł���A���z�L�[�R�[�h�֕ϊ��B
' �@�@�@���̉��z�L�[�R�[�h�́A���E�̃L�[����ʁB
' �@�@2 wCode �͉��z�L�[�R�[�h�ł���A�߂�l�̉��ʃ��[�h�ɃV�t�g�Ȃ��� ASCII �l���i�[�B
' �@�@�@�f�b�h�L�[�i ���������j�́A�߂�l�̏�ʃr�b�g���Z�b�g���邱�Ƃɂ�薾�������B
' �@�@3 Windows NT/2000�FuCode �̓X�L�����R�[�h�ł���A���E�̃L�[����ʂ��鉼�z�L�[�R�[�h�֕ϊ��B
'
' �@�@�@���Â���A�ϊ�����Ȃ��Ƃ��́A�֐��� 0 ��Ԃ��B

'�L�[�{�[�h���́A�}�E�X�{�^���̃N���b�N���V�~�����[�g����
Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, _
     pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
'
' nInputs:�\���̂̐����w��
' pInputs:�z��ւ̃|�C���^ INPUT �\��
' �@�@�@�@�e�\���̂ɂ́A�L�[�{�[�h�܂��̓}�E�X���͓���ɑΉ�����C�x���g��\��
' cbSize :�\���̂̃T�C�Y���w��

Public Sub KeyEvent(VkKey As Integer, UpDown As Integer)
'
' �ȗ����̂��߂�API�ւ�1���������́ˍ\���̂͂P��
'
' VkKey:���z�L�[�R�[�h
' UpDown:����(KEY_DOWN/KEY_UP)
'
    Dim inputevents As INPUT_TYPE
    With inputevents
        .dwType = INPUT_KEYBOARD
        With .xi
            .wVk = VkKey        '����L�[�R�[�h
            .wScan = MapVirtualKey(VkKey, 0)  '�X�L�����R�[�h
            If UpDown = KEY_DOWN Then   '�L�[Down
                .dwFlags = KEYEVENTF_EXTENDEDKEY Or 0
            Else                        '�L�[�t�o
                .dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP
            End If
            .time = 0
            .dwExtraInfo = 0
        End With
    End With
    Call SendInput(1, inputevents, Len(inputevents))
End Sub

Sub ���z�L�[����(ByVal a As Integer)  'a��KEYCODE�Ŏw�� ��)RETURN��13 TAB

Call KeyEvent(a, KEY_DOWN)
Call KeyEvent(a, KEY_UP)
Sleep 50

End Sub
