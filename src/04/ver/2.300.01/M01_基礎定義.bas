Attribute VB_Name = "M01_��b��`"
'�X���[�v�C�x���g
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'�}�E�X�C�x���g
Public Declare Sub mouse_event Lib "user32.dll" _
(ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
ByVal dwData As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000&
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_RIGHTDOWN = &H8

'�N���b�v�{�[�h���N���A
Public Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long

'��ʃT�C�Y�̎擾
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nindex As Long) As Long
Public Const SM_cxScreen As Long = 0
Public Const SM_cyScreen As Long = 1

'�����ꂽ�L�[�̎擾
Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

Public strNewText As String
Public EXTES As String


Sub �œK��()
        '�œK��
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
End Sub

Sub �œK�����ǂ�()
        '�œK�����ǂ�
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub

Sub �œK��2()
        Application.ScreenUpdating = False
        Application.EnableEvents = False
End Sub

Sub �œK��2���ǂ�()
        Application.ScreenUpdating = True
        Application.EnableEvents = True
End Sub

Sub �œK��3()
        '�œK��
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
End Sub

Sub �œK��3���ǂ�()
        '�œK�����ǂ�
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub


Sub �œK��4()
        '�œK��
        Application.EnableEvents = False
End Sub

Sub �œK��4���ǂ�()
        '�œK�����ǂ�
        Application.EnableEvents = True
End Sub

Sub �œK��5()
        '�œK��
        Application.ScreenUpdating = False
End Sub

Sub �œK��5���ǂ�()
        '�œK�����ǂ�
        Application.ScreenUpdating = True
End Sub


Public Function RoundUp(x As Currency, S As Integer) As Currency

Dim w As Currency
T = 10 ^ Abs(S)

If x > 0 Then
If S > 0 Then
RoundUp = -Int(-x * T) / T
Else
RoundUp = -Int(-x / T) * T
End If
Else
If S > 0 Then
RoundUp = Int(x * T) / T
Else
RoundUp = Int(x / T) * T
End If
End If

End Function

Public Function RoundDown(x As Currency, S As Integer) As Currency

Dim T As Currency
T = 10 ^ Abs(S)

If x > 0 Then
If S > 0 Then
RoundDown = Int(x * T) / T
Else
RoundDown = Int(x / T) * T
End If
Else
If S > 0 Then
RoundDown = -Int(-x * T) / T
Else
RoundDown = -Int(-x / T) * T
End If
End If

End Function

Sub �摜�Ƃ��ďo��(myPicName)

    Selection.Copy
    
    ActiveSheet.Pictures.Paste.Select
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Set obj = Selection
    Dim myWidth As Single: myWidth = Selection.Width
    Dim myHeight As Single: myHeight = Selection.Height
     '�摜�\��t���p�̖��ߍ��݃O���t���쐬
    Set cht = ActiveSheet.ChartObjects.add(0, 0, myWidth, myHeight).Chart
     '���ߍ��݃O���t�ɓ\��t����
    cht.Paste
    cht.PlotArea.Fill.Visible = mesofalse
    cht.ChartArea.Fill.Visible = msoFalse
    cht.ChartArea.Border.LineStyle = 0
    
    '�T�C�Y����
    ActiveWindow.Zoom = 100
    '��l = 1000
    �{�� = 1
    �{�� = 192 / Selection.Width
    ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleWidth �{��, False, msoScaleFromTopLeft
    ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "�O���t "))).ScaleHeight �{��, False, msoScaleFromTopLeft
    
    cht.Export fileName:=ActiveWorkbook.path & "\" & myPicName & ".bmp", filtername:="bmp"
    
     '���ߍ��݃O���t���폜
    ActiveSheet.Activate
    cht.Parent.Delete
    obj.Delete
    
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True

End Sub
