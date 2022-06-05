Attribute VB_Name = "M01_基礎定義"
'スリープイベント
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'マウスイベント
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

'クリップボードをクリア
Public Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long

'画面サイズの取得
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_cxScreen As Long = 0
Public Const SM_cyScreen As Long = 1

'押されたキーの取得
Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

Public strNewText As String
Public EXTES As String


Sub 最適化()
        '最適化
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
End Sub

Sub 最適化もどす()
        '最適化もどす
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub

Sub 最適化2()
        Application.ScreenUpdating = False
        Application.EnableEvents = False
End Sub

Sub 最適化2もどす()
        Application.ScreenUpdating = True
        Application.EnableEvents = True
End Sub

Sub 最適化3()
        '最適化
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
End Sub

Sub 最適化3もどす()
        '最適化もどす
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub


Sub 最適化4()
        '最適化
        Application.EnableEvents = False
End Sub

Sub 最適化4もどす()
        '最適化もどす
        Application.EnableEvents = True
End Sub

Sub 最適化5()
        '最適化
        Application.ScreenUpdating = False
End Sub

Sub 最適化5もどす()
        '最適化もどす
        Application.ScreenUpdating = True
End Sub


Public Function RoundUp(X As Currency, s As Integer) As Currency

Dim w As Currency
T = 10 ^ Abs(s)

If X > 0 Then
If s > 0 Then
RoundUp = -Int(-X * T) / T
Else
RoundUp = -Int(-X / T) * T
End If
Else
If s > 0 Then
RoundUp = Int(X * T) / T
Else
RoundUp = Int(X / T) * T
End If
End If

End Function

Public Function RoundDown(X As Currency, s As Integer) As Currency

Dim T As Currency
T = 10 ^ Abs(s)

If X > 0 Then
If s > 0 Then
RoundDown = Int(X * T) / T
Else
RoundDown = Int(X / T) * T
End If
Else
If s > 0 Then
RoundDown = -Int(-X * T) / T
Else
RoundDown = -Int(-X / T) * T
End If
End If

End Function
