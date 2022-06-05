Attribute VB_Name = "M91_SendInput"
Public Const KEY_DOWN = 0   'キー押下
Public Const KEY_UP = 1     'キーアップ

' 指定時間Wait（ミリ秒）
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


'仮想キーコード                    EXTESの割り当て
Public Const クリア = &H21         '33  PAGEUP
Public Const タブ = &H22           '34  PAGEDOWN
Public Const ホーム = &H24         '36  HOME
Public Const エンター = &HD        '13  NUM_RETURN
Public Const ダウン = &H28         '40  不要
Public Const コントロール = &H11   '17  不要
Public Const オルト = &H12         '18  不要
Public Const コピー = &H2D         '45  INSERT
Public Const スベテセンタク = &H6F '111 NUM_/
Public Const シフト = &H10         'SHIFT

Public Const エー = &H41         '65 不要
Public Const シー = &H43         '67 不要
Public Const イー = &H45         '69 不要

Private Const KEYEVENTF_KEYUP = &H2 'キーアップ
Private Const KEYEVENTF_EXTENDEDKEY = &H1   'スキャンコードは拡張コード
Private Const INPUT_KEYBOARD = 1    '入力タイプ：キーボード

'仮想キーコード・ASCII値・スキャンコード間でコードを変換する
Declare Function MapVirtualKey Lib "user32" _
    Alias "MapVirtualKeyA" (ByVal wCode As Long, _
    ByVal wMapType As Long) As Long
'
' 仮想キーコードをスキャンコード、または文字の値（ASCII 値）へ変換。
' また、スキャンコードを仮想コードへ変換も可。
'
'［入力
' 　wCode：キーの仮想キーコード、またはスキャンコードを指定。
'　　　　　この値の解釈方法は、wMapType パラメータの値に依存。
'
' 　uMapType:実行したい変換の種類を指定。
' 　このパラメータの値に基づいて、uCode パラメータの値は次のように解釈。
'
'　　値 意味
' 　　0 wCode は仮想キーコードであり、スキャンコードへ変換。
' 　　　左右のキーを区別しない仮想キーコードのときは、関数は左側のスキャンコードを返却。
' 　　1 wCode はスキャンコードであり、仮想キーコードへ変換。
' 　　　この仮想キーコードは、左右のキーを区別。
' 　　2 wCode は仮想キーコードであり、戻り値の下位ワードにシフトなしの ASCII 値が格納。
' 　　　デッドキー（ 分音符号）は、戻り値の上位ビットをセットすることにより明示される。
' 　　3 Windows NT/2000：uCode はスキャンコードであり、左右のキーを区別する仮想キーコードへ変換。
'
' 　　　いづれも、変換されないときは、関数は 0 を返す。

'キーボード入力、マウスボタンのクリックをシミュレートする
Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, _
     pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
'
' nInputs:構造体の数を指定
' pInputs:配列へのポインタ INPUT 構造
' 　　　　各構造体には、キーボードまたはマウス入力動作に対応するイベントを表す
' cbSize :構造体のサイズを指定

Public Sub KeyEvent(VkKey As Integer, UpDown As Integer)
'
' 簡略化のためにAPIへは1文字ずつ入力⇒構造体は１つ
'
' VkKey:仮想キーコード
' UpDown:動作(KEY_DOWN/KEY_UP)
'
    Dim inputevents As INPUT_TYPE
    With inputevents
        .dwType = INPUT_KEYBOARD
        With .xi
            .wVk = VkKey        '操作キーコード
            .wScan = MapVirtualKey(VkKey, 0)  'スキャンコード
            If UpDown = KEY_DOWN Then   'キーDown
                .dwFlags = KEYEVENTF_EXTENDEDKEY Or 0
            Else                        'キーＵＰ
                .dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP
            End If
            .time = 0
            .dwExtraInfo = 0
        End With
    End With
    Call SendInput(1, inputevents, Len(inputevents))
End Sub

Sub 仮想キー入力(ByVal a As Integer)  'aはKEYCODEで指定 例)RETURNは13 TAB

Call KeyEvent(a, KEY_DOWN)
Call KeyEvent(a, KEY_UP)
Sleep 50

End Sub
