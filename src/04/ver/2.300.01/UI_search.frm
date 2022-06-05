VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_search 
   Caption         =   "検索"
   ClientHeight    =   8490
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8730
   OleObjectBlob   =   "UI_search.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








































































'APIフォーム透過
 Private Declare Function GetParent _
 Lib "user32" _
 (ByVal hWnd As Long) As Long
 Private Declare Function GetWindowLong _
 Lib "user32" Alias "GetWindowLongA" ( _
 ByVal hWnd As Long, ByVal nindex As Long) As Long
 Private Declare Sub SetWindowLong _
 Lib "user32" Alias "SetWindowLongA" ( _
 ByVal hWnd As Long, ByVal nindex As Long _
 , ByVal dwNewLong As Long)
 Private Declare Sub SetLayeredWindowAttributes _
 Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long _
 , ByVal bAlpha As Long, ByVal dwFlags As Long)
 Private Declare Sub DrawMenuBar Lib "user32" (ByVal hWnd As Long)
 Private Const GWL_EXSTYLE As Long = -20&
 Private Const WS_EX_LAYERED As Long = &H80000
 Private Const LWA_ALPHA As Long = &H2&

Private Const PICTURE_BACK  As String = "\T13_Back.Jpg"
Private Const PICTURE_CHARA  As String = "\T13_Chara.GIF"
Private Const PICTURE_MASK  As String = "\T13_CharaMask.GIF"


Private Sub lion_Click()

End Sub

Private Sub UserForm_Initialize()
'透過
    Dim myFrame As MSForms.Control
    Dim myHwnd As Long
    Dim myWindowLong As Long
    Dim myAlpha As Long
    myAlpha = 248 '透明度（0～255の整数値、0で透明）
    Set myFrame = Me.Controls.add("Forms.Frame.1")
    myHwnd = GetParent(GetParent(myFrame.[_GethWnd]))
    Me.Controls.Remove myFrame.Name
    Set myFrame = Nothing
    myWindowLong = GetWindowLong(myHwnd, GWL_EXSTYLE)
    myWindowLong = myWindowLong Or WS_EX_LAYERED
    SetWindowLong myHwnd, GWL_EXSTYLE, myWindowLong
    SetLayeredWindowAttributes myHwnd, 0&, myAlpha, LWA_ALPHA
    DrawMenuBar myHwnd '念のため
End Sub

Private Sub t品番_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 27 Then Unload Me
    
    If KeyCode <> 13 Then Exit Sub
    If Shift = 1 Then
        検索車種リスト.Clear
        t品番 = ""
        Exit Sub
    End If
    
    Dim myDic As Object, myKey, myItem
    Dim myVal, myVal2, myVal3
    Dim i As Long, x As Long
    Dim lastgyo As Long
    Dim バイト数 As Long
    Dim 検索 As String

    検索 = StrConv(t品番.Value, vbNarrow)
    検索 = UCase(検索)
    検索 = Replace(検索, "-", "")
    検索str = "種類,工程,部品品番,部材詳細"
    検索strsp = Split(検索str, ",")
    Dim 検索RAN()
    ReDim 検索RAN(6, 0)
    Dim 検索x()
    ReDim 検索x(UBound(検索strsp))
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws
        Set myKey = .Cells.Find("部品品番", , , 1)
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For x = 0 To UBound(検索strsp)
            検索x(x) = .Cells.Find(検索strsp(x), , , 1).Column
        Next x
        
        For y = myKey.Row + 1 To lastRow
            ReDim Preserve 検索RAN(6, UBound(検索RAN, 2) + 1)
            For x = 0 To UBound(検索x)
                検索RAN(x, UBound(検索RAN, 2)) = .Cells(y, 検索x(x))
            Next x
            検索RAN(5, UBound(検索RAN, 2)) = y
        Next y
    End With
    
    検索車種リスト.RowSource = ""
    検索車種リスト.Clear
    Dim C As Long
    For i = LBound(検索RAN, 2) + 1 To UBound(検索RAN, 2)
        For x = LBound(検索strsp) To UBound(検索strsp)
            If UCase(StrConv(Replace(検索RAN(x, i), "-", ""), vbNarrow)) Like "*" & 検索 & "*" Then
                検索車種リスト.AddItem ""
                検索車種リスト.List(C, 0) = 検索RAN(0, i)
                検索車種リスト.List(C, 1) = 検索RAN(1, i)
                検索車種リスト.List(C, 2) = 検索RAN(2, i)
                検索車種リスト.List(C, 3) = 検索RAN(3, i)
                C = C + 1
                Exit For
            End If
        Next x
    Next i
    If C > 0 Then
        検索車種リスト.ListIndex = 0
    Else
        検索車種リスト.ListIndex = -1
        検索車種リスト.AddItem ""
        検索車種リスト.List(0, 2) = "みつかりません。"
    End If
        
    t品番.SetFocus
    Me!hippo.Visible = False
    Exit Sub
    On Error GoTo 0
err:
    Stop
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    a = KeyCode
End Sub
Private Sub 検索車種リスト_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload UI_search
    If KeyCode <> 13 Then Exit Sub
    If 検索車種リスト.ListIndex = -1 Then Exit Sub
        'gyo
    Dim 番号, gyo As Long, retsu As Long, 製品品番str As String
    
    製品品番str = 検索車種リスト.List(検索車種リスト.ListIndex, 2)
    
    gyo = ActiveSheet.Cells.Find(製品品番str, , , 1).Row
    
    Unload Me
    
    retsu = ActiveCell.Column
    
    Cells(gyo, retsu).Activate
    'ActiveWindow.ScrollColumn = retsu
    ActiveWindow.ScrollRow = gyo
End Sub
Private Function 略図の表示(myVal)
    Dim 画像URL, 部品品番 As String
    Dim 面視 As Long

     If 検索車種リスト.ListIndex = -1 Then
        編集b.Visible = False
        Exit Function
    End If
    '編集b.Visible = True
    部品品番 = 検索車種リスト.List(検索車種リスト.ListIndex, 2)
    
    '面視
    If OptionButton0.Value = True Then
        面視 = 0
    ElseIf OptionButton1.Value = True Then
        面視 = 1
    Else
        OptionButton1.Value = True
        面視 = 1
    End If
    '面視
    If OptionButton2.Value = True Then
    ElseIf OptionButton3.Value = True Then
    Else
        OptionButton2.Value = True
    End If
    With Sheets("設定")
        画像アドレス = .Cells.Find("部材一覧+_", , , 1).Offset(0, 1).Value
    End With
    '略図or写真
    If OptionButton2.Value = True Then
        画像URL = 画像アドレス & "\202_略図\" & 部品品番 & "_" & 面視 & "_" & Format(myVal, "000") & ".emf"
        If Dir(画像URL) <> "" Then
            '対象のファイル数を調べる
            Dim buf As String, cnt As Long
            buf = Dir(画像アドレス & "\202_略図\" & 部品品番 & "_" & 面視 & "_*.emf")
            Do While buf <> ""
                cnt = cnt + 1
                buf = Dir()
            Loop
            'RyakuNo
            RyakuNo.Caption = myVal & "/" & cnt
            '画像の表示
            Ryakuzu.Picture = LoadPicture(画像URL)
            Me!URL = 画像URL
        Else
            Ryakuzu.Picture = LoadPicture(画像アドレス & "\202_略図\NotFound.bmp")
            RyakuNo.Caption = ""
            Me!URL = ""
        End If
    ElseIf OptionButton3.Value = True Then
        'RyakuNo
        RyakuNo.Caption = myVal

        画像URL = 画像アドレス & "\201_写真\" & 部品品番 & "_" & 面視 & "_" & Format(myVal, "000") & ".jpg"
        If Dir(画像URL) <> "" Then
'            Stop
'            On Error Resume Next
'            Me.Hide
'            DoEvents
'            ans = Application.GetOpenFilename(画像URL)
'            WEB.navigate "https://weathernews.jp/onebox/34.72/137.75/temp=c&q=静岡県浜松市中区茄子町&v=d557950e6acf01150531ba2532d9ac7fb4f1d05cb75a0d2f65fdfe5a63cba653"
'            Me.Show
            Ryakuzu.Picture = LoadPicture(画像URL)
            Me.Repaint
            Me!URL = 画像URL
        Else
            Ryakuzu.Picture = LoadPicture(画像アドレス & "\202_略図\NotFound.bmp")
            RyakuNo.Caption = myVal
            Me!URL = ""
        End If
    End If
    Me.Repaint
End Function
Private Sub 検索車種リスト_Click()    'CAV    '
    Call 略図の表示(1)
End Sub
Private Sub OptionButton0_Click()
    Dim temp As String: temp = RyakuNo.Caption
    Dim myVal As Long
    If temp <> "" Then
        If InStr(temp, "/") > 0 Then
            myVal = Left(temp, InStr(temp, "/") - 1)
        Else
            myVal = temp
        End If
    Else
        myVal = 1
    End If
    Call 略図の表示(myVal)
    Me.Repaint
End Sub
Private Sub OptionButton1_Click()
    Dim temp As String: temp = RyakuNo.Caption
    Dim myVal As Long
    If temp <> "" Then
        If InStr(temp, "/") > 0 Then
            myVal = Left(temp, InStr(temp, "/") - 1)
        Else
            myVal = temp
        End If
    Else
        myVal = 1
    End If
    Call 略図の表示(myVal)
    Me.Repaint
End Sub
Private Sub OptionButton2_Click()
    Call 略図の表示(1)
End Sub
Private Sub OptionButton3_Click()
    If OptionButton3 = True Then
'        Me!left.Visible = True
'        Me!right.Visible = True
'        Me!center.Visible = True
    Else
'        Me!left.Visible = False
'        Me!right.Visible = False
'        Me!center.Visible = False
    End If
    Call 略図の表示(1)
End Sub
Private Sub left_Click()
    Dim myVal As Long
    myVal = RyakuNo.Caption
    myVal = myVal + 1
    If myVal > 9 Then myVal = 2
    Call 略図の表示(myVal)
End Sub
Private Sub right_Click()
    Dim myVal As Long
    myVal = RyakuNo.Caption
    myVal = myVal - 1
    If myVal < 2 Then myVal = 9
    Call 略図の表示(myVal)
End Sub
Private Sub center_Click()
    Dim myVal As Long
    Call 略図の表示(1)
End Sub
Private Sub Spin_SpinUp()
    Dim temp As String: temp = RyakuNo.Caption
    If temp = "" Then Exit Sub
    Dim myVal As Long: myVal = Left(temp, InStr(temp, "/") - 1)
    Dim myMax As Long: myMax = Mid(temp, InStr(temp, "/") + 1)
    If myVal < myMax Then
        RyakuNo.Caption = myVal + 1 & "/" & myMax
        Call 略図の表示(myVal + 1)
        Me.Repaint
    End If
End Sub
Private Sub Spin_SpinDown()
    Dim temp As String: temp = RyakuNo.Caption
    If temp = "" Then Exit Sub
    Dim myVal As Long: myVal = Left(temp, InStr(temp, "/") - 1)
    If 1 < myVal Then
        RyakuNo.Caption = myVal - 1 & "/" & myMax
        Call 略図の表示(myVal - 1)
        Me.Repaint
    End If
End Sub
Private Sub Ryakuzu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Me!URL.Caption = "" Then Exit Sub
    'mspaintで開く
    画像URL = Me!URL.Caption
    If InStr(画像URL, ".jpg") > 0 Then
        画像URL = Replace(画像URL, "_jpg", "")
        画像URL = Replace(画像URL, ".jpg", ".png")
    End If
    Shell "C:\WINDOWS\system32\mspaint.exe" & " " & Chr(34) & 画像URL & Chr(34), vbNormalFocus
Exit Sub

    'Dim 画像URL As String
    If 検索車種リスト.ListIndex = -1 Then Exit Sub
    部品品番 = 検索車種リスト.List(検索車種リスト.ListIndex, 1)
    'CAV
    If OptionButton0.Value = True Then
        面視 = 0
    ElseIf OptionButton1.Value = True Then
        面視 = 1
    End If
    'RyakuNo
    Dim temp As String: temp = RyakuNo.Caption
    Dim myVal As Long
    If InStr(temp, "/") > 0 Then
        myVal = Left(temp, InStr(temp, "/") - 1)
    Else
        myVal = temp
    End If
    '略図or写真
    If OptionButton2.Value = True Then
        画像URL = ActiveWorkbook.path & "\部材一覧作成システム_略図\" & 部品品番 & "_" & 面視 & "_" & Format(myVal, "000") & ".emf"
    Else
        画像URL = ActiveWorkbook.path & "\部材一覧作成システム_写真\" & 部品品番 & "_" & 面視 & "_" & myVal & ".bmp"
    End If
    If Dir(画像URL) = "" Then 画像URL = ActiveWorkbook.path & "\部材一覧作成システム_略図\NotFound.bmp"
    'mspaintで開く
    Shell "C:\WINDOWS\system32\mspaint.exe" & " " & Chr(34) & 画像URL & Chr(34), vbNormalFocus
End Sub



