VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_search 
   Caption         =   "����"
   ClientHeight    =   8490
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8730
   OleObjectBlob   =   "UI_search.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








































































'API�t�H�[������
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
'����
    Dim myFrame As MSForms.Control
    Dim myHwnd As Long
    Dim myWindowLong As Long
    Dim myAlpha As Long
    myAlpha = 248 '�����x�i0�`255�̐����l�A0�œ����j
    Set myFrame = Me.Controls.add("Forms.Frame.1")
    myHwnd = GetParent(GetParent(myFrame.[_GethWnd]))
    Me.Controls.Remove myFrame.Name
    Set myFrame = Nothing
    myWindowLong = GetWindowLong(myHwnd, GWL_EXSTYLE)
    myWindowLong = myWindowLong Or WS_EX_LAYERED
    SetWindowLong myHwnd, GWL_EXSTYLE, myWindowLong
    SetLayeredWindowAttributes myHwnd, 0&, myAlpha, LWA_ALPHA
    DrawMenuBar myHwnd '�O�̂���
End Sub

Private Sub t�i��_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 27 Then Unload Me
    
    If KeyCode <> 13 Then Exit Sub
    If Shift = 1 Then
        �����Ԏ탊�X�g.Clear
        t�i�� = ""
        Exit Sub
    End If
    
    Dim myDic As Object, myKey, myItem
    Dim myVal, myVal2, myVal3
    Dim i As Long, x As Long
    Dim lastgyo As Long
    Dim �o�C�g�� As Long
    Dim ���� As String

    ���� = StrConv(t�i��.Value, vbNarrow)
    ���� = UCase(����)
    ���� = Replace(����, "-", "")
    ����str = "���,�H��,���i�i��,���ޏڍ�"
    ����strsp = Split(����str, ",")
    Dim ����RAN()
    ReDim ����RAN(6, 0)
    Dim ����x()
    ReDim ����x(UBound(����strsp))
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws
        Set myKey = .Cells.Find("���i�i��", , , 1)
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For x = 0 To UBound(����strsp)
            ����x(x) = .Cells.Find(����strsp(x), , , 1).Column
        Next x
        
        For y = myKey.Row + 1 To lastRow
            ReDim Preserve ����RAN(6, UBound(����RAN, 2) + 1)
            For x = 0 To UBound(����x)
                ����RAN(x, UBound(����RAN, 2)) = .Cells(y, ����x(x))
            Next x
            ����RAN(5, UBound(����RAN, 2)) = y
        Next y
    End With
    
    �����Ԏ탊�X�g.RowSource = ""
    �����Ԏ탊�X�g.Clear
    Dim C As Long
    For i = LBound(����RAN, 2) + 1 To UBound(����RAN, 2)
        For x = LBound(����strsp) To UBound(����strsp)
            If UCase(StrConv(Replace(����RAN(x, i), "-", ""), vbNarrow)) Like "*" & ���� & "*" Then
                �����Ԏ탊�X�g.AddItem ""
                �����Ԏ탊�X�g.List(C, 0) = ����RAN(0, i)
                �����Ԏ탊�X�g.List(C, 1) = ����RAN(1, i)
                �����Ԏ탊�X�g.List(C, 2) = ����RAN(2, i)
                �����Ԏ탊�X�g.List(C, 3) = ����RAN(3, i)
                C = C + 1
                Exit For
            End If
        Next x
    Next i
    If C > 0 Then
        �����Ԏ탊�X�g.ListIndex = 0
    Else
        �����Ԏ탊�X�g.ListIndex = -1
        �����Ԏ탊�X�g.AddItem ""
        �����Ԏ탊�X�g.List(0, 2) = "�݂���܂���B"
    End If
        
    t�i��.SetFocus
    Me!hippo.Visible = False
    Exit Sub
    On Error GoTo 0
err:
    Stop
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    a = KeyCode
End Sub
Private Sub �����Ԏ탊�X�g_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload UI_search
    If KeyCode <> 13 Then Exit Sub
    If �����Ԏ탊�X�g.ListIndex = -1 Then Exit Sub
        'gyo
    Dim �ԍ�, gyo As Long, retsu As Long, ���i�i��str As String
    
    ���i�i��str = �����Ԏ탊�X�g.List(�����Ԏ탊�X�g.ListIndex, 2)
    
    gyo = ActiveSheet.Cells.Find(���i�i��str, , , 1).Row
    
    Unload Me
    
    retsu = ActiveCell.Column
    
    Cells(gyo, retsu).Activate
    'ActiveWindow.ScrollColumn = retsu
    ActiveWindow.ScrollRow = gyo
End Sub
Private Function ���}�̕\��(myVal)
    Dim �摜URL, ���i�i�� As String
    Dim �ʎ� As Long

     If �����Ԏ탊�X�g.ListIndex = -1 Then
        �ҏWb.Visible = False
        Exit Function
    End If
    '�ҏWb.Visible = True
    ���i�i�� = �����Ԏ탊�X�g.List(�����Ԏ탊�X�g.ListIndex, 2)
    
    '�ʎ�
    If OptionButton0.Value = True Then
        �ʎ� = 0
    ElseIf OptionButton1.Value = True Then
        �ʎ� = 1
    Else
        OptionButton1.Value = True
        �ʎ� = 1
    End If
    '�ʎ�
    If OptionButton2.Value = True Then
    ElseIf OptionButton3.Value = True Then
    Else
        OptionButton2.Value = True
    End If
    With Sheets("�ݒ�")
        �摜�A�h���X = .Cells.Find("���ވꗗ+_", , , 1).Offset(0, 1).Value
    End With
    '���}or�ʐ^
    If OptionButton2.Value = True Then
        �摜URL = �摜�A�h���X & "\202_���}\" & ���i�i�� & "_" & �ʎ� & "_" & Format(myVal, "000") & ".emf"
        If Dir(�摜URL) <> "" Then
            '�Ώۂ̃t�@�C�����𒲂ׂ�
            Dim buf As String, cnt As Long
            buf = Dir(�摜�A�h���X & "\202_���}\" & ���i�i�� & "_" & �ʎ� & "_*.emf")
            Do While buf <> ""
                cnt = cnt + 1
                buf = Dir()
            Loop
            'RyakuNo
            RyakuNo.Caption = myVal & "/" & cnt
            '�摜�̕\��
            Ryakuzu.Picture = LoadPicture(�摜URL)
            Me!URL = �摜URL
        Else
            Ryakuzu.Picture = LoadPicture(�摜�A�h���X & "\202_���}\NotFound.bmp")
            RyakuNo.Caption = ""
            Me!URL = ""
        End If
    ElseIf OptionButton3.Value = True Then
        'RyakuNo
        RyakuNo.Caption = myVal

        �摜URL = �摜�A�h���X & "\201_�ʐ^\" & ���i�i�� & "_" & �ʎ� & "_" & Format(myVal, "000") & ".jpg"
        If Dir(�摜URL) <> "" Then
'            Stop
'            On Error Resume Next
'            Me.Hide
'            DoEvents
'            ans = Application.GetOpenFilename(�摜URL)
'            WEB.navigate "https://weathernews.jp/onebox/34.72/137.75/temp=c&q=�É����l���s����֎q��&v=d557950e6acf01150531ba2532d9ac7fb4f1d05cb75a0d2f65fdfe5a63cba653"
'            Me.Show
            Ryakuzu.Picture = LoadPicture(�摜URL)
            Me.Repaint
            Me!URL = �摜URL
        Else
            Ryakuzu.Picture = LoadPicture(�摜�A�h���X & "\202_���}\NotFound.bmp")
            RyakuNo.Caption = myVal
            Me!URL = ""
        End If
    End If
    Me.Repaint
End Function
Private Sub �����Ԏ탊�X�g_Click()    'CAV    '
    Call ���}�̕\��(1)
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
    Call ���}�̕\��(myVal)
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
    Call ���}�̕\��(myVal)
    Me.Repaint
End Sub
Private Sub OptionButton2_Click()
    Call ���}�̕\��(1)
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
    Call ���}�̕\��(1)
End Sub
Private Sub left_Click()
    Dim myVal As Long
    myVal = RyakuNo.Caption
    myVal = myVal + 1
    If myVal > 9 Then myVal = 2
    Call ���}�̕\��(myVal)
End Sub
Private Sub right_Click()
    Dim myVal As Long
    myVal = RyakuNo.Caption
    myVal = myVal - 1
    If myVal < 2 Then myVal = 9
    Call ���}�̕\��(myVal)
End Sub
Private Sub center_Click()
    Dim myVal As Long
    Call ���}�̕\��(1)
End Sub
Private Sub Spin_SpinUp()
    Dim temp As String: temp = RyakuNo.Caption
    If temp = "" Then Exit Sub
    Dim myVal As Long: myVal = Left(temp, InStr(temp, "/") - 1)
    Dim myMax As Long: myMax = Mid(temp, InStr(temp, "/") + 1)
    If myVal < myMax Then
        RyakuNo.Caption = myVal + 1 & "/" & myMax
        Call ���}�̕\��(myVal + 1)
        Me.Repaint
    End If
End Sub
Private Sub Spin_SpinDown()
    Dim temp As String: temp = RyakuNo.Caption
    If temp = "" Then Exit Sub
    Dim myVal As Long: myVal = Left(temp, InStr(temp, "/") - 1)
    If 1 < myVal Then
        RyakuNo.Caption = myVal - 1 & "/" & myMax
        Call ���}�̕\��(myVal - 1)
        Me.Repaint
    End If
End Sub
Private Sub Ryakuzu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Me!URL.Caption = "" Then Exit Sub
    'mspaint�ŊJ��
    �摜URL = Me!URL.Caption
    If InStr(�摜URL, ".jpg") > 0 Then
        �摜URL = Replace(�摜URL, "_jpg", "")
        �摜URL = Replace(�摜URL, ".jpg", ".png")
    End If
    Shell "C:\WINDOWS\system32\mspaint.exe" & " " & Chr(34) & �摜URL & Chr(34), vbNormalFocus
Exit Sub

    'Dim �摜URL As String
    If �����Ԏ탊�X�g.ListIndex = -1 Then Exit Sub
    ���i�i�� = �����Ԏ탊�X�g.List(�����Ԏ탊�X�g.ListIndex, 1)
    'CAV
    If OptionButton0.Value = True Then
        �ʎ� = 0
    ElseIf OptionButton1.Value = True Then
        �ʎ� = 1
    End If
    'RyakuNo
    Dim temp As String: temp = RyakuNo.Caption
    Dim myVal As Long
    If InStr(temp, "/") > 0 Then
        myVal = Left(temp, InStr(temp, "/") - 1)
    Else
        myVal = temp
    End If
    '���}or�ʐ^
    If OptionButton2.Value = True Then
        �摜URL = ActiveWorkbook.path & "\���ވꗗ�쐬�V�X�e��_���}\" & ���i�i�� & "_" & �ʎ� & "_" & Format(myVal, "000") & ".emf"
    Else
        �摜URL = ActiveWorkbook.path & "\���ވꗗ�쐬�V�X�e��_�ʐ^\" & ���i�i�� & "_" & �ʎ� & "_" & myVal & ".bmp"
    End If
    If Dir(�摜URL) = "" Then �摜URL = ActiveWorkbook.path & "\���ވꗗ�쐬�V�X�e��_���}\NotFound.bmp"
    'mspaint�ŊJ��
    Shell "C:\WINDOWS\system32\mspaint.exe" & " " & Chr(34) & �摜URL & Chr(34), vbNormalFocus
End Sub



