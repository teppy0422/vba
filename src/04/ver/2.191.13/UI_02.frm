VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_02 
   Caption         =   "入力シートの作成"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6990
   OleObjectBlob   =   "UI_02.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




































Private Sub CommandButton1_Click()
    PlaySound "じっこう"
    RLTFサブ = RLTFサブcbx.Value
    Unload UI_02
    Call 製品別端末一覧のシート作成_2009
End Sub

Private Sub CommandButton2_Click()
    PlaySound "じっこう"
    Unload UI_02
    Call 部品リストの作成_Ver2040
    MsgBox "シート[" & ActiveSheet.Name & "] を作成しました。"
End Sub

Private Sub CommandButton3_Click()
    PlaySound "じっこう"
    Unload UI_02
    mytime = ポイント一覧のシート作成_2190
    PlaySound "じっこう"
    Call MsgBox(mytime & "s 作成しました。", vbOKOnly, "ポイント一覧")
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("もどる")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    PlaySound "じっこう"
    Unload Me
    MsgBox 冶具シートの作成
End Sub

Private Sub CommandButton6_Click()
    PlaySound "じっこう"
    Unload Me
    mytime = CAV一覧作成2190
    PlaySound "じっこう"
    Call MsgBox(mytime & "s 作成しました。", vbOKOnly, "CAV一覧")

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub
