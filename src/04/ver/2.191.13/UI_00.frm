VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_00 
   Caption         =   "1.データインポート"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   OleObjectBlob   =   "UI_00.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private Sub CommandButton1_Click()
    Unload UI_00
    PlaySound ("じっこう")
    Call PVSWcsv_csvのインポート_2029
    PlaySound ("じっこう")
    Call ログ出力("test", "test", "PVSWインポート実行")
    MsgBox "[PVSW_RLTF]へのPVSW.csvのインポートが完了しました。"
    Sheets("製品品番").Activate
End Sub

Private Sub CommandButton2_Click()
    PlaySound ("じっこう")
    RLTFサブ = RLTFサブcbx.Value
    Unload UI_00
    Call PVSWcsvにRLTFAから回路条件取得_Ver2026
    Call PVSWcsvにRLTFBから回路条件取得
    PlaySound ("かんせい")
    Call ログ出力("test", "test", "RLTFインポート実行")
    MsgBox "取得が完了しました。"
    Sheets("PVSW_RLTF").Activate
End Sub

Private Sub CommandButton3_Click()
    PlaySound ("じっこう")
    Unload UI_00
    Sheets("PVSW_RLTF").Activate
    Sleep 10
    'Call 最適化
    Call PVSWcsvの共通化_Ver1944
    Call PVSW_RLTFのサブ0に他製品のサブを割り当てる_2048
    'Call 最適化もどす
    PlaySound ("かんせい")
    Call ログ出力("test", "test", "PVSW_RLTF最適化")
    MsgBox "処理が完了しました。"
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("もどる")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub
