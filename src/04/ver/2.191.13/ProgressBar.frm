VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "処理中"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False































Public flgStop As Boolean
Private Sub CommandButton1_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub stopBtn_Click()
    Call 最適化もどす
    End
End Sub

Private Sub UserForm_Initialize()
    ProgressBar.startupposition = 2
    With ProgBar0
        .min = 0
        .Max = 100
        .Value = 0
    End With
    msg.Caption = ""
    msg0.Caption = ""
End Sub
