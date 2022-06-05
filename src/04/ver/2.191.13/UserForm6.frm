VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "通知書の取得"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
























































































































Private Sub Label2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    Dim RAN() As String
    ReDim RAN(車種list.ListCount - 1)
    For i = 0 To 車種list.ListCount - 1
        If 車種list.Selected(i) = True Then
            車種str = 車種str & "_" & 車種list.List(i)
        End If
    Next i
    Unload UserForm6
    Call アドレスセット(ActiveWorkbook)
    Call ie_通知書を取得(車種str)
    Call ログ出力("test", "test", "通知書を取得")
End Sub

Private Sub UserForm_Initialize()

    Set myBook = ActiveWorkbook
    
    With myBook.ActiveSheet
        Dim key As Range: Set key = .Cells.Find("key_", , , 1)
        Dim firstCol As Long: firstCol = key.Column + 1
        Dim lastCol As Long: lastCol = .UsedRange.Columns.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim 車種Row As Long: 車種Row = .Cells.Find("車種_", , , 1).Row
        Dim RAN() As String: ReDim RAN(0)
        Dim j As Long
    End With
    
    'リストの設定
    With 車種list
        .RowSource = ""
        .Clear
    End With
    
    With ActiveSheet
        For X = firstCol To lastCol
            車種 = .Cells(車種Row, X)
            flg = False
            For jj = 0 To j - 1
                If .Cells(車種Row, X) = RAN(jj) Then
                    flg = True
                    Exit For
                End If
            Next jj
            '新規追加
            If flg = False Then
                ReDim Preserve RAN(j)
                RAN(j) = 車種
                j = j + 1
            End If
        Next X
    End With
    
    With 車種list
        For i = LBound(RAN) To UBound(RAN)
            If RAN(i) <> "" Then
                車種list.AddItem RAN(i)
            End If
        Next i
    End With
    
End Sub

Private Sub 車種list_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        For i = 0 To 車種list.ListCount - 1
            車種list.Selected(i) = False
        Next i
    End If
End Sub
