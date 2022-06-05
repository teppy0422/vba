VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_91 
   Caption         =   "通知書の取得"
   ClientHeight    =   4305
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   2890
   OleObjectBlob   =   "UI_91.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_91"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



















































































































































































































Private Sub Label2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    Dim ran() As String
    ReDim ran(車種list.ListCount - 1)
    For i = 0 To 車種list.ListCount - 1
        If 車種list.Selected(i) = True Then
            車種str = 車種str & "_" & 車種list.List(i)
        End If
    Next i
    Unload UI_91
    Call addressSet(ActiveWorkbook)
    Call ie_通知書を取得(車種str)
    Call ログ出力("test", "test", "通知書を取得")
End Sub

Private Sub UserForm_Initialize()

    Set wb(0) = ThisWorkbook
    myIP = GetIPAddress
    addressSet wb(0)
    
    With wb(0).ActiveSheet
        Dim key As Range: Set key = .Cells.Find("製品品番", , , 1)
        Dim firstCol As Long: firstCol = key.Column + 1
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Dim 車種Row As Long: 車種Row = .Cells.Find("型式", , , 1).Row
        Dim ran() As String: ReDim ran(0)
        Dim j As Long
    End With
    
    'リストの設定
    With 車種list
        .RowSource = ""
        .Clear
    End With
    
    With ActiveSheet
        For x = firstCol To lastCol
            車種 = .Cells(車種Row, x)
            flg = False
            For jj = 0 To j - 1
                If .Cells(車種Row, x) = ran(jj) Then
                    flg = True
                    Exit For
                End If
            Next jj
            '新規追加
            If flg = False Then
                ReDim Preserve ran(j)
                ran(j) = 車種
                j = j + 1
            End If
        Next x
    End With
    
    With 車種list
        For i = LBound(ran) To UBound(ran)
            車種list.AddItem ran(i)
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
