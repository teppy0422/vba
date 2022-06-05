VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "製品品番"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum GYO
    typeR = 3
    部位r = 4
    下3r = 5
    製品品番R = 6
End Enum
Public 製品品番db As New Collection

Public 製品品番v As String
Public タイプv As String

Public Function 製品品番Set(book, sheet, 製品品番count)
    Dim sh As Worksheet
    Set sh = Workbooks(book).Sheets(sheet)
    With sh
        Dim i As Long
        For i = 1 To 製品品番count
            With New 製品品番
                .製品品番v = sh.Cells(GYO.製品品番R, i)
                .タイプv = sh.Cells(GYO.typeR, i)
                製品品番db.add .Self
            End With
        Next
    End With
    Set sh = Nothing
End Function

Public Property Get Self() As 製品品番
    Set Self = Me
End Property

Public Sub 製品品番seek(outBookName, outSheetName, GYO, タイプ)
    Dim a As Long: a = 1
    For Each n In 製品品番db
        'If Replace(n.タイプv, " ", "") = タイプ Then
            With Workbooks(outBookName).Sheets(outSheetName)
                .Columns(a + 1).EntireColumn.Insert
                .Cells(GYO, a + 1) = Mid(n.製品品番v, 8, 3)
                .Cells(GYO - 1, a + 1) = Mid(n.製品品番v, 3, 3)
            End With
            a = a + 1
        'End If
    Next n
    '見つからない場合
End Sub


