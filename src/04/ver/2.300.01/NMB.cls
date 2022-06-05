VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NMBdb As New Collection

Public 製品品番nmB As String
Public 構成nmB As String
Public 品種nmB As String
Public サイズnmB As String
Public サイズ呼nmB As String
Public 色nmB As String
Public 色呼nmB As String
Public 生区nmb As String
Public 線長nmB As Long
Public 行nmb As Long
Public 回符1nmb As String
Public 回符2nmb As String
Public 端子11nmb As String
Public 端子21nmb As String
Public 端末1nmb As String
Public 端末2nmb As String
Public Felt1nmb As String
Public Felt2nmb As String
Public titleNmb As Range

Public Function NMBset(book, sheet)
    Dim sh As Worksheet
    Set sh = Workbooks(book).Sheets(sheet)
    With Workbooks(book).Sheets(sheet)
        Dim nmbLastRow As Long
        nmbLastRow = .UsedRange.Rows.count
        Dim nmbRow As Long: nmbRow = .Cells.Find("No.").Row
        Dim nmbCol As Long: nmbCol = .Cells.Find("No.").Column
        Set titleNmb = .Range(.Cells(nmbRow, nmbCol), .Cells(nmbRow, .Columns.count))
        製品c = titleNmb.Find("製品", , , xlWhole).Column
        設変c = titleNmb.Find("設変", , , xlWhole).Column
        構成c = titleNmb.Find("構成", , , xlWhole).Column
        品種c = titleNmb.Find("品種", , , xlWhole).Column
        サイズc = titleNmb.Find("ｻｲｽﾞ", , , xlWhole).Column
        サイズ呼c = titleNmb.Find("サ呼", , , xlWhole).Column
        色c = titleNmb.Find("色", , , xlWhole).Column
        生区c = titleNmb.Find("生区", , , xlWhole).Column
        線長c = titleNmb.Find("線長", , , xlWhole).Column
        色呼c = titleNmb.Find("色呼", , , xlWhole).Column
        回符1c = titleNmb.Find("回符1", , , xlWhole).Column
        回符2c = titleNmb.Find("回符2", , , xlWhole).Column
        端子11c = titleNmb.Find("部品11", , , xlWhole).Column
        端子21c = titleNmb.Find("部品21", , , xlWhole).Column
        端末1c = titleNmb.Find("端末1", , , xlWhole).Column
        端末2c = titleNmb.Find("端末2", , , xlWhole).Column
        Felt1c = titleNmb.Find("ﾏ呼1", , , xlWhole).Column
        Felt2c = titleNmb.Find("ﾏ呼2", , , xlWhole).Column
    End With

    Dim i As Long
    For i = 2 To nmbLastRow
        With New NMB
            .製品品番nmB = sh.Cells(i, 製品c)
            .構成nmB = sh.Cells(i, 構成c)
            .品種nmB = sh.Cells(i, 品種c)
            .サイズnmB = sh.Cells(i, サイズc)
            .サイズ呼nmB = sh.Cells(i, サイズ呼c)
            .色nmB = sh.Cells(i, 色c)
            .色呼nmB = sh.Cells(i, 色呼c)
            .生区nmb = sh.Cells(i, 生区c)
            .線長nmB = sh.Cells(i, 線長c)
            .行nmb = sh.Cells(i).Row
            .回符1nmb = sh.Cells(i, 回符1c)
            .回符2nmb = sh.Cells(i, 回符2c)
            .端子11nmb = sh.Cells(i, 端子11c)
            .端子21nmb = sh.Cells(i, 端子21c)
            .端末1nmb = sh.Cells(i, 端末1c)
            .端末2nmb = sh.Cells(i, 端末2c)
            .Felt1nmb = sh.Cells(i, Felt1c)
            .Felt2nmb = sh.Cells(i, Felt2c)
            NMBdb.add .Self
        End With
    Next
    Set sh = Nothing
End Function

Public Property Get Self() As NMB
    Set Self = Me
End Property

Public Sub NMBseek(product, cons, ByRef found)
    For Each n In NMBdb
        If n.構成nmB = cons Then
            If Replace(n.製品品番nmB, " ", "") = product Then
                品種nmB = n.品種nmB
                サイズnmB = n.サイズnmB
                サイズ呼nmB = n.サイズ呼nmB
                色nmB = n.色nmB
                色呼nmB = n.色呼nmB
                線長nmB = n.線長nmB
                生区nmb = n.生区nmb
                行nmb = n.行nmb
                回符1nmb = n.回符1nmb
                回符2nmb = n.回符2nmb
                端子11nmb = n.端子11nmb
                端子21nmb = n.端子21nmb
                端末1nmb = n.端末1nmb
                端末2nmb = n.端末2nmb
                Felt1nmb = n.Felt1nmb
                Felt2nmb = n.Felt2nmb
                found = 1
                Exit Sub
            End If
        End If
    Next n
    '見つからない場合
    'Stop
    found = 0
End Sub

