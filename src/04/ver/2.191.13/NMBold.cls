VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NMBold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ディレクトリ As String
Public ブック名 As String
Public シート名 As String

Public Sub Class_Initialize()
    Workbooks.Open Left(ActiveWorkbook.Path, InStrRev(ActiveWorkbook.Path, "\")) & nmbDirName & "\" & nmbBookName
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbStartGyo As Long: nmbStartGyo = 2
        Dim nmbLastGyo As Long: nmbLastGyo = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim nmbタイトルRamge As Range: Set nmbタイトルRange = .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.count).End(xlToLeft).Column))
    End With
    For Each r In nmbタイトルRange
        Dim nmb製品品番Retsu As Long: If r = "製品" Then nmb製品品番Retsu = r.Column
        Dim nmb構成Retsu As Long: If r = "構成" Then nmb構成Retsu = r.Column
        Dim nmb電線品種Retsu As Long: If r = "品種" Then nmb電線品種Retsu = r.Column
        Dim nmb電線サイズRetsu As Long: If r = "ｻｲｽﾞ" Then nmb電線サイズRetsu = r.Column
        Dim nmb電線色Retsu As Long: If r = "色" Then nmb電線色Retsu = r.Column
        Dim nmb電線線長Retsu As Long: If r = "線長" Then nmb電線線長Retsu = r.Column
    Next r
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmb製品品番Range As Range: Set nmb製品品番Range = .Range(.Cells(nmbStartGyo, nmb製品品番Retsu), .Cells(nmbLastGyo, nmb製品品番Retsu))
        Dim nmb構成Range As Range: Set nmb構成Range = .Range(.Cells(nmbStartGyo, nmb構成Retsu), .Cells(nmbLastGyo, nmb構成Retsu))
    End With
End Sub

Public Property Get 条件抽出(製品品番, 構成)
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
    For n = nmbStartGyo To nmslastgyo
        If .Cells(n, nmb製品品番Retsu) = 製品品番 Then
            If .Cells(n, nmb構成Retsu) = 構成 Then
                電線線長 = .Cells(n, nmb電線線長Retsu).Value
                Exit Property
            End If
        End If
    Next n
    End With
End Property
