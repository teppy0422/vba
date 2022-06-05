Attribute VB_Name = "Module2"
 Sub CommandButtonを横で2個()
    ActiveSheet.Range("D2").Select
'（*この部分省略*記録マクロから変更2　のコードをそのまま貼り付ける）
    ActiveCell.Offset(0, 3).Activate
    With ActiveSheet.OLEObjects.add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False)
        .Object.Caption = "menu"
        .Object.Font.Name = "メイリオ"
        .Object.Font.Size = 13
        .Width = ActiveCell.Width * 2
        .Height = ActiveCell.Height  'ActiveCell.Height * 2
    End With
End Sub
Sub SetButtonsOnActiveSheet(codeName As String)
    'For Excel VBA
    'Buid by Q11Q From Qiita
    'https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11106096531
    '[マクロでボタンを挿入する - moug](https://www.moug.net/tech/exvba/0150104.html)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim shp As Excel.Shape
    'Dim r As Range: Set r = ws.Range("G9:E9") '連続したセルでも結合したセルでもよい
    Dim r As Range: Set r = ws.Range("c1")
    Dim obj As Object
    Dim shpRng As Excel.ShapeRange
    With r
        Set obj = ws.Buttons.add(.Left, .Top, Application.CentimetersToPoints(3.65), Application.CentimetersToPoints(0.56)) 'プロパティシートはCm単位なので、Cm単位で指定している。しかし実際は誤差が出るため、出力されたボタンのサイズは同じにならない。
        obj.Caption = "MENU" 'Captionでも指定できる
        obj.OnAction = codeName 'マクロの名前を入れる
        'obj.Enable = True '使用可能にする しかしエラーになるため使えない
        obj.PrintObject = False 'False 印刷しない
        obj.Visible = True '可視
        obj.Font.Name = "メイリオ"
        obj.Font.FontStyle = "標準"
        obj.Font.Size = 12
        obj.Font.Underline = xlUnderlineStyleNone
        obj.Locked = True
        obj.LockedText = True
        obj.Width = .Width
        obj.Height = .Height * 2
    End With
    Set shp = ws.Shapes(ws.Shapes.count)
    shp.ZOrder msoBringToFront '最前面に移動
    Set shpRng = ws.Shapes.Range(Array(shp.Name))
    'Stop
End Sub
Sub InsertButtonOnSheet()
' For Excel VBA
' 自動記録を参考に作成したバージョン
'F3にボタンを設置する
'A1に今日の日付を入力するマクロが起動するようにする
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = ActiveSheet
Dim Rng As Range
Dim obj As Object
Dim shp As Excel.Shape
Dim shpRng As Excel.ShapeRange
Set Rng = ws.Range("F3")
Rng.Select
Set obj = ws.Buttons.add(273.75, 40.5, 90.75, 15.75)
With obj
    obj.Characters.Text = "text editing"
    .OnAction = "sampleInsertToday"
    With obj.Characters(Start:=1, Length:=12).Font
        .Name = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
End With
Set shp = ws.Shapes(ws.Shapes.count)
shp.ZOrder msoBringToFront '最前面に移動
Set shpRng = ws.Shapes.Range(Array(shp.Name))
shpRng.Select
    With Selection.Font
        .Name = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 10
        .Strikethrough = True
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
        .AddIndent = False
    End With
    With obj
        .Placement = xlFreeFloating
        .PrintObject = True 'objectが印刷される
    End With
End Sub


Sub sampleInsertToday()
'A1に今日の日付を入力するマクロ
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = ActiveSheet
Dim Rng As Range: Set Rng = ws.Range("A1"): Rng.Activate
Rng.Value = Date
End Sub

