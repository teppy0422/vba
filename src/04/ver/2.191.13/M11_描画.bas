Attribute VB_Name = "M11_描画"
'スリープイベント
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' 定数の宣言

Public GYO As Long
Public retsu As Long

Public ColorVal() As String     ' カラー値格納用
Public ColorValFont() As Long ' フォント用
Public ColorCode() As Long    ' カラー値格納用
Public ColorName() As String  ' 色記号格納用

Public WhiteLineFrg As Boolean      ' 白線フラグ（塗り色が黒の場合線の色を白にする）
Public 端末 As String
Public 端末図 As String
Public cav As Long
Public 端末cav集合 As String
Public my幅 As Single


Public Function ColorNameToColorValue(color, filcolor, strcolor, fontcolor)
    
        If InStr(color, "/") > 0 Then
            filc = Left(color, InStr(color, "/") - 1)
            strc = Mid(color, InStr(color, "/") + 1)
        Else
            filc = color
            strc = filc
        End If
        
        colorFind = False
        For qq = LBound(ColorCode) To UBound(ColorCode)
            If ColorName(qq) = filc Then
                filcolor = ColorVal(qq)
                fontcolor = ColorValFont(qq)
                colorFind = True
                Exit For
            End If
        Next qq
        If colorFind = False Then Stop '色の登録が無い
        
        If filc <> strc Then
            colorFind = False
            For qq = LBound(ColorCode) To UBound(ColorCode)
                If ColorName(qq) = strc Then
                    strcolor = ColorVal(qq)
                    colorFind = True
                    Exit For
                End If
            Next qq
            If colorFind = False Then Stop '色の登録が無い
        Else
            strcolor = filcolor
        End If
End Function

Sub Init2()
    With myBook.Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        
        ReDim ColorName(lastRow - key.Row)
        ReDim ColorCode(lastRow - key.Row)
        ReDim ColorVal(lastRow - key.Row)
        ReDim ColorValFont(lastRow - key.Row)
        
        For i = key.Row + 1 To lastRow
            ColorName(i - key.Row - 1) = .Cells(i, 1)
            ColorCode(i - key.Row - 1) = .Cells(i, 2)
            ColorVal(i - key.Row - 1) = .Cells(i, 3)
            ColorValFont(i - key.Row - 1) = .Cells(i, 4)
        Next i
    End With

End Sub

Sub Init()
    
    ' *** 初期化処理
    ' カラー値（ＰＣ上での表現値）
    ColorVal(0) = -1                   ' 無し
    ColorVal(1) = RGB(20, 20, 20)      ' 黒
    ColorVal(2) = RGB(252, 252, 252)   ' 白
    ColorVal(3) = RGB(240, 0, 0)       ' 赤
    ColorVal(4) = RGB(0, 186, 84)      ' 緑
    ColorVal(5) = RGB(255, 255, 0)     ' 黄
    ColorVal(6) = RGB(162, 89, 0)      ' 茶
    ColorVal(7) = RGB(0, 110, 255)     ' 青
    ColorVal(8) = RGB(255, 160, 177)   ' ピンク
    ColorVal(9) = RGB(186, 186, 186)   ' 灰
    ColorVal(10) = RGB(170, 255, 0)    ' 若葉
    ColorVal(11) = RGB(101, 226, 255)  ' 空
    ColorVal(12) = RGB(186, 68, 255)   ' 紫
    ColorVal(13) = RGB(255, 130, 17)    ' オレンジ
    ColorVal(14) = RGB(205, 152, 0)    ' チョコレート
    ColorVal(15) = RGB(255, 179, 102)  'ベージュ
    ColorVal(16) = RGB(100, 100, 100)  'ZZ(色が特定出来ない)
    ColorVal(17) = RGB(93, 93, 93)     '深灰
    ColorVal(18) = RGB(173, 173, 173)  '銀灰
    ColorVal(19) = RGB(203, 203, 203)  '銀
    ColorVal(20) = RGB(6, 52, 6)       '深緑
    ColorVal(21) = RGB(255, 239, 143)  'クリーム
    ColorVal(22) = RGB(234, 234, 89)   '黄褐
    ColorVal(23) = RGB(6, 6, 74)       '深青
    ColorVal(24) = RGB(63, 6, 0)       'ダークチョコ
    ColorVal(25) = RGB(214, 182, 65)   '琥珀
    ColorVal(26) = RGB(100, 74, 141)   '紫青
    ColorVal(27) = RGB(184, 101, 204)  'ラベンダー
    ColorVal(28) = RGB(230, 90, 0)     'からし色(空栓で使用)
    ColorVal(29) = RGB(186, 186, 186)      '空栓_実物確認してないからとりあえず赤
    
    ' フォント色
    ColorValFont(0) = -1
    ColorValFont(1) = RGB(255, 255, 255)
    ColorValFont(2) = RGB(0, 0, 0)
    ColorValFont(3) = RGB(255, 255, 255)
    ColorValFont(4) = RGB(255, 255, 255)
    ColorValFont(5) = RGB(0, 0, 0)
    ColorValFont(6) = RGB(255, 255, 255)
    ColorValFont(7) = RGB(255, 255, 255)
    ColorValFont(8) = RGB(0, 0, 0)
    ColorValFont(9) = RGB(0, 0, 0)
    ColorValFont(10) = RGB(0, 0, 0)
    ColorValFont(11) = RGB(0, 0, 0)
    ColorValFont(12) = RGB(255, 255, 255)
    ColorValFont(13) = RGB(0, 0, 0)
    ColorValFont(14) = RGB(255, 255, 255)
    ColorValFont(15) = RGB(0, 0, 0)
    ColorValFont(16) = RGB(0, 0, 0)
    ColorValFont(17) = RGB(255, 255, 255)
    ColorValFont(18) = RGB(0, 0, 0)
    ColorValFont(19) = RGB(0, 0, 0)
    ColorValFont(20) = RGB(255, 255, 255)
    ColorValFont(21) = RGB(0, 0, 0)
    ColorValFont(22) = RGB(0, 0, 0)
    ColorValFont(23) = RGB(255, 255, 255)
    ColorValFont(24) = RGB(255, 255, 255)
    ColorValFont(25) = RGB(0, 0, 0)
    ColorValFont(26) = RGB(255, 255, 255)
    ColorValFont(27) = RGB(0, 0, 0)
    ColorValFont(28) = RGB(255, 0, 0)
    ColorValFont(29) = RGB(0, 0, 0)
    
    ' 色コード（規格上の色コード）
    ColorCode(0) = -1   ' 無し
    ColorCode(1) = 30   ' 黒
    ColorCode(2) = 40   ' 白
    ColorCode(3) = 50   ' 赤
    ColorCode(4) = 60   ' 緑
    ColorCode(5) = 70   ' 黄
    ColorCode(6) = 80   ' 茶
    ColorCode(7) = 90   ' 青
    ColorCode(8) = 52   ' ピンク
    ColorCode(9) = 41   ' 灰
    ColorCode(10) = 61  ' 若葉
    ColorCode(11) = 91  ' 空
    ColorCode(12) = 92  ' 紫
    ColorCode(13) = 51  ' オレンジ
    ColorCode(14) = 81  'チョコレート
    ColorCode(15) = 11
    ColorCode(16) = 0
    ColorCode(17) = 21
    ColorCode(18) = 12
    ColorCode(19) = 16
    ColorCode(20) = 17
    ColorCode(21) = 19
    ColorCode(22) = 20
    ColorCode(23) = 22
    ColorCode(24) = 23
    ColorCode(25) = 24
    ColorCode(26) = 25
    ColorCode(27) = 27
    ColorCode(28) = 99
    ColorCode(29) = 99
    
    ' 色記号
    ColorName(0) = "Notting"    ' 無し
    ColorName(1) = "B"   ' 黒
    ColorName(2) = "W"   ' 白
    ColorName(3) = "R"   ' 赤
    ColorName(4) = "G"   ' 緑
    ColorName(5) = "Y"   ' 黄
    ColorName(6) = "BR"  ' 茶
    ColorName(7) = "L"   ' 青
    ColorName(8) = "P"   ' ピンク
    ColorName(9) = "GY"  ' 灰
    ColorName(10) = "LG" ' 若葉
    ColorName(11) = "SB" ' 空
    ColorName(12) = "V"  ' 紫
    ColorName(13) = "O"  ' オレンジ
    ColorName(14) = "CH" ' チョコレート
    ColorName(15) = "BE"
    ColorName(16) = "ZZ"
    ColorName(17) = "DGY"
    ColorName(18) = "SGY"
    ColorName(19) = "SI"
    ColorName(20) = "DG"
    ColorName(21) = "C"
    ColorName(22) = "TA"
    ColorName(23) = "DL"
    ColorName(24) = "DCH"
    ColorName(25) = "AM"
    ColorName(26) = "VL"
    ColorName(27) = "LA"
    ColorName(28) = "RI" 'からし色(空栓で使用される)
    ColorName(29) = "LY" 'からし色(空栓で使用される)
    
End Sub

Function BoxBaseColor(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, i As Long) As String
    
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' *** 正方形ベースカラー描画
    
    ' 白線フラグ解除
    WhiteLineFrg = False
    
    ' 正方形描画
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.left, ActiveCell.Top, ActiveCell.Height, ActiveCell.Height).Select
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5 * 0.7, 5 * 0.71, 60 * 0.76, 60 * 0.76).Select   '0
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select   '1
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 73 * 0.74, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '2
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 132 * 0.746, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '3
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 194 * 0.745, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '4
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 253 * 0.752, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '5
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 314 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select      '6
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 375 * 0.749, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '7
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 435 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '8
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 496 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '9
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select   '10
    
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (50 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (125 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (185 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (246 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
     
    'xxxx = "12.72.132.193.253.314.375.435.496" '7282-5833
    'xxxx = "50.125.185.246" '7283-2055
    'xxx = Split(xxxx, ".")
    'For Each xx In xxx
        'x = (xx * 0.747) ^ 1.0006
        'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select  'xの座標
    'Next xx
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    'Selection.OnAction = "先後CH"
    ' ベース色が黒だった
    If filcolor = 1315860 Then
        ' 白線フラグセット
        WhiteLineFrg = True
        ' 線の色を白に変更
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    End If
    
    'ダブリの為の処理_同じ名前が無いか確認

    Selection.Name = 端末図 & "_" & cav
    BoxBaseColor = Selection.Name
    
End Function
Function BoxBaseColor2(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, strcolor As Long, i As Long) As String
    
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' *** 正方形ベースカラー描画
    
    ' 白線フラグ解除
    WhiteLineFrg = False
    
    ' 正方形描画
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.left, ActiveCell.Top, ActiveCell.Height, ActiveCell.Height).Select
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5 * 0.7, 5 * 0.71, 60 * 0.76, 60 * 0.76).Select   '0
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select   '1
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 73 * 0.74, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '2
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 132 * 0.746, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '3
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 194 * 0.745, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '4
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 253 * 0.752, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '5
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 314 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select      '6
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 375 * 0.749, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '7
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 435 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '8
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 496 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '9
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select   '10
    
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (50 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (125 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (185 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (246 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
     
    'xxxx = "12.72.132.193.253.314.375.435.496" '7282-5833
    'xxxx = "50.125.185.246" '7283-2055
    'xxx = Split(xxxx, ".")
    'For Each xx In xxx
        'x = (xx * 0.747) ^ 1.0006
        'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select  'xの座標
    'Next xx
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Adjustments.Item(1) = 0.15

    If strcolor = -1 Then
        If filcolor = 13355979 Then '色呼がSIの時
            Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(255, 51, 153), 0
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(255, 102, 51), 0.25
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(255, 255, 0), 0.5
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(1, 167, 143), 0.75
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(51, 102, 255), 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        Else
            'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
            Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.401
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.599
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        End If
    Else
        Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.401
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.599
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.999
        Selection.ShapeRange.Fill.GradientStops.Delete 1
        Selection.ShapeRange.Fill.GradientStops.Delete 1
    End If
    'Selection.OnAction = "先後CH"
    ' ベース色が黒だった
    If filcolor = 1315860 Then
        ' 白線フラグセット
        WhiteLineFrg = True
        ' 線の色を白に変更
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(250, 250, 250)
    Else
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    End If
    Selection.ShapeRange.Line.Weight = 1
    
    'ダブリの為の処理_同じ名前が無いか確認

    Selection.Name = 端末図 & "_" & cav
    BoxBaseColor2 = Selection.Name
    
End Function
Function TerBaseColor2(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, strcolor As Long, i As Long) As String
    
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' *** 正方形ベースカラー描画
    
    ' 白線フラグ解除
    WhiteLineFrg = False
    
    ' 正方形描画
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.left, ActiveCell.Top, ActiveCell.Height, ActiveCell.Height).Select
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5 * 0.7, 5 * 0.71, 60 * 0.76, 60 * 0.76).Select   '0
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select   '1
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 73 * 0.74, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '2
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 132 * 0.746, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '3
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 194 * 0.745, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '4
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 253 * 0.752, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '5
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 314 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select      '6
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 375 * 0.749, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '7
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 435 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '8
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 496 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '9
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select   '10
    
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (50 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (125 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (185 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (246 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
     
    'xxxx = "12.72.132.193.253.314.375.435.496" '7282-5833
    'xxxx = "50.125.185.246" '7283-2055
    'xxx = Split(xxxx, ".")
    'For Each xx In xxx
        'x = (xx * 0.747) ^ 1.0006
        'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select  'xの座標
    'Next xx
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Adjustments.Item(1) = 0.15
    If strcolor = -1 Then
        'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
        Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.401
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.599
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99
        Selection.ShapeRange.Fill.GradientStops.Delete 1
        Selection.ShapeRange.Fill.GradientStops.Delete 1
    Else
        Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.401
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.599
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.999
        Selection.ShapeRange.Fill.GradientStops.Delete 1
        Selection.ShapeRange.Fill.GradientStops.Delete 1
    End If
    
    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    'Selection.OnAction = "先後CH"
    ' ベース色が黒だった
    If filcolor = 1315860 Then
        ' 白線フラグセット
        WhiteLineFrg = True
        ' 線の色を白に変更
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    End If
    
    'ダブリの為の処理_同じ名前が無いか確認

    Selection.Name = 端末図 & "_" & cav
    TerBaseColor2 = Selection.Name
    
End Function


Function BoxStrColor(i As Long, ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long) As String
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' *** 正方形ストライプカラー描画
    xLeft = (xLeft * 0.747) ^ 1.0006
    yTop = (yTop * 0.747) ^ 1.0006
    myWidth = (myWidth * 0.747) ^ 1.0006
    myHeight = (myHeight * 0.747) ^ 1.0006

    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, xLeft + myWidth * 0.011 + myWidth * 0.8, yTop + myHeight * 0.01)
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.993, yTop + myHeight * 0.01
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.993, yTop + myHeight * 0.01 + myHeight * 0.28
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.28, yTop + myHeight * -0.01 + myHeight
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.011, yTop + myHeight * -0.011 + myHeight
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.011, yTop + myHeight * -0.011 + myHeight * 0.8
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.011 + myWidth * 0.8, yTop + myHeight * 0.005
        .ConvertToShape.Select
    End With
    
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Visible = msoFalse
    '線の色を白に変更
    If WhiteLineFrg = True Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    
    'オートシェイプの名前を返す
    Selection.Name = 端末図 & "_" & cav & "_str"
    BoxStrColor = Selection.Name
        
End Function

Sub DeleteAllShapes()
    
    ' *** ワークシート上のすべての図形を消去する
    
    ' エラーが発生してもスキップする
    On Error Resume Next
    
    ' アクティブなワークシートのオートシェイプ等の図形を消去
    ActiveSheet.Shapes.SelectAll
    Selection.ShapeRange.Delete
    
    ' エラーのスキップを解除
    On Error GoTo 0
    
End Sub

Public Function BoxFill(xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, _
                 FilColor1 As Variant, i As Long, Optional FilColor2 As Variant = 0, Optional マルマ1, _
                 Optional シールドフラグ As String, Optional 選択出力 As String, Optional ByVal サイズ呼 As String, _
                 Optional ハメ As String) As String
    
    ' *** 四角形描画
    ' 変数の宣言
    Dim BaseName As String      ' ベース色のオブジェクト名
    Dim StrName As String       ' ストライプ色のオブジェクト名
    Dim clocode1 As Long        ' 色１格納用
    Dim clocode2 As Long        ' 色２格納用
    Dim CloCode3 As Long
    Dim CloCodeFont1 As Long    ' CloCode1に対するフォント色
    Dim BufSize As Single       ' サイズ仮保存用
    Dim sFontSize As Long
    Dim myFontColor As Long
    Dim baseSize As Single

    ハメs = Split(ハメ, "!")
    If InStr(ハメs(5), "金") > 0 Then
        金 = "Au" & vbLf
        金a = 4
    Else
        金 = ""
        金a = 1
    End If
    色呼 = FilColor1
    If FilColor2 <> 0 Then 色呼 = 色呼 & "/" & FilColor2
    Call 色変換(色呼, clocode1, clocode2, clofont)
    
    'If 色で判断 = True Then clocode1 = 16777215: clofont = 0
    If マルマ1 <> "" Then Call 色変換(マルマ1, CloCode3, 0, 0)
    
    ' サイズ値取得
    BufSize = Size
    
    ' サイズが１．２以下の場合は補正する(Excel2000用のバグ対策)
    If BufSize < 1.2 Then BufSize = 2
    BaseName = BoxBaseColor2(xLeft, yTop, myWidth, myHeight, clocode1, clocode2, i)
    'テキストを図形からはみ出して表示する
    Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
    Selection.Font.Name = myFont
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Selection.ShapeRange.TextFrame2.MarginLeft = 0
    Selection.ShapeRange.TextFrame2.MarginRight = 0
    Selection.ShapeRange.TextFrame2.MarginTop = 0
    Selection.ShapeRange.TextFrame2.MarginBottom = 0
    Select Case ハメ図タイプ
    Case "チェッカー用", "回路符号", "構成", "相手端末"
        If InStr(選択出力, "!") = 0 Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 金 & 選択出力 & vbLf & サイズ呼
            '両端先ハメなら1行目にアンダーバー
            If ハメs(1) = "1" Then
                Selection.ShapeRange.TextFrame2.TextRange.Characters(金a, Len(選択出力)).Font.UnderlineStyle = msoUnderlineSingleLine
            End If
            '両端が同じ端子なら1行目が斜体
            If ハメs(2) = "1" Then
                Selection.ShapeRange.TextFrame2.TextRange.Characters(金a, Len(選択出力)).Font.Italic = msoTrue
            End If
        Else
            選択出力A = Split(選択出力, "!")
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 金 & 選択出力A(0) & vbLf & 選択出力A(1)
            選択出力 = 選択出力A(0)
        End If
    
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * 0.48
            gyospace = 0.8
        Else
            sFontSize = Selection.Width * 0.48
            gyospace = 0.8
        End If
        If Len(選択出力) = 4 Then
            sFontSize = sFontSize * 0.87
        End If
        
        myFontColor = clofont
        'ストライプは光彩を使う
        If clocode1 <> clocode2 Or 色で判断 = True Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                If 色で判断 = True Then
                    .color = 16777215 '白
                    .Radius = 11
                Else
                    .color = clocode1
                    .Radius = 8
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clofont
                End If
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
            End With
        Else
            Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clofont
        End If
        
        'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB 'フォントを反対色にする
    
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat.SpaceWithin = 0.1 '行間
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = gyospace  '行間
    Case Else
        If Len(サイズ呼) > 3 Then サイズ呼t = Left(サイズ呼, 3) Else サイズ呼t = サイズ呼
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 金 & サイズ呼t
       
        myFontColor = clofont
        'ストライプは光彩を使う
        If clocode1 <> clocode2 Or 色で判断 = True Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                If 色で判断 = True Then
                    .color = 16777215 '白
                    .Radius = 10
                Else
                    .color = clocode1
                    .Radius = 8
                End If
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
            End With
        End If
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB 'フォントを反対色にする

        If Len(サイズ呼t) < Len(金) - 2 Then サイズ呼t = "Au"
        Select Case Len(サイズ呼t)
            Case 1
                baseSize = 0.8
            Case 2
                baseSize = 0.9
            Case Else
                baseSize = 0.9
        End Select
        
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * (1.6 / Len(Replace(サイズ呼t, ".", ""))) * baseSize
            If InStr(サイズ呼t, ".") > 0 Then sFontSize = sFontSize - (Selection.Height * 0.3)
        Else
            sFontSize = Selection.Width * (1.6 / Len(Replace(サイズ呼t, ".", ""))) * baseSize
            If InStr(サイズ呼t, ".") > 0 Then sFontSize = sFontSize - (Selection.Width * 0.3)
        End If
        
    End Select
    '主に端末経路で文字が小さいから少し大きくする
    If Len(選択出力) = 4 Then
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
    Else
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize * 1.2
    End If
        
    
    If 金 <> "" Then
        myLen = Len(Selection.ShapeRange.TextFrame2.TextRange.Characters.Text)
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 2).Font.Size = sFontSize * 1
        gyospace = 0.7
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 金a).ParagraphFormat.SpaceWithin = gyospace
    End If
    
'    ライン = (Selection.Width + Selection.Height) / 60
'    If ライン < 0.25 Then ライン = 0.25
'    If ライン > 2 Then ライン = 2
    'line.weightは1に固定(部材一覧+のハメ図作成と同じにする)
    Selection.ShapeRange.Line.Weight = 1
    
    ハメs = Split(ハメ, "!")
    
    If 色で判断 = True Then
        For i2 = 1 To UBound(ハメ色設定, 2)
            If ハメs(0) = ハメ色設定(2, i2) Then
                Selection.Font.color = ハメ色設定(1, i2)
                Exit For
            End If
        Next i2
    End If
    
    If ハメs(4) <> "" Then
        If 二重係止flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If
        
    If ハメ作業表現 <> "" Then
        If CLng(ハメs(3)) < CLng(ハメ作業表現) Then
            Select Case ハメ表現
            Case "1"
                Call 先後CH_先ハメ
            Case "2"
                Call 先後CH_後ハメ_2
            Case "3"
                Call 先後CH_後ハメ_3
            End Select
        End If
    Else
        If ハメs(0) = "先ハメ" Then
            Select Case ハメ表現
            Case "1"
                Call 先後CH_先ハメ
            Case "2"
                Call 先後CH_後ハメ_2
            Case "3"
                Call 先後CH_後ハメ_3
            End Select
        End If
    End If
    
    If マルマ1 <> "" Then
        Set obtemp = Selection.ShapeRange
        StrName = BoxFeltTip(xLeft, yTop, myWidth, myHeight, CloCode3, clocode1, myFontColor)
        'グループ化
        obtemp.Select False
        'ActiveSheet.Shapes.Range(BaseName).Select False
        Selection.Group.Select
        Selection.Name = 端末図 & "_" & cav & "_g"
    End If
    
    '自分と同じ名前が2個以上無いか確認_あったらダブリ
    Dim 同じ名前の数 As Long: 同じ名前の数 = 0
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        If objShp.Name = 端末図 & "_" & cav Or objShp.Name = 端末図 & "_" & cav & "_g" Then
            objShp.Select
            同じ名前の数 = 同じ名前の数 + 1
        End If
    Next
    
    'ダブリの画像サイズ変更
    If 同じ名前の数 > 1 Then
        Dim ダブリ1本目 As Long: ダブリ1本目 = 0
        For Each objShp In ActiveSheet.Shapes
            If objShp.Name = Selection.Name Then
                '2行目がポイントかどうか判断
                On Error Resume Next
                zz = objShp.TextFrame2.TextRange.Characters.Text
                If Err <> 0 Then 'マルマとグループ化している場合
                    'objShp.Ungroup
                    For Each objShp2 In objShp.GroupItems
                        'ActiveSheet.Shapes.Range(端末図 & "_" & cav & "_g").Ungroup
                        'ActiveSheet.Shapes.Range(端末図 & "_" & CAV).Ungroup
                        If Not objShp2.Name Like "*_Felt" Then
                            zz = objShp2.TextFrame2.TextRange.Characters.Text
                            Set objShp3 = objShp2
                        End If
                    Next
                Else
                    Set objShp3 = objShp
                End If
                On Error GoTo 0
Return0:
                zzz = InStr(zz, vbLf)
                cc = ""
                If zzz > 0 Then
                    aa = Left(zz, zzz - 1)
                    bb = Replace(Mid(zz, zzz), vbLf, "")
                    If Len(bb) > 3 Then cc = 1 Else cc = 0
                End If
                If ダブリ1本目 = 0 Then
                    objShp.Height = objShp.Height / 2
                    ダブリ1本目 = 1
                    objShp.Name = 端末図 & "_" & cav & "_w1"
                    If cc = "" Then
                        objShp3.TextFrame2.TextRange.Characters.Font.Size = objShp3.TextFrame2.TextRange.Characters.Font.Size / 2
                        objShp3.TextFrame2.TextRange.Characters.Text = zz
                    Else
                        objShp3.TextFrame2.TextRange.Characters.Text = aa
                    End If
                Else
                    objShp.Height = objShp.Height / 2
                    objShp.Top = objShp.Top + objShp.Height
                    objShp.Name = 端末図 & "_" & cav & "_w2"
                    If cc = "" Then
                        objShp3.TextFrame2.TextRange.Characters.Font.Size = objShp3.TextFrame2.TextRange.Characters.Font.Size / 2
                        objShp3.TextFrame2.TextRange.Characters.Text = zz
                    Else
                        objShp3.TextFrame2.TextRange.Characters.Text = aa
                    End If
                End If
            End If
        Next
        ActiveSheet.Shapes.Range(端末図 & "_" & cav & "_w1").Select False
        Selection.Group.Select
        Selection.Name = 端末図 & "_" & cav
    End If

    If 端末cav集合 = "" Then
        端末cav集合 = Selection.Name
    Else
        端末cav集合 = 端末cav集合 & "," & Selection.Name
    End If
Exit Function

End Function

Function TerFill(xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, _
                 FilColor1 As Variant, i As Long, Optional FilColor2 As Variant = 0, Optional マルマ1, Optional シールドフラグ As String) As String
    
    ' *** 四角形描画
    ' 変数の宣言
    Dim BaseName As String      ' ベース色のオブジェクト名
    Dim StrName As String       ' ストライプ色のオブジェクト名
    Dim clocode1 As Long        ' 色１格納用
    Dim clocode2 As Long        ' 色２格納用
    Dim CloCode3 As Long
    Dim BufSize As Single       ' サイズ仮保存用
    Dim sFontSize As Long
    
    色呼 = FilColor1
    If FilColor2 <> 0 Then 色呼 = 色呼 & "/" & FilColor2
    Call 色変換(色呼, clocode1, clocode2, clofont)
    
    If マルマ1 <> "" Then
        Call 色変換(マルマ1, CloCode3, 0, 0)
    Else
        CloCode3 = -1
    End If
    
    ' サイズ値取得
    BufSize = Size
    
    ' サイズが１．２以下の場合は補正する(Excel2000用のバグ対策)
    If BufSize < 1.2 Then BufSize = 2
    ' （ちょっと一言）
    ' Excel2000ってばヒドぃんだって！
    ' マクロのコードから極小サイズのフリーフォームを
    ' 作成しようとするとエラー起こすんだってば！
    ' 「マクロの記録」をしながら極小サイズの
    ' フリーフォームを描画するのは何ら問題ないのに
    ' 記録されたコードを実行してフリーフォームを
    ' 描画しようとすると「オートメーションエラー」
    ' なんてのが発生すんの！！何よこれ！？
    ' ↑
    ' ...
    'CloCode1
    'BaseName = BoxBaseColor(xLeft, yTop, myWidth, myHeight, CloCode1, i)
    BaseName = TerBaseColor2(xLeft, yTop, myWidth, myHeight, clocode1, clocode2, i)
    
    myFontColor = CloCodeFont1 'フォント色をベース色で決める
    'ストライプは光彩を使う
    If clocode2 <> clocode1 Then
        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
            .color = clocode1
            .color.TintAndShade = 0
            .color.Brightness = 0
            .Transparency = 0#
            .Radius = 8
        End With
    End If
'If シールドフラグ = "S" Then
'    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "S"
'    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
'    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
'    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
'    Selection.ShapeRange.TextFrame2.MarginLeft = 0
'    Selection.ShapeRange.TextFrame2.MarginRight = 0
'    Selection.ShapeRange.TextFrame2.MarginTop = 0
'    Selection.ShapeRange.TextFrame2.MarginBottom = 0
'    If Selection.Width > Selection.Height Then
'        sFontSize = Selection.Height * 1.5
'    Else
'        sFontSize = Selection.Width * 1.5
'    End If
'    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
'    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB
'End If
    'CloCode2
'    If CloCode2 >= 0 And CloCode1 <> CloCode2 Then
'        StrName = BoxStrColor(i, xLeft, yTop, myWidth, myHeight, CloCode2)
'        'グループ化
'        ActiveSheet.Shapes.Range(BaseName).Select False
'        Selection.Group.Select
'        'ActiveSheet.Shapes.Range(Array(BaseName, StrName)).Group.Select
'        Selection.Name = 端末図 & "_" & Cav
'        BaseName = Selection.Name
'    End If
    'CloCode
    If CloCode3 >= 0 Then
        StrName = BoxFeltTip(xLeft, yTop, myWidth, myHeight, CloCode3, clocode1, myFontColor)
        'グループ化
        ActiveSheet.Shapes.Range(BaseName).Select False
        Selection.Group.Select
        Selection.Name = 端末図 & "_" & cav
    End If
    
    '自分と同じ名前が2個以上無いか確認_あったらダブリ
    Dim 同じ名前の数 As Long: 同じ名前の数 = 0
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        If objShp.Name = Selection.Name Then
            同じ名前の数 = 同じ名前の数 + 1
        End If
    Next
    
    'ダブリの画像サイズ変更
    If 同じ名前の数 > 1 Then
        Dim ダブリ1本目 As Long: ダブリ1本目 = 0
        For Each objShp In ActiveSheet.Shapes
            If objShp.Name = Selection.Name Then
                If ダブリ1本目 = 0 Then
                    objShp.Width = objShp.Width / 2
                    objShp.Name = Selection.Name & "_w1"
                    ダブリ1本目 = ダブリ1本目 + 1
                ElseIf ダブリ1本目 = 1 Then
                    objShp.Width = objShp.Width / 2
                    objShp.Left = objShp.Left + objShp.Width
                    objShp.Name = Selection.Name & "_w2"
                    ダブリ1本目 = ダブリ1本目 + 1
                Else
                    Stop
                End If
            End If
        Next
        ActiveSheet.Shapes.Range(端末図 & "_" & cav & "_w1").Select False
        Selection.Group.Select
        Selection.Name = 端末図 & "_" & cav
    End If
        
    'line.weightは1に固定(部材一覧+のハメ図作成と同じにする)
    Selection.ShapeRange.Line.Weight = 1
    
    '端子なので電線を最背面に移動
    Selection.ShapeRange.ZOrder msoSendToBack
    
    
    If 端末cav集合 = "" Then
        端末cav集合 = 端末図 & "_" & cav
    Else
        端末cav集合 = 端末cav集合 & "," & 端末図 & "_" & cav
    End If
    ' 作成されたオートシェイプの名前を返す
    'ActiveSheet.Shapes.Range(端末図).Select False
    'Selection.Group.Select
    'Selection.Name = 端末
    'Set target = Union(ActiveSheet.Shapes.Range(端末), ActiveSheet.Shapes.Range(BoxFill))
    'ActiveSheet.Shapes.Range(Array(端末, BoxFill)).Group.Select
    'Selection.Name = 端末図
    'BoxFill = Selection.Name
    
End Function


Function BonFill(xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, Optional RowStr) As String
    
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    
    ' *** 四角形描画
    ' 変数の宣言
    Dim BaseName As String      ' ベース色のオブジェクト名
    Dim StrName As String       ' ストライプ色のオブジェクト名
    Dim clocode1 As Long        ' 色１格納用
    Dim clocode2 As Long        ' 色２格納用
    Dim CloCode3 As Long
    Dim BufSize As Single       ' サイズ仮保存用
    Dim sFontSize As Long


    Dim filc As String, strc As String
    ' データの初期化がされていない時は初期化する
    
    'サイズ値取得
    BufSize = Size
    
    'サイズが1.2以下の場合は補正する(Excel2000用のバグ対策)
    If BufSize < 1.2 Then BufSize = 2
    
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Adjustments.Item(1) = 0.05
    Selection.ShapeRange.Fill.OneColorGradient 1, 1, 1
    
    a = 1 / (UBound(RowStr) + 1)
    For q = LBound(RowStr) To UBound(RowStr)
        V = Split(RowStr(q), "_")
        Call 色変換(V(4), filcolor, strcolor, fontcolor)
        If q = LBound(ColorCode) Then Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0 / (UBound(RowStr) + 1)
'        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.05 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.3 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.301 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.699 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.7 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert fontcolor, 0.999 * a + (q / (UBound(RowStr) + 1))
    Next q
    Selection.ShapeRange.Fill.GradientStops.Delete 1
    Selection.ShapeRange.Fill.GradientStops.Delete 1

    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
        
    'フォント色をベース色で決める
    myFontColor = fontcolor
    'ストライプは光彩を使う
    If clocode1 = clocode2 Then
        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
            .color = clocode1
            .color.TintAndShade = 0
            .color.Brightness = 0
            .Transparency = 0#
            .Radius = 8
        End With
    End If
        
    'lineのサイズ変更
    Selection.ShapeRange.Line.Weight = 1
    
    'bondaなので電線を最背面に移動
    Selection.ShapeRange.ZOrder msoSendToBack
    
    Selection.Name = 端末図 & "_" & cav
    
    If 端末cav集合 = "" Then
        端末cav集合 = 端末図 & "_" & cav
    Else
        端末cav集合 = 端末cav集合 & "," & 端末図 & "_" & cav
    End If
    
End Function

Function CircleBaseColor(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, i As Long) As String
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' 白線フラグ解除
    WhiteLineFrg = False
    
    ' 正円形描画＆色設定
    ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    'Selection.OnAction = "先後CH"
    ' ベース色が黒だった
    If filcolor = 1315860 Then
        ' 白線フラグセット
        WhiteLineFrg = True
        ' 線の色を白に変更
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    End If
    
    ' オートシェイプの名前を返す
    Selection.Name = 端末図 & "_" & cav
    CircleBaseColor = Selection.Name
   
End Function

Function CircleBaseColor2(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, strcolor As Long, i As Long) As String
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' 白線フラグ解除
    WhiteLineFrg = False
    
    ' 正円形描画＆色設定
    ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
'    If 色で判断 = True Then
'
'    Else
        If strcolor = -1 Then
            'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
            Selection.ShapeRange.Fill.OneColorGradient 3, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.401
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.599
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        Else
            Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
            Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.401
            Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.599
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.999
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        End If
        ' ベース色が黒だった
        If filcolor = 1315860 Then
            ' 白線フラグセット
            WhiteLineFrg = True
            ' 線の色を白に変更
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
        Else
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
        End If
'    End If
    Selection.ShapeRange.Line.Weight = 1
    'Selection.OnAction = "先後CH"

    ' オートシェイプの名前を返す
    Selection.Name = 端末図 & "_" & cav
    CircleBaseColor2 = Selection.Name
   
End Function


Function CircleStrColor(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, i As Long) As String
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' *** 正円形ストライプカラー
    Dim セル As Object
    Set セル = ActiveCell
    
    xLeft = (xLeft * 0.747) ^ 1.0006
    yTop = (yTop * 0.747) ^ 1.0006
    myWidth = (myWidth * 0.747) ^ 1.0006
    myHeight = (myHeight * 0.747) ^ 1.0006
    
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, xLeft + myWidth * 0.733, yTop + myHeight * 0.068)
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.818, yTop + myHeight * 0.115, _
                                                     xLeft + myWidth * 0.887, yTop + myHeight * 0.185, _
                                                     xLeft + myWidth * 0.933, yTop + myHeight * 0.265
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.265, yTop + myHeight * 0.933
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.185, yTop + myHeight * 0.887, _
                                                     xLeft + myWidth * 0.115, yTop + myHeight * 0.818, _
                                                     xLeft + myWidth * 0.068, yTop + myHeight * 0.733
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.733, yTop + myHeight * 0.068
        .ConvertToShape.Select
    End With
    
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Visible = msoFalse
    
    '線の色を白に変更
    If WhiteLineFrg = True Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    
    ' オートシェイプの名前を返す
    Selection.Name = 端末図 & "_" & cav & "_str"
    CircleStrColor = Selection.Name
    
End Function
Function BoxNull(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, _
                 ByVal 選択出力 As String, ByVal サイズ呼 As String, ByVal EmptyPlug As String, _
                    ByVal PlugColor As String) As String
    ' 変数の宣言
    Dim BaseName As String      ' ベース色のオブジェクト名
    Dim BufSize As Single       ' サイズ仮保存用
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    
    '空栓がある時
    If EmptyPlug <> "" Then
        'ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "*"
        'Selection.Font.Name = myFont
        sFontSize = Selection.Width * 1.2
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
        Call 色変換(PlugColor, clocode1, clocode2, clofont)
        myFontColor = clocode1
         Selection.ShapeRange.TextFrame2.MarginLeft = 0
         Selection.ShapeRange.TextFrame2.MarginRight = 0
         Selection.ShapeRange.TextFrame2.MarginTop = 0
         Selection.ShapeRange.TextFrame2.MarginBottom = 0
         Selection.ShapeRange.TextFrame2.Orientation = msoTextOrientationHorizontalRotatedFarEast
         Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
         Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
         Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
         Selection.ShapeRange.TextFrame2.TextRange.Font.Line.Visible = True
         Selection.ShapeRange.TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0
    Else '空き
        If InStr(選択出力, "!") = 0 Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 選択出力 & vbCrLf & " "
            If Selection.Width > Selection.Height Then
                sFontSize = Selection.Height * 0.4
                gyospace = 0.8
            Else
                sFontSize = Selection.Width * 0.4
                gyospace = 0.8
            End If
        Else
            選択出力A = Split(選択出力, "!")
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 選択出力A(0) & vbCrLf & 選択出力A(1)
            選択出力 = 選択出力A(0)
            If Selection.Width > Selection.Height Then
                sFontSize = Selection.Height * 0.4
                gyospace = 0.8
            Else
                sFontSize = Selection.Width * 0.4
                gyospace = 0.8
            End If
        End If
        If Len(選択出力) = 4 Then sFontSize = sFontSize * 0.87
        Selection.Font.Name = myFont
        Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(選択出力) + 1, Len(" ") + 1).ParagraphFormat.SpaceWithin = gyospace  '行間
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Selection.ShapeRange.TextFrame2.MarginLeft = 0
        Selection.ShapeRange.TextFrame2.MarginRight = 0
        Selection.ShapeRange.TextFrame2.MarginTop = 0
        Selection.ShapeRange.TextFrame2.MarginBottom = 0
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
        Selection.ShapeRange.Fill.BackColor.RGB = RGB(230, 230, 230)
    End If
           
    'lineのサイズ変更
    ライン = (Selection.Width + Selection.Height) / 60
    If ライン < 0.25 Then ライン = 0.25
    If ライン > 2 Then ライン = 2
    Selection.Name = 端末図 & "_" & cav & ""
    Name1 = Selection.Name
        
    Selection.ShapeRange.Line.Weight = ライン
    
    
    If フォームからの呼び出し = False Then
        If 二重係止flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If
        
    If 端末cav集合 = "" Then
        端末cav集合 = 端末図 & "_" & cav
    Else
        端末cav集合 = 端末cav集合 & "," & 端末図 & "_" & cav
    End If
    
End Function
Function CircleNull(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, _
                    ByVal 選択出力 As String, ByVal サイズ呼 As String, ByVal EmptyPlug As String, _
                    ByVal PlugColor As String) As String

    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    ' 変数の宣言
    Dim BaseName As String      ' ベース色のオブジェクト名
    Dim BufSize As Single       ' サイズ仮保存用
    
    ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
        
    If ハメ図タイプ = "チェッカー用" And 選択出力 <> "" Then
        'フォントサイズを決める
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * 0.32
            gyospace = 0.7
        Else
            sFontSize = Selection.Width * 0.32
            gyospace = 0.8
        End If
        If Len(選択出力) = 4 Then
            sFontSize = sFontSize * 0.87
        End If
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 選択出力 & vbLf & " "
        'Selection.Font.Name = myFont
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
        Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(選択出力) + 1, Len(" ") + 1).ParagraphFormat.SpaceWithin = gyospace  '行間
    End If
    
    If ハメ図タイプ <> "チェッカー用" And CStr(ハメ表現) <> "4" Then
       'ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
       Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "*"
       'Selection.Font.Name = myFont
       sFontSize = Selection.Width * 1.2
       Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
       'Selection.ShapeRange.Fill.UserPicture "D:\18_部材一覧\部材一覧作成システム_パーツ\NullCircle.png"
       Call 色変換(PlugColor, clocode1, clocode2, clofont)
       myFontColor = clocode1
        Selection.ShapeRange.TextFrame2.MarginLeft = 0
        Selection.ShapeRange.TextFrame2.MarginRight = 0
        Selection.ShapeRange.TextFrame2.MarginTop = 0
        Selection.ShapeRange.TextFrame2.MarginBottom = 0
        Selection.ShapeRange.TextFrame2.Orientation = msoTextOrientationHorizontalRotatedFarEast
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        
        Selection.ShapeRange.TextFrame2.TextRange.Font.Line.Visible = True
        Selection.ShapeRange.TextFrame2.TextRange.Font.Line.ForeColor.RGB = clofont
        myred = myFontColor Mod 256
        myGreen = Int(myFontColor / 256) Mod 256
        myBlue = Int(myFontColor / 256 / 256)
        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(myred * 1.1, myGreen * 1.1, myBlue * 1.1)
    Else
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Selection.ShapeRange.TextFrame2.MarginLeft = 0
        Selection.ShapeRange.TextFrame2.MarginRight = 0
        Selection.ShapeRange.TextFrame2.MarginTop = 0
        Selection.ShapeRange.TextFrame2.MarginBottom = 0
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
        Selection.ShapeRange.Fill.BackColor.RGB = RGB(220, 220, 220)
    End If
    'lineのサイズ変更
'    ライン = (Selection.Width + Selection.Height) / 60
'    If ライン < 0.25 Then ライン = 0.25
'    If ライン > 2 Then ライン = 2
    Selection.ShapeRange.Line.Weight = 1
'    Selection.ShapeRange.Line.Weight = ライン
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    Selection.Name = 端末図 & "_" & cav
    
    xLeft = (xLeft * 0.747) ^ 1.0006
    yTop = (yTop * 0.747) ^ 1.0006
    myWidth = (myWidth * 0.747) ^ 1.0006
    myHeight = (myHeight * 0.747) ^ 1.0006
    
    If 端末cav集合 = "" Then
        端末cav集合 = 端末図 & "_" & cav
    Else
        端末cav集合 = 端末cav集合 & "," & 端末図 & "_" & cav
    End If
    
    If フォームからの呼び出し = False Then
        If 二重係止flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If

    '空栓品番の表記
    If EmptyPlug <> "" Then
        空栓追加flg = 0
        For a = 0 To 空栓c
            If 空栓表記(0, a) = EmptyPlug Then
                空栓表記(1, a) = 空栓表記(1, a) + 1
                空栓追加flg = 1
            End If
        Next a
        If 空栓追加flg = 0 Then
            空栓c = 空栓c + 1
            ReDim Preserve 空栓表記(2, 空栓c)
            空栓表記(0, 空栓c) = EmptyPlug
            空栓表記(1, 空栓c) = 1
            空栓表記(2, 空栓c) = PlugColor
        End If
    End If
End Function
Public Function CircleFill(i As Long, xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, _
                 FilColor1 As Variant, Optional FilColor2 As Variant = 0, Optional マルマ1 As String, _
                 Optional シールドフラグ As String, Optional 選択出力 As String, Optional ByVal サイズ呼 As String, _
                 Optional ハメ As String) As String
    ' 変数の宣言
    Dim BaseName As String      ' ベース色のオブジェクト名
    Dim StrName As String       ' ストライプ色のオブジェクト名
    Dim clocode1 As Long        ' 色１格納用
    Dim clocode2 As Long        ' 色２格納用
    Dim CloCode3 As Long
    Dim CloCodeFont1 As Long    ' CloCode1に対するフォント色
    Dim BufSize As Single       ' サイズ仮保存用
    
    ハメs = Split(ハメ, "!")
    
    If InStr(ハメs(5), "金") > 0 Then
        金 = "Au" & vbLf
        金a = 4
    Else
        金 = ""
        金a = 1
    End If
    色呼 = FilColor1
    If FilColor2 <> 0 Then 色呼 = 色呼 & "/" & FilColor2
    Call 色変換(色呼, clocode1, clocode2, clofont)
    
    'If 色で判断 = True Then clocode1 = 16777215: clofont = 0
    
    If マルマ1 <> "" Then Call 色変換(マルマ1, CloCode3, 0, 0)
    
    BufSize = Size
    
    ' サイズが100以下の場合は補正する(小サイズではフリーフォームでの誤差が大きいため)
    'If BufSize < 100 Then BufSize = 100
    
    ' ベース色描画
    'BaseName = CircleBaseColor(xLeft, yTop, myWidth, myHeight, CloCode1, i)
    BaseName = CircleBaseColor2(xLeft, yTop, myWidth, myHeight, clocode1, clocode2, i)
    'テキストを図形からはみ出して表示する
    Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame2.WordWrap = msgfalse
    Selection.Font.Name = myFont
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Selection.ShapeRange.TextFrame2.MarginLeft = 0
    Selection.ShapeRange.TextFrame2.MarginRight = 0
    Selection.ShapeRange.TextFrame2.MarginTop = 0
    Selection.ShapeRange.TextFrame2.MarginBottom = 0
        
    Select Case ハメ図タイプ
    Case "チェッカー用", "回路符号", "構成", "相手端末"

        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 金 & 選択出力 & vbLf & サイズ呼
        '両端先ハメなら1行目にアンダーバー
        If ハメs(1) = "1" Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters(金a, Len(選択出力)).Font.UnderlineStyle = msoUnderlineSingleLine
        End If
        '両端が同じ端子なら1行目が斜体
        If ハメs(2) = "1" Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters(金a, Len(選択出力)).Font.Italic = msoTrue
        End If
        
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * 0.32
            gyospace = 0.6
        Else
            sFontSize = Selection.Width * 0.32
            gyospace = 0.6
        End If
        
        myFontColor = CloCodeFont1 'フォント色をベース色で決める
        'ストライプは光彩を使う
        If clocode1 <> clocode2 Or 色で判断 = True Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                If 色で判断 = True Then
                    .color = 16777215 '白
                    .Radius = 12
                Else
                    .color = clocode1
                    .Radius = 8
                End If
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
            End With
        End If
        
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clofont
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat.SpaceWithin = 0.1 '行間
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(選択出力) + 1, Len(サイズ呼) + 1).ParagraphFormat.SpaceWithin = gyospace  '行間
    Case Else
        If Len(サイズ呼) > 3 Then サイズ呼t = Left(サイズ呼, 3) Else サイズ呼t = サイズ呼
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = 金 & サイズ呼t
        
        myFontColor = clofont 'フォント色をベース色で決める
        'ストライプは光彩を使う
        If clocode2 <> -1 Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                .color = clocode1
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
                .Radius = 8
            End With
        End If
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB 'フォントを反対色にする
        
        Select Case Len(サイズ呼t)
            Case 1
                baseSize = 1.2
            Case Else
                baseSize = 1.1
        End Select
        
        If Selection.Width > Selection.Height Then
            sFontSize = baseSize * Selection.Height * 0.4
            gyospace = 0.7
        Else
            sFontSize = baseSize * Selection.Width * 0.4
            gyospace = 0.8
        End If
    '    If Len(ポイント) = 4 Then
    '        sFontSize = sFontSize * 0.87
    '    End If
        
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat.SpaceWithin = 0.1 '行間
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(ポイント) + 1, Len(サイズ呼) + 1).ParagraphFormat.SpaceWithin = gyoSpace  '行間
    '    Stop
    End Select
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
    If 金 <> "" Then
        myLen = Len(Selection.ShapeRange.TextFrame2.TextRange.Characters.Text)
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 2).Font.Size = sFontSize * 1
        gyospace = 0.7
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 金a).ParagraphFormat.SpaceWithin = gyospace
    End If
        
    'line.weightは1に固定(部材一覧+のハメ図作成と同じにする)
    Selection.ShapeRange.Line.Weight = 1
    
    ハメs = Split(ハメ, "!")
    If 色で判断 = True Then
        For i2 = 1 To UBound(ハメ色設定, 2)
            If ハメs(0) = ハメ色設定(2, i2) Then
                Selection.Font.color = ハメ色設定(1, i2)
                Exit For
            End If
        Next i2
    Else
        If ハメs(0) = "先ハメ" Then
            Select Case ハメ表現
            Case "1"
                Call 先後CH_先ハメ
            Case "2"
                Call 先後CH_後ハメ_2
            Case "3"
                Call 先後CH_後ハメ_3
            End Select
        End If
    End If
    
    If ハメs(4) <> "" Then
        If 二重係止flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If
    
    If マルマ1 <> "" Then
        StrName = CircleFeltTip(xLeft, yTop, myWidth, myHeight, CloCode3, clocode1, myFontColor)
        'グループ化
        ActiveSheet.Shapes.Range(BaseName).Select False
        Selection.Group.Select
        Selection.Name = 端末図 & "_" & cav & "_g"
    End If
    
    If 端末cav集合 = "" Then
        端末cav集合 = Selection.Name
    Else
        端末cav集合 = 端末cav集合 & "," & Selection.Name
    End If
    ' 作成されたオートシェイプの名前を返す
    'CircleFill = Selection.Name
    'ActiveSheet.Shapes.Range(端末図).Select False
    'Selection.Group.Select
    'ActiveSheet.Shapes.Range(Array(端末, CircleFill)).Group.Select
    'Selection.Name = 端末図
    
End Function
Function CircleFeltTip(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, Basecolor As Long, myFontColor) As String
    
    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    
    Dim feltSize As Single
    If myWidth >= myHeight Then
        feltSize = myHeight * 0.4
    Else
        feltSize = myWidth * 0.4
    End If
    
    ' 白線フラグ解除
    WhiteLineFrg = False
    
    ' 正円形描画＆色設定
    If ハメ図タイプ = "チェッカー用" Or ハメ図タイプ = "回路符号" Or ハメ図タイプ = "構成" Then
        feltSize = feltSize * 0.7
        ActiveSheet.Shapes.AddShape(マルマ形状, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.05) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    Else
        ActiveSheet.Shapes.AddShape(マルマ形状, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.05) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    End If
    Selection.ShapeRange.Glow.color.RGB = filcolor
    Selection.ShapeRange.Glow.Radius = 2
    Selection.ShapeRange.Glow.Transparency = 0.5
    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 0.59
    
    'マジック色が黒の時見づらいのでラインを白にする
    If filcolor = 1315860 Then Selection.ShapeRange.Line.ForeColor.RGB = 16777215
    
    ' オートシェイプの名前を返す
    Selection.Name = 端末図 & "_" & cav & "_Felt"
    CircleFeltTip = Selection.Name
    
End Function
Function BoxFeltTip(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, Basecolor As Long, myFontColor) As String

    xLeft = xLeft * my幅
    yTop = yTop * my幅
    myWidth = myWidth * my幅
    myHeight = myHeight * my幅
    Dim feltSize As Single
        
    '自分と同じ名前が2個以上無いか確認_あったらダブリのマルマ
    Dim 同じ名前の数 As Long: 同じ名前の数 = 0
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        If objShp.Name = 端末図 & "_" & cav & "Felt" Then
            同じ名前の数 = 同じ名前の数 + 1
        End If
    Next
    
    If 同じ名前の数 > 1 Then Stop
    
    If myWidth >= myHeight Then
        feltSize = myHeight * 0.38
    Else
        feltSize = myWidth * 0.38
    End If
    ' 白線フラグ解除
    WhiteLineFrg = False

    ' 正円形描画＆色設定
    If ハメ図タイプ = "チェッカー用" Or ハメ図タイプ = "回路符号" Or ハメ図タイプ = "構成" Then
        feltSize = feltSize * 0.7
        ActiveSheet.Shapes.AddShape(マルマ形状, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.02) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    Else
        ActiveSheet.Shapes.AddShape(マルマ形状, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.02) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    End If

    Selection.ShapeRange.Glow.color.RGB = filcolor
    Selection.ShapeRange.Glow.Radius = 2
    Selection.ShapeRange.Glow.Transparency = 0.5
    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 0.7

    'マジック色が黒の時見づらいのでラインを白にする
    If filcolor = 1315860 Then Selection.ShapeRange.Line.ForeColor.RGB = 16777215

    ' オートシェイプの名前を返す
    Selection.Name = 端末図 & "_" & cav & "_Felt"
    BoxFeltTip = Selection.Name

End Function
Function ExCngColor2(TgtName As Variant) As Long
    
    
End Function

Function ExCngColorFont(TgtName As Variant) As Long
    
    ' *** 色コード変換
    
    ' 変数の宣言
    Dim i As Integer                ' 汎用変数
    Dim j As Integer                ' 汎用変数
    Dim BufCloCode As Integer       ' 色コード格納バッファ
    Dim BufCloName As String        ' 色記号格納バッファ
    Dim Errfrg As Boolean           ' エラーフラグ設定用
    
    ' エラーのトラップ
    On Error GoTo ErrSet
    
    ' ターゲットを数値化してみる
    BufCloCode = CInt(TgtName)
    
    ' エラーのトラップを解除
    On Error GoTo 0
    
    ' エラーは発生しなかった
    If Errfrg = False Then
        
        ' 数値は０～２９
        If BufCloCode >= 0 And BufCloCode < 30 Then
            
            ' 通常の色変換
            ExCngColorFont = ColorVal(BufCloCode)
            
        ' 数値は３０以上
        Else
            
            ' ポインタ初期化
            j = 0
            
            ' ループ（コード検索）
            For i = 1 To MaxRng
                
                ' 同一コードが見つかった
                If ColorCode(i) = BufCloCode Then
                    
                    ' ポインタ取得
                    j = i
                    
                    ' ループを抜ける
                    Exit For
                    
                End If
                
            Next
            
            ' コードによる色変換
            ExCngColorFont = ColorVal(j)
            
        End If
        
    ' エラーが発生した（文字による色指定）
    Else
        
        ' 色記号を取得
        BufCloName = TgtName
        
        ' 全角文字の半角化（全角混じりだとうまく変換出来ない為）
        ' BufCloName = WorksheetFunction.Asc(BufCloName)
        
        ' ポインタ初期化
        j = 0
        
        ' ループ（記号検索）
        For i = 1 To MaxRng
            
            ' 同一の記号が見つかった
            If ColorName(i) = BufCloName Then
            
                ' ポインタ取得
                j = i
                
                ' ループを抜ける
                Exit For
            
            End If
            
        Next
        
        ' 記号による色変換
        ExCngColorFont = ColorValFont(j)
    End If
    
    ' このプロシージャを抜ける
    Exit Function
    
' エラー時の処理
ErrSet:
    
    ' エラーフラグセット
    Errfrg = True
    
    ' エラー行の次の行へ
    Resume Next
    
End Function


Public Function ColorMark3(端末, xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, 色呼, 種類, 形状, マルマ1, シールドフラグ, 選択出力, サイズ呼, ハメ, EmptyPlug, PlugColor, RowStr)
'    Dim 端末 As String
'    Dim xLeft As Single
'    Dim yTop As Single
'    Dim myWidth As Single
'    Dim myHeight As Single
'    Dim 色呼 As String
'    Dim 種類 As String
'    dim
    ' 変数の宣言
    Dim i As Long               ' 汎用変数
    Dim j As Long               ' 汎用変数
    Dim LstCnt As Long          ' リスト数取得用
    Dim BufMarkStr As String    ' カラーコード取得用
    Dim BufClo1 As String       ' カラーコード１処理用
    Dim BufClo2 As String       ' カラーコード２処理用
    Dim BufSize As Single       ' サイズ値取得用
    Dim BufRes As String        ' 描画結果取得用
    Dim Errfrg As Boolean       ' エラーフラグ
    Dim lastgyo As Long
    
    'Dim myLastRow As Long: myLastRow = Cells(Rows.Count, 2).End(xlUp).Row
    'Dim 座標範囲 As Range: Set 座標範囲 = Range(Cells(41, 2), Cells(myLastRow, 9))
    'Dim 座標範囲c As Object
    
    '図以外のオブジェクトを削除
    'With Worksheets("座標")
    '    For i = .Shapes.Count To 1 Step -1
    '        If .Shapes(i).Type <> msoPicture Then
    '            If .Shapes(i).Type <> msoOLEControlObject Then
    '                .Shapes(i).Delete
    '            End If
    '        End If
    '    Next i
    'End With
        
    'For Each 座標範囲c In 座標範囲
    '    i = i + 1
        'Dim myProduct As String: myProduct = 座標範囲(i, 1)
        'Dim myCav As Long: myCav = 座標範囲(i, 2)
        'Dim xLeft As Single: xLeft = 座標範囲(i, 3)
        'Dim yTop As Single: yTop = 座標範囲(i, 4)
        'Dim myWidth As Single: myWidth = 座標範囲(i, 5)
        'Dim myHeight As Single: myHeight = 座標範囲(i, 6)
        'Dim 色呼 As String: 色呼称 = 座標範囲(i, 7)
        'Dim 形状 As String: 形状 = 座標範囲(i, 8)
        
        'If myProduct <> myProductBack And myProductBack <> "" Then GoTo line90
        'myProductBack = myProduct
        ' 変数の初期化
        'Call Init
        ハメs = Split(ハメ, "!")
            ' 設定されているカラーコード取得
            BufMarkStr = 色呼
            If ハメ作業表現 <> "" And 色呼 <> "" Then
                If CLng(ハメs(3)) > CLng(ハメ作業表現) Then
                    BufMarkStr = ""
                End If
            End If
            ' 設定値が存在した
            If BufMarkStr <> "" Then
                ' 色コード組合せ数
                If InStr(1, BufMarkStr, "/") = 0 Then
                    BufClo1 = BufMarkStr
                    BufClo2 = "0"
                Else
                    BufClo1 = Left$(BufMarkStr, InStr(1, BufMarkStr, "/") - 1)
                    BufClo2 = Mid$(BufMarkStr, InStr(1, BufMarkStr, "/") + 1)
                End If
                If 形状 = "Box" Then BufRes = BoxFill(xLeft, yTop, myWidth, myHeight, BufClo1, GYO, BufClo2, マルマ1, CStr(シールドフラグ), CStr(選択出力), CStr(サイズ呼), CStr(ハメ))
                If 形状 = "Cir" Then BufRes = CircleFill(GYO, xLeft, yTop, myWidth, myHeight, BufClo1, BufClo2, CStr(マルマ1), CStr(シールドフラグ), CStr(選択出力), CStr(サイズ呼), CStr(ハメ))
                If 形状 = "Ter" Then BufRes = TerFill(xLeft, yTop, myWidth, myHeight, BufClo1, GYO, BufClo2, マルマ1, CStr(シールドフラグ))
                If 形状 = "Bon" Then
                    BufRes = BonFill(xLeft, yTop, myWidth, myHeight, RowStr)
                End If
                ' 実行良好
                If InStr(1, BufRes, "Err") = 0 Then
                    ' カウンタインクリメント
                    j = j + 1
                Else
                    Errfrg = True
                    Stop ' "変換できないカラーコードが存在します。ＯＫボタンをクリックして修正してください。
                    'ActiveCell.Select
                End If
            Else
                If 形状 = "Box" Then BufRes = BoxNull(xLeft, yTop, myWidth, myHeight, CStr(選択出力), CStr(サイズ呼), CStr(EmptyPlug), CStr(PlugColor))
                If 形状 = "Cir" Then BufRes = CircleNull(xLeft, yTop, myWidth, myHeight, CStr(選択出力), CStr(サイズ呼), CStr(EmptyPlug), CStr(PlugColor))
            End If
        If Errfrg = False Then DoEvents
    'Next 座標範囲c

line90:
    'Set 座標範囲 = Nothing
    
End Function '

Sub ハメ図作成_temp()

    'dictionary用宣言
    Dim i As Long
    Dim myDic As Object, myKey, myItem
    Dim myVal, myVal2, myVal3
    Set myDic = CreateObject("Scripting.Dictionary")
    
    'カレントフォルダ変更
    myCurDir = CurDir
    On Error Resume Next
    保存場所a = ActiveWorkbook.Path
    ブック名 = ActiveWorkbook.Name
    CreateObject("WScript.Shell").CurrentDirectory = 保存場所a & "\200_PVSW_RLTF"
    On Error GoTo 0
        
    '対象フォルダ名の取得
    With Application.FileDialog(msoFileDialogFilePicker)
    If .Show = True Then 対象ファイル = .SelectedItems(1)
    対象ファイル名 = Mid(対象ファイル, InStrRev(対象ファイル, "\") + 1, Len(対象ファイル))
    End With
    If 対象ファイル = "" Then End
  
    '■対象フォルダのファイル毎に処理
    Dim objFSO As FileSystemObject  ' FSO
    Set objFSO = New FileSystemObject
    'Set JJF = objFSO.GetFolder(対象ファイル).Files 'CreateObject("Scripting.FileSystemObject")
GoTo line10
    Call 最適化
        ファイルパス = 対象ファイル
        'ファイル名 = JJ.Name
         
        '■各行毎に読み込んで追加
            Dim strfn As String '必要
            strfn = ファイルパス
    
            Dim LngLoop As Long
            Dim IntFlNo As Integer
                
            IntFlNo = FreeFile
            Open strfn For Input As #IntFlNo
            
            Y = 0
                Do Until EOF(IntFlNo)
                    Y = Y + 1
                    Line Input #IntFlNo, aa
                    temp = Split(aa, ",")
                    For X = LBound(temp) To UBound(temp)
                        '■新規ファイルを作成
                        If Y = 1 And X = 0 Then
                            新規ブック名 = "部材一覧作成システム_ハメ図_原紙.xlsm"
                            シート名 = "Sheet1"
                            Workbooks.Open 保存場所a & "\部材一覧作成システム_パーツ\" & 新規ブック名
                            Dim 書式() As String: ReDim 書式(UBound(temp))
                            For xx = LBound(temp) To UBound(temp)
                                If Len(temp(xx)) = 15 Then 書式(xx) = "0": 製品品番count = xx + 1
                                If temp(xx) = "始点側キャビティNo." Then 書式(xx) = "0":
                                If temp(xx) = "終点側キャビティNo." Then 書式(xx) = "0"
                                If temp(xx) = "始点側端末識別子" Then 書式(xx) = "0"
                                If temp(xx) = "終点側端末識別子" Then 書式(xx) = "0"
                                If temp(xx) = "線長" Then 書式(xx) = "0"
                                With Workbooks(新規ブック名).Sheets(シート名)
                                    .Cells(Y, xx + 1).NumberFormat = "@"
                                    .Cells(Y, xx + 1) = temp(xx)
                                End With
                            Next xx
                            Exit For
                        Else
                            With Workbooks(新規ブック名).Sheets(シート名)
                                If 書式(X) = "0" Then
                                    .Cells(Y, X + 1).NumberFormat = 0
                                Else
                                    .Cells(Y, X + 1).NumberFormat = "@"
                                End If
                                .Cells(Y, X + 1) = Replace(temp(X), vbLf, "")
                            End With
                        End If
                    Next X
                Loop
                
        Call 最適化もどす
line10:
                With Workbooks(新規ブック名).Sheets(シート名)
                    Dim タイトル範囲 As Range: Set タイトル範囲 = .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.count).End(xlToLeft).Column))
                    Dim 電線識別名Col As Long: 電線識別名Col = タイトル範囲.Find("電線識別名").Column
                    Dim 電線品種Col As Long: 電線品種Col = タイトル範囲.Find("電線品種").Column
                    Dim 電線サイズCol As Long: 電線サイズCol = タイトル範囲.Find("電線サイズ").Column
                    Dim 電線色Col As Long: 電線色Col = タイトル範囲.Find("電線色").Column
                    Dim 線長Col  As Long: 線長Col = タイトル範囲.Find("線長").Column
                    Dim 始点端末Col As Long: 始点端末Col = タイトル範囲.Find("始点側端末識別子").Column
                    Dim 終点端末Col As Long: 終点端末Col = タイトル範囲.Find("終点側端末識別子").Column
                    Dim 始点cavCol As Long: 始点cavCol = タイトル範囲.Find("始点側キャビティNo.").Column
                    Dim 終点cavCol As Long: 終点cavCol = タイトル範囲.Find("終点側キャビティNo.").Column
                    Dim 始点端末品番Col As Long: 始点端末品番Col = タイトル範囲.Find("始点側端末矢崎品番").Column
                    Dim 終点端末品番Col As Long: 終点端末品番Col = タイトル範囲.Find("終点側端末矢崎品番").Column
                    Dim 始点マルCol As Long: 始点マルCol = タイトル範囲.Find("始点側マルマ色１").Column
                    Dim 終点マルCol As Long: 終点マルCol = タイトル範囲.Find("終点側マルマ色２").Column
                    Set タイトル範囲 = Nothing
                End With
                Worksheets.add after:=Worksheets(Worksheets.count)
                追加シート名 = ActiveSheet.Name
                lastgyo = Y
                For i = 1 To lastgyo
                    With Workbooks(新規ブック名).Sheets(シート名)
                        Set 製品品番範囲 = .Range(.Cells(i, 1), .Cells(i, 製品品番count))
                        電線識別名 = .Cells(i, 電線識別名Col)
                        電線品種 = .Cells(i, 電線品種Col)
                        電線サイズ = .Cells(i, 電線サイズCol)
                        電線色 = .Cells(i, 電線色Col)
                        線長 = .Cells(i, 線長Col)
                        始点端末 = .Cells(i, 始点端末Col)
                        終点端末 = .Cells(i, 終点端末Col)
                        始点cav = .Cells(i, 始点cavCol)
                        終点cav = .Cells(i, 終点cavCol)
                        始点端末品番 = .Cells(i, 始点端末品番Col)
                        終点端末品番 = .Cells(i, 終点端末品番Col)
                        始点マル = .Cells(i, 始点マルCol)
                        終点マル = .Cells(i, 終点マルCol)
                    End With
                    With Workbooks(新規ブック名).Sheets(追加シート名)
                        If i = 1 Then
                            .Range(.Cells(i, 1), .Cells(i, 製品品番count)).Value = 製品品番範囲.Value
                            .Cells(1, 製品品番count + 1) = "電線識別名": .Columns(製品品番count + 1).NumberFormat = "@"
                            .Cells(1, 製品品番count + 2) = "電線品種": .Columns(製品品番count + 2).NumberFormat = "@"
                            .Cells(1, 製品品番count + 3) = "電線サイズ": .Columns(製品品番count + 3).NumberFormat = "@"
                            .Cells(1, 製品品番count + 4) = "電線色": .Columns(製品品番count + 4).NumberFormat = "@"
                            .Cells(1, 製品品番count + 5) = "線長": .Columns(製品品番count + 5).NumberFormat = 0
                            .Cells(1, 製品品番count + 6) = "端末№": .Columns(製品品番count + 6).NumberFormat = 0
                            .Cells(1, 製品品番count + 7) = "cav": .Columns(製品品番count + 7).NumberFormat = 0
                            .Cells(1, 製品品番count + 8) = "マルマ": .Columns(製品品番count + 8).NumberFormat = "@"
                            .Cells(1, 製品品番count + 9) = "部品品番": .Columns(製品品番count + 9).NumberFormat = "@"
                        Else
                            addgyo = .Cells(.Rows.count, 製品品番count + 1).End(xlUp).Row + 1
                            .Range(.Cells(addgyo, 1), .Cells(addgyo + 1, 製品品番count)).Value = 製品品番範囲.Value
                            .Cells(addgyo, 製品品番count + 1) = 電線識別名
                            .Cells(addgyo, 製品品番count + 2) = 電線品種
                            .Cells(addgyo, 製品品番count + 3) = 電線サイズ
                            .Cells(addgyo, 製品品番count + 4) = 電線色
                            .Cells(addgyo, 製品品番count + 5) = 線長
                            .Cells(addgyo, 製品品番count + 6) = 始点端末
                            .Cells(addgyo, 製品品番count + 7) = 始点cav
                            .Cells(addgyo, 製品品番count + 8) = 始点マル
                            .Cells(addgyo, 製品品番count + 9) = 始点端末品番
                            .Cells(addgyo + 1, 製品品番count + 1) = 電線識別名
                            .Cells(addgyo + 1, 製品品番count + 2) = 電線品種
                            .Cells(addgyo + 1, 製品品番count + 3) = 電線サイズ
                            .Cells(addgyo + 1, 製品品番count + 4) = 電線色
                            .Cells(addgyo + 1, 製品品番count + 5) = 線長
                            .Cells(addgyo + 1, 製品品番count + 6) = 終点端末
                            .Cells(addgyo + 1, 製品品番count + 7) = 終点cav
                            .Cells(addgyo + 1, 製品品番count + 8) = 終点マル
                            .Cells(addgyo + 1, 製品品番count + 9) = 終点端末品番
                        End If
                    End With
                Next i
            Stop
                With Workbooks(新規ブック名).Sheets(追加シート名)
                    .Range(.Cells(1, 1), .Cells(addgyo + 1, 製品品番count + 9)).Sort _
                        key1:=Range("j2"), Order1:=xlAscending, _
                        key2:=Range("k2"), order2:=xlAscending, _
                        Header:=xlYes
                    lastgyo = .Cells(.Rows.count, 製品品番count + 1).End(xlUp).Row
                End With
              Stop
                With Workbooks(ブック名).Sheets("座標")
                    Set 座標範囲 = .Range(.Cells(41, 2), .Cells(.Cells(.Rows.count, 2).End(xlUp).Row, 11))
                    座標lastgyo = .Range("b" & .Rows.count).End(xlUp).Row
                End With
                For i = 2 To lastgyo
                    With Workbooks(新規ブック名).Sheets(追加シート名)
                        電線サイズ = .Cells(i, 製品品番count + 3)
                        電線色 = .Cells(i, 製品品番count + 4)
                        端末 = .Cells(i, 製品品番count + 6)
                        cav = .Cells(i, 製品品番count + 7)
                        マルマ = .Cells(i, 製品品番count + 8)
                        部品品番 = .Cells(i, 製品品番count + 9)
                    End With
                    With Workbooks(ブック名).Sheets("座標")
                        For j = 1 To 座標lastgyo
                            If Replace(座標範囲(j, 1), "-", "") = 部品品番 Then Stop
                        Next
                        If 部品品番 <> 部品品番back And 部品品番back <> "" Then
                            '図の読み込み
                            URL = 保存場所a & "\部材一覧作成システム_略図\" & 部品品番 & "_0_00"
                            ActiveSheet.Pictures.Insert("D:\18_部材一覧\部材一覧作成システム_略図\7009-1323_0_001.emf").Select
                        End If
                        部品品番back = 部品品番
                    End With
                Next i
              
                                品番 = Replace(Mid(temp(X), 1, 15), " ", "")
                                設変 = Mid(temp(X), 19, 3)
                            '品番が異なる場合
                            If 品番 <> 品番bak Then GoSub 部品リスト分割
                            With Workbooks(新規ブック名).Sheets(新規シート名) 'タイトル
                                    lastgyo = .Range("a" & .Rows.count).End(xlUp).Row + 1
                                If lastgyo = 2 And .Cells(1, 1) = "" Then
                                    .Range(.Cells(1, 1), .Cells(1, .Columns.count)).NumberFormat = "@"
                                    .Cells(1, 1) = "製品品番"
                                    .Cells(1, 2) = "設変"
                                    .Cells(1, 3) = "部品品番"
                                    .Cells(1, 4) = "-"
                                    .Cells(1, 5) = "ｻｲｽﾞ"
                                    .Cells(1, 7) = "色"
                                    .Cells(1, 8) = "長さ"
                                    .Cells(1, 9) = "箇所数"
                                    .Cells(1, 10) = "作業工程"
                                    .Cells(1, 11) = "ﾋｻﾌﾞ"
                                    .Cells(1, 13) = "呼称"
                                End If
                                '特殊な分析,チューブ類
                                If Mid(temp(X), 27, 1) = "T" Then
                                    .Cells(lastgyo, 1) = Mid(temp(X), 1, 15)
                                    .Cells(lastgyo, 2) = Mid(temp(X), 19, 3)
                                    .Cells(lastgyo, 3) = Mid(temp(X), 375, 8)
                                    .Cells(lastgyo, 4) = Replace(Mid(temp(X), 383, 6), " ", "")
                                    .Cells(lastgyo, 5) = Replace(Mid(temp(X), 389, 4), " ", "")
                                    .Cells(lastgyo, 6) = Replace(Mid(temp(X), 393, 4), " ", "")
                                    If .Cells(lastgyo, 5) = "" And .Cells(lastgyo, 6) = "" Then
                                        サイズ = .Cells(lastgyo, 5) & .Cells(lastgyo, 6)
                                    ElseIf .Cells(lastgyo, 5) <> "" And .Cells(lastgyo, 6) = "" Then
                                        サイズ = "D" & String(3 - Len(.Cells(lastgyo, 5)), " ") & .Cells(lastgyo, 5)
                                    ElseIf .Cells(lastgyo, 5) = "" And .Cells(lastgyo, 6) <> "" Then
                                        サイズ = .Cells(lastgyo, 5)
                                    Else
                                        aaaa = String(3 - Len(Replace(.Cells(lastgyo, 5), ".", "")), " ") & .Cells(lastgyo, 5)
                                        aaaB = String(3 - Len(Replace(.Cells(lastgyo, 6), ".", "")), " ") & .Cells(lastgyo, 6)
                                        サイズ = aaaa & "×" & aaaB
                                    End If
                                    .Cells(lastgyo, 7) = Mid(temp(X), 397, 5)
                                    .Cells(lastgyo, 8) = Mid(temp(X), 403, 5)
                                        長さ = "L=" & String(4 - Len(.Cells(lastgyo, 8)), " ") & .Cells(lastgyo, 8)
                                    .Cells(lastgyo, 9) = 1
                                    .Cells(lastgyo, 10) = Mid(temp(X), 153, 2)
                                    .Cells(lastgyo, 11).NumberFormat = "@"
                                    .Cells(lastgyo, 11) = Left(Mid(temp(X), 544, 4), 1)
                                    .Cells(lastgyo, 12) = Mid(temp(X), 544, 4)
                                    .Cells(lastgyo, 13) = Left(.Cells(lastgyo, 3), 3) & "-" & サイズ & " " & 長さ
                                'コネクタ類の分析
                                ElseIf Mid(temp(X), 27, 1) = "B" Then
                                    For bs = 0 To 9
                                        If Mid(temp(X), 174 + (bs * 20) + 11, 3) = "ATO" Then
                                            lastgyo = .Range("a" & .Rows.count).End(xlUp).Row + 1
                                            .Cells(lastgyo, 1) = Mid(temp(X), 1, 15)
                                            .Cells(lastgyo, 2) = Mid(temp(X), 19, 3)
                                            .Cells(lastgyo, 3) = Replace(Mid(temp(X), 175 + (bs * 20), 10), " ", "")
                                            .Cells(lastgyo, 4) = " "
                                            .Cells(lastgyo, 9) = Val(Mid(temp(X), 175 + (bs * 20) + 15, 4))
                                            .Cells(lastgyo, 10) = Mid(temp(X), 558 + (bs * 2), 2)
                                            .Cells(lastgyo, 11).NumberFormat = "@"
                                            .Cells(lastgyo, 11) = Left(Mid(temp(X), 544, 4), 1)
                                            .Cells(lastgyo, 12) = Mid(temp(X), 544, 4)
                                            Select Case Len(.Cells(lastgyo, 3))
                                            Case 8
                                            .Cells(lastgyo, 13) = Mid(.Cells(lastgyo, 3), 1, 4) & "-" & Mid(.Cells(lastgyo, 3), 5, 4)
                                            Case 10
                                            .Cells(lastgyo, 13) = Mid(.Cells(lastgyo, 3), 1, 4) & "-" & Mid(.Cells(lastgyo, 3), 5, 4) & "-" & Mid(.Cells(lastgyo, 3), 9, 2)
                                            Case Else
                                            .Cells(lastgyo, 13) = .Cells(lastgyo, 3)
                                            End Select
                                        End If
                                    Next bs
                                End If
                            End With
                            If EOF(IntFlNo) Then GoSub 部品リスト分割
                                '■新規ファイルを作成
                                If Y = 1 And X = 0 Then
                                    Workbooks.add
                                    新規ブック名 = ActiveWorkbook.Name
                                    新規シート名 = ActiveSheet.Name
                                End If
                                
                                With Workbooks(新規ブック名).Sheets(新規シート名)
                                '■最初の設定
                                If Y = 書式決定行 Then
                                        .Range(.Cells(1, 1), .Cells(1, .Columns.count)).NumberFormat = 書式行書式
                                        '●書式の設定
                                        '数字に含まれる文字列がある場合は書式=0
                                        Select Case True
                                            Case 数字 Like "*_" & temp(X) & "_*"
                                            If X < 256 Then .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = 0
                                            Case 特殊 Like "*_" & temp(X) & "_*"
                                            If X < 256 Then .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = 特殊書式
                                            Case Else
                                                If X < 256 Then
                                                .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = その他書式
                                                End If
                                        End Select
                                        'LEN15の場合は書式=0
                                        If Len(temp(X)) = 15 And X < 256 Then
                                            .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = 0
                                            品番一覧 = 品番一覧 & Replace(temp(X), " ", "") & vbLf
                                        End If
                                End If
                                    '●値を出力
                                    If X < 256 Then .Cells(Y, X + 1) = temp(X)
                                End With
            If タイトル <> "部品リスト" Then
                Workbooks(新規ブック名).Sheets(新規シート名).Columns.AutoFit
                '■保存
                 保存場所 = 保存場所a & "\002_エクセルデータ\" & 対象フォルダ名 & "\" & 対象フォルダ名 & "_" & タイトル & ".xls"
                 If タイトル = "製品別回路マトリクス" Then ファイル名p = 保存場所
                 If Dir(保存場所a & "\002_エクセルデータ\" & 対象フォルダ名, vbDirectory) = "" Then MkDir (保存場所a & "\002_エクセルデータ\" & 対象フォルダ名)
                 Application.DisplayAlerts = False
                 ActiveWorkbook.SaveAs fileName:=保存場所, FileFormat:=xlExcel8
                 Application.DisplayAlerts = True
                 ActiveWorkbook.Close
             End If
        Close #IntFlNo
        
line20:

    Set objFSO = Nothing
    Set JJF = Nothing
  
    Cells(5, 3) = 品番一覧
    Call 最適化もどす
Exit Sub

部品リスト分割:
With Workbooks(新規ブック名).Sheets(新規シート名)
    '●並べ替え
    .Range("a2:z" & .Range("a" & .Rows.count).End(xlUp).Row).Sort _
    key1:=.Range("m2"), Order1:=xlAscending, _
    Header:=xlGuess
    .Range("a2:z" & .Range("a" & .Rows.count).End(xlUp).Row).Sort _
    key1:=.Range("j2"), Order1:=xlAscending, _
    key2:=.Range("k2"), order2:=xlAscending, _
    key3:=.Range("d2"), Order3:=xlAscending, _
    Header:=xlGuess
    .Columns.AutoFit
End With
        '●工程50
        With Workbooks(新規ブック名).Sheets(新規シート名)
            製品品番 = .Range("a2") & "_" & .Range("b2")
            '●元データを配列に格納
            myVal = .Range("A2", .Range("A" & .Rows.count).End(xlUp)).Resize(, 20).Value
                'myDicへデータを格納
                For i = 1 To UBound(myVal, 1)
                    If myVal(i, 10) = "50" Or myVal(i, 10) = "60" Or myVal(i, 10) = "70" Or myVal(i, 10) = "80" Then
                        myVal2 = myVal(i, 13) & "_" & myVal(i, 10)
                        If Not myDic.exists(myVal2) Then
                            myDic.add myVal2, myVal(i, 9)
                        Else
                            myDic(myVal2) = myDic(myVal2) + myVal(i, 9)
                        End If
                    End If
                Next i
        End With
        
        Workbooks(新規ブック名).Sheets("原紙A3").Copy Workbooks(新規ブック名).Sheets("原紙A3")
        Workbooks(新規ブック名).ActiveSheet.Name = "Auto(サブ別)"
        新規シート名2 = ActiveSheet.Name
        Workbooks(新規ブック名).Sheets(新規シート名2).Range("s4") = "製品品番:    " & 品番bak & "  " & 設変 & "     "
        Workbooks(新規ブック名).Sheets(新規シート名2).Range("ak1") = "作 成 日：'" & Right(Year(Date), 2) & "年 " & Format((Date), "mm") & "月 " & Format(Date, "dd") & "日"
        Workbooks(新規ブック名).Sheets(新規シート名2).Range("ae5") = "※品番別リスト"
        '●Key,Itemの書き出し
        With Workbooks(新規ブック名).Sheets(新規シート名2)
        myKey = myDic.keys
        myItem = myDic.items
            lastcolumn = 19: lastRow = 8: co = 0
            For i = 0 To UBound(myKey)
                myVal3 = Split(myKey(i), "_")
                '50が終わったら次の列
                If 作業工程bak = "50" And 作業工程bak <> myVal3(1) And co <> 0 Then lastcolumn = lastcolumn + 6: lastRow = 8: co = 0
                .Cells(lastRow + co, lastcolumn).Value = myVal3(0)
                .Cells(lastRow + co, lastcolumn + 1).Value = myItem(i)
                Select Case myVal3(1)
                Case "50": cc = 2
                Case "60": cc = 3
                Case "70": cc = 4
                Case "80": cc = 5
                Case Else: Stop  '作業工程が上記以外
                End Select
                .Cells(lastRow + co, lastcolumn + cc).Value = "●"
                co = co + 1
                If co = 30 Then lastcolumn = lastcolumn + 6: lastRow = 8: co = 0
                作業工程bak = myVal3(1)
            Next i
        End With
        'dictionaryを再セット
         Set myDic = Nothing
         Set myDic = CreateObject("Scripting.Dictionary")
                
        '●工程40_サブ順
        With Workbooks(新規ブック名).Sheets(新規シート名)
            '●元データを配列に格納
            myVal = .Range("A2", .Range("A" & .Rows.count).End(xlUp)).Resize(, 20).Value
                'myDicへデータを格納
                For i = 1 To UBound(myVal, 1)
                    If myVal(i, 10) = "40" Then
                        myVal2 = myVal(i, 13) & "_" & myVal(i, 12) & "_" & myVal(i, 10)
                        If Not myDic.exists(myVal2) Then
                            myDic.add myVal2, myVal(i, 9)
                        Else
                            myDic(myVal2) = myDic(myVal2) + myVal(i, 9)
                        End If
                    End If
                Next i
        End With
        
        Workbooks(新規ブック名).Sheets(新規シート名2).Range("b4") = "製品品番:    " & 品番bak & "  " & 設変 & "     "
        Workbooks(新規ブック名).Sheets(新規シート名2).Range("n1") = "作 成 日：'" & Right(Year(Date), 2) & "年 " & Format((Date), "mm") & "月 " & Format(Date, "dd") & "日"
        Workbooks(新規ブック名).Sheets(新規シート名2).Range("j6") = "※サブ別リスト"
        '●Key,Itemの書き出し
        With Workbooks(新規ブック名).Sheets(新規シート名2)
        myKey = myDic.keys
        myItem = myDic.items
            lastcolumn = 2: lastRow = 8: co = 0: cc = 0
            For i = 0 To UBound(myKey)
                myVal3 = Split(myKey(i), "_")
                'サブL1が前回と異なる場合、1行空ける
                If サブL1bak <> Left(myVal3(1), 1) And co <> 0 Then co = co + 1
                '1行空けて30行を超えた場合は次の列
                If co = 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 8: co = 0
                .Cells(lastRow + co, lastcolumn).Value = myVal3(0)
                .Cells(lastRow + co, lastcolumn + 1).Value = myItem(i)
                .Cells(lastRow + co, lastcolumn + 2).Value = Replace(myVal3(1), " ", "")
                co = co + 1
                If co = 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 8: co = 0
                サブL1bak = Left(myVal3(1), 1)
            Next i
        End With
        'dictionaryを再セット
         Set myDic = Nothing
         Set myDic = CreateObject("Scripting.Dictionary")
         
        '●工程40_品番順
        With Workbooks(新規ブック名).Sheets(新規シート名)
            '●並べ替え
            .Range("a2:z" & .Range("a" & .Rows.count).End(xlUp).Row).Sort _
            key1:=.Range("d2"), Order1:=xlAscending, _
            key2:=.Range("m2"), order2:=xlAscending, _
            key3:=.Range("l2"), Order3:=xlAscending, _
            Header:=xlGuess
            .Columns.AutoFit
        End With
        
        Workbooks(新規ブック名).Sheets(新規シート名2).Copy Workbooks(新規ブック名).Sheets(新規シート名2)
        ActiveSheet.Name = "Auto(品番別)"
        新規シート名3 = ActiveSheet.Name
        Workbooks(新規ブック名).Sheets(新規シート名2).Columns("s:ap").Delete
        Workbooks(新規ブック名).Sheets(新規シート名3).Range("b8:p" & Workbooks(新規ブック名).Sheets(新規シート名2).Rows.count).ClearContents
        Workbooks(新規ブック名).Sheets(新規シート名3).Range("j6") = "※品番別リスト"
        With Workbooks(新規ブック名).Sheets(新規シート名)
            '●元データを配列に格納
            myVal = .Range("A2", .Range("A" & .Rows.count).End(xlUp)).Resize(, 20).Value
                'myDicへデータを格納
                For i = 1 To UBound(myVal, 1)
                    If myVal(i, 10) = "40" Then
                        myVal2 = myVal(i, 13) & "_" & myVal(i, 12) & "_" & myVal(i, 10)
                        If Not myDic.exists(myVal2) Then
                            myDic.add myVal2, myVal(i, 9)
                        Else
                            myDic(myVal2) = myDic(myVal2) + myVal(i, 9)
                        End If
                    End If
                Next i
        End With
        
        '●Key,Itemの書き出し
        With Workbooks(新規ブック名).Sheets(新規シート名3)
        myKey = myDic.keys
        myItem = myDic.items
            lastcolumn = 2: lastRow = 7: co = 0: cc = 0
            For i = 0 To UBound(myKey)
                myVal3 = Split(myKey(i), "_")
                'サブL1が前回と異なる場合、1行空ける
                If 呼称L4bak <> Left(myVal3(0), 4) And co <> 0 Then co = co + 1
                '1行空けて30行を超えた場合は次の列
                If co >= 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 7: co = 0
                '呼称の4文字目が-の場合は次の列
                If 呼称M4_1 <> Mid(myVal3(0), 4, 1) And Mid(myVal3(0), 4, 1) = "-" And lastcolumn <> 14 And co <> 0 Then lastcolumn = lastcolumn + 3: lastRow = 7: co = 0
                '部品品番が同じ場合、同じ行にまとめる
                If 呼称bak = myVal3(0) And co <> 0 Then
                .Cells(lastRow + co, lastcolumn + 1) = .Cells(lastRow + co, lastcolumn + 1).Value + myItem(i)
                .Cells(lastRow + co, lastcolumn + 2) = .Cells(lastRow + co, lastcolumn + 2) & "･" & Replace(myVal3(1), " ", "")
                Else
                co = co + 1
                .Cells(lastRow + co, lastcolumn).Value = myVal3(0)
                .Cells(lastRow + co, lastcolumn + 1).Value = .Cells(lastRow + co, lastcolumn + 1).Value + myItem(i)
                .Cells(lastRow + co, lastcolumn + 2).Value = Replace(myVal3(1), " ", "")
                End If
                If co = 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 7: co = 0
                呼称L4bak = Left(myVal3(0), 4)
                呼称bak = myVal3(0)
                呼称M4_1 = Mid(myVal3(0), 4, 1)
            Next i
        End With
        'dictionaryを再セット
         Set myDic = Nothing
         Set myDic = CreateObject("Scripting.Dictionary")
        
        Application.DisplayAlerts = False
        Workbooks(新規ブック名).Sheets("原紙A3").Delete
        Application.DisplayAlerts = True
            '■保存
             保存場所 = 保存場所a & "\002_エクセルデータ\" & 対象フォルダ名 & "\部品リスト\" & 対象フォルダ名 & "_" & タイトル & "_" & 品番bak & "_" & 設変bak & ".xls"
             If タイトル = "製品別回路マトリクス" Then ファイル名p = 保存場所
             If Dir(保存場所a & "\002_エクセルデータ\" & 対象フォルダ名, vbDirectory) = "" Then MkDir (保存場所a & "\002_エクセルデータ\" & 対象フォルダ名)
             If Dir(保存場所a & "\002_エクセルデータ\" & 対象フォルダ名 & "\部品リスト", vbDirectory) = "" Then MkDir (保存場所a & "\002_エクセルデータ\" & 対象フォルダ名 & "\部品リスト")
             
             Application.DisplayAlerts = False
             ActiveWorkbook.SaveAs fileName:=保存場所, FileFormat:=xlExcel8
             Application.DisplayAlerts = True
             ActiveWorkbook.Close
             
             品番bak = 品番
             設変bak = 設変
            
            '■新規ファイルを作成
            If EOF(IntFlNo) = False Then
                新規ブック名 = "単線分析変換システム_原紙.xls"
                新規シート名 = "一覧"
                Workbooks.Open ActiveWorkbook.Path & "\000_原紙\" & 新規ブック名
                'dictionaryを再セット
                Set myDic = Nothing
                Set myDic = CreateObject("Scripting.Dictionary")
            End If
Return


End Sub
Function 先後CH_Ver182_MoreLater()
    
    'Debug.Print Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    Application.StatusBar = Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    If ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor = RGB(255, 80, 80) Then
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight / 6
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(20, 20, 20)
        Dim filcolor As Long
        filcolor = ActiveSheet.Shapes.Range(Application.Caller).Fill.ForeColor
        If filcolor = 1315860 Then ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight * 6
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 80, 80)
        'Selection.Name = Selection.Name & "s"
    End If
    'Selection.Unselect
    'Cells(1, 1).Select
    'ベース色が黒の場合、白にする
        
    'selectを解除する為に写真を選択remake
    Dim myShapeName As String
    Dim a, b, i As Long
    myShapeName = Application.Caller
    For i = 1 To Len(myShapeName)
        If Mid(myShapeName, i, 1) = "_" Then a = a + 1
        If a = 2 Then
            b = i
            Exit For
        End If
    Next i
    myShapeName = Left(myShapeName, b - 1)
    ActiveSheet.Shapes.Range(myShapeName).Select
    
End Function
Function 先後CH()
    
    'Debug.Print Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    'Application.StatusBar = Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    If ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor = RGB(255, 80, 80) Then
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight / 5
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(20, 20, 20)
        Dim filcolor As Long
        filcolor = ActiveSheet.Shapes.Range(Application.Caller).Fill.ForeColor
        If filcolor = 1315860 Then ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight * 5
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 80, 80)
        'Selection.Name = Selection.Name & "s"
    End If
    'Selection.Unselect
    'Cells(1, 1).Select
    'ベース色が黒の場合、白にする
        
    'selectを解除する為に写真を選択remake
    Dim myShapeName As String
    Dim a, b, i As Long
    myShapeName = Application.Caller
    For i = 1 To Len(myShapeName)
        If Mid(myShapeName, i, 1) = "_" Then a = a + 1
        If a = 2 Then
            b = i
            Exit For
        End If
    Next i
    myShapeName = Left(myShapeName, b - 1)
    ActiveSheet.Shapes.Range(myShapeName).Select
    
End Function
Function 先後CH_先ハメ()
    
    'Selection.ShapeRange.Line.DashStyle = 1
'    'Selection.ShapeRange.Line.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
'    'Selection.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line.Weight = 8
'    Selection.ShapeRange.Adjustments.Item(1) = 0.15
'    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
'    Selection.ShapeRange.Line.OneColorGradient msoGradientDiagonalUp, 1, 1
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.4
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.401
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.599
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.6
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.99
'    Selection.ShapeRange.Line.GradientStops.Delete 1
'    Selection.ShapeRange.Line.GradientStops.Delete 1


    Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 250, 250)

    If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
    Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 3
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
    
End Function


Function 先後CH_後ハメ_2()
    With Selection
        .Left = .Left + .Width * 0.25
        .Top = .Top + .Height * 0.25
        .Width = .Width * 0.5
        .Height = .Height * 0.5
        '.Line.Weight = .Line.Weight
        .ShapeRange.TextFrame2.TextRange.Characters.Text = ""
    End With
End Function
Function 先後CH_後ハメ_3()
    With Selection.ShapeRange
        .Fill.Patterned msoPatternWideDownwardDiagonal
        .Fill.ForeColor.RGB = RGB(255, 255, 250)
        .Fill.BackColor.RGB = RGB(20, 20, 20)
        .TextFrame2.TextRange.Characters.Text = ""
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
End Function
Function 先後Make(Optional 画像名 As String)
    If 画像名 = "" Then 画像名 = Application.Caller
    With ActiveWorkbook.Sheets("製品品番")
        Dim 後ハメ図表現 As Long: 後ハメ図表現 = .Range("s4")
        If Not (後ハメ図表現 = 1 Or 後ハメ図表現 = 2) Then MsgBox "Sheets(製品品番)のCells(S4)を1か2で選択してください。": End
    End With
    
    Call 先後描画_SH_season2(画像名)
    Call 先後描画_AH(後ハメ図表現, 画像名)
'Exit Function
    'If ActiveSheet.Name = "PVSW_ハメ図" Then Call 先後CH_season2
    
End Function

Function 先後CH_season2()
    'Debug.Print Application.Caller
    Call 最適化
    
    With ActiveWorkbook.Sheets("PVSW_ハメ図")
        Dim 製品品番点数 As Long: 製品品番点数 = .Cells.Find("端末矢崎品番").Column - 1
    End With
    
    Dim a, b, c, z, i As Long
    Dim yoso
    Dim 端末図 As String: 端末図 = Left(Application.Caller, Len(Application.Caller) - 2)
    Dim 端末図Len As Long: 端末図Len = Len(端末図)
    Dim objShp As Shape
        For Each objShp In ActiveSheet.Shapes
            If 端末図 = Left(objShp.Name, 端末図Len) Then
                a = 0: b = 0: c = 0: z = 0
                For i = 1 To Len(objShp.Name)
                    If Mid(objShp.Name, i, 1) = "_" Then z = z + 1
                    If z = 1 And a = 0 Then a = i
                    If z = 2 And b = 0 Then b = i
                    If z = 3 And c = 0 Then c = i
                Next i
                If InStr(objShp.Name, "w") > 0 Then
                    myCav = Mid(objShp.Name, b + 1, c - b - 1)
                Else
                    myCav = Mid(objShp.Name, b + 1, Len(objShp.Name) - b)
                End If
                If IsNumeric(myCav) Then
                    yoso = Split(objShp.Name, "_")
                    Dim 端末 As String: 端末 = yoso(0)
                    Dim 組合せ As String: 組合せ = yoso(1)
                    Dim cav As String: cav = yoso(2)
                    Dim ハメ順 As String
                    If ActiveSheet.Shapes.Range(objShp.Name).Line.ForeColor.RGB = RGB(255, 80, 80) Then
                        ハメ順 = "先"
                    Else
                        ハメ順 = "後"
                    End If
                    'sheets("PVSW_RLTF両端")への反映
                    With Sheets("PVSW_RLTF両端")
                        Dim 電線識別名Col As Long: 電線識別名Col = .Cells.Find("電線識別名").Column
                        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 5).End(xlUp).Row
                        Dim 端末Col As Long: 端末Col = .Cells.Find("端末識別子").Column
                        Dim cavCol As Long: cavCol = .Cells.Find("キャビティNo.").Column
                        Dim ハメ順Col As Long: ハメ順Col = .Cells.Find("ハメ順").Column
                        Dim X As Long
                        
                        Dim varBinary As Variant: varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                        "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
                        Dim strH As String: strH = 組合せ
                        Dim HtoB As String
                        ReDim strhtob(1 To Len(strH)) As String
                        For i = 1 To Len(strH)
                            strhtob(i) = varBinary(Val("&H" & Mid$(strH, i, 1)))
                        Next i
                        HtoB = Join$(strhtob, vbNullString)
                        組合せ = Right(HtoB, 10)
                        
                        For X = 1 To Len(組合せ)
                            If Mid(組合せ, X, 1) = 1 Then
                                For i = 2 To lastRow
                                    If .Cells(i, 電線識別名Col) <> "" Then
                                        If .Cells(i, 端末Col) = 端末 Then
                                            If .Cells(i, cavCol) = cav Then
                                                .Cells(i, ハメ順Col) = ハメ順
                                            End If
                                        End If
                                    End If
                                Next i
                            End If
                        Next X
                    End With
                    'このシートへの反映
                    With ActiveSheet
                        Dim 構成Col As Long: 構成Col = .Cells.Find("構成").Column
                        Dim 色 As Long
                        端末Col = .Cells.Find("端末№").Column
                        cavCol = .Cells.Find("Cav").Column
                        lastRow = .Cells(.Rows.count, 構成Col).End(xlUp).Row
                        For i = 2 To lastRow
                            If 端末 = .Cells(i, 端末Col) Then
                                If cav = .Cells(i, cavCol) Then
                                    For X = 1 To Len(Right(組合せ, 製品品番点数))
                                        If .Cells(i, X) <> "" Then
                                            If Mid(Right(組合せ, 製品品番点数), X, 1) = Val(Replace(.Cells(i, X), " ", "")) Then
                                                If ハメ順 = "先" Then
                                                    色 = RGB(240, 150, 150)
                                                Else
                                                    色 = RGB(150, 150, 240)
                                                End If
                                                .Cells(i, X).Interior.color = 色
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next i
                    End With
                End If
                'objShp.Duplicate
                'Selection.Name = objShp.Name & "s"
            End If
        Next
        Call 最適化もどす
End Function

Public Function 先後描画_AH(後ハメ図表現 As Long, Optional 画像名 As String)
Dim pTime As Single: pTime = Timer
    Dim 端末図 As String: 端末図 = Left(画像名, Len(画像名) - 2)
    Dim 端末図Len As Long: 端末図Len = Len(端末図)
    Dim myLeft, myTop, myWidth, myHeight As Long
    Dim myLeftM As Long
    Dim myTopM As Long
    Dim myWidthM As Long
    Dim myHieghtM As Long
    Dim myLineWeight As Long
    Dim myFlag As Long
    Dim myFlagName As String
    Dim objShp As Shape
    Dim aa As Boolean
    Dim myFlagRange As Variant
    Dim c As Variant
    Dim skipFlag As Long
    Dim fontDel As Long
    On Error Resume Next
    aa = IsObject(ActiveSheet.Shapes.Range(端末図 & "_AH"))
    On Error GoTo 0
    If aa = True Then ActiveSheet.Shapes.Range(端末図 & "_AH").Delete
        
        For Each objShp In ActiveSheet.Shapes
            If 端末図 = Left(objShp.Name, 端末図Len) And Not objShp.Name Like "*SH*" Then
            'Debug.Print objShp.Name
                skipFlag = 0
                If myFlagName <> "" Then
                    myFlagRange = Split(myFlagName, ",")
                    For Each c In myFlagRange
                        If c <> "" Then
                            If Left(objShp.Name, InStrRev(objShp.Name, "_")) & "e" = c Then
                                skipFlag = 1
                                Exit For
                            End If
                        End If
                    Next c
                    'kkk = InStr(myFlagName, objShp.Name)
                End If
                If skipFlag = 1 And 後ハメ図表現 = 2 Then GoTo line20
                
                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    myFlagName = myFlagName & "," & objShp.Name & "_e"
                End If
                
                fontDel = 0
                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    Select Case 後ハメ図表現
                    Case 1
                        myLeft = objShp.Left + (objShp.Width * 0.25)
                        myTop = objShp.Top + (objShp.Height * 0.25)
                        myWidth = objShp.Width * 0.5
                        myHeight = objShp.Height * 0.5
                        myLineWeight = objShp.Line.Weight / 5
                        fontDel = 1
                    Case 2
                        myLeft = objShp.Left
                        myTop = objShp.Top
                        myWidth = objShp.Width
                        myHeight = objShp.Height
                    End Select
                Else
                    myLeft = objShp.Left
                    myTop = objShp.Top
                    myWidth = objShp.Width
                    myHeight = objShp.Height
                End If
                'If myLeft < myLeftM Or myLeftM = 0 Then myLeftM = myLeft
                'If myTop < myTopM Or myTopM = 0 Then myTopM = myTop
                objShp.Copy
                Sleep 5
                ActiveSheet.Paste
                Selection.Left = myLeft
                Selection.Top = myTop
                Selection.Width = myWidth
                Selection.Height = myHeight
                Selection.Name = Selection.Name & "a"
                If fontDel = 1 Then Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                '対象のタイプがオートシェイプ(1)の時
                If Selection.ShapeRange.Type = 1 Then
                    'テキストがSだった時にフォントサイズを小さくする
                    If Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "S" Then
                        If Selection.Width > Selection.Height Then
                            sFontSize = Selection.Height * 1.05
                        Else
                            sFontSize = Selection.Width * 1.05
                        End If
                        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
                    End If
                End If
                'マルマの色が黒の時、外枠の色を白にする
                If Right(Selection.Name, 5) = "Felta" And 後ハメ図表現 <> 2 Then
                    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
                End If
                mygroup = mygroup & "," & Selection.Name
                
                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    ActiveSheet.Shapes.Range(Selection.Name).Line.Weight = myLineWeight
                    If ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = 1315860 Then
                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(255, 255, 255)
                    Else
                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(20, 20, 20)
                    End If
                    If 後ハメ図表現 = 2 Then
                        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
                        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 250)
                        Selection.ShapeRange.Fill.BackColor.RGB = RGB(20, 20, 20)
                        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                    End If
                End If
                
                'ActiveSheet.Shapes.Range(objShp.Name).Select False
                'myFlagName = ""
            End If
line20:
        Next
        g = Split(mygroup, ",")
        For i = LBound(g) To UBound(g)
            If g(i) <> "" Then
                If 端末図 = Left(g(i), 端末図Len) And g(i) Like "*a*" Then
                    ActiveSheet.Shapes.Range(g(i)).Select False
                End If
            End If
        Next i
        Selection.OnAction = ""
        'Selection.Group.Select
        Selection.Left = Selection.Left + Selection.Width + Selection.Width + 6
        Selection.Name = 端末図 & "_AH"
Selection.Placement = xlMove 'セルに合わせて移動はするがサイズ変更はしない
'Debug.Print Round(Timer - pTime, 2)
End Function

Public Function 先後描画_SH_season2(Optional 画像名 As String)
If 画像名 = "" Then 画像名 = Application.Caller
Dim pTime As Single: pTime = Timer
    Dim 端末図 As String: 端末図 = Left(画像名, Len(画像名) - 2)
    Dim 端末図Len As Long: 端末図Len = Len(端末図)
    Dim myLeft, myTop, myWidth, myHeight As Long
    Dim myLeftM As Long
    Dim myTopM As Long
    Dim myWidthM As Long
    Dim myHieghtM As Long
    Dim myLineWeight As Long
    Dim myFlag As Long
    Dim myFlagName As String
    Dim objShp As Shape
    Dim aa As Boolean
    Dim myFlagRange As Variant
    Dim c As Variant
    Dim skipFlag As Long
    Dim fontDel As Long
    On Error Resume Next
    aa = IsObject(ActiveSheet.Shapes.Range(端末図 & "_SH"))
    On Error GoTo 0
    If aa = True Then ActiveSheet.Shapes.Range(端末図 & "_SH").Delete
        
        For Each objShp In ActiveSheet.Shapes
            If 端末図 = Left(objShp.Name, 端末図Len) And Not objShp.Name Like "*AH*" Then
            'Debug.Print objShp.Name
                skipFlag = 0
                If myFlagName <> "" Then
                    myFlagRange = Split(myFlagName, ",")
                    For Each c In myFlagRange
                        If c <> "" Then
                            If Left(objShp.Name, InStrRev(objShp.Name, "_")) & "e" = c Then
                                skipFlag = 1
                                Exit For
                            End If
                        End If
                    Next c
                    'kkk = InStr(myFlagName, objShp.Name)
                End If
                'If skipFlag = 1 Then GoTo line20
                
                'If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    myFlagName = myFlagName & "," & objShp.Name & "_e"
                'End If
                
                fontDel = 0
                    myLeft = objShp.Left
                    myTop = objShp.Top
                    myWidth = objShp.Width
                    myHeight = objShp.Height
'                End If
                'If myLeft < myLeftM Or myLeftM = 0 Then myLeftM = myLeft
                'If myTop < myTopM Or myTopM = 0 Then myTopM = myTop
                objShp.Copy
                Sleep 5
                ActiveSheet.Paste
                Selection.Left = myLeft
                Selection.Top = myTop
                Selection.Width = myWidth
                Selection.Height = myHeight
                Selection.Name = Selection.Name & "s"
                mygroup = mygroup & "," & Selection.Name
                
'                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
'                    ActiveSheet.Shapes.Range(Selection.Name).Line.Weight = myLineWeight
'                    If ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = 1315860 Then
'                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(255, 255, 255)
'                    Else
'                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(20, 20, 20)
'                    End If
'                    If 後ハメ図表現 = 2 Then
'                        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
'                        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(20, 20, 20)
'                        Selection.ShapeRange.Fill.BackColor.RGB = RGB(255, 255, 255)
'                        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
'                    End If
'                End If
                
                'ActiveSheet.Shapes.Range(objShp.Name).Select False
                'myFlagName = ""
            End If
line20:
        Next
        g = Split(mygroup, ",")
        For i = LBound(g) To UBound(g)
            If g(i) <> "" Then
                If 端末図 = Left(g(i), 端末図Len) And g(i) Like "*s" Then
                    ActiveSheet.Shapes.Range(g(i)).Select False
                End If
            End If
        Next i
        Selection.OnAction = ""
        'Selection.Group.Select
        Selection.Left = Selection.Left + Selection.Width + 3
        Selection.Name = 端末図 & "_SH"
Selection.Placement = xlMove 'セルに合わせて移動はするがサイズ変更はしない
Debug.Print Round(Timer - pTime, 2)
End Function


Public Function 先後描画_SH()
Dim pTime As Single: pTime = Timer
    Dim 端末図 As String: 端末図 = Left(Application.Caller, Len(Application.Caller) - 2)
    Dim 端末図Len As Long: 端末図Len = Len(端末図)
    Dim myLeft, myTop, myWidth, myHeight As Long
    Dim myLeftM As Long
    Dim myTopM As Long
    Dim myWidthM As Long
    Dim myHieghtM As Long
    Dim myLineWeight As Long
    Dim myLineColor As Long
    Dim myFlag(5) As Long
    Dim myFlagName As String
    Dim myFlagName2 As String
    Dim objShp As Shape
    Dim aa As Boolean
    
    '同じ先ハメ図があったら削除
    On Error Resume Next
    aa = IsObject(ActiveSheet.Shapes.Range(端末図 & "_SH"))
    On Error GoTo 0
    If aa = True Then ActiveSheet.Shapes.Range(端末図 & "_SH").Delete
    
    
    For Each objShp In ActiveSheet.Shapes
        If 端末図 = Left(objShp.Name, 端末図Len) Then
        'Debug.Print objShp.Name
            If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                myFlagName = objShp.Name
            End If
            
            For i = LBound(myFlag) To UBound(myFlag)
                myFlag(i) = 0
            Next i
            not後ハメ = 0
            If objShp.Name Like "*Felt*" Then myFlag(0) = 1: not後ハメ = 1
            If objShp.Name Like "*n*" Then myFlag(1) = 1: not後ハメ = 1
            If objShp.Name Like "*t*" Then myFlag(2) = 1: not後ハメ = 1
            If 端末図 = objShp.Name Then myFlag(3) = 1: not後ハメ = 1
            If myFlagName = Left(objShp.Name, Len(myFlagName)) And myFlagName <> "" Then myFlag(4) = 1: not後ハメ = 1
            If not後ハメ = 0 Then
                myFlagName2 = objShp.Name
                GoTo line20
            End If
            If myFlagName2 = Left(objShp.Name, Len(myFlagName2)) And myFlagName2 <> "" Then
                GoTo line20
            End If
            mygroup = mygroup & "," & Selection.Name
            GoTo line20
            
            myLeft = objShp.Left
            myTop = objShp.Top
            myWidth = objShp.Width
            myHeight = objShp.Height
            myLineWeight = objShp.Line.Weight / 6
            objShp.Duplicate.Select
            'objShp.Copy
            'ActiveSheet.Paste
            Selection.Left = myLeft
            Selection.Top = myTop
            Selection.Width = myWidth
            Selection.Height = myHeight
            Selection.Name = Selection.Name & "s"
            mygroup = mygroup & "," & Selection.Name
            If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                ActiveSheet.Shapes.Range(Selection.Name).Line.Weight = myLineWeight
                If ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = 1315860 Then
                    ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(20, 20, 20)
                End If
            End If
            '端末№の時Fillの色変更
            'If objShp.Name Like "*_t" Then
                'ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = RGB(120, 120, 200)
            'End If
            'ActiveSheet.Shapes.Range(objShp.Name).Select False
            'myFlagName = ""
        End If
line20:
    Next
    g = Split(mygroup, ",")
    For i = LBound(g) To UBound(g)
        If g(i) <> "" Then
            If 端末図 = Left(g(i), 端末図Len) And g(i) Like "*s*" Then
                ActiveSheet.Shapes.Range(g(i)).Select False
            End If
        End If
    Next i
    Selection.OnAction = ""
    Selection.Group.Select
    Selection.Left = Selection.Left + Selection.Width + 3
    Selection.Name = 端末図 & "_SH"
    Selection.Placement = xlMove 'セルに合わせて移動はするがサイズ変更はしない
Debug.Print Round(Timer - pTime, 2)
End Function

Function 電線色でセルを塗る(myRow As Long, myCol As Long, 色呼 As String)
        Dim clocode1 As Long        ' 色１格納用
        Dim clocode2 As Long        ' 色２格納用
        Dim CloCode3 As Long
        
        If 色呼 = "SI" Then
            'Stop
        Else
            Call 色変換(色呼, clocode1, clocode2, clofont)
            With Cells(myRow, myCol).Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 45
                .Gradient.ColorStops.Clear
                .Gradient.ColorStops.add(0).color = clocode1
                .Gradient.ColorStops.add(0.4).color = clocode1
                .Gradient.ColorStops.add(0.401).color = clocode2
                .Gradient.ColorStops.add(0.599).color = clocode2
                .Gradient.ColorStops.add(0.6).color = clocode1
                .Gradient.ColorStops.add(1).color = clocode1
            End With
            Cells(myRow, myCol).Font.color = clofont
        End If
    
End Function

Function 色変換(色呼, clocode1, clocode2, clofont)

    Set mysel = Selection
    Dim 色呼a As String, 色呼b As String
    Dim 変換前 As String
    With myBook.Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        色呼 = Replace(色呼, " ", "")
        If InStr(色呼, "/") = 0 Then
            色呼a = 色呼
            色呼b = ""
        Else
            色呼a = Left(色呼, InStr(色呼, "/") - 1)
            色呼b = Mid(色呼, InStr(色呼, "/") + 1)
        End If
        
        If 色呼 = "" Then
            clocode1 = RGB(255, 255, 255)
            clocode2 = RGB(255, 255, 255)
            clofont = RGB(255, 255, 255)
            mysel.Select
            Exit Function
        End If
        '色の登録確認
        検索色 = 色呼a
        Set 検索x = .Columns(key.Column).Find(検索色, , , 1)
        If 検索x Is Nothing Then GoTo errFlg
        
        変換前 = 検索x.Offset(0, 2)
        clocode1s = Split(変換前, ",")
        clocode1 = RGB(clocode1s(0), clocode1s(1), clocode1s(2))
        変換前 = 検索x.Offset(0, 3)
        clofonts = Split(変換前, ",")
        clofont = RGB(clofonts(0), clofonts(1), clofonts(2))
        
        clocode2 = clocode1
        If 色呼b <> "" Then
            '色の登録確認
            検索色 = 色呼b
            Set 検索x = .Columns(key.Column).Find(検索色, , , 1)
            If 検索x Is Nothing Then GoTo errFlg
            
            変換前 = 検索x.Offset(0, 2)
            clocode2s = Split(変換前, ",")
            clocode2 = RGB(clocode2s(0), clocode2s(1), clocode2s(2))
        End If
    End With
    mysel.Select
    色変換 = clocode1
Exit Function
errFlg:
    MsgBox "登録されていない色 " & 色呼a & " を含んでいます。登録してください。"
    Call 最適化もどす
    With myBook.Sheets("color")
        .Select
        .Cells(.Cells(.Rows.count, key.Column).End(xlUp).Row + 1, key.Column) = 検索色
    End With
    
    End
Return
End Function
Function 色変換css(色呼, clocode1, clocode2, clofont)
    Set mysel = Selection
    Dim 色呼a As String, 色呼b As String
    Dim 変換前 As String
    With myBook.Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        色呼 = Replace(色呼, " ", "")
        If InStr(色呼, "/") = 0 Then
            色呼a = 色呼
            色呼b = ""
        Else
            色呼a = Left(色呼, InStr(色呼, "/") - 1)
            色呼b = Mid(色呼, InStr(色呼, "/") + 1)
        End If
        
        If 色呼 = "" Then
            clocode1 = "FFF"
            clocode2 = "FFF"
            clofont = "000"
            mysel.Select
            Exit Function
        End If
        '色の登録確認
        検索色 = 色呼a
        Set 検索x = .Columns(key.Column).Find(検索色, , , 1)
        If 検索x Is Nothing Then GoTo errFlg
        
        変換前 = 検索x.Offset(0, 2)
        clocode1s = Split(変換前, ",")
        clocode1 = Format(Hex(clocode1s(0)), "00") & Format(Hex(clocode1s(1)), "00") & Format(Hex(clocode1s(2)), "00")
        変換前 = 検索x.Offset(0, 3)
        clofonts = Split(変換前, ",")
        clofont = Format(Hex(clofonts(0)), "00") & Format(Hex(clofonts(1)), "00") & Format(Hex(clofonts(2)), "00")
        
        clocode2 = clocode1
        If 色呼b <> "" Then
            '色の登録確認
            検索色 = 色呼b
            Set 検索x = .Columns(key.Column).Find(検索色, , , 1)
            If 検索x Is Nothing Then GoTo errFlg
            
            変換前 = 検索x.Offset(0, 2)
            clocode2s = Split(変換前, ",")
            clocode2 = Hex(clocode2s(0)) & Hex(clocode2s(1)) & Hex(clocode2s(2))
        End If
    End With
    mysel.Select
Exit Function
errFlg:
    MsgBox "登録されていない色 " & 色呼a & " を含んでいます。登録してください。"
    Call 最適化もどす
    With myBook.Sheets("color")
        .Select
        .Cells(.Cells(.Rows.count, key.Column).End(xlUp).Row + 1, key.Column) = 検索色
    End With
    
    End
Return
End Function
