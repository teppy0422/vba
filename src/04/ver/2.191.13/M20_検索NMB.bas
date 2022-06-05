Attribute VB_Name = "M20_検索NMB"
Public sheetArray As Variant
Public 製品c As Long, 構成c As Long, 品種c As Long, サイズc As Long, 色c As Long, 品種呼c As Long, サイズ呼c As Long, 色呼c As Long, 線長c As Long, 生区c As Long, JCDFc As Long, 製品品番c As Long
Public 回符1c As Long, 部品11c As Long, 端末1c As Long, getFelt1c As Long, 回符2c As Long, 部品21c As Long, 端末2c As Long, getFelt2c As Long
Public 製品val As String, 構成val As String, 品種val As String, サイズval As String, 色val As String, 品種呼val As String, サイズ呼val As String, 色呼val As String, 線長val As String, 生区val As String, JCDFval As String, 製品品番val As String
Public 回符1val As String, 部品11val As String, 端末1val As String, getFelt1val As String, 回符2val As String, 部品21val As String, 端末2val As String, getFelt2val As String

Public Function NMBset(book, sheet)
    Dim maxRow As Long, maxCol As Long
    Dim nmbTitle As Range
    With Workbooks(book).Sheets(sheet)
        maxRow = .Cells(.Rows.count, 1).End(xlUp).Row
        maxCol = .Cells(1, 1).End(xlToRight).Column
        sheetArray = .Range(.Cells(1, 1), .Cells(maxRow, maxCol))
        Set nmbTitle = .Range(.Cells(1, 1), .Cells(1, maxCol))
        '電線用
        製品c = nmbTitle.Find("製品", , , xlWhole).Column: If 製品c = 0 Then Stop
        構成c = nmbTitle.Find("構成", , , xlWhole).Column: If 構成c = 0 Then Stop
        品種c = nmbTitle.Find("品種", , , xlWhole).Column: If 品種c = 0 Then Stop
        サイズc = nmbTitle.Find("ｻｲｽﾞ", , , xlWhole).Column: If サイズc = 0 Then Stop
        色c = nmbTitle.Find("色", , , xlWhole).Column: If 色c = 0 Then Stop
        品種呼c = nmbTitle.Find("品呼", , , xlWhole).Column: If 品種呼c = 0 Then Stop
        サイズ呼c = nmbTitle.Find("サ呼", , , xlWhole).Column: If サイズ呼c = 0 Then Stop
        色呼c = nmbTitle.Find("色呼", , , xlWhole).Column: If 色呼c = 0 Then Stop
        線長c = nmbTitle.Find("線長", , , xlWhole).Column: If 線長c = 0 Then Stop
        生区c = nmbTitle.Find("生区", , , xlWhole).Column: If 生区c = 0 Then Stop
        JCDFc = nmbTitle.Find("JCDF", , , xlWhole).Column: If JCDFc = 0 Then Stop
        '電線端末用
        回符1c = nmbTitle.Find("回符1", , , xlWhole).Column: If 回符1c = 0 Then Stop
        部品11c = nmbTitle.Find("部品11", , , xlWhole).Column: If 部品11c = 0 Then Stop
        端末1c = nmbTitle.Find("端末1", , , xlWhole).Column: If 端末1c = 0 Then Stop
        getFelt1c = nmbTitle.Find("ﾏ呼1", , , xlWhole).Column: If getFelt1c = 0 Then Stop
        回符2c = nmbTitle.Find("回符2", , , xlWhole).Column: If 回符2c = 0 Then Stop
        部品21c = nmbTitle.Find("部品21", , , xlWhole).Column: If 部品21c = 0 Then Stop
        端末2c = nmbTitle.Find("端末2", , , xlWhole).Column: If 端末2c = 0 Then Stop
        getFelt2c = nmbTitle.Find("ﾏ呼2", , , xlWhole).Column: If getFelt2c = 0 Then Stop
    End With
    Set nmbTitle = Nothing
End Function

Public Function NMBrelease()
    sheetArray = ""
End Function

Public Function NMBseek_電線(product, cons, ByRef found, z)
    Dim i As Long
    found = 0
    For i = 2 To UBound(sheetArray)
        If Replace(sheetArray(i, 構成c), " ", "") = cons Then
            If Replace(sheetArray(i, 製品c), " ", "") = product Then
                製品品番val = sheetArray(i, 製品c)
                品種val = sheetArray(i, 品種c)
                サイズval = sheetArray(i, サイズc)
                色val = sheetArray(i, 色c)
                品種呼val = Replace(sheetArray(i, 品種呼c), " ", "")
                サイズ呼val = Replace(sheetArray(i, サイズ呼c), " ", "")
                色呼val = Replace(sheetArray(i, 色呼c), " ", "")
                線長val = sheetArray(i, 線長c)
                生区val = Replace(sheetArray(i, 生区c), " ", "")
                JCDFval = sheetArray(i, JCDFc)
                found = 1
                Exit Function
            End If
        End If
    Next i
End Function

Public Function NMBseek_電線端末(product, cons, ByRef found)
    Dim i As Long
    found = 0
    For i = 2 To UBound(sheetArray)
        If Replace(sheetArray(i, 構成c), " ", "") = cons Then
            If Replace(sheetArray(i, 製品c), " ", "") = product Then
                品種val = sheetArray(i, 品種c)
                サイズval = sheetArray(i, サイズc)
                色val = sheetArray(i, 色c)
                品種呼val = Replace(sheetArray(i, 品種呼c), " ", "")
                サイズ呼val = Replace(sheetArray(i, サイズ呼c), " ", "")
                色呼val = Replace(sheetArray(i, 色呼c), " ", "")
                線長val = sheetArray(i, 線長c)
                '電線端末用
                回符1val = sheetArray(i, 回符1c)
                部品11val = sheetArray(i, 部品11c)
                端末1val = sheetArray(i, 端末1c)
                getFelt1val = sheetArray(i, getFelt1c)
                回符2val = sheetArray(i, 回符2c)
                部品21val = sheetArray(i, 部品21c)
                端末2val = sheetArray(i, 端末2c)
                getFelt2val = sheetArray(i, getFelt2c)
                found = "1"
                Exit Function
            End If
        End If
    Next i
End Function
