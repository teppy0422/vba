Attribute VB_Name = "hsf"
Option Explicit

Function readTextToArray_hsf(ByVal myPath As String)
    Dim myRan As Variant
    Dim target As New FileSystemObject
    If Dir(myPath) = "" Then readTextToArray_hsf = False: Exit Function
    Dim intFino As Variant, aa As Variant, temp As Variant, a As Long, x As Long, maxaa As Long
    intFino = FreeFile
    Open myPath For Input As #intFino
    'カンマの数を確認
    Do Until EOF(intFino)
        Line Input #intFino, aa
        temp = Split(aa, ",")
        If maxaa < UBound(temp) Then maxaa = UBound(temp)
    Loop
    Close #intFino
    Open myPath For Input As #intFino
    
    Do Until EOF(intFino)
        Line Input #intFino, aa
        temp = Split(aa, ",")
        If IsEmpty(myRan) Then
            a = UBound(temp)
            ReDim myRan(maxaa, 0)
        End If
        ReDim Preserve myRan(maxaa, UBound(myRan, 2) + 1)
        For x = LBound(temp) To UBound(temp)
            myRan(x, UBound(myRan, 2)) = temp(x)
        Next x
    Loop
    Close #intFino
    readTextToArray_hsf = myRan
End Function

Sub export_ArrayToSheet_v2(ByVal myArray As Variant, ByVal sheetName As String, ByVal transpose As Boolean)
    
    Dim a As Variant
    On Error Resume Next
    a = UBound(myArray, 2)
    On Error GoTo 0
    
    If transpose And Not IsEmpty(a) Then myArray = WorksheetFunction.transpose(myArray)
    
    If hsf.check_Sheet_existence(sheetName, ActiveWorkbook) Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    Dim newWs As Worksheet
    Set newWs = Worksheets.add
    newWs.Name = sheetName
    newWs.Cells.NumberFormat = "@"

    With ActiveWorkbook.Sheets(sheetName)
        .Cells.ClearContents
        If Not IsEmpty(myArray) Then
            If IsEmpty(a) Then a = 0 '配列が1次元
            .Range(.Cells(1, 1), .Cells(UBound(myArray) + 1, a + 1)) = myArray
        End If
    End With
    
End Sub

Public Function check_Sheet_existence(ByVal targetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet, flag As Boolean
    For Each ws In wb.Worksheets
        If ws.Name = targetName Then
            check_Sheet_existence = True
            Exit Function
        End If
    Next ws
    check_Sheet_existence = False
End Function

Function checkMD_moreNew(ByVal path As String, ByVal pNumber As String) As String
    Dim FSO, folder, folders
    Set FSO = CreateObject("scripting.filesystemobject")
    Set folders = FSO.getfolder(path)
    
    pNumber = Replace(pNumber, " ", "")
    Dim moreNew As String, thisRevision As String, sp As Variant
    For Each folder In folders.SubFolders
        If Left(folder.Name, Len(pNumber)) = pNumber Then
            If right(folder.Name, 2) = "MD" Then
                sp = Split(folder.Name, "_")
                thisRevision = sp(1)
                moreNew = revision_Compare(moreNew, thisRevision)
                If moreNew = thisRevision Then checkMD_moreNew = folder.path
            End If
        End If
    Next
End Function

'MD専用_getしたらusedRanの値を削除して重複しないようにしてる
    
Public Function searchRan_forMD(ByVal myRan As Variant, ByRef usedRan As Variant, ByVal searchWords As String, _
    ByVal searchFields As String, ByVal getFields As String) As String

    Dim searchWord As Variant, searchField  As Variant, getField As Variant, x As Long, xx As Long
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim f As Variant
    ReDim f(UBound(searchField))
    '検索フィールド番号の取得
    For x = LBound(searchField) To UBound(searchField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 1) = searchField(x) Then
                f(x) = xx
                GoTo line10
            End If
        Next xx
line10:
    Next x
    Dim g As Variant
    ReDim g(UBound(getField))
    '取得フィールド番号の取得
    For x = LBound(getField) To UBound(getField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 1) = getField(x) Then
                g(x) = xx
                GoTo line20
            End If
        Next xx
line20:
    Next x
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(f) To UBound(f)
            If searchWord(x) = myRan(f(x), i) Then
                If x = UBound(f) Then
                    i_Anser = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                If usedRan(g(x), i_Anser) <> "" Then
                    searchRan_forMD = myRan(g(x), i_Anser)
                    usedRan(g(x), i_Anser) = ""
                    Exit Function
                End If
            Next x
            Exit Function
        End If
    Next i
    searchRan_forMD = False
End Function

'MD専用_getしたらusedRanの値を削除して重複しないようにしてる
'製品の部品数によってフィールド数が異なるので､フィールドが無くなるまで参照する
    
Public Function searchRan_forMD_v2(ByVal myRan As Variant, ByRef usedRan As Variant, ByVal searchWords As String, _
    ByVal searchFields As String, ByVal getFieldHeader As String) As String
    Dim x As Long, fieldCount As Long, getFields As String
    For x = LBound(myRan) To UBound(myRan)
        If myRan(x, 1) Like "*" & getFieldHeader & "*" Then
            fieldCount = fieldCount + 1
            getFields = getFields & "," & getFieldHeader & fieldCount
        End If
    Next x
    getFields = Mid(getFields, 2)
    
    Dim searchWord As Variant, searchField  As Variant, getField As Variant, xx As Long
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim f As Variant
    ReDim f(UBound(searchField))
    '検索フィールド番号の取得
    For x = LBound(searchField) To UBound(searchField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 1) = searchField(x) Then
                f(x) = xx
                GoTo line10
            End If
        Next xx
line10:
    Next x
    Dim g As Variant
    ReDim g(UBound(getField))
    '取得フィールド番号の取得
    For x = LBound(getField) To UBound(getField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 1) = getField(x) Then
                g(x) = xx
                GoTo line20
            End If
        Next xx
line20:
    Next x
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(f) To UBound(f)
            If searchWord(x) = myRan(f(x), i) Then
                If x = UBound(f) Then
                    i_Anser = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                If usedRan(g(x), i_Anser) <> "" Then
                    searchRan_forMD_v2 = myRan(g(x), i_Anser)
                    usedRan(g(x), i_Anser) = ""
                    Exit Function
                End If
            Next x
            Exit Function
        End If
    Next i
    searchRan_forMD_v2 = False
End Function
'MDから取得した取り出し方向が記号の為、数字に変換する
Function converting_to_formatting_Direction(ByVal formatting As String) As Variant
    Select Case formatting
        Case "U"
            converting_to_formatting_Direction = "0"
        Case "L"
            converting_to_formatting_Direction = "90"
        Case "D"
            converting_to_formatting_Direction = "180"
        Case "R"
            converting_to_formatting_Direction = "270"
        Case "S"
            converting_to_formatting_Direction = Empty
        Case Else
            converting_to_formatting_Direction = Empty
    End Select
End Function
