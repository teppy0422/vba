VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_11 
   Caption         =   "部品箱表示_作成"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_11.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














































































Private Sub CB0_Change()
    Dim 項目(1) As String
    Dim 項目2(1) As String
    'CB0.Text
    With ActiveWorkbook.Sheets("製品品番")
        Set myKey = .Cells.Find("型式", , , 1)
        Set myKey = .Rows(myKey.Row).Find(CB0.Text, , , 1)
        Set mykey2 = .Rows(myKey.Row).Find("結き", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For y = myKey.Row + 1 To lastRow
            If InStr(項目(0), "," & .Cells(y, myKey.Column)) & "," = 0 Then
                項目(0) = 項目(0) & "," & .Cells(y, myKey.Column) & ","
                項目2(0) = 項目2(0) & "," & .Cells(y, mykey2.Column) & ","
            End If
        Next y
        If Len(項目(0)) <= 2 Then
            項目(0) = ""
            項目s = Empty
        Else
            項目(0) = Mid(項目(0), 2)
            項目(0) = Left(項目(0), Len(項目(0)) - 1)
            項目s = Split(項目(0), ",,")
            項目2(0) = Mid(項目2(0), 2)
            項目2(0) = Left(項目2(0), Len(項目2(0)) - 1)
            項目2s = Split(項目2(0), ",,")
        End If
    End With
    
    With CB1
        .RowSource = ""
        .Clear
        If Not IsEmpty(項目s) Then
            For i = LBound(項目s) To UBound(項目s)
                .AddItem
                .List(i, 0) = 項目s(i)
                .List(i, 1) = 項目2s(i)
            Next i
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub CB1_Change()
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    If 製品品番RANc <> 1 Then
        myLabel.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        myLabel.Caption = ""
    End If
End Sub

Private Sub CommandButton4_Click()
    PlaySound "もどる"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "じっこう"
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    
    Dim fileName As String: fileName = Replace(wb(0).Name, ".xlsm", "") & "_部品箱表示_" & CB0.Value & "_" & CB1.Value & ".xlsx"
    Unload Me
    
    Dim i As Long, pNumbers As String
    For i = LBound(製品品番Ran, 2) + 1 To UBound(製品品番Ran, 2)
        pNumbers = pNumbers & "," & 製品品番Ran(製品品番RAN_read(製品品番Ran, "メイン品番"), i)
    Next i
    
    Dim setWords As String, setWordsSP As Variant
    setWords = "部品品番,呼称,d,D,W,L,色,部品名称,部材詳細,種類,工程,工程a"
    setWordsSP = Split(setWords, ",")
    
    Set ws(1) = wb(0).Sheets("部品リスト")
    msg = checkFieldName("部品品番", ws(1), setWords)
    If msg <> "" Then
        msg = "[部品リスト]に次のフィールドが見つかりません。" & msg & vbCrLf & vbCrLf & _
                   "この機能を使用するにはVer2.200.70以降で作成した[部品リスト]である必要があります。" & vbCrLf & _
                   "作成を中止します。"
        MsgBox msg, vbOKOnly, "PLUS+"
        End
    End If
    
    Dim Array_部品リスト As Variant
    Array_部品リスト = readSheetToRan2(ws(1), "部品品番", setWords & pNumbers, "")
    
    '製品品番毎に使用が無い部品を削除
    Dim x As Long, skipFlag As Boolean
    For i = LBound(Array_部品リスト, 2) + 1 To UBound(Array_部品リスト, 2)
        skipFlag = True
        For x = UBound(setWordsSP) + 1 To UBound(setWordsSP) + UBound(製品品番Ran, 2)
            If Array_部品リスト(x, i) <> "" Then
                skipFlag = False
                Exit For
            End If
        Next x
        If skipFlag = True Then
            Debug.Print i, Array_部品リスト(0, i)
            Array_部品リスト = delete_RanVer2(Array_部品リスト, i)
            i = i - 1
        End If
        If i + 1 > UBound(Array_部品リスト, 2) Then Exit For
    Next i
    
    '出力するデータのまとめ
    Dim addArray() As Variant
    ReDim addArray(2, UBound(Array_部品リスト, 2))
    addArray(0, 0) = "A"
    addArray(1, 0) = "B"
    addArray(2, 0) = "C"
    
    For i = LBound(Array_部品リスト, 2) + 1 To UBound(Array_部品リスト, 2)
        If Array_部品リスト(9, i) = "B" Then
            addArray(0, i) = Array_部品リスト(0, i)
            If Array_部品リスト(2, i) <> "" Then
                addArray(1, i) = Replace(Array_部品リスト(2, i) & " L=" & Array_部品リスト(5, i), ".0", "")
            End If
        ElseIf Array_部品リスト(9, i) = "T" Then
            addArray(2, i) = Replace(Array_部品リスト(0, i), "-", " ")
            addArray(1, i) = Replace(Array_部品リスト(1, i), " ", "") & "-" & Array_部品リスト(6, i)
            If Replace(Array_部品リスト(3, i), " ", "") <> "" Then
                addArray(0, i) = "D" & Replace(Replace(Array_部品リスト(2, i) & "×" & Array_部品リスト(3, i), ".0", ""), " ", "") & " L=" & Replace(Array_部品リスト(5, i), " ", "")
            ElseIf Replace(Array_部品リスト(2, i), " ", "") <> "" Then
                addArray(0, i) = "D" & Replace(Replace(Array_部品リスト(2, i), ".0", ""), " ", "") & " L=" & Replace(Replace(Array_部品リスト(5, i), ".0", ""), " ", "")
            ElseIf Replace(Array_部品リスト(4, i), " ", "") <> "" Then
                addArray(0, i) = Replace(Replace("W" & Array_部品リスト(4, i), ".0", ""), " ", "") & " L=" & Replace(Replace(Array_部品リスト(5, i), ".0", ""), " ", "")
            End If
        Else
            
        End If
    Next i
    
    '長さLを0埋め4桁にする
    Dim array_temp
    For i = LBound(Array_部品リスト, 2) To UBound(Array_部品リスト, 2)
        array_temp = Array_部品リスト(5, i)
        If array_temp <> "" Then
            If IsNumeric(array_temp) Then
                array_temp = Int(array_temp)
                If (Len(array_temp) <= 4) Then
                    array_temp = String(4 - Len(array_temp), "0") & array_temp
                    Array_部品リスト(5, i) = array_temp
                    
                End If
            End If
        End If
    Next
     
    Array_部品リスト = merge_Array(addArray, Array_部品リスト)

    export_ArrayToSheet Array_部品リスト, "部品箱表示", True
    
    Dim outputDirectory As String
    outputDirectory = wb(0).path & "\42_部品箱表示"
    If Dir(outputDirectory, vbDirectory) = "" Then MkDir outputDirectory
    
    'テプラのSPC100のcsvとtxtは、カンマを区切り文字として認識する為、カンマ区切りで出力するとテキスト内のカンマで列ズレが発生する _
    なのでxlsxで出力
    wb(0).Sheets("部品箱表示").Move
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=outputDirectory & "\" & fileName, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Call 最適化もどす
    PlaySound "かんせい"
    
    Dim myMsg As String: myMsg = "作成しました" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "部品箱表示_作成")
End Sub

Private Sub UserForm_Initialize()
    Dim 項目(1) As String
    With wb(0).Sheets("製品品番")
        Set myKey = .Cells.Find("型式", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For x = myKey.Column To lastCol
            項目(0) = 項目(0) & "," & .Cells(myKey.Row, x)
        Next x
        項目(0) = Mid(項目(0), 2)
    End With
    項目s = Split(項目(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
            If 項目s(i) = "結き" Then myindex = i
        Next i
        .ListIndex = myindex
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub
