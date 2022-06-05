Attribute VB_Name = "M99_Verup"
Function VBC_Export(Optional Path)
    Dim VBC
    Dim myCount As Long

    With ActiveWorkbook.VBProject
        For Each VBC In .VBComponents
            Debug.Print VBC.Type, VBC.Name
            If VBC.Type <> 100 And _
               VBC.CodeModule.CountOfDeclarationLines <> VBC.CodeModule.CountOfLines Then
                If VBC.Type = 1 Then VBC.Export Path & "\" & VBC.Name & ".bas"
                If VBC.Type = 2 Then VBC.Export Path & "\" & VBC.Name & ".cls"
                If VBC.Type = 3 Then VBC.Export Path & "\" & VBC.Name & ".frm"
                myCount = myCount + 1
            End If
        Next VBC
    End With
    VBC_Export = myCount
End Function
'
'Sub VBC_Remove() 'このモジュールは削除しない
'    Dim VBC
'    With ActiveWorkbook.VBProject
'        For Each VBC In .VBComponents
'            Debug.Print VBC.Type, VBC.Name
'            If VBC.Type <> 100 Then
'                .VBComponents.Remove .VBComponents(VBC.Name)
'            End If
'        Next VBC
'    End With
'End Sub
'Public Function VBC_Import()
'    Path = "D:\04_生産準備+\000_システムパーツ\ver\2.125"
'    Dim buf As String, cnt As Long
'    Path = Path & "\"
'    buf = Dir(Path & "*.*")
'    Do While buf <> ""
'        'Debug.Print buf
'        If Right(buf, 3) <> "frx" And Right(buf, 3) <> "log" Then
'            ActiveWorkbook.VBProject.VBComponents.Import (Path & buf)
'        End If
'        buf = Dir()
'    Loop
'End Function

Sub Sheet_Ver_Export(Path)
    'シート[Ver]のエクスポート
    ThisWorkbook.Sheets("Ver").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=Path & "\sheet_Ver", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    'シート[フィールド名]のエクスポート
    ThisWorkbook.Sheets("フィールド名").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=Path & "\sheet_FieldName", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    'シート[color]のエクスポート
    ThisWorkbook.Sheets("color").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=Path & "\sheet_color", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    'シート[設定]のエクスポート
    ThisWorkbook.Sheets("設定").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=Path & "\sheet_setting", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
End Sub

 Sub fjkajfdaljdfka()
    PlaySound "けってい"
    
    Call アドレスセット(ActiveWorkbook)
    Path = アドレス(0) & "\ver"
    If Dir(Path, vbDirectory) = "" Then MkDir (Path)

    Path = Path & "\" & Mid(ThisWorkbook.Name, 6, InStr(ThisWorkbook.Name, "_") - 6)
    If Dir(Path, vbDirectory) = "" Then MkDir (Path)
    
    myCount = VBC_Export(Path)
    Call Sheet_Ver_Export(Path)
    
    DoEvents
    
    If myCount = 0 Then
        MsgBox "エクスポート出来るコードがありませんでした。"
    Else
        MsgBox myCount & " 点のコードをエクスポートしました。"
    End If
    
    Unload UI_04
End Sub
