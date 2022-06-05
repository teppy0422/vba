Attribute VB_Name = "StringUtility"
'文字列<->バイト列変換ユーティリティ

Option Explicit

Private Const adTypeBinary As Integer = 1
Private Const adTypeText As Integer = 2
Private Const adStateOpen = 1

'バイト列をShift-JISで文字列に変換する
Public Function byteToString(ByRef byteData() As Byte) As String
    Dim objStream As Object

    On Error GoTo ErrorHandler
    Set objStream = CreateObject("ADODB.Stream")

    objStream.Open
    objStream.Type = adTypeBinary
    objStream.Write byteData

    objStream.position = 0
    objStream.Type = adTypeText
    objStream.Charset = "shift-jis"

    byteToString = objStream.ReadText
    Exit Function

ErrorHandler:
    If Not objStream Is Nothing Then
        If (objStream.State And adStateOpen) = adStateOpen Then
            objStream.Close
        End If
    End If
    Set objStream = Nothing
    MsgBox "エラーコード:" & err.number & vbCrLf & err.Description
End Function

'文字列をShift-JISでバイト列に変換する
Public Function stringToByte(ByVal strData As String) As Byte()
    Dim objStream As Object

    On Error GoTo ErrorHandler
    Set objStream = CreateObject("ADODB.Stream")

    objStream.Open
    objStream.Type = adTypeText
    objStream.Charset = "shift-jis"
    objStream.WriteText strData

    objStream.position = 0
    objStream.Type = adTypeBinary

    stringToByte = objStream.Read
    Exit Function

ErrorHandler:
    If Not objStream Is Nothing Then
        If (objStream.State And adStateOpen) = adStateOpen Then
            objStream.Close
        End If
    End If
    Set objStream = Nothing
    MsgBox "エラーコード:" & err.number & vbCrLf & err.Description
End Function
