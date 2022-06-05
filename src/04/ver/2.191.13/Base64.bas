Attribute VB_Name = "Base64"
'Base64エンコーダ/デコーダ

Option Explicit

'Base64文字列をバイト列にデコードする
Public Function decode(ByVal strData As String) As Byte()
    Dim objBase64 As Object

    Set objBase64 = CreateObject("MSXML2.DOMDocument").createElement("b64")
    objBase64.DataType = "bin.base64"
    objBase64.Text = strData
    decode = objBase64.nodeTypedValue

    Set objBase64 = Nothing
End Function

'バイト列をBase64文字列にエンコードする
Public Function encode(ByRef byteData() As Byte) As String
    Dim objBase64 As Object

    Set objBase64 = CreateObject("MSXML2.DOMDocument").createElement("b64")
    objBase64.DataType = "bin.base64"
    objBase64.nodeTypedValue = byteData
    encode = objBase64.Text

    Set objBase64 = Nothing
End Function
