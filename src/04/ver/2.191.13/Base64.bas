Attribute VB_Name = "Base64"
'Base64�G���R�[�_/�f�R�[�_

Option Explicit

'Base64��������o�C�g��Ƀf�R�[�h����
Public Function decode(ByVal strData As String) As Byte()
    Dim objBase64 As Object

    Set objBase64 = CreateObject("MSXML2.DOMDocument").createElement("b64")
    objBase64.DataType = "bin.base64"
    objBase64.Text = strData
    decode = objBase64.nodeTypedValue

    Set objBase64 = Nothing
End Function

'�o�C�g���Base64������ɃG���R�[�h����
Public Function encode(ByRef byteData() As Byte) As String
    Dim objBase64 As Object

    Set objBase64 = CreateObject("MSXML2.DOMDocument").createElement("b64")
    objBase64.DataType = "bin.base64"
    objBase64.nodeTypedValue = byteData
    encode = objBase64.Text

    Set objBase64 = Nothing
End Function
