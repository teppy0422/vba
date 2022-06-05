Attribute VB_Name = "StringUtility"
'������<->�o�C�g��ϊ����[�e�B���e�B

Option Explicit

Private Const adTypeBinary As Integer = 1
Private Const adTypeText As Integer = 2
Private Const adStateOpen = 1

'�o�C�g���Shift-JIS�ŕ�����ɕϊ�����
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
    MsgBox "�G���[�R�[�h:" & err.number & vbCrLf & err.Description
End Function

'�������Shift-JIS�Ńo�C�g��ɕϊ�����
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
    MsgBox "�G���[�R�[�h:" & err.number & vbCrLf & err.Description
End Function
