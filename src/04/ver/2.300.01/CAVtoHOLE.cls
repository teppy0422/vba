VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAVtoHOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum RetsuCa
    '�[���i��ca = 1.11111111111111E+17
    cavCa
    HOLEca
    ���lca
End Enum
Dim CAVtoHOLEdb As New Collection

Public �[���i��cav As String
Public cavCav As String
Public HOLEcav As String
Public ���lcav As String



Public Property Get Self() As CAVtoHOLE
    Set Self = Me
End Property

Public Sub CAVtoHOLEseek(Ter, cav)
    For Each C In CAVtoHOLEdb
        If Format(Replace(C.�[���i��cav, " ", ""), "@") = Format(Ter, "@") Then
            If val(C.cavCav) = val(cav) Then
                HOLEcav = C.HOLEcav
                ���lcav = C.���lcav
                Exit Sub
            End If
        End If
    Next C
    '������Ȃ��ꍇ
    Debug.Print "CavNotFound" & Ter & "_" & cav
    Stop
End Sub