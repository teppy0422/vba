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
    For Each c In CAVtoHOLEdb
        If Format(Replace(c.�[���i��cav, " ", ""), "@") = Format(Ter, "@") Then
            If Val(c.cavCav) = Val(cav) Then
                HOLEcav = c.HOLEcav
                ���lcav = c.���lcav
                Exit Sub
            End If
        End If
    Next c
    '������Ȃ��ꍇ
    Debug.Print "CavNotFound" & Ter & "_" & cav
    Stop
End Sub