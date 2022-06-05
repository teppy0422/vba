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
    '端末品番ca = 1.11111111111111E+17
    cavCa
    HOLEca
    備考ca
End Enum
Dim CAVtoHOLEdb As New Collection

Public 端末品番cav As String
Public cavCav As String
Public HOLEcav As String
Public 備考cav As String



Public Property Get Self() As CAVtoHOLE
    Set Self = Me
End Property

Public Sub CAVtoHOLEseek(Ter, cav)
    For Each c In CAVtoHOLEdb
        If Format(Replace(c.端末品番cav, " ", ""), "@") = Format(Ter, "@") Then
            If Val(c.cavCav) = Val(cav) Then
                HOLEcav = c.HOLEcav
                備考cav = c.備考cav
                Exit Sub
            End If
        End If
    Next c
    '見つからない場合
    Debug.Print "CavNotFound" & Ter & "_" & cav
    Stop
End Sub
