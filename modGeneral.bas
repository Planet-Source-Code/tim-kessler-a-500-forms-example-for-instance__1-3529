Attribute VB_Name = "Module1"
Option Explicit

' This program is a sample program
' Which demonstrates Instancing in VB
' Wrote by Tim Kessler
' takessl@rocketmail.com
' There is no warranty expressed or
' implied.


Global FormCount As Integer

' Set Test number of forms here:
Public Const MAX_FORMS = 500

Public Sub Main()
    Dim frmNew As New Form1
    Set frmNew = New Form1
    frmNew.Show
End Sub
