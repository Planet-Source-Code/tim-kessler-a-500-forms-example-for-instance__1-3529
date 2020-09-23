VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hang On"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This is an instance example I wrote
' and that I do not warrant.
' Tim Kessler
' See module1 for more instructions


Private Sub Form_Load()
On Error GoTo ErrorHandler:
    Dim frmNew As New Form1
    Show
    Caption = "Instance " & Str$(FormCount + 1)
    Refresh
    Let FormCount = FormCount + 1
    If FormCount < MAX_FORMS Then
        Set frmNew = New Form1
        Load frmNew
        Set frmNew = Nothing
    End If
    Timer1.Enabled = True

    Exit Sub
ErrorHandler:
    MsgBox "Your only have memory for  " & Str$(FormCount) & " forms." & vbCrLf _
        & "Click OK.....forms will unload within a few seconds after a you click ok."
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
