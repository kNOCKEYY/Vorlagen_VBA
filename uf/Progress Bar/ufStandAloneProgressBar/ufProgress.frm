VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Progress"
   ClientHeight    =   1305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "ufProgress.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    With Me
        .lblProgress.BackColor = vbRed
        .lblProgress.Width = 0
    End With

End Sub

Public Sub SetDescription(Description As String)

    Me.lblDescription.Caption = Description

End Sub

Public Sub UpdateProgress(PctDone As Double)

    With Me
        .frmProgress.Caption = Format(PctDone, "0%")
        .lblProgress.Width = PctDone * (.frmProgress.Width - 10)
        .lblProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        .Repaint
    End With

End Sub



