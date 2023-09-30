VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFillListBoxNoDupes 
   Caption         =   "Select an item"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   OleObjectBlob   =   "ufFillListBoxNoDupes.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ufFillListBoxNoDupes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OKButton_Click()
    If ListBox1.Value <> "" Then MsgBox "You chose " & ListBox1.Value
    Unload Me
End Sub

