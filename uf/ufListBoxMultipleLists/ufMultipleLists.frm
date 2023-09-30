VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMultipleLists 
   Caption         =   "ListBox Demo"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   OleObjectBlob   =   "ufMultipleLists.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ufMultipleLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub optMonths_Click()
    Me.ListBox1.RowSource = "Sheet1!Months"
End Sub
Private Sub optCars_Click()
    Me.ListBox1.RowSource = "Sheet1!Cars"
End Sub
Private Sub optColors_Click()
    Me.ListBox1.RowSource = "Sheet1!Colors"
End Sub

Private Sub cmdOK_Click()
    MsgBox Me.ListBox1.Value
    Unload Me
End Sub

