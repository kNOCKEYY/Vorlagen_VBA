VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRenameRange 
   Caption         =   "Rename Range"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   OleObjectBlob   =   "ufRenameRange.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ufRenameRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub UserForm_Initialize()

    Me.refRange.Text = ActiveWindow.RangeSelection.Address
    
End Sub

Private Sub cmdOK_Click()
    Dim UserRange As Range
    Dim WorkRange As Range
    Dim Operand As Double
    Dim cell As Range
    
    ' Validate range entry
    On Error Resume Next
    Set UserRange = Range(Me.refRange.Text)
    If Err <> 0 Then
        MsgBox "Invalid range selected"
        Me.refRange.SetFocus
        Exit Sub
    End If
    On Error GoTo 0
    
    ActiveWorkbook.Names.Add Name:=Me.tbxRangeName.Text, RefersTo:=UserRange

    Unload Me
    
End Sub


Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

