VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAppendNameAndGender 
   Caption         =   "Get Name and Gender"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ufAppendNameAndGender.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ufAppendNameAndGender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim lNextRow As Long
    Dim ws As Worksheet
    Dim wf As WorksheetFunction
    
    Set ws = ActiveSheet
    Set wf = Application.WorksheetFunction
    
    ' Make sure a name is entered
    If Len(Me.tbxName.Text) = 0 Then
        MsgBox "You must enter a name."
        Me.tbxName.SetFocus
    Else
        ' Determine the next empty row
        lNextRow = wf.CountA(ws.Range("A:A")) + 1
        ' Transfer the name
        ws.Cells(lNextRow, 1) = Me.tbxName.Text
        
        ' Transer the gener
        With ws.Cells(lNextRow, 2)
            If Me.optMale.Value Then .Value = "Male"
            If Me.optFemale.Value Then .Value = "Female"
            If Me.optOther.Value Then .Value = "Other"
        End With
        
        ' Clear the controls for the next entry
        Me.tbxName.Text = vbNullString
        Me.optOther.Value = True
        Me.tbxName.SetFocus
    End If
            
End Sub

Private Sub UserForm_Click()

End Sub
