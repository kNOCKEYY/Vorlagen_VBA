VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListTransfer 
   Caption         =   "ListBox Transfer Demo"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   OleObjectBlob   =   "ufListTransfer.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ufListTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    'Add the value
    Me.lbxTo.AddItem Me.lbxFrom.Value
    If Not Me.chkDuplicates.Value Then
        'If duplicates aren't allowed, remove the value
        Me.lbxFrom.RemoveItem Me.lbxFrom.ListIndex
    End If
    EnableButtons
End Sub

Private Sub cmdRemove_Click()
    If Not Me.chkDuplicates.Value Then
        Me.lbxFrom.AddItem Me.lbxTo.Value
    End If
    Me.lbxTo.RemoveItem Me.lbxTo.ListIndex
    EnableButtons
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim Msg As String
    Dim i As Long
    
    Msg = ""

    MsgBox "The 'To list' contains " & Me.lbxTo.ListCount & " items."

    For i = 0 To Me.lbxTo.ListCount - 1
        Msg = Msg & Me.lbxTo.List(i) & vbNewLine
    Next i

    MsgBox "You selected: " & vbNewLine & Msg
    
    Unload Me
End Sub

Private Sub lbxFrom_Change()
    EnableButtons
End Sub

Private Sub lbxTo_Change()
    EnableButtons
End Sub

Private Sub EnableButtons()
    Me.cmdAdd.Enabled = Me.lbxFrom.ListIndex > -1
    Me.cmdRemove.Enabled = Me.lbxTo.ListIndex > -1
End Sub

