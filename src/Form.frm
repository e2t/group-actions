VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Пакетные действия"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4770
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelBut_Click()

  Unload Me
  
End Sub

Private Sub okBut_Click()

  Execute
  Unload Me
  
End Sub

Private Sub UserForm_Initialize()

  Me.fastenerBox.ControlTipText = "Действует только на детали"
  If gCurrentFolderPath = "" Then
    Me.RootOnlyBox.value = False
    Me.RootOnlyBox.Enabled = False
  Else
    Me.RootOnlyBox.value = True
    Me.CurFolderTextBox.value = gCurrentFolderPath
    Me.CurFolderTextBox.SelStart = Len(gCurrentFolderPath)
    Me.CurFolderTextBox.SetFocus
  End If
  
End Sub
