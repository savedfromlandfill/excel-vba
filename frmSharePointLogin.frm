VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSharePointLogin 
   Caption         =   "SharePoint Login"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "frmSharePointLogin.frx":0000
End
Attribute VB_Name = "frmSharePointLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
  Me.Tag = "Cancel"
  Me.Hide
End Sub

Private Sub btnOk_Click()
  Me.Tag = "Ok"
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
   txtLoginID.Text = Environ("Username")
   txtPassword.SetFocus
   Me.Left = Application.Left + 100
   Me.Top = Application.Top + 300
End Sub
