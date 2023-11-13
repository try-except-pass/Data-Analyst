VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pe_form 
   Caption         =   "UserForm2"
   ClientHeight    =   156
   ClientLeft      =   -301
   ClientTop       =   -1393
   ClientWidth     =   49
   OleObjectBlob   =   "pe_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pe_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        End
    End If
End Sub

