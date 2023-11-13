VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents CmdAdd As MSForms.CommandButton
Attribute CmdAdd.VB_VarHelpID = -1
Public WithEvents CmdUp As MSForms.CommandButton
Attribute CmdUp.VB_VarHelpID = -1
Public WithEvents CmdDown As MSForms.CommandButton
Attribute CmdDown.VB_VarHelpID = -1
Public WithEvents CmdDel As MSForms.CommandButton
Attribute CmdDel.VB_VarHelpID = -1

Private Sub CmdAdd_Click()
    Add_Item
End Sub

Private Sub CmdUp_Click()
    MoveItem_Up
End Sub

Private Sub CmdDown_Click()
    MoveItem_Down
End Sub

Private Sub CmdDel_Click()
    Delete_Item
End Sub
