VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents BoxSelect As MSForms.ComboBox
Attribute BoxSelect.VB_VarHelpID = -1
Private Sub BoxSelect_Change()
    Update_ListBox
End Sub
