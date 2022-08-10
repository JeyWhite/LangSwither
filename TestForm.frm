VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "TestForm.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "ex1"
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()
    If ComboBox1.ListIndex = -1 Then Exit Sub
    gLang.setLang ComboBox1.List(ComboBox1.ListIndex)
    gLang.executeUserForm Me
End Sub

Private Sub UserForm_Activate()
    
End Sub

Private Sub UserForm_Initialize()
    Me.ComboBox1.List = gLang.getLangArr
    gLang.executeUserForm Me
End Sub
