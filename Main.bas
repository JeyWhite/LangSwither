Attribute VB_Name = "Main"
Option Explicit


Global gLang As LangSwither

Sub doStuff()
    Set gLang = New LangSwither
'    Call gLang.init(ThisWorkbook.Sheets(WS_NAME))
    
    gLang.setLang ("EN")
    Dim oForm As UserForm1
    Set oForm = New UserForm1
    oForm.Show
    
End Sub
