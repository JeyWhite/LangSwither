VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LangSwither"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const KEY_NAME As String = "Keys_title"
Public Enum Methods
    byTag
    byCaption
End Enum
Private height
Private keyPositions As Dictionary
Private languages As Dictionary
Private content As Dictionary

Private Sub class_initialize()
    Dim KeyCount As Long, langCount As Long
    KeyCount = initKeys
    langCount = initLanguages
    Debug.Print ("Lang Switcer  created. Keys:" & KeyCount & " languages:" & langCount)
End Sub

Public Function getLangArr() As Variant
    getLangArr = languages.Keys
End Function

Public Function setLang(language As String, Optional mode As Methods = Methods.byTag) As Long
    If Not languages.Exists(language) Then
        setLang = False
        Exit Function
    End If
    Select Case mode
        Case Methods.byTag
            Dim arr
            Set content = New Dictionary
            
            arr = keysTitle.Offset(1, languages(language)).Resize(height, 1)
            Dim i As Long
            For i = LBound(keyPositions.Keys) To UBound(keyPositions.Keys)
                content.Add keyPositions.Keys(i), arr(keyPositions(keyPositions.Keys(i)), 1)
            Next i
            setLang = True
        Case Methods.byCaption
            Debug.Print ("WIP! Please waiy for updates")
            setLang = False
    End Select
End Function

Public Sub executeUserForm(form As Object)
    Dim element As Object
    executeControl form
    For Each element In form.Controls
        executeControl element
    Next element
End Sub

Private Sub executeControl(element As Object)
    If Not IsEmpty(element.Tag) Then
        If content.Exists(element.Tag) Then
            element.Caption = content(element.Tag)
        End If
    End If
End Sub

Private Function keysTitle() As Range
    Set keysTitle = ThisWorkbook.Names(KEY_NAME).RefersToRange
End Function

Private Function initKeys() As Long
    Dim keysRange As Range
    Set keysRange = undertitle(keysTitle.Cells(1))
    height = keysRange.Rows.Count
    Set keyPositions = New Dictionary
    Dim i As Long
    For i = 1 To keysRange.Rows.Count
        If Not IsEmpty(keysRange.Cells(i).Value2) Then
            If Not keyPositions.Exists(keysRange.Cells(i)) Then
                keyPositions.Add keysRange.Cells(i).Value2, i
            End If
        End If
    Next i
    initKeys = keyPositions.Count
End Function

Private Function initLanguages() As Long
     Dim langRange As Range
     Set langRange = rightTitle(keysTitle.Cells(1))
     Set languages = New Dictionary
     Dim i As Long
     For i = 1 To langRange.Columns.Count
        If Not IsEmpty(langRange.Cells(1, i)) Then
            languages.Add langRange.Cells(1, i).Value2, i
        End If
     Next i
     initLanguages = langRange.Count
End Function


Private Function rightTitle(r As Range)
    Set rightTitle = r.Parent.Range(r.Offset(0, 1), _
    r.EntireRow.Columns(r.EntireRow.Columns.Count).End(xlToLeft))
End Function


Private Function undertitle(r As Range) As Range
     Set undertitle = r.Parent.Range(r.Offset(1, 0), _
     r.EntireColumn.Rows(r.EntireColumn.Rows.Count).End(xlUp))
End Function

