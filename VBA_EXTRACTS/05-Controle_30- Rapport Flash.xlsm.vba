' --- Module: ThisWorkbook.cls ---
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


' --- Module: ShGuide.cls ---
Attribute VB_Name = "ShGuide"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


' --- Module: ShFlash.cls ---
Attribute VB_Name = "ShFlash"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


' --- Module: ShSettings.cls ---
Attribute VB_Name = "ShSettings"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit

Sub Delete1()
    Dim RngName As String
    Dim i As Integer
    RngName = "Statut"
    i = 0
    If ActiveCell.Row < Range(RngName).Row + Range(RngName).Rows.Count And ActiveCell.Row >= Range(RngName).Row _
      And (ActiveCell.Column = Range(RngName).Column Or ActiveCell.Column = Range(RngName).Column + i) Then
        If MsgBox("Vous voullez supprimer ce champs?", vbYesNo + vbQuestion, _
            Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)) = vbYes Then
                Range(Cells(ActiveCell.Row, Range(RngName).Column), Cells(ActiveCell.Row, Range(RngName).Column + i)).Delete shift:=xlUp
        End If
    End If
End Sub


Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnFlash": ShFlash.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Public Sub startCode()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub stopCode()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



' --- Module: Langue.bas ---
Attribute VB_Name = "Langue"
Option Explicit

Public Sub change_langue()
    Dim dict As Object
    Dim dictKey As Variant
    Dim e As Range
    Dim trad As String
    Dim tabEtat As ListObject
    Dim sh As Shape
    Dim originalText As String
    Dim sht As Worksheet
    startCode
    Set dict = CreateLangDict
    
    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("G2:N6"), .Range("B7"), .Range("L7"), .Range("D13:D20"), .Range("D11"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShFlash
        trad = vbNullString
        For Each e In Union(.Range("E2:J6"), .Range("B7:J48"), .Range("C7"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
    End With
    
    
    With ShGuide
        trad = vbNullString
        For Each e In Union(.Range("E1:M5"), .Range("C6"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    For Each sht In ActiveWorkbook.Sheets
        trad = vbNullString
        For Each sh In sht.Shapes
            If sh.Type = 1 Then
                originalText = sh.TextFrame2.TextRange.Text
                trad = dict(originalText)
                If trad <> vbNullString Then
                    sh.TextFrame2.TextRange.Text = trad
                End If
            End If
        Next sh
        trad = dict(sht.Name)
        If trad <> vbNullString Then
            sht.Name = trad
        End If
    Next sht
    stopCode
End Sub

Private Function CreateLangDict() As Object
    Dim dict As Object
    Dim tabLang As ListObject
    Dim key As String
    Dim value As String
    Dim tblRow As ListRow
    Dim dictKey As Variant
    Dim CELL_KEY As Integer
    Dim CELL_VALUE As Integer
    
    CELL_KEY = ShSettings.Range("Langue")
    CELL_VALUE = IIf(CELL_KEY = 1, 2, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    Set tabLang = ShLangue.ListObjects("TabLangue")
    
    For Each tblRow In tabLang.ListRows
        key = tblRow.Range.Cells(1, CELL_KEY).value
        value = tblRow.Range.Cells(1, CELL_VALUE).value
        dict(key) = value
    Next tblRow
    

'    For Each dictKey In dict.Keys
'        Debug.Print "Key: " & dictKey & ", Value: " & dict(dictKey)
'    Next dictKey
    Set CreateLangDict = dict
End Function







' --- Module: ShLangue.cls ---
Attribute VB_Name = "ShLangue"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



