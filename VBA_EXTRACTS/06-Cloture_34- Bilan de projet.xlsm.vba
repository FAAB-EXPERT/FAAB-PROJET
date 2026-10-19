' --- Module: ThisWorkbook.cls ---
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: ShBilanProject.cls ---
Attribute VB_Name = "ShBilanProject"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



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



' --- Module: Langue.bas ---
Attribute VB_Name = "Langue"
Option Explicit

Sub MeasureExecutionTime()
    Dim startTime As Double
    Dim endTime As Double
    Dim executionTime As Double
    
    ' Record the start time
    startTime = Timer
    
    ' Your code or subroutine to be measured
    ' Replace this with the code you want to measure
    For i = 1 To 1000000
        ' Code to be measured
    Next i
    
    ' Record the end time
    endTime = Timer
    
    ' Calculate the execution time
    executionTime = endTime - startTime
    
    ' Display the result
    MsgBox "Execution Time: " & executionTime & " seconds"
End Sub

Public Sub change_langue()

    Dim dict As Object
    Dim dictKey As Variant
    Dim e As Range
    Dim trad As String
    Dim tabEtat As ListObject
    Dim sh As Shape
    Dim originalText As String
    Dim sht As Worksheet

    Dim startTime As Double
    Dim endTime As Double
    
    startTime = Timer
    startCode

    Set dict = CreateLangDict


    'ShBilanProject SHEET-----------------------------------------------
    With ShBilanProject
        trad = vbNullString
        For Each e In Union(.Range("B5"), .Range("M2"), .Range("B7:O146"))
            If e <> vbNullString Then
                trad = dict(e.value)
                If trad <> vbNullString Then e = trad
            End If
        Next e
    End With
    
    'ShLangue -----------------------------------------------
    With ShLangue
        trad = vbNullString
        For Each e In .ListObjects("TabYesNo").Range
            If e <> vbNullString Then
                trad = dict(e.value)
                If trad <> vbNullString Then e = trad
            End If
        Next e
    End With
    
    'ActiveWorkbook-----------------------------------------------
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
    endTime = Timer
   ' MsgBox "Execution Time: " & endTime - startTime & " seconds"
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
    
    CELL_KEY = ShBilanProject.Range("Langue")
    CELL_VALUE = IIf(CELL_KEY = 1, 2, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    Set tabLang = ShLangue.ListObjects("TabLangue")
    
    For Each tblRow In tabLang.ListRows
        key = tblRow.Range.Cells(1, CELL_KEY).value
        value = tblRow.Range.Cells(1, CELL_VALUE).value
        dict(key) = value
    Next tblRow

    Set CreateLangDict = dict
End Function










' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit

Public Sub startCode()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub stopCode()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


