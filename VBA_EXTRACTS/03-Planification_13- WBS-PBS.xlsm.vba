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



' --- Module: ShPBS.cls ---
Attribute VB_Name = "ShPBS"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: ShWBS.cls ---
Attribute VB_Name = "ShWBS"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Public Function GetIdColTable(ByVal toFind As String) As Integer
    Dim e As Range
    
    GetIdColTable = 0
    For Each e In ShWBS.ListObjects("TabWBS").HeaderRowRange
        If e = toFind Then
            GetIdColTable = e.Column '- ShGantt.ListObjects("TabProjet").Range.Column
        End If
    Next e
End Function

Public Sub getLevel()
    Dim Tablevel(1 To 4) As Integer
    Dim myTab As ListObject
    Dim e As Range
    Dim i As Integer
    Dim COL_LEVEL As Integer
    
    COL_LEVEL = GetIdColTable("Niveau")
    
    Set myTab = ShWBS.ListObjects("TabWBS")
    
    For i = LBound(Tablevel) To UBound(Tablevel)
        Tablevel(i) = 0
    Next i

    For Each e In myTab.ListColumns("ID").DataBodyRange
        e = vbNullString
        If Not Cells(e.Row, COL_LEVEL) = vbNullString And IsNumeric(Cells(e.Row, COL_LEVEL)) Then
            Tablevel(Cells(e.Row, COL_LEVEL)) = Tablevel(Cells(e.Row, COL_LEVEL)) + 1
            If Cells(e.Row, COL_LEVEL) < 4 Then
                For i = Cells(e.Row, COL_LEVEL) + 1 To UBound(Tablevel)
                    Tablevel(i) = 0
                Next i
            End If
            For i = LBound(Tablevel) To Cells(e.Row, COL_LEVEL)
                e = IIf(e = vbNullString, "", e & ".") & Tablevel(i)
            Next i
        End If
    Next e
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    startCode
    'If Target.Count = 1 Then
    Call getLevel
    Call GetCharge
    'End If
    stopCode
End Sub


' --- Module: ShDonnees.cls ---
Attribute VB_Name = "ShDonnees"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit

Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnPBS": ShPBS.Activate
        Case "BtnWBS": ShWBS.Activate
        Case "BtnDonnees": ShDonnees.Activate
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

    With ShPBS
        trad = vbNullString
        For Each e In Union(.Range("H4"), .Range("D7:D80"), .Range("B6:E6"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    
    With ShWBS
        trad = vbNullString
        For Each e In Union(.Range("D7:D80"), .Range("B6:E6"))
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
    
    CELL_KEY = ShPBS.Range("Langue")
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







' --- Module: WBS.bas ---
Attribute VB_Name = "WBS"
Option Explicit

Sub GetCharge()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long, j As Long
    Dim totalCharge As Double
    Dim countProgression As Long
    Dim startRow As Long
    Dim lastRow As Long
    Dim currentNiveau As Long

    Set tbl = ShWBS.ListObjects("TabWBS")

    ' Get the last row of the table
    lastRow = tbl.ListRows.Count

    ' Loop through each row in the table
    For i = 1 To lastRow
        currentNiveau = tbl.DataBodyRange(i, tbl.ListColumns("Niveau").Index).value
        ' Check if it's a Niveau 1 row
        If currentNiveau = 1 Then
            totalCharge = 0
            startRow = i
            ' Loop through the subsequent rows to find the min/max dates, sum the charge, and average the progression
            For j = i + 1 To lastRow
                If tbl.DataBodyRange(j, tbl.ListColumns("Niveau").Index).value = 1 Then
                    Exit For ' Stop when we reach the next Niveau 1
                End If
                ' Add to totalCharge for Charge de travail only for Niveau 2 rows
                If tbl.DataBodyRange(j, tbl.ListColumns("Niveau").Index).value = 2 Then
                    If IsNumeric(tbl.DataBodyRange(j, tbl.ListColumns("Charge de travail").Index).value) Then
                        totalCharge = totalCharge + tbl.DataBodyRange(j, tbl.ListColumns("Charge de travail").Index).value
                    End If
                End If
            Next j
            
            ' Fill the total charge in the "Charge de travail" column of the Niveau 1 row (for Niveau 2 rows only)
            tbl.DataBodyRange(i, tbl.ListColumns("Charge de travail").Index).value = totalCharge
        End If
    Next i
End Sub



