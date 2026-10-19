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



' --- Module: ShRegRisques.cls ---
Attribute VB_Name = "ShRegRisques"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: ShSettings.cls ---
Attribute VB_Name = "ShSettings"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Deactivate()
    Call RefreshCondFormat
End Sub



' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit

Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnEchelles": ShEchelles.Activate
        Case "BtnRisques": ShRegRisques.Activate
        Case "BtnProblemes": ShRegProblemes.Activate
        Case "BtnMatrice": ShMatrice.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub


Public Sub AddRow()
    If Application.Caller = "BtnAddStatut" Then
        Call AddRowHelper("TabStatut")
    ElseIf Application.Caller = "BtnAddNature" Then
        Call AddRowHelper("TabNature")
    ElseIf Application.Caller = "BtnAddProbabilite" Then
        Call AddRowHelper("TabProbabilite")
    ElseIf Application.Caller = "BtnAddResponsable" Then
        Call AddRowHelper("TabResponsable")
    ElseIf Application.Caller = "BtnAddConsequence" Then
        Call AddRowHelper("TabConseq")
    ElseIf Application.Caller = "BtnAddGravite" Then
        Call AddRowHelper("TabGravite")
    ElseIf Application.Caller = "BtnAddPriorite" Then
        Call AddRowHelper("TabPriorite")
    ElseIf Application.Caller = "BtnAddTendance" Then
        Call AddRowHelper("TabTendance")
    ElseIf Application.Caller = "BtnAddCriticite" Then
        Call AddRowHelper("TabCriticite")
        
    ElseIf Application.Caller = "BtnAddRisk" Then
        Call AddRowHelper("TabRegistresRisques")
    ElseIf Application.Caller = "BtnAddProblem" Then
        Call AddRowHelper("TabRegistresProblemes")
    End If
End Sub

Public Sub DeleteRow()
    If Application.Caller = "BtnDeleteStatut" Then
        Call DeleteRowHelper("TabStatut")
    ElseIf Application.Caller = "BtnDeleteNature" Then
        Call DeleteRowHelper("TabNature")
    ElseIf Application.Caller = "BtnDeleteProbabilite" Then
        Call DeleteRowHelper("TabProbabilite")
    ElseIf Application.Caller = "BtnDeleteResponsable" Then
        Call DeleteRowHelper("TabResponsable")
    ElseIf Application.Caller = "BtnDeleteConsequence" Then
        Call DeleteRowHelper("TabConseq")
    ElseIf Application.Caller = "BtnDeleteGravite" Then
        Call DeleteRowHelper("TabGravite")
    ElseIf Application.Caller = "BtnDeletePriorite" Then
        Call DeleteRowHelper("TabPriorite")
    ElseIf Application.Caller = "BtnDeleteTendance" Then
        Call DeleteRowHelper("TabTendance")
    ElseIf Application.Caller = "BtnDeleteCriticite" Then
        Call DeleteRowHelper("TabCriticite")
        
    ElseIf Application.Caller = "BtnRemoveRisk" Then
        Call DeleteRowHelper("TabRegistresRisques")
    ElseIf Application.Caller = "BtnDeleteProblem" Then
        Call DeleteRowHelper("TabRegistresProblemes")
    End If
End Sub

Public Sub DeleteRowHelper(tabName As String)
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set myTable = ActiveSheet.ListObjects(tabName)
    
    Set SelectedRange = Intersect(Selection, Selection.Columns(1))
    If myTable.ListRows.count > 0 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de " & SelectedRange.count & " ligne(s) ?", vbOKCancel)
        If MsgAnswer = vbOK Then
            For count = SelectedRange.count To 1 Step -1
                If myTable.ListRows.count > 0 Then
                    If Not Intersect(SelectedRange(count, 1), myTable.DataBodyRange) Is Nothing Then
                           myTable.ListRows(SelectedRange(count).Row - myTable.DataBodyRange.Row + 1).Delete
                    End If
                End If
            Next count
        End If
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub AddRowHelper(tabName As String)
    Dim myTable As ListObject
    Dim MyCell As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set myTable = ActiveSheet.ListObjects(tabName)
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
        myTable.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'Public Sub DeleteRow()
'    Dim myTable As ListObject
'    Dim MyCell As Range
'    Dim MsgAnswer As VbMsgBoxResult
'    Dim count As Long
'    Dim SelectedRange As Range
'
'    StartCode
'    Application.Calculation = xlCalculationManual
'    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
'    Set myTable = ActiveSheet.ListObjects(MyCell.ListObject.Name)
'
'    Set SelectedRange = Intersect(Selection, Selection.Columns(1))
'    If myTable.ListRows.count > 0 Then
'        MsgAnswer = MsgBox("Confirmez-vous la suppression de " & SelectedRange.count & " ligne(s) ?", vbOKCancel)
'        If MsgAnswer = vbOK Then
'            For count = SelectedRange.count To 1 Step -1
'                If myTable.ListRows.count > 0 Then
'                    If Not Intersect(SelectedRange(count, 1), myTable.DataBodyRange) Is Nothing Then
'                           myTable.ListRows(SelectedRange(count).Row - myTable.DataBodyRange.Row + 1).Delete
'                    End If
'                End If
'            Next count
'        End If
'    End If
'    Application.Calculation = xlCalculationAutomatic
'    StopCode
'End Sub
'
'Public Sub AddRow()
'    Dim myTable As ListObject
'    Dim MyCell As Range
'
'    StartCode
'    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
'    Set myTable = MyCell.ListObject
'
'    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
'        myTable.ListRows.Add
'    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
'        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
'    End If
'    StopCode
'End Sub

Public Function FindValueInRange(lookupValue As Double, tableRange As Range) As Long
    Dim statusRange As Range
    Dim minRange As Range
    Dim maxRange As Range
    Dim i As Long

    Application.Volatile
    Set statusRange = tableRange.Columns(2)
    Set minRange = tableRange.Columns(1)
    Set maxRange = tableRange.Columns(3)

    For i = 1 To tableRange.Rows.count
        
        If lookupValue > IIf(minRange.Cells(i).value = "-" Or minRange.Cells(i).value = "", -10000, minRange.Cells(i).value) And _
            lookupValue <= IIf(maxRange.Cells(i).value = "-" Or maxRange.Cells(i).value = "", 10000, maxRange.Cells(i).value) Then
            FindValueInRange = i
            Exit Function
        End If
    Next i
End Function

Public Sub StartCode()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub StopCode()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub SetupNavBar()
    Selection.ShapeRange.Top = 63.6
    Selection.ShapeRange.Left = 255
End Sub




' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit

Public Sub RefreshCondFormat()
    StartCode
    
    If ShRegRisques.ListObjects("TabRegistresRisques").ListRows.count > 0 Then
        ShRegRisques.Cells.FormatConditions.Delete
        Call RefreshCondFormListCriticite
    End If
    If ShRegProblemes.ListObjects("TabRegistresProblemes").ListRows.count > 0 Then
        ShRegProblemes.Cells.FormatConditions.Delete
        Call RefreshCondFormListStatut
        Call RefreshCondFormListPriorite
    End If
    StopCode
End Sub




' --- Module: ShRegProblemes.cls ---
Attribute VB_Name = "ShRegProblemes"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: MRegRisques.bas ---
Attribute VB_Name = "MRegRisques"
Option Explicit

Sub RefreshCondFormListCriticite()
    Dim TableList As ListObject
    Dim e As Range
    Dim myFormula As String
    
    Set TableList = ShSettings.ListObjects("TabCriticite")
    myFormula = "=" & Mid(ShRegRisques.ListObjects("TabRegistresRisques").ListColumns("Criticité").DataBodyRange(1).Address, 1, 2) & "14="
    
    For Each e In TableList.ListColumns("Criticité").DataBodyRange
        With ShRegRisques.ListObjects("TabRegistresRisques").ListColumns("Gravité").DataBodyRange.Resize(, 3)
            .FormatConditions.Add Type:=xlExpression, Formula1:=myFormula & """" & e.value & """"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Font
                On Error Resume Next
                .Color = e.Font.Color
                If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
                .Bold = e.Font.Bold
                .Italic = e.Font.Italic
                .TintAndShade = e.Font.TintAndShade
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = e.Interior.PatternColorIndex
                If e.Interior.ThemeColor > 0 Then
                    .ThemeColor = e.Interior.ThemeColor
                    .Pattern = e.Interior.Pattern
                    .PatternColorIndex = e.Interior.PatternColorIndex
                Else
                    .Color = e.Interior.Color
                End If
                .TintAndShade = e.Interior.TintAndShade

            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Application.CutCopyMode = False
    Next e
End Sub



' --- Module: MRegProblemes.bas ---
Attribute VB_Name = "MRegProblemes"
Option Explicit

Sub RefreshCondFormListStatut()
    Dim TableList As ListObject
    Dim e As Range
 
    Set TableList = ShSettings.ListObjects("TabStatut")

    For Each e In TableList.DataBodyRange
        With ShRegProblemes.ListObjects("TabRegistresProblemes").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$I16=""" & e.value & """"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Font
                On Error Resume Next
                .Color = e.Font.Color
                If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
                .Bold = e.Font.Bold
                .Italic = e.Font.Italic
                .TintAndShade = e.Font.TintAndShade
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = e.Interior.PatternColorIndex
                If e.Interior.ThemeColor > 0 Then
                    .ThemeColor = e.Interior.ThemeColor
                    .Pattern = e.Interior.Pattern
                    .PatternColorIndex = e.Interior.PatternColorIndex
                Else
                    .Color = e.Interior.Color
                End If
                .TintAndShade = e.Interior.TintAndShade
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Application.CutCopyMode = False
    Next e
End Sub

Sub RefreshCondFormListPriorite()
    Dim TableList As ListObject
    Dim e As Range
 
    Set TableList = ShSettings.ListObjects("TabPriorite")

    For Each e In TableList.DataBodyRange
        With ShRegProblemes.ListObjects("TabRegistresProblemes").ListColumns("Priorité").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$G16=""" & e.value & """"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Font
                On Error Resume Next
                .Color = e.Font.Color
                If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
                .Bold = e.Font.Bold
                .Italic = e.Font.Italic
                .TintAndShade = e.Font.TintAndShade
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = e.Interior.PatternColorIndex
                If e.Interior.ThemeColor > 0 Then
                    .ThemeColor = e.Interior.ThemeColor
                    .Pattern = e.Interior.Pattern
                    .PatternColorIndex = e.Interior.PatternColorIndex
                Else
                    .Color = e.Interior.Color
                End If
                .TintAndShade = e.Interior.TintAndShade

            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Application.CutCopyMode = False
    Next e
End Sub



' --- Module: ShGuide.cls ---
Attribute VB_Name = "ShGuide"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: ShMatrice.cls ---
Attribute VB_Name = "ShMatrice"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
    StartCode
    Call GetLegende
    Call GetMatrice
    StopCode
End Sub



' --- Module: MMatrice.bas ---
Attribute VB_Name = "MMatrice"
Option Explicit

Sub GetLegende()
    Dim tabRisques As ListObject
    Dim tabLegende As ListObject
    
    Set tabRisques = ShRegRisques.ListObjects("TabRegistresRisques")
    Set tabLegende = ShMatrice.ListObjects("TabLegendeMatrice")
    
    If tabRisques.ListRows.count = 0 Then Exit Sub
    tabRisques.ListColumns("Description").DataBodyRange.Copy tabLegende.ListColumns("Description").DataBodyRange
    tabLegende.Resize tabLegende.Range.Resize(tabRisques.ListRows.count + tabLegende.HeaderRowRange.Rows.count)
    tabLegende.DataBodyRange.WrapText = False
    ShMatrice.Range("C" & tabRisques.ListRows.count + tabRisques.DataBodyRange.Row + 1 & ":D100").ClearContents
End Sub

Sub GetMatrice()
    Call InitMatrice
    Call CreateMatrice
    Call FillMatrice
End Sub

Private Sub InitMatrice()
    Dim shp As Shape
    
    For Each shp In ShMatrice.Shapes
        If Left(shp.Name, 3) = "Mat" Or Left(shp.Name, 3) = "Pro" Or Left(shp.Name, 3) = "Gra" Then
            If shp.Name <> "Gravite1" And shp.Name <> "Probabilite1" And shp.Name <> "Matrice11" Then
                shp.Delete
            End If
        End If
    Next shp
End Sub

Private Sub CreateMatrice()
    Dim tabGravite As ListObject
    Dim tabProbabilite As ListObject
    Dim tabCriticite As ListObject
    Dim shpProbabilite As Shape
    Dim shpGravite As Shape
    Dim shpMatrice As Shape
    Dim shp As Shape
    Dim count As Long
    Dim countMat As Long
    Dim e As Range
    
    Set tabGravite = ShSettings.ListObjects("tabGravite")
    Set tabProbabilite = ShSettings.ListObjects("tabProbabilite")
    Set tabCriticite = ShSettings.ListObjects("TabCriticite")
    Set shpProbabilite = ShMatrice.Shapes("Probabilite1")
    Set shpGravite = ShMatrice.Shapes("Gravite1")
    Set shpMatrice = ShMatrice.Shapes("Matrice11")
    count = tabProbabilite.ListRows.count
    shpProbabilite.TextFrame.Characters.Text = tabProbabilite.ListColumns("Probabilité").DataBodyRange.Rows(count)
    shpMatrice.TextFrame.Characters.Text = ""
    Call FormatMatrice(shpMatrice, 1, count)
    count = count - 1
    While count > 0
        'Duplicate Probabilite shape
        Set shp = shpProbabilite.Duplicate
        shp.Name = "Probabilite" & (tabGravite.ListRows.count - count + 1)
        shp.Top = shpProbabilite.Top
        shp.Left = shpProbabilite.Left + (tabProbabilite.ListRows.count - count) * (shpProbabilite.Width + 5)
        shp.TextFrame.Characters.Text = tabProbabilite.ListColumns("Probabilité").DataBodyRange.Rows(count)
        
        'Duplicate Matrice shape
        Set shp = shpMatrice.Duplicate
        shp.Name = "Matrice1" & (tabProbabilite.ListRows.count - count + 1)
        shp.Top = shpMatrice.Top
        shp.Left = shpMatrice.Left + (tabProbabilite.ListRows.count - count) * (shpMatrice.Width + 5)
        Call FormatMatrice(shp, 1, count)
        count = count - 1
    Wend

    shpGravite.TextFrame.Characters.Text = tabGravite.ListColumns("Gravité").DataBodyRange.Rows(1)
    
    For count = 2 To tabGravite.ListRows.count
        'Duplicate Gravité shape
        Set shp = shpGravite.Duplicate
        shp.Name = "Gravite" & count
        shp.Top = shpGravite.Top + (count - 1) * (shpGravite.Height + 5)
        shp.Left = shpGravite.Left
        shp.TextFrame.Characters.Text = tabGravite.ListColumns("Gravité").DataBodyRange.Rows(count)
        
        'Duplicate Matrice shape
        For countMat = 1 To tabProbabilite.ListRows.count
            Set shp = shpMatrice.Duplicate
            shp.Name = "Matrice" & count & countMat
            shp.Top = shpMatrice.Top + (count - 1) * (shpMatrice.Height + 5)
            shp.Left = shpMatrice.Left + (countMat - 1) * (shpMatrice.Width + 5)
            Call FormatMatrice(shp, count, tabProbabilite.ListRows.count + 1 - countMat)
        Next countMat
    Next count
End Sub

Sub FormatMatrice(shp As Shape, indexGravite As Long, indexProba As Long)
    Dim tabGravite As ListObject
    Dim tabProbabilite As ListObject
    Dim tabCriticite As ListObject
    Dim weight As Double
    
    Set tabGravite = ShSettings.ListObjects("tabGravite")
    Set tabProbabilite = ShSettings.ListObjects("tabProbabilite")
    Set tabCriticite = ShSettings.ListObjects("TabCriticite")
    
    weight = tabGravite.ListColumns("Poids").DataBodyRange.Rows(indexGravite) * tabProbabilite.ListColumns("Poids").DataBodyRange.Rows(indexProba)
    With tabCriticite.ListColumns("Criticité").DataBodyRange.Rows(FindValueInRange(weight, tabCriticite.DataBodyRange))
        shp.Fill.ForeColor.RGB = .Interior.Color
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = .Font.Color
    End With
End Sub

Sub FillMatrice()
    Dim tabRisques As ListObject
    Dim tabCriticite As ListObject
    Dim e As Range
    Dim currentIndex As Long
    Dim sep As String
    
    Set tabRisques = ShRegRisques.ListObjects("TabRegistresRisques")
    Set tabCriticite = ShSettings.ListObjects("TabCriticite")
    
    If tabRisques.ListRows.count = 0 Then Exit Sub
    For Each e In tabRisques.ListColumns("Matrice").DataBodyRange
        currentIndex = e.Row - tabRisques.HeaderRowRange.Row
        If e <> vbNullString Then
'            If Len(ShMatrice.Shapes(e).TextFrame.Characters.Text) = 0 Then
'                sep = ""
'            Else
'                sep = vbNewLine
'            End If
            With ShMatrice.Shapes(e)
                If .TextFrame.Characters.Text <> vbNullString Then .TextFrame.Characters.Text = .TextFrame.Characters.Text & "  "
                .TextFrame.Characters.Text = .TextFrame.Characters.Text & sep & tabRisques.ListColumns("ID").DataBodyRange.Rows(currentIndex)
            End With
        End If
    Next e
End Sub



' --- Module: ShEchelles.cls ---
Attribute VB_Name = "ShEchelles"
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

Public Sub change_langue()
    Dim dict As Object
    Dim dictKey As Variant
    Dim e As Range
    Dim trad As String
    Dim myTab As ListObject
    Dim sh As Shape
    Dim originalText As String
    Dim sht As Worksheet
    StartCode
    Set dict = CreateLangDict
    
    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("E1:L5"), .Range("C6"), .Range("L6"), .Range("C14:M14"), .Range("C17:H31"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        For Each myTab In .ListObjects
            For Each e In myTab.DataBodyRange
                trad = dict(e.value)
                If e <> vbNullString And trad <> vbNullString Then
                    e = trad
                End If
            Next e
        Next myTab
    End With
    
    With ShEchelles
        trad = vbNullString
        For Each e In .Range("C1:I18")
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        
    End With

    With ShRegRisques
        trad = vbNullString
        For Each e In Union(.Range("E6"), .Range("F1:M5"), .Range("B12:N12"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabRegistresRisques").DataBodyRange
            trad = dict(e.value)
            If e.HasFormula = False And e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    With ShRegProblemes
        trad = vbNullString
        For Each e In Union(.Range("C6"), .Range("D1:K5"), .Range("C13:M14"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabRegistresProblemes").DataBodyRange
            trad = dict(e.value)
            If e.HasFormula = False And e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    With ShMatrice
        trad = vbNullString
        For Each e In Union(.Range("J1:J5"), .Range("C6"), .Range("D8:D10"), .Range("J8:J10"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
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
    StopCode
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








