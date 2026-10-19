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



' --- Module: ShRACI.cls ---
Attribute VB_Name = "ShRACI"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Function IsRACI(ByVal str) As Boolean
    Dim i As Long
    
    For i = 1 To Len(str)
        If InStr(1, "RACI", Mid(str, i, 1)) = 0 Then
            IsRACI = False
            Exit Function
        End If
    Next i
    IsRACI = True
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
        Dim TabRACI As ListObject
    Dim count As Long
    
    StartCode
    
    Set TabRACI = ShRACI.ListObjects("TabRACI")
    count = TabRACI.ListColumns.count
    If Target.count = 1 And Not Intersect(TabRACI.ListColumns(6).DataBodyRange.Resize(, count - 5), Target) Is Nothing Then
        If Not IsRACI(Target) Then
            MsgBox (Target & " n'est pas valide!")
            Target = vbNullString
            Target.Activate
        End If
    End If
    StopCode
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim TabRACI As ListObject
    Dim count As Long
    
    StartCode
    
    Set TabRACI = ShRACI.ListObjects("TabRACI")
    count = TabRACI.ListColumns.count
    If TabRACI.ListRows.count > 0 Then
        If Target.count = 1 And Not Intersect(TabRACI.ListColumns(6).DataBodyRange.Resize(, count - 5), Target) Is Nothing Then
            With ActiveSheet
                .Shapes("BtnR").ShapeStyle = msoShapeStylePreset39
                .Shapes("BtnA").ShapeStyle = msoShapeStylePreset39
                .Shapes("BtnC").ShapeStyle = msoShapeStylePreset39
                .Shapes("BtnI").ShapeStyle = msoShapeStylePreset39
                If (InStr(1, Target, "R")) Then .Shapes("BtnR").ShapeStyle = msoShapeStylePreset42
                If (InStr(1, Target, "A")) Then .Shapes("BtnA").ShapeStyle = msoShapeStylePreset41
                If (InStr(1, Target, "C")) Then .Shapes("BtnC").ShapeStyle = msoShapeStylePreset40
                If (InStr(1, Target, "I")) Then .Shapes("BtnI").ShapeStyle = msoShapeStylePreset38
            End With
        End If
    End If
    StopCode
End Sub


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

Public Sub StartCode()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub StopCode()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnRACI": ShRACI.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Private Sub RefreshFormulaTabRACI()
    Dim myTable As ListObject
    Dim currentIndex As Long
    
    Set myTable = ShRACI.ListObjects("TabRACI")
    If myTable.ListRows.count > 0 Then
        myTable.ListColumns("IDPhase").DataBodyRange.Formula = "=IF(ISNUMBER($C12), IF([@[IsPhase]] <>"""",$C12+1, $C12), 1)"
        myTable.ListColumns("IDTask").DataBodyRange.Formula = "=IFERROR(IF([@[IsPhase]]= """",$E12+1,""""),1)"
    End If
End Sub

Public Sub DeleteRow()
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    Dim DeleteConfirmation As String
    
    StartCode
    Application.Calculation = xlCalculationManual
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    Set myTable = ShRACI.ListObjects("TabRACI")
    
    Set SelectedRange = Intersect(Selection, Selection.Columns(1))
    If myTable.ListRows.count > 0 Then
        DeleteConfirmation = "Confirmez-vous la suppression de " & SelectedRange.count & " ligne(s) ?"
        If ShRACI.Range("Langue") = 1 Then DeleteConfirmation = "Confirm the deletion of " & SelectedRange.count & " row(s)?"
        MsgAnswer = MsgBox(DeleteConfirmation, vbOKCancel)
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
    Call RefreshFormulaTabRACI
    Application.Calculation = xlCalculationAutomatic
    StopCode
End Sub

Public Sub AddRow()
    Dim MyCell As Range
    Dim newRow As ListRow
    
    StartCode
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    
    With ShRACI.ListObjects("TabRACI")
        If ActiveCell.ListObject Is Nothing Or .ListRows.count = 0 Then
            Set newRow = .ListRows.Add
        ElseIf ActiveCell.ListObject.Name = .Name Then
            Set newRow = .ListRows.Add(ActiveCell.Row - .DataBodyRange.Row + 1)
        End If
        If Application.Caller = "BtnAddPhase" Then .ListColumns("IsPhase").DataBodyRange.Rows(newRow.index) = "X"
    End With
    Call RefreshFormulaTabRACI
    StopCode
End Sub





' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit

Public Sub RefreshCondFormat()
    StartCode
    ShRACI.Cells.FormatConditions.Delete
    Call RefreshCondFormListStatut
    Call FormatCondIsPhase
    StopCode
End Sub




' --- Module: RACI.bas ---
Attribute VB_Name = "RACI"
Option Explicit

Private Sub AddRACI(Target As Range, Prefixe As String)
    If Target = vbNullString Then
        Target = Prefixe
    Else
        Target = Target & Prefixe
    End If

End Sub

Private Sub RemoveRACI(Target As Range, Prefixe As String)
    Dim index As Long
    
    index = InStr(1, Target, Prefixe)
    If Len(Target) = 1 Then
        Target = vbNullString
    ElseIf index = 1 Then
        Target = Right(Target, Len(Target) - 1)
    ElseIf index = Len(Target) Then
        Target = Left(Target, Len(Target) - 1)
    Else
        Target = Left(Target, index - 1) & Right(Target, Len(Target) - index)
    End If
    ActiveSheet.Shapes("Btn" & Prefixe).ShapeStyle = msoShapeStylePreset39
End Sub



Public Sub GetRACI()
    Dim TabRACI As ListObject
    Dim count As Long
    Dim sep As String
    
    StartCode
    
    Set TabRACI = ShRACI.ListObjects("TabRACI")
    count = TabRACI.ListColumns.count
    If Not Intersect(TabRACI.ListColumns(6).DataBodyRange.Resize(, count - 5), ActiveCell) Is Nothing Then
    
        Select Case Application.Caller
            Case "BtnR"
                If (InStr(1, ActiveCell, "R")) Then
                    Call RemoveRACI(ActiveCell, "R")
                Else
                    Call AddRACI(ActiveCell, "R")
                    ActiveSheet.Shapes("BtnR").ShapeStyle = msoShapeStylePreset42
                End If
            Case "BtnA"
                If (InStr(1, ActiveCell, "A")) Then
                    Call RemoveRACI(ActiveCell, "A")
                Else
                    Call AddRACI(ActiveCell, "A")
                    ActiveSheet.Shapes("BtnA").ShapeStyle = msoShapeStylePreset41
                End If
            Case "BtnC"
                If (InStr(1, ActiveCell, "C")) Then
                    Call RemoveRACI(ActiveCell, "C")
                Else
                    Call AddRACI(ActiveCell, "C")
                    ActiveSheet.Shapes("BtnC").ShapeStyle = msoShapeStylePreset40
                End If
            Case "BtnI"
                If (InStr(1, ActiveCell, "I")) Then
                    Call RemoveRACI(ActiveCell, "I")
                Else
                    Call AddRACI(ActiveCell, "I")
                    ActiveSheet.Shapes("BtnI").ShapeStyle = msoShapeStylePreset38
                End If
        End Select
    End If
    StopCode
End Sub

Sub RefreshCondFormListStatut()
    Dim TableList As ListObject
    Dim e As Range
    Dim TabRACI As ListObject
    
    Set TableList = ShSettings.ListObjects("TabResp")
    Set TabRACI = ShRACI.ListObjects("TabRACI")
    For Each e In TableList.ListColumns("Préfixe").DataBodyRange
        With TabRACI.ListColumns(6).DataBodyRange.Resize(, TabRACI.ListColumns.count - 5)
            .FormatConditions.Add Type:=xlCellValue, Operator:=IIf(e.Offset(, -1) = "Mixte", xlNotEqual, xlEqual), _
                Formula1:=IIf(e.Offset(, -1) = "Mixte", """""", "=""" & e.value & """")
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
            .FormatConditions(1).StopIfTrue = True
        End With
        Application.CutCopyMode = False
    Next e
End Sub

Sub FormatCondIsPhase()
    With ShRACI.ListObjects("TabRACI").DataBodyRange
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$D13=""X"""
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Font
            .Bold = True
            .Italic = False
            .Color = RGB(117, 113, 113)
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = RGB(237, 237, 237)
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
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
    StartCode

    Set dict = CreateLangDict


    'ShRACI SHEET-----------------------------------------------
    With ShRACI
        trad = vbNullString
        For Each e In Union(.Range("G1:G3"), .Range("M1:M3"), .Range("R1:R3"), .Range("T6"), .Range("G10:V10"), .Range("F11:X11"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabRACI").ListColumns("Livrable ou tâche").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    'ShGuide SHEET-----------------------------------------------
    With ShGuide
        trad = vbNullString
        For Each e In .Range("G1:G5")
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
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
    
    CELL_KEY = ShRACI.Range("Langue")
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









