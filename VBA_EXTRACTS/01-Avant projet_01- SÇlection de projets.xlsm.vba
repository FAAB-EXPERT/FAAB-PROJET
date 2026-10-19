' --- Module: MCriteresSelection.bas ---
Attribute VB_Name = "MCriteresSelection"
Option Explicit

Public Sub SaveUpdateCS()
    Dim TableCriteressSelection As ListObject
    Dim TableBECriteresSelection As ListObject
    Dim TableBEListCriteria As ListObject
    Dim CountCode As Long
    Dim CountScore As Long
    Dim i As Long
    
    StartCode
    Set TableCriteressSelection = ShCriteresSelection.ListObjects("TableCriteressSelection")
    Set TableBECriteresSelection = ShBackEnd.ListObjects("TableBECriteresSelection")
    Set TableBEListCriteria = ShBackEnd.ListObjects("TableListCriteria")
    
    CountScore = TableCriteressSelection.ListColumns.count - 3
    TableBECriteresSelection.DataBodyRange.ClearContents
    TableBECriteresSelection.Resize TableBECriteresSelection.Range.Resize(CountScore + 1, TableCriteressSelection.ListRows.count + 1)
    TableBECriteresSelection.ListColumns(2).DataBodyRange.Resize(CountScore, TableCriteressSelection.ListRows.count).value = Application.Transpose(TableCriteressSelection.ListColumns(3).DataBodyRange.Resize(, CountScore))
    TableBECriteresSelection.HeaderRowRange(2).Resize(, TableCriteressSelection.ListRows.count).value = Application.Transpose(TableCriteressSelection.ListColumns("Code").DataBodyRange)
    TableBECriteresSelection.ListColumns("Points").DataBodyRange.value = Application.Transpose(TableCriteressSelection.HeaderRowRange(, 3).Resize(, CountScore))
    
    CountCode = TableCriteressSelection.ListRows.count
    For i = 1 To CountScore
        TableBEListCriteria.ListColumns("Points").DataBodyRange.Rows(CountCode * i - CountCode + 1 & ":" & CountCode * i).value = TableBECriteresSelection.ListColumns("Points").DataBodyRange.Rows(i).value
        TableBEListCriteria.ListColumns("Critere").DataBodyRange.Rows(CountCode * i - CountCode + 1 & ":" & CountCode * i).value = TableCriteressSelection.ListColumns(i + 2).DataBodyRange.value
        TableBEListCriteria.ListColumns("Code").DataBodyRange.Rows(CountCode * i - CountCode + 1 & ":" & CountCode * i).value = TableCriteressSelection.ListColumns("Code").DataBodyRange.value
    
    Next i
    Call UpdateEvaluationTable

    StopCode
End Sub

Public Sub DeleteCritère()
    Dim TableCS As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim TableEvaluation As ListObject
    Dim e As Range
    
    StartCode
    ShEvaluation.Unprotect
    Set TableCS = ActiveSheet.ListObjects("TableCriteressSelection")
    Set TableEvaluation = ShEvaluation.ListObjects("TableEvaluation")
    
    If Not Intersect(ActiveCell, TableCS.DataBodyRange) Is Nothing Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de cette ligne ?", vbOKCancel)
        If MsgAnswer = vbOK Then
            TableCS.ListRows(ActiveCell.Row - TableCS.DataBodyRange.Row + 1).Delete
            TableEvaluation.ListRows(ActiveCell.Row - TableCS.DataBodyRange.Row + 1).Delete
            For Each e In TableCS.ListColumns("Code").DataBodyRange
                e = "C" & e.Row - TableCS.DataBodyRange.Row + 1
            Next e
            Call UpdateEvaluationTable
        End If
    End If
    StopCode
    ShEvaluation.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
End Sub

Public Sub AddNewCritère()
    Dim TableCS As ListObject
    Dim TableEvaluation As ListObject
    Dim MyCell As Range
    Dim e As Range
    
    StartCode
    ShEvaluation.Unprotect
    Set TableCS = ShCriteresSelection.ListObjects("TableCriteressSelection")
    Set TableEvaluation = ShEvaluation.ListObjects("TableEvaluation")
    
    If ActiveCell.ListObject Is Nothing Then
        TableCS.ListRows.Add
        TableEvaluation.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = TableCS.Name Then
        TableCS.ListRows.Add ActiveCell.Row - TableCS.DataBodyRange.Row + 2
        TableEvaluation.ListRows.Add ActiveCell.Row - TableCS.DataBodyRange.Row + 2
    End If
    For Each e In TableCS.ListColumns("Code").DataBodyRange
        e = "C" & e.Row - TableCS.DataBodyRange.Row + 1
    Next e
        Call UpdateEvaluationTable
    StopCode
    ShEvaluation.Protect
End Sub

Public Sub AddNewScore()
    Dim TableCS As ListObject
    Dim PositionToAdd As Long
    Dim MsgAnswer As VbMsgBoxResult
    
    StartCode
    Set TableCS = ShCriteresSelection.ListObjects("TableCriteressSelection")
    
    PositionToAdd = ActiveCell.Column - TableCS.Range.Column
    If Not ActiveCell.ListObject Is Nothing Then
        If ActiveCell.ListObject.Name = TableCS.Name Then
            If PositionToAdd > 1 And PositionToAdd < TableCS.ListColumns.count Then
                MsgAnswer = MsgBox("Voulez vous ajouter un nouvelle colonne à gauche de " & TableCS.HeaderRowRange(PositionToAdd + 1) & "?", vbOKCancel)
                If MsgAnswer = vbOK Then
                    TableCS.ListColumns.Add PositionToAdd + 1
                    If IsNumeric(TableCS.HeaderRowRange(PositionToAdd + 2)) Then
                        TableCS.HeaderRowRange(PositionToAdd + 1) = TableCS.ListColumns(PositionToAdd + 2) - 1
                    ElseIf IsNumeric(TableCS.HeaderRowRange(PositionToAdd)) Then
                        TableCS.HeaderRowRange(PositionToAdd + 1) = TableCS.ListColumns(PositionToAdd) + 1
                    End If
                    TableCS.HeaderRowRange(, 3).Offset(-1).Resize(, TableCS.ListColumns.count - 3).Interior.Color = RGB(185, 225, 203)
                End If
             Else
                MsgBox ("La cellule sélectionnée n'est pas valide")
            End If
        End If
    End If
    StopCode
End Sub

Public Sub RemoveScore()
    Dim TableCS As ListObject
    Dim PositionToDelete As Long
    Dim MsgAnswer As VbMsgBoxResult
    
    StartCode
    Set TableCS = ShCriteresSelection.ListObjects("TableCriteressSelection")
    
    PositionToDelete = ActiveCell.Column - TableCS.Range.Column + 1

    If Not ActiveCell.ListObject Is Nothing Then
        If ActiveCell.ListObject.Name = TableCS.Name And PositionToDelete > 2 And PositionToDelete < TableCS.ListColumns.count Then
            MsgAnswer = MsgBox("Voulez vous supprimer " & TableCS.HeaderRowRange(PositionToDelete) & "?", vbOKCancel)
            If MsgAnswer = vbOK Then
                TableCS.ListColumns(PositionToDelete).Delete
                With TableCS.HeaderRowRange(, 3).Offset(-1).Resize(, TableCS.ListColumns.count - 3)
                    .EntireRow.Interior.Color = xlNone
                 '   .HorizontalAlignment = xlCenterAcrossSelection
                    .Interior.Color = RGB(185, 225, 203)
                End With
            End If
        Else
            MsgBox ("Sélection n'est pas valide")
        End If
    Else
        MsgBox ("Sélection n'est pas valide")
    End If
    StopCode
End Sub

Public Sub UpdateEvaluationTable()
    Dim TableCS As ListObject
    Dim TableEvaluation As ListObject
    
    
    Set TableCS = ShCriteresSelection.ListObjects("TableCriteressSelection")
    Set TableEvaluation = ShEvaluation.ListObjects("TableEvaluation")
    'Code & Critère
    TableEvaluation.ListColumns(1).DataBodyRange.Resize(TableCS.ListRows.count, 2).value = TableCS.ListColumns(1).DataBodyRange.Resize(, 2).value
    'Coefficient de pondération
    TableEvaluation.ListColumns(3).DataBodyRange.value = TableCS.ListColumns(TableCS.ListColumns.count).DataBodyRange.value
    With ShEvaluation.Shapes("ExampleShape")
        .Top = ShEvaluation.Range("PriorityLevel").Top + 35
        .Left = ShEvaluation.Range("PriorityLevel").Left
    End With
End Sub



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



' --- Module: ShGuide.cls ---
Attribute VB_Name = "ShGuide"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


' --- Module: ShCriteresSelection.cls ---
Attribute VB_Name = "ShCriteresSelection"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Deactivate()
    Call SaveUpdateCS
End Sub


' --- Module: ShEvaluation.cls ---
Attribute VB_Name = "ShEvaluation"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim TableCS As ListObject
    Dim TableEvaluation As ListObject
    
    Set TableCS = ShCriteresSelection.ListObjects("TableCriteressSelection")
    Set TableEvaluation = ShEvaluation.ListObjects("TableEvaluation")
    StartCode

    If Not Intersect(Target, TableEvaluation.ListColumns(3).DataBodyRange) Is Nothing Then
        TableCS.ListColumns(TableCS.ListColumns.count).DataBodyRange.value = TableEvaluation.ListColumns(3).DataBodyRange.value
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
    StartCode
    ShEvaluation.Unprotect
    Call RefreshCondFormPriority
    ShEvaluation.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
    Call SaveUpdateCS
    StopCode
End Sub


' --- Module: ShBackEnd.cls ---
Attribute VB_Name = "ShBackEnd"
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
        Case "BtnCriteresSelection": ShCriteresSelection.Activate
        Case "BtnEvaluation": ShEvaluation.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub
'
'Public Sub DeleteRow()
'    Dim MyTable As ListObject
'    Dim MyCell As Range
'    Dim MsgAnswer As VbMsgBoxResult
'    Dim i As Long
'
'    StartCode
'
'    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(1, -1)
'    Set MyTable = ActiveSheet.ListObjects(MyCell.ListObject.Name)
'
'    If Not Intersect(ActiveCell, MyTable.DataBodyRange) Is Nothing Then
'        MsgAnswer = MsgBox("Confirmez-vous la suppression de cette ligne ?", vbOKCancel)
'        If MsgAnswer = vbOK Then
'            MyTable.ListRows(ActiveCell.Row - MyTable.DataBodyRange.Row + 1).Delete
'            If MyTable.Name = "TablePriority" Then
'                For i = 2 To MyTable.Range.Rows.Count
'                    MyTable.ListColumns("Max").DataBodyRange(i).Formula = "=" & MyTable.ListColumns("Min").DataBodyRange(i - 1).Address(False, False)
'                Next i
'            End If
'        End If
'    End If
'    StopCode
'End Sub
'
'Public Sub AddRow()
'    Dim MyTable As ListObject
'    Dim MyCell As Range
'    Dim i As Long
'
'    StartCode
'    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(1, -1)
'    Set MyTable = MyCell.ListObject
'
'    If ActiveCell.ListObject Is Nothing Then
'        MyTable.ListRows.Add
'    ElseIf ActiveCell.ListObject.Name = MyTable.Name Then
'        MyTable.ListRows.Add ActiveCell.Row - MyTable.DataBodyRange.Row + 2
'    End If
'
'    If MyTable.Name = "TablePriority" Then
'        For i = 2 To MyTable.Range.Rows.Count
'            MyTable.ListColumns("Max").DataBodyRange(i).Formula = "=" & MyTable.ListColumns("Min").DataBodyRange(i - 1).Address(False, False)
'        Next i
'    End If
'    StopCode
'End Sub

Public Sub AddRow()
    If Application.Caller = "BtnAddTablePriority" Then
        Call AddRowHelper("TablePriority")
    ElseIf Application.Caller = "BtnAddTableCoeffPonderation" Then
        Call AddRowHelper("TableCoeffPonderation")
    ElseIf Application.Caller = "BtnAddRessource" Then
        Call AddRowHelper("TabRessources")
    End If
End Sub

Public Sub DeleteRow()
    If Application.Caller = "BtnDeleteTablePriority" Then
        Call DeleteRowHelper("TablePriority")
    ElseIf Application.Caller = "BtnDeleteTableCoeffPonderation" Then
        Call DeleteRowHelper("TableCoeffPonderation")
    ElseIf Application.Caller = "BtnDeleteRessource" Then
        Call DeleteRowHelper("TabRessources")
    End If
End Sub

Public Sub DeleteRowHelper(tabName As String)
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    Dim i As Long
    
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
                        If tabName = "TablePriority" Then
                            For i = 2 To myTable.Range.Rows.count
                                myTable.ListColumns("Max").DataBodyRange(i).Formula = "=" & myTable.ListColumns("Min").DataBodyRange(i - 1).Address(False, False)
                            Next i
                        End If
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
    Dim i As Long
    
    StartCode
    Set myTable = ActiveSheet.ListObjects(tabName)
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
        myTable.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
    End If
    If tabName = "TablePriority" Then
        For i = 2 To myTable.Range.Rows.count
            myTable.ListColumns("Max").DataBodyRange(i).Formula = "=" & myTable.ListColumns("Min").DataBodyRange(i - 1).Address(False, False)
        Next i
    End If
    StopCode
End Sub



' --- Module: MEvaluation.bas ---
Attribute VB_Name = "MEvaluation"
Option Explicit

Public Sub AddNewProject()
    Dim TableCS As ListObject
    
    StartCode
    ShEvaluation.Unprotect
    Set TableCS = ShEvaluation.ListObjects("TableEvaluation")
    
    If Not ActiveCell.Offset(1).ListObject Is Nothing Then
        If ActiveCell.Offset(1).ListObject.Name = TableCS.Name Then
            If ActiveCell.Column - TableCS.Range.Column > 2 Then
                TableCS.ListColumns.Add ActiveCell.Column - TableCS.Range.Column + 1
                Call RefreshEvaluationTable
                With TableCS.ListColumns(ActiveCell.Column - TableCS.Range.Column + 1).DataBodyRange.Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=INDIRECT(""TableBECriteresSelection["" & $D17 & ""]"")"
                End With
                ActiveCell.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
             Else
                MsgBox ("La cellule sélectionnée n'est pas valide")
            End If
        End If
    Else
        TableCS.ListColumns.Add
        Call RefreshEvaluationTable
        With TableCS.ListColumns(TableCS.ListColumns.count).DataBodyRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=INDIRECT(""TableBECriteresSelection["" & $D17 & ""]"")"
        End With
    End If
    StopCode
    ShEvaluation.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
End Sub

Public Sub RemoveProject()
    Dim TableCS As ListObject
    Dim i As Long
    Dim MsgAnswer As VbMsgBoxResult
    
    StartCode
    ShEvaluation.Unprotect
    Set TableCS = ShEvaluation.ListObjects("TableEvaluation")

    If Not ActiveCell.Offset(1).ListObject Is Nothing Then
        If ActiveCell.Offset(1).ListObject.Name = TableCS.Name And ActiveCell.Column - TableCS.Range.Column + 1 > 3 Then
            MsgAnswer = MsgBox("Voulez vous supprimer " & ActiveCell.value & "?", vbOKCancel)
            If MsgAnswer = vbOK Then
                TableCS.ListColumns(ActiveCell.Column - TableCS.Range.Column + 1).Delete
                Call RefreshEvaluationTable
                ActiveCell.Delete Shift:=xlToLeft
            End If
        Else
            MsgBox ("Sélection n'est pas valide")
        End If
    Else
        MsgBox ("Sélection n'est pas valide")
    End If
    StopCode
    ShEvaluation.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
End Sub


Public Sub RefreshEvaluationTable()
    Dim TableCS As ListObject
    Dim i As Long
    Dim Sep As String
    Dim Coeff As String
    
    Sep = "," 'Application.International(xlListSeparator)
    Set TableCS = ShEvaluation.ListObjects("TableEvaluation")
    StartCode
    Coeff = IIf(ShSettings.Range("Langue") = 2, "Coefficient de pondération", "Weighting coefficient")
    
    TableCS.TotalsRowRange.Offset(1).EntireRow.ClearContents
    TableCS.TotalsRowRange(, 2).Offset(1) = "Niveaux de priorité"
    TableCS.TotalsRowRange(, 1).Offset(1) = "-"
    TableCS.TotalsRowRange(, 3).Offset(1) = "-"
    For i = 4 To TableCS.ListColumns.count
        TableCS.HeaderRowRange(i) = "Project " & i - 3
        TableCS.TotalsRowRange(, i).FormulaR1C1 = _
        "=SUM(SUMPRODUCT(SUMIFS(TableListCriteria[Points]" & Sep & " TableListCriteria[Code]" & Sep & " [Code]" & Sep & " TableListCriteria[Critere]" & Sep & " [Project " & i - 3 & "]" & Sep & " TableListCriteria[Critere]" & Sep & " ""<> """""")" & Sep & " [" & Coeff & "]))"
        TableCS.TotalsRowRange(, i).Offset(1).FormulaArray = _
        "=INDEX(TablePriority[Priorité]" & Sep & " MATCH(1" & Sep & " (TableEvaluation[[#Totals]" & Sep & " [Project " & i - 3 & "]]>=TablePriority[Min])*(TableEvaluation[[#Totals]" & Sep & " [Project " & i - 3 & "]]<TablePriority[Max])" & Sep & " 0))"
    Next i
    Call RefreshCondFormPriority
    StopCode
End Sub

Public Sub RefreshCondFormPriority()
    Dim TableList As ListObject
    Dim TableCS As ListObject
    Dim e As Range

    StopCode

    Set TableList = ShSettings.ListObjects("TablePriority")
    Set TableCS = ShEvaluation.ListObjects("TableEvaluation")
    ShEvaluation.Cells.FormatConditions.Delete
    For Each e In TableList.ListColumns("Priorité").DataBodyRange
        With Union(TableCS.DataBodyRange, TableCS.TotalsRowRange)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=D$" & TableCS.TotalsRowRange(, 1).Offset(1).Row & "=""" & e.value & """"
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
    StartCode
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
    StartCode
    Set dict = CreateLangDict
    
    ShEvaluation.Unprotect
    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("E1:N5"), .Range("E6"), .Range("L6"), .Range("L6"), .Range("F11"), .Range("J11"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShCriteresSelection
        trad = vbNullString
        For Each e In Union(.Range("F1:F5"), .Range("C6"), .Range("C8:G10"), .Range("F15"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TableCriteressSelection").Range
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    
    End With
    
    With ShEvaluation
        trad = vbNullString
        For Each e In Union(.Range("E8:K12"), .Range("LevelPriority"), .Range("LevelPriority").Offset(-1), .Range("HeaderEvaluation"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TableEvaluation").Range
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    
    End With
    
    With ShGuide
        trad = vbNullString
        For Each e In .Range("G1:G5")
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
    Call UpdateEvaluationTable
    ShEvaluation.Protect
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





' --- Module: Module1.bas ---
Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Unprotect
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWorkbook.Save
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Unprotect
    Selection.Delete Shift:=xlToLeft
End Sub


