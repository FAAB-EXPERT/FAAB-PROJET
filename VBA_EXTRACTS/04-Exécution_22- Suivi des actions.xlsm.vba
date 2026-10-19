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



' --- Module: ShActions.cls ---
Attribute VB_Name = "ShActions"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tabActions As ListObject
    Dim currentRow As Long
    Dim currentCol As Long
    
    StartCode
    Set tabActions = ShActions.ListObjects("TabActions")
    If Not Intersect(Target, tabActions.ListColumns("Réalisation").DataBodyRange) Is Nothing Then
        currentRow = Target.Row - tabActions.HeaderRowRange.Row
        If Target = 1 Then
           tabActions.ListColumns("Etat").DataBodyRange(currentRow) = IIf(ShSettings.Range("Langue") = 2, "Réalisé", "Completed")
        ElseIf Target < 1 Then
           tabActions.ListColumns("Etat").DataBodyRange(currentRow) = IIf(ShSettings.Range("Langue") = 2, "En cours", "In Progress")
        End If
    ElseIf Not Intersect(Target, tabActions.ListColumns("Etat").DataBodyRange) Is Nothing Then
        currentRow = Target.Row - tabActions.HeaderRowRange.Row
        If Target = "Réalisé" Then
           tabActions.ListColumns("Réalisation").DataBodyRange(currentRow) = 1
        ElseIf Target <> "Réalisé" Then
           tabActions.ListColumns("Réalisation").DataBodyRange(currentRow) = 0
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

Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnActions": ShActions.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Public Sub DeleteRow()
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    
    StartCode
    Application.Calculation = xlCalculationManual
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    Set myTable = ActiveSheet.ListObjects(MyCell.ListObject.Name)
    
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
    StopCode
End Sub

Public Function GetSep()
    GetSep = Application.International(xlListSeparator)
End Function

Public Function GetOr()

    GetOr = "Or"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetOr = "Ou"
    
'    GetOr = "OR"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetOr & "(1,2)"
'    If Err.Number = 0 Then GetOr = "Ou"
'    On Error GoTo 0
End Function

'Public Function GetOr()

'End Function

Public Function GetAnd()
    GetAnd = "And"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetAnd = "Et"
    
'    GetAnd = "And"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetAnd & "(1,2)"
'    If Err.Number = 0 Then GetAnd = "Et"
'    On Error GoTo 0
End Function

Public Function GetIsoWeek()

    GetIsoWeek = "ISOWEEKNUM()"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetIsoWeek = "NO.SEMAINE.ISO()"
'
'    GetIsoWeek = "ISOWEEKNUM"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetIsoWeek & "(1,2)"
'    If Err.Number = 0 Then GetIsoWeek = "NO.SEMAINE.ISO"
'    On Error GoTo 0
End Function

Public Function GetToday()

    GetToday = "Today()"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetToday = "AUJOURDHUI()"

'    GetToday = "Today"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetToday & "(1,2)"
'    If Err.Number = 0 Then GetToday = "AUJOURDHUI"
'    On Error GoTo 0
End Function


Public Function GetYear()

    GetYear = "Year()"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetYear = "ANNEE()"
'    GetYear = "Year"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetYear & "(1,2)"
'    If Err.Number = 0 Then GetYear = "ANNEE"
'    On Error GoTo 0
End Function

Public Sub AddRow()
    Dim myTable As ListObject
    Dim MyCell As Range
    
    StartCode
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    Set myTable = MyCell.ListObject
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
        myTable.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
    End If
    StopCode
End Sub

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
    ShActions.Cells.FormatConditions.Delete
    Call RefreshCondFormListEtat


    Call RefreshCondFormListWarningPlanEndDateRed
    Call RefreshCondFormListWarningEndDate
    Call RefreshCondFormListWarningPlanEndDate


    StopCode
End Sub




' --- Module: Actions.bas ---
Attribute VB_Name = "Actions"
Option Explicit

Sub RefreshCondFormListEtat()
    Dim TableList As ListObject
    Dim e As Range
    Dim myFormula As String
    Set TableList = ShSettings.ListObjects("TabEtat")

    myFormula = "=" & Mid(ShActions.ListObjects("TabActions").ListColumns("Etat").DataBodyRange(1).Address, 1, 2) & "12="
    For Each e In TableList.ListColumns("Etat").DataBodyRange
        With ShActions.ListObjects("TabActions").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:=myFormula & """" & e.value & """"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Font
                On Error Resume Next
                .Color = e.Font.Color
                If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
                .Bold = e.Font.Bold
                .Italic = e.Font.Italic
                .TintAndShade = e.Font.TintAndShade
                .Strikethrough = e.Font.Strikethrough
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

Sub test()
    Dim e As Range
    Dim colEcheancePlanifiee As String
    Dim colEtat As String
    Dim tabActions As ListObject
    
    Set e = ShSettings.Range("SeuilAvertissement")
    Set tabActions = ShActions.ListObjects("TabActions")
    colEcheancePlanifiee = Mid(tabActions.ListColumns("Échéance planifiée").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    colEtat = Mid(tabActions.ListColumns("Etat").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    
    MsgBox (colEcheancePlanifiee)
    MsgBox (colEtat)
    
End Sub

Sub RefreshCondFormListWarningPlanEndDate()
    Dim e As Range
    Dim CellEcheancePlanifiee As String
    Dim cellEtat As String
    Dim tabActions As ListObject
    
    Set e = ShSettings.Range("SeuilAvertissement")
    Set tabActions = ShActions.ListObjects("TabActions")
    CellEcheancePlanifiee = Mid(tabActions.ListColumns("Échéance planifiée").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    cellEtat = Mid(tabActions.ListColumns("Etat").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    
    With Union(tabActions.ListColumns("Échéance planifiée").DataBodyRange, tabActions.ListColumns("Date de fin réelle").DataBodyRange)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & GetAnd & "(" & CellEcheancePlanifiee & " <= " & GetToday & " +" & e & GetSep & CellEcheancePlanifiee & " >= " & GetToday & GetSep & cellEtat & " <> ""Completed""" & GetSep & cellEtat & " <> ""Cancelled""" & GetSep & cellEtat & " <> ""Réalisé""" & GetSep & cellEtat & " <> ""Annulé"")"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).StopIfTrue = True
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
    End With
End Sub

Sub RefreshCondFormListWarningPlanEndDateRed()
    Dim e As Range
    Dim CellPlanEndDate As String
    Dim cellEtat As String
    Dim tabActions As ListObject
    
    Set e = ShSettings.Range("RedFlag")
    Set tabActions = ShActions.ListObjects("TabActions")
    CellPlanEndDate = Mid(tabActions.ListColumns("Échéance planifiée").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    cellEtat = Mid(tabActions.ListColumns("Etat").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    
    With Union(tabActions.ListColumns("Échéance planifiée").DataBodyRange, tabActions.ListColumns("Date de fin réelle").DataBodyRange)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & GetAnd & "(" & CellPlanEndDate & " < " & GetToday & GetSep & cellEtat & " <> ""Réalisé""" & GetSep & cellEtat & " <> ""Annulé"")"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).StopIfTrue = True
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
    End With
End Sub

Sub RefreshCondFormListWarningEndDate()
    Dim e As Range
    Dim CellPlanEndDate As String
    Dim CellEndDate As String
    Dim tabActions As ListObject
    
    Set e = ShSettings.Range("RedFlag")
    Set tabActions = ShActions.ListObjects("TabActions")
    CellPlanEndDate = Mid(tabActions.ListColumns("Échéance planifiée").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    CellEndDate = Mid(tabActions.ListColumns("Date de fin réelle").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    
    With ShActions.ListObjects("TabActions").ListColumns("Date de fin réelle").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:="= " & CellPlanEndDate & " < " & CellEndDate & ""
            .FormatConditions(.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).StopIfTrue = True
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
    End With
End Sub

Sub RefreshCondFormListEmptyEndDate()
    Dim e As Range
    Dim CellPlanEndDate As String
    Dim CellEndDate As String
    Dim tabActions As ListObject
    
    Set e = ShSettings.Range("RedFlag")
    Set tabActions = ShActions.ListObjects("TabActions")
    CellPlanEndDate = Mid(tabActions.ListColumns("Échéance planifiée").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    CellEndDate = Mid(tabActions.ListColumns("Date de fin réelle").DataBodyRange(1).Address, 1, 2) & tabActions.DataBodyRange.Row
    
    With ShActions.ListObjects("TabActions").ListColumns("Date de fin réelle").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:="= """
            .FormatConditions(.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).StopIfTrue = True
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
    End With
End Sub
Public Sub AddColumn()
    Dim tabActions As ListObject
    Dim currentCol As Long
    
    StartCode
    Set tabActions = ShActions.ListObjects("tabActions")
    If Not Intersect(ActiveCell, tabActions.HeaderRowRange.Offset(-1)) Is Nothing Then
        currentCol = ActiveCell.Column - tabActions.DataBodyRange.Column + 1
        ShActions.Cells(ShActions.Range("HeaderRange").Row, currentCol + 1).Insert Shift:=xlToRight
        tabActions.ListColumns.Add currentCol
        ShActions.Columns(currentCol + 1).EntireColumn.Validation.Delete
    ElseIf Intersect(ActiveCell, tabActions.DataBodyRange) Is Nothing Then
        tabActions.ListColumns.Add
        Cells(ShActions.Range("HeaderRange").Row, tabActions.ListColumns.count + 1).Insert Shift:=xlToRight
        With Cells(ShActions.Range("HeaderRange").Row, tabActions.ListColumns.count + 1)
            .Offset(, -1).Copy
            .PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End With
        ShActions.Columns(tabActions.ListColumns.count + 1).EntireColumn.Validation.Delete
    Else
        MsgBox ("La selection n'est pas valide !")
    End If
    StopCode
    RefreshCondFormat
End Sub

Public Sub DeleteColumn()
    Dim tabActions As ListObject
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim currentCol As Long
    Dim SelectedRange As Range
    
    StartCode
    Set tabActions = ShActions.ListObjects("tabActions")
    
    Set SelectedRange = Intersect(Selection, Selection.Rows(1), tabActions.HeaderRowRange.Offset(-1))
    If Not SelectedRange Is Nothing And tabActions.ListColumns.count > 1 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de " & SelectedRange.count & " colonne(s) ?", vbOKCancel)
        If MsgAnswer = vbOK Then
            For count = SelectedRange.count To 1 Step -1
                currentCol = SelectedRange(count).Column - tabActions.DataBodyRange.Column + 1
                If tabActions.ListColumns.count > 2 Then
                    If Not Intersect(SelectedRange(1, count), ShActions.Range("HeaderRange")) Is Nothing Then
                        ShActions.Cells(ShActions.Range("HeaderRange").Row, currentCol + 1).Delete Shift:=xlToLeft
                        tabActions.ListColumns(currentCol).Delete
                    End If
                End If
            Next count
        End If
    Else
        MsgBox ("La selection n'est pas valide!")
    End If
    StopCode
    RefreshCondFormat
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
    
    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("E1:N5"), .Range("J6"), .Range("C9"), .Range("C6"), .Range("I9"), .Range("I11"), .Range("E9"), .Range("G9"), .Range("L11"))
            trad = dict(e.value)
            If trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabEtat").DataBodyRange
            trad = dict(e.value)
            If trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShActions
        trad = vbNullString
        For Each e In Union(.Range("E1:M5"), .Range("C10:K10"), .Range("C6"))
            trad = dict(e.value)
            If trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabActions").ListColumns("Etat").DataBodyRange
            trad = dict(e.value)
            If trad <> vbNullString Then
                e = trad
            End If
        Next e
    
    End With
    
    With ShGuide
        trad = vbNullString
        For Each e In Union(.Range("E1:M5"), .Range("C6"))
            trad = dict(e.value)
            If trad <> vbNullString Then
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




