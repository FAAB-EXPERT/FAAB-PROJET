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

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ShBackend.Range("IsMaximize") = False
    Application.DisplayFullScreen = False
    ActiveWindow.DisplayHeadings = True
End Sub

Private Sub Workbook_Open()
    Application.AutoCorrect.AutoFillFormulasInLists = False
End Sub


' --- Module: ShGantt.cls ---
Attribute VB_Name = "ShGantt"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_ColumnInsert(ByVal Col As Range)

    MsgBox "test"

End Sub


' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit
Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnGantt": ShGantt.Activate
        Case "BtnHistorique": ShHistorique.Activate
        Case "BtnHistorique": ShHistorique.Activate
        Case "BtnCouts": ShSuiviCouts.Activate
        Case "BtnDashboard": ShDashboard.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Public Sub StartCode()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ShGantt.Unprotect
End Sub

Public Sub StopCode()
    ShGantt.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub DetectExcelLanguage()
    Dim functionName As String
    Dim isFrench As Boolean
    
    functionName = "SUM"
    
    ' V�rifier si la fonction existe en fran�ais
    On Error Resume Next
    Application.WorksheetFunction.Evaluate functionName & "(1,2)"
    If Err.Number = 0 Then
        isFrench = True
    Else
        isFrench = False
    End If
    On Error GoTo 0
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

    GetIsoWeek = "ISOWEEKNUM"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetIsoWeek = "NO.SEMAINE.ISO"

'    GetIsoWeek = "ISOWEEKNUM"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetIsoWeek & "(1,2)"
'    If Err.Number = 0 Then GetIsoWeek = "NO.SEMAINE.ISO"
'    On Error GoTo 0
End Function

Public Function GetToday()

    GetToday = "Today"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetToday = "AUJOURDHUI"

'    GetToday = "Today"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetToday & "(1,2)"
'    If Err.Number = 0 Then GetToday = "AUJOURDHUI"
'    On Error GoTo 0
End Function


Public Function GetYear()

    GetYear = "Year"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetYear = "ANNEE"
    
    
'    GetYear = "Year"
'
'    On Error Resume Next
'    Application.WorksheetFunction.Evaluate GetYear & "(1,2)"
'    If Err.Number = 0 Then GetYear = "ANNEE"
'    On Error GoTo 0
End Function

Public Sub AddRow()
    If Application.Caller = "BtnAddHolidays" Then
        Call AddRowHelper("TabOffDays")
    ElseIf Application.Caller = "BtnAddStatut" Then
        Call AddRowHelper("TabTheme")
    End If
End Sub

Public Sub DeleteRow()
    If Application.Caller = "BtnDeleteHolidays" Then
        Call DeleteRowHelper("TabOffDays")
    ElseIf Application.Caller = "BtnDeleteStatut" Then
        Call DeleteRowHelper("TabTheme")
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


' --- Module: Gantt.bas ---
Attribute VB_Name = "Gantt"
Option Explicit

Public Sub maximize()
    ShBackend.Range("IsMaximize") = Not ShBackend.Range("IsMaximize")
    Application.DisplayFullScreen = ShBackend.Range("IsMaximize")
    ActiveWindow.DisplayHeadings = Not ShBackend.Range("IsMaximize")
End Sub

Public Sub DeleteColumn()
    Dim MsgAnswer As VbMsgBoxResult
    
    StartCode
    If Selection.count = 1 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de la colonne s�lectionn�e?", vbOKCancel)
        If MsgAnswer = vbOK Then ActiveSheet.Columns(ActiveCell.Column).Delete
    Else
        MsgBox ("Veuillez s�lectionner une seule colonne � supprimer")
    End If
    StopCode
End Sub

Public Sub AddColumn()
    Dim myTab As ListObject
    Dim currentCol As Long
    Dim tableName As String
    
    StartCode
    
    Set myTab = ShGantt.ListObjects("TabProjet")
    
    currentCol = ActiveCell.Column - myTab.DataBodyRange.Column + 1
    ActiveSheet.Columns(ActiveCell.Column).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    With myTab.ListColumns(currentCol).DataBodyRange
        .ClearFormats
        .Cells.Locked = False
        .Validation.Delete
    End With
    StopCode
End Sub

'-----------------------FUNCTION_ADD 26/02/2024---------------------------------------

Public Function GetIdColTable(ByVal toFind As String) As Integer
    Dim e As Range
    
    GetIdColTable = 0
    For Each e In ShGantt.ListObjects("TabProjet").HeaderRowRange
        If e = toFind Then
            GetIdColTable = e.Column '- ShGantt.ListObjects("TabProjet").Range.Column
        End If
    Next e
End Function

Public Function MaxDeadlineDate(Id As Long) As Long
    Dim myTable As ListObject
    Dim e As Range
    Dim currentIndex As Long
    Dim COL_DEADLINE As Integer
    Dim COL_IS_PROJECT As Integer
    
    COL_DEADLINE = GetIdColTable("Deadline")
    COL_IS_PROJECT = GetIdColTable("Is Project")
    MaxDeadlineDate = 0
    Application.Volatile
    Set myTable = ShGantt.ListObjects("TabProjet")
    With myTable
        For Each e In myTable.ListColumns("IDProj").DataBodyRange
            currentIndex = e.Row - myTable.Range.Row
            If e = Id And Cells(e.Row, COL_IS_PROJECT) <> "X" Then
                MaxDeadlineDate = WorksheetFunction.Max(Cells(e.Row, COL_DEADLINE), IIf(MaxDeadlineDate = 0, Cells(e.Row, COL_DEADLINE), MaxDeadlineDate))
            End If
        Next e
    End With
    If MaxDeadlineDate = 0 Then MaxDeadlineDate = ""
End Function

Public Function MinStartDate(Id As Long) As Long
    Dim myTable As ListObject
    Dim e As Range
    Dim currentIndex As Long
    Dim COL_DEBUT As Integer
    Dim COL_IS_PROJECT As Integer

    COL_DEBUT = GetIdColTable("D�but")
    COL_IS_PROJECT = GetIdColTable("Is Project")

    MinStartDate = 0
    Application.Volatile
    Set myTable = ShGantt.ListObjects("TabProjet")
    With myTable
        For Each e In myTable.ListColumns("IDProj").DataBodyRange
            currentIndex = e.Row - myTable.Range.Row
            If e = Id And Cells(e.Row, COL_IS_PROJECT) <> "X" Then
                MinStartDate = WorksheetFunction.Min(Cells(e.Row, COL_DEBUT), IIf(MinStartDate = 0, Cells(e.Row, COL_DEBUT), MinStartDate))
            End If
        Next e
    End With
    If MinStartDate = 0 Then MinStartDate = ""

End Function

Public Sub GoToToday()
    StartCode
    ShGantt.Range("ProjectStart") = Date
    UpdateMonths
    StopCode
End Sub

Public Sub UpdateStartingDate()
    Dim isNextOrPrev As Integer
    Dim intervalType As String
    
    StartCode
    isNextOrPrev = IIf(Left(Application.Caller, 7) = "BtnPrev", -1, 1)
    
    With ShBackend
        If .Range("IsDaysGantt") And Right(Application.Caller, 5) = "Speed" Then
            intervalType = "ww"
        ElseIf .Range("IsDaysGantt") Then
            intervalType = "d"
        ElseIf .Range("IsWeeksGant") And Right(Application.Caller, 5) = "Speed" Then
            intervalType = "m"
        ElseIf .Range("IsWeeksGant") Then
            intervalType = "ww"
        ElseIf .Range("IsMonthsGantt") And Right(Application.Caller, 5) = "Speed" Then
            intervalType = "q"
        ElseIf .Range("IsMonthsGantt") Then
            intervalType = "m"
        End If
    End With
    ShGantt.Range("ProjectStart") = DateAdd(intervalType, isNextOrPrev, ShGantt.Range("ProjectStart"))
    Call UpdateMonths
    StopCode
End Sub

Private Sub RefreshFormulaTabProjects()
    Dim myTable As ListObject
    Dim currentIndex As Long
    
    Set myTable = ShGantt.ListObjects("TabProjet")
    If myTable.ListRows.count > 0 Then
        myTable.ListColumns("IDProj").DataBodyRange.Formula = "=IF(ISNUMBER($D13),IF([@[Is Project]] <>"""",$D13+1,$D13),1)"
        myTable.ListColumns("Point").DataBodyRange.Formula = "=IF([@[Is Project]] <> """","""",""."")"
        myTable.ListColumns("IDSubProj").DataBodyRange.Formula = "=IFERROR(IF([@[Is Project]]= """",$F13+1,""""),1)"
    End If
End Sub


Public Sub UpdateMonths()
    Dim e As Range
    Dim count As Long
    Dim oldValue As Long
    Dim currentMonth As Long
 '   StartCode
    With ShGantt.ListObjects("TabProjet")
        count = .ListColumns("C0").index + .HeaderRowRange.Column - 1
    End With
    oldValue = 0
    Application.Calculation = xlCalculationManual
    ShGantt.Range("GanttHeaderMonthsWeeks").ClearContents
    ShGantt.Range("GanttHeaderYears").ClearContents
    With ShGantt
        For Each e In .Range("GanttHeaderDateWeeks")
            currentMonth = Month(DateSerial(Year(e), Month(e), Day(e) - Weekday(e, 2) + 1))
            If oldValue <> currentMonth Then
                .Cells(5, count).Formula = "=UPPER(TEXT(DATE(YEAR(" & e.Address & "), MONTH(" & e.Address & "), DAY(" & e.Address & ")-WEEKDAY(" & e.Address & ", 2) + 1),""mmmm""))"
                .Cells(6, count).Formula = "=Right(""0000"" & Year(" & e.Address & "), 4)"
            End If
            If .Cells(6, count) = vbNullString And oldValue <> 0 Then
                .Cells(7, count) = oldValue
            Else
                .Cells(7, count) = currentMonth
            End If
            
            If .Cells(7, count) Mod 2 = 0 Then
                .Cells(5, count).Resize(7).Interior.Color = ShSettings.Range("FirstMonthTheme").Interior.Color
            Else
                .Cells(5, count).Resize(7).Interior.Color = ShSettings.Range("SecondMonthTheme").Interior.Color
            End If
            count = count + 1
            oldValue = currentMonth
        Next e
    End With
 '   StopCode
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub SwitchDaysWeeks()
    StartCode
    With ShGantt
        If ShBackend.Range("IsDaysGantt") Then
            .Range("GanttHeader").NumberFormat = "d"
            .Range("GanttHeader").EntireRow.Hidden = False
        ElseIf ShBackend.Range("IsWeeksGant") Then
            .Range("GanttHeader").NumberFormat = "0"
            .Range("GanttHeader").EntireRow.Hidden = False
        Else
            .Range("GanttHeader").NumberFormat = "0"
            .Range("GanttHeader").EntireRow.Hidden = True
        End If
    End With
    Call RefreshCondFormat
    StopCode
End Sub

Public Sub AddProject()
    Dim myTable As ListObject
    Dim NewRow As ListRow
    
    StartCode
    Set myTable = ShGantt.ListObjects("TabProjet")
    
    With myTable
        If ActiveCell.ListObject Is Nothing Or .ListRows.count = 0 Then
            Set NewRow = .ListRows.Add
        ElseIf ActiveCell.ListObject.Name = myTable.Name Then
            Set NewRow = .ListRows.Add(ActiveCell.Row - .DataBodyRange.Row + 1)
        End If
        If Application.Caller = "BtnAddProject" Then
            .ListColumns("Is Project").DataBodyRange.Rows(NewRow.index) = "X"
            .ListColumns("Progression").DataBodyRange.Rows(NewRow.index).Formula = "=IFERROR(AVERAGEIFS([Progression],[Is Project],"""",[IDProj],[@IDProj]), 0)"
            '.ListColumns("D�but").DataBodyRange.Rows(NewRow.Index).Formula = "=IF(MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]) = 0, """", MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]))"
            .ListColumns("D�but").DataBodyRange.Rows(NewRow.index).Formula = "=IFERROR(MinStartDate([@ID]), """")"
            .ListColumns("Charge en J").DataBodyRange.Rows(NewRow.index).Formula = "=IFERROR(SUMIFS([Charge en J],[Is Project],"""",[IDProj],[@IDProj]), """")"
            .ListColumns("Deadline").DataBodyRange.Rows(NewRow.index).Formula = "=IFERROR(MaxDeadlineDate([@ID]), """")"
        Else
            .ListColumns("Progression").DataBodyRange.Rows(NewRow.index) = 0
            .ListColumns("D�but").DataBodyRange.Rows(NewRow.index).Formula = ""
            .ListColumns("Charge en J").DataBodyRange.Rows(NewRow.index).Formula = ""
            .ListColumns("Deadline").DataBodyRange.Rows(NewRow.index).Formula = "=IF([@D�but]="""","""",WORKDAY(IF(WEEKDAY([@D�but],1)=7,[@D�but]+2,IF(WEEKDAY([@D�but],1)=1,[@D�but]+1,[@D�but])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),TabOffDays[Date]))"
        End If
        .DataBodyRange.Rows.AutoFit
        Call RefreshFormulaTabProjects
        Call RefreshCondFormat
    End With
    StopCode
End Sub

Public Sub DeleteProject()
    Dim myTable As ListObject
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    
    StartCode
    Application.Calculation = xlCalculationManual
    Set myTable = ShGantt.ListObjects("TabProjet")
    
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
    
    If myTable.ListRows.count > 0 Then myTable.DataBodyRange.Rows.AutoFit
    Call RefreshFormulaTabProjects
    Application.Calculation = xlCalculationAutomatic
    StopCode
End Sub

Public Sub ConvertToProject()
    Dim myTable As ListObject
    Dim currentIndex As Long
    
    StartCode
    Set myTable = ShGantt.ListObjects("TabProjet")
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
        MsgBox ("La s�lection n'est pas valide")
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        currentIndex = ActiveCell.Row - myTable.DataBodyRange.Row + 1
        With myTable.ListColumns("Is Project").DataBodyRange
            If .Rows(currentIndex) = "X" Then
                .Rows(currentIndex) = ""
                myTable.ListColumns("Progression").DataBodyRange.Rows(currentIndex) = 0
                myTable.ListColumns("D�but").DataBodyRange.Rows(currentIndex).Formula = ""
                myTable.ListColumns("Charge en J").DataBodyRange.Rows(currentIndex).Formula = ""
                myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IF([@D�but]="""","""",WORKDAY(IF(WEEKDAY([@D�but],1)=7,[@D�but]+2,IF(WEEKDAY([@D�but],1)=1,[@D�but]+1,[@D�but])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),TabOffDays[Date]))"
            Else
                .Rows(currentIndex) = "X"
                myTable.ListColumns("Progression").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(AVERAGEIFS([Progression],[Is Project],"""",[IDProj],[@IDProj]), 0)"
                'myTable.ListColumns("D�but").DataBodyRange.Rows(currentIndex).Formula = "=IF(MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]) = 0, """", MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]))"
                myTable.ListColumns("D�but").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(MinStartDate([@ID]), """")"
                myTable.ListColumns("Charge en J").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(SUMIFS([Charge en J],[Is Project],"""",[IDProj],[@IDProj]), """")"
                myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(MaxDeadlineDate([@ID]), """")"
            End If
            'myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IF([@D�but]="""","""",WORKDAY(IF(WEEKDAY([@D�but],1)=7,[@D�but]+2,IF(WEEKDAY([@D�but],1)=1,[@D�but]+1,[@D�but])),roundup([@[Charge en J]],0)-1,TabOffDays[Date]))"
        End With
        myTable.DataBodyRange.Rows.AutoFit
        Call RefreshFormulaTabProjects
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
    Call RefreshCondFormat
    StopCode
End Sub


' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit
Public Sub RefreshCondFormat()
    If ShGantt.ListObjects("TabProjet").ListRows.count > 0 Then
        ShGantt.Cells.FormatConditions.Delete
        Call RefreshCondFormIsProject
        Call RefreshCondFormListGantt
        Call RefreshCondFormThemeCol
        Call RefreshCondFormToday
        Call UpdateMonths
        Call RefreshCondFormProgression
    End If
End Sub
Public Sub test()

    MsgBox GetAdrressColTable("Th�me")
    'MsgBox Cells(1, 1).Address
    'MsgBox ShGantt.ListObjects("TabProjet").Range.Column
End Sub

Public Function GetAdrressColTable(ByVal toFind As String) As String
    Dim e As Range
    
    GetAdrressColTable = 0
    For Each e In ShGantt.ListObjects("TabProjet").HeaderRowRange
        If e = toFind Then
            GetAdrressColTable = Mid(e.Address, 2, 1)
        End If
    Next e
    
End Function

Sub RefreshCondFormThemeCol()
    Dim TableList As ListObject
    Dim e As Range
    Dim COL_THEME As String
    
    COL_THEME = "$" & GetAdrressColTable("Th�me") & 14
    
    Set TableList = ShSettings.ListObjects("TabTheme")

    For Each e In TableList.ListColumns("Libell�").DataBodyRange
        With ShGantt.ListObjects("TabProjet").ListColumns("Th�me").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & COL_THEME & "=""" & e.value & """"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Font
                On Error Resume Next
                .Color = e.Offset(, 1).Font.Color
                If e.Offset(, 1).Font.ThemeColor > 0 Then .ThemeColor = e.Offset(, 1).Font.ThemeColor
                .Bold = e.Offset(, 1).Font.Bold
                .Italic = e.Offset(, 1).Font.Italic
                .TintAndShade = e.Offset(, 1).Font.TintAndShade
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = e.Offset(, 1).Interior.PatternColorIndex
                If e.Interior.ThemeColor > 0 Then
                    .ThemeColor = e.Offset(, 1).Interior.ThemeColor
                    .Pattern = e.Offset(, 1).Interior.Pattern
                    .PatternColorIndex = e.Offset(, 1).Interior.PatternColorIndex
                Else
                    .Color = e.Offset(, 1).Interior.Color
                End If
                .TintAndShade = e.Offset(, 1).Interior.TintAndShade
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Application.CutCopyMode = True
    Next e
End Sub

Sub RefreshCondFormListGantt()
    Dim TableList As ListObject
    Dim e As Range
    Dim myFormula As String
    Dim cdtFormat As FormatCondition
    Dim C_C0 As String
    Dim C_DEBUT As String
    Dim C_PROG As String
    Dim C_DA_PROG As String
    Dim C_STATUT As String
    Dim C_DEADLINE As String
    
    C_C0 = GetAdrressColTable("C0") 'P " & C_C0 & "
    C_DEBUT = GetAdrressColTable("D�but") 'K " & C_DEBUT & "
    C_PROG = GetAdrressColTable("Progression") 'N " & C_PROG & "
    C_DA_PROG = GetAdrressColTable("DateProgression") 'O " & C_DA_PROG & "
    C_STATUT = GetAdrressColTable("Th�me") 'H " & C_STATUT & "
    C_DEADLINE = GetAdrressColTable("Deadline") 'M " & C_DEADLINE & "
    
    Set TableList = ShSettings.ListObjects("TabTheme")

    For Each e In TableList.ListColumns("Couleur du % r�alis�").DataBodyRange.Resize(, 2)
        If ShBackend.Range("IsDaysGantt") Then
            If e.Column = TableList.DataBodyRange.Column + 1 Then
                'myFormula = "=" & GetAnd & "(P$9>=$K14" & GetSep & " $K14 <> """"" & GetSep & " P$9<=$O14" & GetSep & " $N14 <> 0" & GetSep & " $H14 =""" & e.Offset(, -1) & """)"
                myFormula = "=" & GetAnd & "(" & C_C0 & "$9>=$" & C_DEBUT & "14" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep & " " & C_C0 & "$9 <= $" & C_DA_PROG & "14" & GetSep & " $" & C_PROG & "14 <> 0" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -1) & """)"
            
            Else
                'myFormula = "=" & GetAnd & "(P$9>=$K14" & GetSep & " $K14 <> """"" & GetSep & " P$9 <= $M14" & GetSep & " $M14 <> """"" & GetSep & " $H14 =""" & e.Offset(, -2) & """)"
                myFormula = "=" & GetAnd & "(" & C_C0 & "$9>=$" & C_DEBUT & "14" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep & " " & C_C0 & "$9 <= $" & C_DEADLINE & "14" & GetSep & " $" & C_DEADLINE & "14 <> """"" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -2) & """)"

            End If
        Else
            If e.Column = TableList.DataBodyRange.Column + 1 Then
                myFormula = "=" & GetOr & "(" & GetAnd & "(" & C_C0 & "$8>=$" & C_DEBUT & "14" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep & " " & C_C0 & "$8<=$" & C_DA_PROG & "14" & GetSep & " $" & C_PROG & "14 <> 0" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -1) & """)" & GetSep & " " & GetAnd & "(" & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8)" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) = " & GetIsoWeek & "($" & C_DEBUT & "14)" & GetSep & " $" & C_PROG & "14 <> 0" & GetSep & " $" & C_STATUT & "14 =""" & e & """))"
            Else
                myFormula = "=" & GetOr & "(" & GetAnd & "(" & C_C0 & "$8>=$" & C_DEBUT & "14" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep & " " & C_C0 & "$8<=$" & C_DEADLINE & "14" & GetSep & " $" & C_DEADLINE & "14 <> """"" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -2) & """)" & GetSep & " " & GetAnd & "(" & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8)" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) = " & GetIsoWeek & "($" & C_DEBUT & "14)" & GetSep & " $" & C_STATUT & "14 =""" & e & """))"
            End If
        End If
        With ShGantt.ListObjects("TabProjet").ListColumns("C0").DataBodyRange.Resize(, 56)
            Set cdtFormat = .FormatConditions.Add(Type:=xlExpression, Formula1:=myFormula)
                cdtFormat.SetFirstPriority
            With .FormatConditions(1).Font
                'On Error Resume Next
                .Color = e.Font.Color
                'If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
'                .Bold = e.Font.Bold
'                .Italic = e.Font.Italic
'                .TintAndShade = e.Font.TintAndShade
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
            With .FormatConditions(1)
                .Borders(xlLeft).LineStyle = xlNone
                .Borders(xlRight).LineStyle = xlNone
            End With
            .FormatConditions(1).StopIfTrue = True
            If e.Column <> TableList.DataBodyRange.Column + 1 Then
                cdtFormat.SetLastPriority
            End If
        End With
        Application.CutCopyMode = False
    Next e
End Sub

Sub RefreshCondFormIsProject()
    Dim e As Range
    Dim COL_IS_PROJECT As String
    
    COL_IS_PROJECT = "$" & GetAdrressColTable("Is Project") & 14

    Set e = ShSettings.Range("IsProject")
    With ShGantt.ListObjects("TabProjet").ListColumns("ID").DataBodyRange.Resize(, 13)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & COL_IS_PROJECT & "<>"""""
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
End Sub

Sub RefreshCondFormToday()
    Dim e As Range
    Dim myFormula As String
    Dim C_C0 As String
    Dim C_C1 As String
    
    C_C0 = GetAdrressColTable("C0") 'P " & C_C0 & "
    C_C1 = GetAdrressColTable("C1") 'Q " & C_C1 & "


    If ShBackend.Range("IsDaysGantt") Then
        myFormula = "=" & GetAnd & "(" & GetToday & "()>=" & C_C0 & "$9" & GetSep & "" & GetToday & "()<" & C_C1 & "$9)"
    Else
        myFormula = "=" & GetAnd & "(" & GetIsoWeek & "(" & GetToday & "())>=" & C_C0 & "$9" & GetSep & "" & GetIsoWeek & "(" & GetToday & "())<" & C_C1 & "$9)"
    End If
        

    Set e = ShSettings.Range("TodayFormat")
    With ShGantt.ListObjects("TabProjet").ListColumns("C0").DataBodyRange.Resize(, 56)
        .FormatConditions.Add Type:=xlExpression, Formula1:=myFormula
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Font
            On Error Resume Next
            .Color = e.Font.Color
            If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
            .Bold = e.Font.Bold
            .Italic = e.Font.Italic
            .TintAndShade = e.Font.TintAndShade
        End With
'        With .FormatConditions(1).Interior
'            .PatternColorIndex = e.Interior.PatternColorIndex
'            If e.Interior.ThemeColor > 0 Then
'                .ThemeColor = e.Interior.ThemeColor
'                .Pattern = e.Interior.Pattern
'                .PatternColorIndex = e.Interior.PatternColorIndex
'            Else
'                .Color = e.Interior.Color
'            End If
'            .TintAndShade = e.Interior.TintAndShade
'        End With
        With .FormatConditions(1).Borders(xlEdgeLeft)
            .LineStyle = e.Borders(xlEdgeLeft).LineStyle
            .ThemeColor = e.Borders(xlEdgeLeft).ThemeColor
            .TintAndShade = e.Borders(xlEdgeLeft).TintAndShade
            .Weight = e.Borders(xlEdgeLeft).Weight
        End With
        With .FormatConditions(1).Borders(xlEdgeRight)
            .LineStyle = e.Borders(xlEdgeRight).LineStyle
            .ThemeColor = e.Borders(xlEdgeRight).ThemeColor
            .TintAndShade = e.Borders(xlEdgeRight).TintAndShade
            .Weight = e.Borders(xlEdgeRight).Weight
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    Application.CutCopyMode = False
End Sub

Sub RefreshCondFormProgression()
    With ShGantt.ListObjects("TabProjet").ListColumns("Progression").DataBodyRange
        .FormatConditions.AddDatabar
        .FormatConditions(.FormatConditions.count).ShowValue = True
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
            .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
            With .BarColor
                .Color = 8700771
                .TintAndShade = 0
            End With
            .BarFillType = xlDataBarFillGradient
            .Direction = xlContext
            .NegativeBarFormat.ColorType = xlDataBarColor
            .BarBorder.Type = xlDataBarBorderSolid
            .NegativeBarFormat.BorderColorType = xlDataBarColor
            With .BarBorder.Color
                .Color = 8700771
                .TintAndShade = 0
            End With
            .AxisPosition = xlDataBarAxisAutomatic
            With .AxisColor
                .Color = 0
                .TintAndShade = 0
            End With
            With .NegativeBarFormat.Color
                .Color = 255
                .TintAndShade = 0
            End With
            With .NegativeBarFormat.BorderColor
                .Color = 255
                .TintAndShade = 0
            End With
        End With
    End With
End Sub





' --- Module: ShBackend.cls ---
Attribute VB_Name = "ShBackend"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
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
Option Explicit



' --- Module: ShSuiviCouts.cls ---
Attribute VB_Name = "ShSuiviCouts"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: ShHistorique.cls ---
Attribute VB_Name = "ShHistorique"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: ShDashboard.cls ---
Attribute VB_Name = "ShDashboard"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
    Dim e As Range
    Dim cc As Range
    
  '  StartCode
    ShDashboard.Range("F17:H22").ClearContents
    Set cc = ShDashboard.Range("F17")
    For Each e In ShGantt.ListObjects("TabProjet").ListColumns("Point").DataBodyRange
        If e = vbNullString Then
            cc = e.Offset(, -1).value
            cc.Offset(, 1).value = e.Offset(, 2).value
            cc.Offset(, 2).value = e.Offset(, 9).value
            Set cc = cc.Offset(1)
        End If
    Next e
    ShBackend.ListObjects("Table_TabStatus").QueryTable.Refresh BackgroundQuery:=False
    Call UpdateGraphColor
    StopCode
End Sub

Private Sub UpdateGraphColor()
    Dim myChart As ChartObject
    Dim i As Long
    Dim index As Long
    
    Set myChart = ShDashboard.ChartObjects("Chart 2")
    i = 1
    
    '  .ForeColor.RGB = RGB(255, 0, 0) ' Set fill color to red
    While i <= ShBackend.ListObjects("Table_TabStatus").ListColumns("Th�me").DataBodyRange.count
        index = FindColorInTheme(ShBackend.Range("A" & 38 + i))
        If index > 0 Then
            With myChart.Chart.SeriesCollection(1).Points(i).Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = ShSettings.Range("C" & index).Interior.Color
                .Solid
            End With
        End If
        i = i + 1
    Wend

End Sub

Private Function FindColorInTheme(toFind As String)
    Dim e As Range
    
    FindColorInTheme = 0
    For Each e In ShSettings.ListObjects("TabTheme").ListColumns("Libell�").DataBodyRange
        If e.value = toFind Then
            FindColorInTheme = e.Row
        End If
    Next e

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
    
    'SETTINGS SHEET-----------------------------------------------
    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("J1:J3"), .Range("B4"), .Range("K4"), .Range("B6:J8"), .Range("I10:I21"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        For Each e In .ListObjects("TabTheme").ListColumns("Libell�").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        For Each e In .ListObjects("TabOffDays").ListColumns("Jour f�ri�").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e

    End With

    'ShCost Tracking-----------------------------------------------
    With ShSuiviCouts
        trad = vbNullString
        For Each e In Union(.Range("G1:G3"), .Range("B4"), .Range("C7:D7"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    'ShHistorique-----------------------------------------------
    With ShHistorique
        trad = vbNullString
        For Each e In Union(.Range("I1:I3"), .Range("B4"), .Range("C8:E8"), .Range("F9"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    'ShDashboard-----------------------------------------------
    With ShDashboard
        trad = vbNullString
        For Each e In Union(.Range("H1:H3"), .Range("B4"), .Range("C17:C20"), .Range("F9"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    'ShBackend-----------------------------------------------
    With ShBackend
        trad = vbNullString
        For Each e In Union(.Range("E28:F31"), .Range("A19"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    'ShGantt SHEET-----------------------------------------------
    With ShGantt
        trad = vbNullString
        For Each e In Union(.Range("H1:M3"), .Range("AB1:AB3"), .Range("B11:N11"), .Range("B12"), .Range("G12"), .Range("G4"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In Union(.ListObjects("TabProjet").ListColumns("WBS - Activit� / t�che").DataBodyRange, .ListObjects("TabProjet").ListColumns("Th�me").DataBodyRange)
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = dict(.OptionButtons("Option Button 3").Caption)
        If trad <> vbNullString Then
            .OptionButtons("Option Button 3").Caption = trad
        End If
        trad = dict(.OptionButtons("Option Button 4").Caption)
        If trad <> vbNullString Then
            .OptionButtons("Option Button 4").Caption = trad
        End If
        trad = dict(.OptionButtons("Option Button 8").Caption)
        If trad <> vbNullString Then
            .OptionButtons("Option Button 8").Caption = trad
        End If
    End With
'
'    'ShGuide SHEET-----------------------------------------------
'    With ShGuide
'        trad = vbNullString
'        For Each e In .Range("G1:G5")
'            trad = dict(e.value)
'            If e <> vbNullString And trad <> vbNullString Then
'                e = trad
'            End If
'        Next e
'    End With
    
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
    
    CELL_KEY = ShSettings.Range("Langue")
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








' --- Module: Version.bas ---
Attribute VB_Name = "Version"
Option Explicit

'Last update 26/02/2024 By SO
'-- Permettre l'ajout et la suppression de colonnes
'-- Permettre de changer la devise



' --- Module: Settings_Devise.bas ---
Attribute VB_Name = "Settings_Devise"
Option Explicit

Public Sub ChangeDevise()
    Dim celDevise As Range
    
    Set celDevise = ShSettings.Range("SETTINGS_DEVISE")
    
    celDevise.Select
    Application.Dialogs(xlDialogFormatNumber).Show
    With ShSuiviCouts
        .Range("TabCo�ts[Co�ts consom�s par mois]").NumberFormat = celDevise.NumberFormat
        .Range("TOTAL_BUDGET_TO_DATE").NumberFormat = celDevise.NumberFormat
    End With
    ShSettings.Range("BUDGET_INI_SETTINGS").NumberFormat = celDevise.NumberFormat
    ShBackend.Range("BUDGET_RANGE").NumberFormat = celDevise.NumberFormat
End Sub


