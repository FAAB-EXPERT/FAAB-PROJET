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

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Check if the Ribbon object is loaded and invalidate it
    If Not ribbonUI Is Nothing Then
        ribbonUI.Invalidate
    End If
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
Private Sub Worksheet_Change(ByVal Target As Range)
    StartCode
    If Target.count = 1 Then
        If Not Intersect(Target, Me.Range("ProjectStart")) Is Nothing Then
            Call UpdateMonths
        End If
        If Not Intersect(Target, ShGantt.ListObjects("TabProjet").ListColumns("Level").DataBodyRange) Is Nothing Then
            Call getLevel
        End If
    End If
    GetResultsMaxMin
    StopCode
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not ribbonUI Is Nothing Then
        ribbonUI.Invalidate
    End If
End Sub


' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit
Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnGantt": ShGantt.Activate
        Case "BtnRessources": ShRessources.Activate
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

Public Function GetSep()
    GetSep = Application.International(xlListSeparator)
End Function

Public Function GetOr()
    GetOr = "Or"

    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetOr = "Ou"
End Function

Public Function GetAnd()
    GetAnd = "And"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetAnd = "Et"
End Function

Public Function GetIsoWeek()
    GetIsoWeek = "ISOWEEKNUM"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetIsoWeek = "NO.SEMAINE.ISO"
End Function

Public Function GetToday()
    GetToday = "Today"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetToday = "AUJOURDHUI"
End Function


Public Function GetYear()
    GetYear = "Year"
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036 Then GetYear = "ANNEE"
End Function

Public Sub AddRow()
    If Application.Caller = "BtnAddHolidays" Then
        Call AddRowHelper("TabOffDays")
    ElseIf Application.Caller = "BtnAddStatut" Then
        Call AddRowHelper("TabTheme")
    ElseIf Application.Caller = "BtnAddRessource" Then
        Call AddRowHelper("TabRessources")
    End If
End Sub

Public Sub DeleteRow()
    If Application.Caller = "BtnDeleteHolidays" Then
        Call DeleteRowHelper("TabOffDays")
    ElseIf Application.Caller = "BtnDeleteStatut" Then
        Call DeleteRowHelper("TabTheme")
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

'------------------------------- Added 22.09.2024 ------------------
Public Sub getLevel()
    Dim Tablevel(1 To 4) As Integer
    Dim myTab As ListObject
    Dim e As Range
    Dim i As Integer
    Dim COL_LEVEL As Integer
    Dim COL_DEADLINE As Integer
    Dim COL_SYMBOL As Integer
    Dim COL_CHARGE As Integer
    
    COL_LEVEL = GetIdColTable("Level")
    COL_DEADLINE = GetIdColTable("Deadline")
    COL_SYMBOL = GetIdColTable("Sym")
    COL_CHARGE = GetIdColTable("Charge en J")
    
    Set myTab = ShGantt.ListObjects("TabProjet")
    
    For i = LBound(Tablevel) To UBound(Tablevel)
        Tablevel(i) = 0
    Next i

    If Not myTab.DataBodyRange Is Nothing Then
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
                Cells(e.Row, COL_SYMBOL) = ChrW(&H25BA)
            ElseIf Cells(e.Row, COL_LEVEL) = "M" Then
                Cells(e.Row, COL_SYMBOL) = ChrW(&H25C6)
                Cells(e.Row, COL_CHARGE) = 0
            End If
        Next e
    End If
End Sub

Public Sub AddRowGantt(typeRow As String)
    Dim myTable As ListObject
    Dim NewRow As ListRow
    Dim COL_LEVEL As Integer
    
    COL_LEVEL = GetIdColTable("Level")
    StartCode
    Set myTable = ShGantt.ListObjects("TabProjet")
    
    With myTable
        If ActiveCell.ListObject Is Nothing Or .ListRows.count = 0 Then
            Set NewRow = .ListRows.Add
        ElseIf ActiveCell.ListObject.Name = myTable.Name Then
            Set NewRow = .ListRows.Add(ActiveCell.Row - .DataBodyRange.Row + 1)
        End If
        NewRow.Range(, COL_LEVEL) = typeRow
        Call getLevel
        Call GetResultsMaxMin
        .DataBodyRange.Rows.AutoFit
    End With
    StopCode
End Sub

Public Sub AddColumn()
    Dim myTab As ListObject
    Dim currentCol As Long
    Dim tableName As String
    
    StartCode
    
    'If Application.Caller = "BtnAddColProject" Then
        Set myTab = ActiveSheet.ListObjects(1)
    'ElseIf Application.Caller = "BtnAddColRessources" Then
        'Set myTab = ShRessources.ListObjects("TabRessources")
    'End If
    'Application.Volatile
    currentCol = ActiveCell.Column - myTab.DataBodyRange.Column + 1
    ActiveSheet.Columns(ActiveCell.Column).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    With myTab.ListColumns(currentCol).DataBodyRange
        .ClearFormats
        .Cells.Locked = False
        .Validation.Delete
    End With
    StopCode
End Sub


'-----------------------FUNCTION_ADD 23/02/2024---------------------------------------

Public Function GetIdColTable(ByVal toFind As String) As Integer
    Dim e As Range
    
    GetIdColTable = 0
    For Each e In ShGantt.ListObjects("TabProjet").HeaderRowRange
        If e = toFind Then
            GetIdColTable = e.Column '- ShGantt.ListObjects("TabProjet").Range.Column
        End If
    Next e
End Function

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

Public Sub UpdateMonths()
    Dim e As Range
    Dim count As Long
    Dim oldValue As Long
    Dim currentMonth As Long
 '   StartCode
    With ShGantt.ListObjects("TabProjet")
        count = .ListColumns("C0").Index + .HeaderRowRange.Column - 1
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

'Public Sub AddProject()
'    Dim myTable As ListObject
'    Dim NewRow As ListRow
'
'    StartCode
'    Set myTable = ShGantt.ListObjects("TabProjet")
'
'    With myTable
'        If ActiveCell.ListObject Is Nothing Or .ListRows.count = 0 Then
'            Set NewRow = .ListRows.Add
'        ElseIf ActiveCell.ListObject.Name = myTable.Name Then
'            Set NewRow = .ListRows.Add(ActiveCell.Row - .DataBodyRange.Row + 1)
'        End If
'        If Application.Caller = "BtnAddProject" Then
'            .ListColumns("Is Project").DataBodyRange.Rows(NewRow.Index) = "X"
'            .ListColumns("Progression").DataBodyRange.Rows(NewRow.Index).Formula = "=IFERROR(AVERAGEIFS([Progression],[Is Project],"""",[IDProj],[@IDProj]), 0)"
'            '.ListColumns("D�but").DataBodyRange.Rows(NewRow.Index).Formula = "=IF(MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]) = 0, """", MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]))"
'            .ListColumns("D�but").DataBodyRange.Rows(NewRow.Index).Formula = "=IFERROR(MinStartDate([@ID]), """")"
'            .ListColumns("Charge en J").DataBodyRange.Rows(NewRow.Index).Formula = "=IFERROR(SUMIFS([Charge en J],[Is Project],"""",[IDProj],[@IDProj]), """")"
'            .ListColumns("Deadline").DataBodyRange.Rows(NewRow.Index).Formula = "=IFERROR(MaxDeadlineDate([@ID]), """")"
'        Else
'            .ListColumns("Progression").DataBodyRange.Rows(NewRow.Index) = 0
'            .ListColumns("D�but").DataBodyRange.Rows(NewRow.Index).Formula = ""
'            .ListColumns("Charge en J").DataBodyRange.Rows(NewRow.Index).Formula = ""
'            .ListColumns("Deadline").DataBodyRange.Rows(NewRow.Index).Formula = "=IF([@D�but]="""","""",WORKDAY(IF(WEEKDAY([@D�but],1)=7,[@D�but]+2,IF(WEEKDAY([@D�but],1)=1,[@D�but]+1,[@D�but])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),TabOffDays[Date]))"
'        End If
'        .DataBodyRange.Rows.AutoFit
'        Call RefreshFormulaTabProjects
'        Call RefreshCondFormat
'    End With
'    StopCode
'End Sub

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
    Call getLevel
    Call GetResultsMaxMin
    If myTable.ListRows.count > 0 Then myTable.DataBodyRange.Rows.AutoFit
    'Call RefreshFormulaTabProjects
    Application.Calculation = xlCalculationAutomatic
    StopCode
End Sub

'Public Sub ConvertToProject()
'    Dim myTable As ListObject
'    Dim currentIndex As Long
'
'    StartCode
'    Set myTable = ShGantt.ListObjects("TabProjet")
'
'    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
'        MsgBox ("La s�lection n'est pas valide")
'    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
'        currentIndex = ActiveCell.Row - myTable.DataBodyRange.Row + 1
'        With myTable.ListColumns("Is Project").DataBodyRange
'            If .Rows(currentIndex) = "X" Then
'                .Rows(currentIndex) = ""
'                myTable.ListColumns("Progression").DataBodyRange.Rows(currentIndex) = 0
'                myTable.ListColumns("D�but").DataBodyRange.Rows(currentIndex).Formula = ""
'                myTable.ListColumns("Charge en J").DataBodyRange.Rows(currentIndex).Formula = ""
'                myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IF([@D�but]="""","""",WORKDAY(IF(WEEKDAY([@D�but],1)=7,[@D�but]+2,IF(WEEKDAY([@D�but],1)=1,[@D�but]+1,[@D�but])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),TabOffDays[Date]))"
'            Else
'                .Rows(currentIndex) = "X"
'                myTable.ListColumns("Progression").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(AVERAGEIFS([Progression],[Is Project],"""",[IDProj],[@IDProj]), 0)"
'                'myTable.ListColumns("D�but").DataBodyRange.Rows(currentIndex).Formula = "=IF(MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]) = 0, """", MINIFS([D�but],[Is Project],"""",[IDProj],[@IDProj]))"
'                myTable.ListColumns("D�but").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(MinStartDate([@ID]), """")"
'                myTable.ListColumns("Charge en J").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(SUMIFS([Charge en J],[Is Project],"""",[IDProj],[@IDProj]), """")"
'                myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(MaxDeadlineDate([@ID]), """")"
'            End If
'            'myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IF([@D�but]="""","""",WORKDAY(IF(WEEKDAY([@D�but],1)=7,[@D�but]+2,IF(WEEKDAY([@D�but],1)=1,[@D�but]+1,[@D�but])),roundup([@[Charge en J]],0)-1,TabOffDays[Date]))"
'        End With
'        myTable.DataBodyRange.Rows.AutoFit
'        Call RefreshFormulaTabProjects
'    End If
'    StopCode
'End Sub
'


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

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not ribbonUI Is Nothing Then
        ribbonUI.Invalidate
    End If
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
        Call RefreshCondFormSymbol
        Call RefreshCondFormIsMilestone
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

Sub RefreshCondFormSymbol()
    Dim TableList As ListObject
    Dim e As Range
    Dim COL_LEVEL As String
    Dim COL_C0 As Integer
    
    COL_LEVEL = "$" & GetAdrressColTable("Level") & 14
    COL_C0 = GetIdColTable("C0") - 3
    
    Set TableList = ShBackend.ListObjects("TabSymbol")

    For Each e In TableList.ListColumns("Symbol").DataBodyRange
        With ShGantt.ListObjects("TabProjet").ListColumns("Sym").DataBodyRange
            If IsNumeric(e.value) Then
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & COL_LEVEL & "=" & e.value
            Else
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & COL_LEVEL & "=""" & e.value & """"
            End If
                .FormatConditions(.FormatConditions.count).SetFirstPriority
                With .FormatConditions(1).Font
                    On Error Resume Next
                    .Color = e.Font.Color
                    If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
                    .Bold = e.Font.Bold
                    .Italic = e.Font.Italic
                    .TintAndShade = e.Font.TintAndShade
                End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Application.CutCopyMode = True
    Next e
End Sub

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
                myFormula = "=" & GetOr & "(" & GetAnd & "(" & C_C0 & "$8>=$" & C_DEBUT & "14" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep & " " & C_C0 & "$8<=$" & C_DA_PROG & "14" & GetSep & " $" & C_PROG & "14 <> 0" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -1) & """)" & GetSep & " " & GetAnd & "(" & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8)" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) = " & GetIsoWeek & "($" & C_DEBUT & "14)" & GetSep & " $" & C_PROG & "14 <> 0" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -1) & """))"
            Else
                myFormula = "=" & GetOr & "(" & GetAnd & "(" & C_C0 & "$8>=$" & C_DEBUT & "14" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep & " " & C_C0 & "$8<=$" & C_DEADLINE & "14" & GetSep & " $" & C_DEADLINE & "14 <> """"" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -2) & """)" & GetSep & " " & GetAnd & "(" & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8)" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) = " & GetIsoWeek & "($" & C_DEBUT & "14)" & GetSep & " $" & C_STATUT & "14 =""" & e.Offset(, -2) & """))"
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
Sub RefreshCondFormIsMilestone()
    Dim e As Range
    Dim COL_LEVEL As String
    Dim COL_C0 As Integer
    
    COL_LEVEL = "$" & GetAdrressColTable("Level") & 14
    COL_C0 = GetIdColTable("C0") - 3
    
    Set e = ShBackend.Range("D2")
    With ShGantt.ListObjects("TabProjet").ListColumns("Sym").DataBodyRange.Resize(, COL_C0)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & COL_LEVEL & "=""M"""
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
Sub RefreshCondFormIsProject()
    Dim e As Range
    Dim COL_LEVEL As String
    Dim COL_C0 As Integer
    
    COL_LEVEL = "$" & GetAdrressColTable("Level") & 14
    COL_C0 = GetIdColTable("C0") - 3
    
    Set e = ShSettings.Range("IsProject")
    With ShGantt.ListObjects("TabProjet").ListColumns("ID").DataBodyRange.Resize(, COL_C0)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & COL_LEVEL & "=1"
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



' --- Module: ShRessources.cls ---
Attribute VB_Name = "ShRessources"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Deactivate()
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not ribbonUI Is Nothing Then
        ribbonUI.Invalidate
    End If
End Sub


' --- Module: Ressources.bas ---
Attribute VB_Name = "Ressources"
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
    Dim Sh As Shape
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
        For Each e In Union(.Range("LBL_STATUS"), .Range("LBL_HOLIDAY_LIST"), .Range("SETT_HEADER"), .Range("SETT_SUBHEADER"), .Range("LBL_TAB_STATUS"), .Range("M4"), .Range("LBL_TAB_HOLIDAYS"))
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

    'ShRessources SHEET-----------------------------------------------
    With ShRessources
        trad = vbNullString
        For Each e In Union(.Range("K1:K3"), .Range("C4"), .Range("H5"), .Range("K5"), .Range("B7:M7"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabRessources").ListColumns("R�le").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    'ShGantt SHEET-----------------------------------------------
    With ShGantt
        trad = vbNullString
        For Each e In Union(.Range("E1:X3"), .Range("B11:O11"), .Range("B12"), .Range("H12"), .Range("G4"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabProjet").ListColumns("WBS - Activit� / t�che").DataBodyRange.Resize(, 2)
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
'        trad = dict(.OptionButtons("Option Button 3").Caption)
'        If trad <> vbNullString Then
'            .OptionButtons("Option Button 3").Caption = trad
'        End If
'        trad = dict(.OptionButtons("Option Button 4").Caption)
'        If trad <> vbNullString Then
'            .OptionButtons("Option Button 4").Caption = trad
'        End If
'        trad = dict(.OptionButtons("Option Button 8").Caption)
'        If trad <> vbNullString Then
'            .OptionButtons("Option Button 8").Caption = trad
'        End If
    End With
'
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
        For Each Sh In sht.Shapes
            If Sh.Type = 1 Then
                originalText = Sh.TextFrame2.TextRange.Text
                trad = dict(originalText)
                If trad <> vbNullString Then
                    Sh.TextFrame2.TextRange.Text = trad
                End If
            End If
        Next Sh
        trad = dict(sht.Name)
        If trad <> vbNullString Then
            sht.Name = trad
        End If
    Next sht
    endTime = Timer
    Call RefreshCondFormat
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



' --- Module: MinMax.bas ---
Attribute VB_Name = "MinMax"
Option Explicit

Sub GetResultsMaxMin()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long, j As Long
    Dim minDate As Date, maxDate As Date
    Dim totalCharge As Double, totalProgression As Double
    Dim countProgression As Long
    Dim startRow As Long
    Dim lastRow As Long
    Dim currentLevel As Long

    Set ws = ShGantt
    Set tbl = ws.ListObjects("TabProjet")

    lastRow = tbl.ListRows.count

    For i = 1 To lastRow
        If IsNumeric(tbl.DataBodyRange(i, tbl.ListColumns("Level").Index).value) Then
            currentLevel = tbl.DataBodyRange(i, tbl.ListColumns("Level").Index).value
            If currentLevel = 1 Then
                minDate = DateValue("31/12/9999")
                maxDate = DateValue("01/01/1900")
                totalCharge = 0
                totalProgression = 0
                countProgression = 0
                startRow = i
                
                For j = i + 1 To lastRow
                    If tbl.DataBodyRange(j, tbl.ListColumns("Level").Index).value = 1 Then
                        Exit For
                    ElseIf IsDate(tbl.DataBodyRange(j, tbl.ListColumns("D�but").Index).value) Then
                        If tbl.DataBodyRange(j, tbl.ListColumns("D�but").Index).value < minDate Then
                            minDate = tbl.DataBodyRange(j, tbl.ListColumns("D�but").Index).value
                        End If
                    End If
                    If tbl.DataBodyRange(j, tbl.ListColumns("D�but").Index).value <> "" Then
                        tbl.DataBodyRange(j, tbl.ListColumns("Deadline").Index).Formula = _
                        "=IF([@D�but]="""","""",WORKDAY(IF(OR(WEEKDAY([@D�but],1)=7,WEEKDAY([@D�but],1)=1,COUNTIF(TabOffDays[Date],[@D�but])>0),WORKDAY([@D�but],1,TabOffDays[Date]),[@D�but]),ROUNDUP([@[Charge en J]],0)-IF([@[Charge en J]]>0,1,0),TabOffDays[Date]))"
                    End If
                    If IsDate(tbl.DataBodyRange(j, tbl.ListColumns("Deadline").Index).value) Then
                        If tbl.DataBodyRange(j, tbl.ListColumns("Deadline").Index).value > maxDate Then
                            maxDate = tbl.DataBodyRange(j, tbl.ListColumns("Deadline").Index).value
                        End If
                    End If
                    If tbl.DataBodyRange(j, tbl.ListColumns("Level").Index).value = 2 Then
                        If IsNumeric(tbl.DataBodyRange(j, tbl.ListColumns("Charge en J").Index).value) Then
                            totalCharge = totalCharge + tbl.DataBodyRange(j, tbl.ListColumns("Charge en J").Index).value
                        End If
                    End If
                
                    If IsNumeric(tbl.DataBodyRange(j, tbl.ListColumns("Progression").Index).value) Then
                        totalProgression = totalProgression + tbl.DataBodyRange(j, tbl.ListColumns("Progression").Index).value
                        countProgression = countProgression + 1
                    End If
    

                Next j
                
                tbl.DataBodyRange(i, tbl.ListColumns("D�but").Index).value = vbNullString
                tbl.DataBodyRange(i, tbl.ListColumns("Deadline").Index).value = vbNullString
                
                ' Fill the minimum D�but date in the "D�but" cell of the Level 1 row
                If minDate <> DateValue("31/12/9999") Then
                    tbl.DataBodyRange(i, tbl.ListColumns("D�but").Index).value = minDate
                End If
                ' Fill the maximum Deadline date in the "Deadline" cell of the Level 1 row
                If maxDate <> DateValue("01/01/1900") Then
                    tbl.DataBodyRange(i, tbl.ListColumns("Deadline").Index).value = maxDate
                End If
                ' Fill the total charge in the "Charge en J" column of the Level 1 row (for Level 2 rows only)
                tbl.DataBodyRange(i, tbl.ListColumns("Charge en J").Index).value = totalCharge
                ' Calculate and fill the average progression in the "Progression" column of the Level 1 row
                If countProgression > 0 Then
                    tbl.DataBodyRange(i, tbl.ListColumns("Progression").Index).value = totalProgression / countProgression
                End If
            End If
        ElseIf tbl.DataBodyRange(i, tbl.ListColumns("Level").Index).value = "M" Then
            If tbl.DataBodyRange(i, tbl.ListColumns("Charge en J").Index).value <> 0 Then
                MsgBox "Milestone est une t�che d'une dur�e �gale � 0."
                tbl.DataBodyRange(i, tbl.ListColumns("Charge en J").Index).value = 0
            End If
            If tbl.DataBodyRange(i, tbl.ListColumns("D�but").Index).value <> "" Then
                tbl.DataBodyRange(i, tbl.ListColumns("Deadline").Index).Formula = _
                "=IF([@D�but]="""","""",WORKDAY(IF(OR(WEEKDAY([@D�but],1)=7,WEEKDAY([@D�but],1)=1,COUNTIF(TabOffDays[Date],[@D�but])>0),WORKDAY([@D�but],1,TabOffDays[Date]),[@D�but]),ROUNDUP([@[Charge en J]],0)-IF([@[Charge en J]]>0,1,0),TabOffDays[Date]))"
            End If
        End If
        
    Next i
End Sub





' --- Module: Import.bas ---
Attribute VB_Name = "Import"
Option Explicit

Public Sub ImportData()
    Dim sourceWb As Workbook
    Dim destWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim sourceTable As ListObject
    Dim destTable As ListObject
    Dim filePath As String
    Dim fd As FileDialog
    Dim e As Range
    
    ' Open the file dialog to select the source workbook
    StartCode

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "S�lectionnez un fichier WBS"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsm"

    ' Afficher la bo�te de dialogue
    If fd.Show = 0 Then
        ' L'utilisateur a annul�, on quitte la proc�dure
        StopCode
        Exit Sub
    End If

    ' R�cup�rer le chemin du fichier s�lectionn�
    filePath = fd.SelectedItems(1)
    
    Set sourceWb = Workbooks.Open(filePath)
    
    Set sourceWs = GetSheetByCodeName(sourceWb, "ShWBS") ' Adjust to the correct sheet name
    
    ' Set references to the tables
    Set sourceTable = sourceWs.ListObjects("TabWBS")
    Set destTable = ShGantt.ListObjects("TabProjet")
    
    With destTable
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
            .ListRows.Add
            .ListColumns("Progression").DataBodyRange.NumberFormat = "0%"
        End If
        If .DataBodyRange Is Nothing Then
            .ListRows.Add
        End If
    End With
    
    ' Copy data from the source columns to the destination columns
    sourceTable.ListColumns("Activit�s").DataBodyRange.Copy
    destTable.ListColumns("WBS - Activit� / t�che").DataBodyRange.PasteSpecial xlPasteValues
    
    'destTable.ListColumns("WBS - Activit� / t�che").DataBodyRange = sourceTable.ListColumns("Activit�s").DataBodyRange
    sourceTable.ListColumns("Charge de travail").DataBodyRange.Copy
    destTable.ListColumns("Charge en J").DataBodyRange.PasteSpecial xlPasteValues
    
    
    sourceTable.ListColumns("Niveau").DataBodyRange.Copy
    destTable.ListColumns("Level").DataBodyRange.PasteSpecial xlPasteValues
    
    
    ' Close the source workbook without saving changes
    sourceWb.Close SaveChanges:=False
    ' Clear clipboard
    Application.CutCopyMode = False

    Call getLevel
    Call GetResultsMaxMin
    Call RefreshCondFormat
    'Call RefreshAllFormula
    StopCode
    MsgBox "Done"
End Sub

Function GetSheetByCodeName(wb As Workbook, codeName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.codeName = codeName Then
            Set GetSheetByCodeName = ws
            Exit Function
        End If
    Next ws
    Set GetSheetByCodeName = Nothing
End Function


' --- Module: MRibbon.bas ---
Attribute VB_Name = "MRibbon"
Option Explicit

Public ribbonUI As IRibbonUI

Function isEnabled(control As IRibbonControl, ByRef returnedVal) As Boolean
    Select Case control.Id
        Case "BtnAddProject", "TasksMenu", "BtnAddMilestone"
            returnedVal = ActiveSheet.Name = ShGantt.Name And (Not ActiveCell.ListObject Is Nothing)
        Case "DateMenu", "BtnToday", "BtnImport"
            returnedVal = (ActiveSheet.Name = ShGantt.Name)
        Case "ColumnsMenu"
            returnedVal = (ActiveSheet.Name = ShGantt.Name Or ActiveSheet.Name = ShRessources.Name)
        Case "RowsMenu"
            returnedVal = Not ActiveCell.ListObject Is Nothing
        Case "BtnAddRow"
            returnedVal = (Not ActiveCell.ListObject Is Nothing) And (ActiveSheet.Name = ShRessources.Name Or ActiveSheet.Name = ShSettings.Name)
    End Select
End Function

Sub Ribbon_Load(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
    ribbon.Invalidate
End Sub

Public Sub myRibbon(control As IRibbonControl)
    Select Case control.Id
        Case "BtnAddProject"
            If ActiveSheet.Name = ShGantt.Name Then
                Call AddRowGantt("1")
            End If
        Case "BtnAddTask2"
            If ActiveSheet.Name = ShGantt.Name Then
                Call AddRowGantt("2")
            End If
        Case "BtnAddTask3"
            If ActiveSheet.Name = ShGantt.Name Then
                Call AddRowGantt("3")
            End If
        Case "BtnAddTask4"
            If ActiveSheet.Name = ShGantt.Name Then
                Call AddRowGantt("4")
            End If
        Case "BtnAddMilestone"
            If ActiveSheet.Name = ShGantt.Name Then
                Call AddRowGantt("M")
            End If
             
        Case "BtnAddRow"
'            If ActiveSheet.Name = ShGantt.Name Then
'                Call AddRowGantt
            If ActiveSheet.Name <> ShGantt.Name And Not ActiveCell.ListObject Is Nothing Then
                Call AddRowHelper(ActiveCell.ListObject.Name)
            End If
        Case "BtnDeleteRow"
            If ActiveSheet.Name = ShGantt.Name Then
                Call DeleteProject
            ElseIf Not ActiveCell.ListObject Is Nothing Then
                Call DeleteRowHelper(ActiveCell.ListObject.Name)
            End If
        Case "BtnAddColumn"
            If ActiveSheet.Name = ShGantt.Name Or ActiveSheet.Name = ShRessources.Name Then
                Call AddColumn
            End If
        Case "BtnDeleteColumn"
            If ActiveSheet.Name = ShGantt.Name Or ActiveSheet.Name = ShRessources.Name Then
                Call DeleteColumn
            End If
        Case "BtnDay"
            ShBackend.Range("c1") = 1
            Call SwitchDaysWeeks
        Case "BtnWeek"
            ShBackend.Range("c1") = 2
            Call SwitchDaysWeeks
        Case "BtnMonth"
            ShBackend.Range("c1") = 3
            Call SwitchDaysWeeks
        Case "BtnToday"
            Call GoToToday
        Case "BtnImport"
            Call ImportData
        Case "BtnFullscreen"
            Call maximize
            
        Case "BtnSettings"
            ShSettings.Activate
        Case "BtnFrench"
            ShSettings.Range("Langue") = 2
            Call change_langue
        Case "BtnEnglish"
            ShSettings.Range("Langue") = 1
            Call change_langue
        Case "BtnHelp"
            ThisWorkbook.FollowHyperlink Address:="https://blog-gestion-de-projet.com/docs/kit-management-projet/planification/planning-gantt/"
            
    End Select
End Sub

Function CheckIfInTable(control As IRibbonControl) As Boolean
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1) ' V�rifie la premi�re table sur la feuille active

    ' Si la cellule active fait partie d'une table, retourne True
    If Not tbl Is Nothing Then
        CheckIfInTable = Not Intersect(ActiveCell, tbl.Range) Is Nothing
    Else
        CheckIfInTable = False
    End If
End Function



