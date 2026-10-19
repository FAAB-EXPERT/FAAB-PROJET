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


' --- Module: ShProjetClassique.cls ---
Attribute VB_Name = "ShProjetClassique"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    StartCode
    If Target.count = 1 Then
        If Not Intersect(Target, Me.Range("ProjectStart")) Is Nothing Then
            Call UpdateMonths
        End If
    End If
    StopCode
End Sub


' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit
Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnClassique": ShProjetClassique.Activate
       ' Case "BtnRessources": ShRessources.Activate
       ' Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Public Sub StartCode()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ShProjetClassique.Unprotect
End Sub

Public Sub StopCode()
    ShProjetClassique.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
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
'
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

Public Sub DeleteRow() 'OK
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
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
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub AddRow() 'OK
    Dim myTable As ListObject
    Dim MyCell As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    Set myTable = MyCell.ListObject
    
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

Public Sub maximize() 'OK
    ShBackend.Range("IsMaximize") = Not ShBackend.Range("IsMaximize")
    Application.DisplayFullScreen = ShBackend.Range("IsMaximize")
    ActiveWindow.DisplayHeadings = Not ShBackend.Range("IsMaximize")
End Sub

Public Sub GoToToday() 'OK
    StartCode
    ShProjetClassique.Range("ProjectStart") = Date
    UpdateMonths
    StopCode
End Sub

Public Sub DeleteColumn()
    Dim MsgAnswer As VbMsgBoxResult
    
    StartCode
    If Selection.count = 1 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de la colonne sélectionnée?", vbOKCancel)
        If MsgAnswer = vbOK Then ActiveSheet.Columns(ActiveCell.Column).Delete
    Else
        MsgBox ("Veuillez sélectionner une seule colonne à supprimer")
    End If
    StopCode
End Sub

Public Sub AddColumn()
    Dim myTab As ListObject
    Dim currentCol As Long
    Dim tableName As String
    
    StartCode
    
    Set myTab = ShProjetClassique.ListObjects("TabProjet")
    
    currentCol = ActiveCell.Column - myTab.DataBodyRange.Column + 1
    ActiveSheet.Columns(ActiveCell.Column).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    With myTab.ListColumns(currentCol).DataBodyRange
        .ClearFormats
        .Cells.Locked = False
        .Validation.Delete
    End With
    StopCode
End Sub

Public Function GetIdColTable(ByVal toFind As String) As Integer
    Dim e As Range
    
    GetIdColTable = 0
    For Each e In ShProjetClassique.ListObjects("TabProjet").HeaderRowRange
        If e = toFind Then
            GetIdColTable = e.Column '- ShGantt.ListObjects("TabProjet").Range.Column
        End If
    Next e
End Function


Public Function MinStartDate(Id As Long) As Long 'OK_ Update 23022024
    Dim myTable As ListObject
    Dim e As Range
    Dim COL_DEBUT As Integer
    Dim COL_IS_PROJECT As Integer
    Dim COL_IS_JALON As Integer
    
    COL_IS_PROJECT = GetIdColTable("Is Project")
    COL_DEBUT = GetIdColTable("Début")
    COL_IS_JALON = GetIdColTable("IsJalon")

    MinStartDate = 0
    Application.Volatile
    Set myTable = ShProjetClassique.ListObjects("TabProjet")
    With myTable
        For Each e In myTable.ListColumns("IDProj").DataBodyRange
            If e = Id And Cells(e.Row, COL_IS_PROJECT) <> "X" And Cells(e.Row, COL_IS_JALON) <> "X" Then
                MinStartDate = WorksheetFunction.Min(Cells(e.Row, COL_DEBUT), IIf(MinStartDate = 0, Cells(e.Row, COL_DEBUT), MinStartDate))
            End If
        Next e
    End With
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


Public Function GetDeadline(Id As Long) As Long
    Dim myTable As ListObject
    Dim Target As Range
    Dim e As Range
    Dim COL_DEADLINE As Integer
    Dim COL_IS_PROJECT As Integer
    Dim COL_IS_JALON As Integer
    Dim COL_DEBUT As Integer
    Dim COL_CHARGES As Integer
    
    COL_IS_PROJECT = GetIdColTable("Is Project") '-7 cells(Target.row, COL_IS_PROJECT)
    COL_DEADLINE = GetIdColTable("Deadline")
    COL_IS_JALON = GetIdColTable("IsJalon")
    COL_DEBUT = GetIdColTable("Début")
    COL_CHARGES = GetIdColTable("Charge en J")
    
    GetDeadline = 0
    Application.Volatile
    Set Target = Application.Caller
    Set myTable = ShProjetClassique.ListObjects("TabProjet")
    With myTable
        If Cells(Target.Row, COL_IS_PROJECT) = "X" Then
            For Each e In myTable.ListColumns("IDProj").DataBodyRange
                If e = Id And Cells(e.Row, COL_IS_PROJECT) <> "X" And Cells(e.Row, COL_IS_JALON) <> "X" Then
                    GetDeadline = WorksheetFunction.Max(Cells(e.Row, COL_DEADLINE), IIf(GetDeadline = 0, Cells(e.Row, COL_DEADLINE), GetDeadline))
                End If
                If e.value > Id Then Exit Function
            Next e
        ElseIf Cells(Target.Row, COL_DEBUT) <> vbNullString Then
            If WorksheetFunction.WorkDay(Cells(Target.Row, COL_DEBUT), 1) = 7 Then
                GetDeadline = Cells(Target.Row, COL_DEBUT) + 2
            ElseIf WorksheetFunction.Weekday(Cells(Target.Row, COL_DEBUT), 1) = 1 Then
                GetDeadline = Cells(Target.Row, COL_DEBUT) + 1
            Else
                GetDeadline = Cells(Target.Row, COL_DEBUT)
            End If
            GetDeadline = WorksheetFunction.WorkDay(GetDeadline, WorksheetFunction.RoundUp(Cells(Target.Row, COL_CHARGES), 0) - IIf(Cells(Target.Row, COL_CHARGES) > 0, 1, 0))
        End If
    End With
End Function

Public Sub UpdateStartingDate() 'OK
    Dim isNextOrPrev As Integer
    Dim intervalType As String
    
    StartCode
    isNextOrPrev = IIf(Left(Application.Caller, 7) = "BtnPrev", -1, 1)
    
    With ShBackend
        If .Range("IsWeeksGant") And Right(Application.Caller, 5) = "Speed" Then
            intervalType = "m"
        ElseIf .Range("IsWeeksGant") Then
            intervalType = "ww"
        End If
    End With
    ShProjetClassique.Range("ProjectStart") = DateAdd(intervalType, isNextOrPrev, ShProjetClassique.Range("ProjectStart"))
    Call UpdateMonths
    StopCode
End Sub

Public Function GetAdrressColTable(ByVal toFind As String) As String
    Dim e As Range
    
    GetAdrressColTable = 0
    For Each e In ShProjetClassique.ListObjects("TabProjet").HeaderRowRange
        If e = toFind Then
            GetAdrressColTable = Mid(e.Address, 2, 1)
        End If
    Next e
    
End Function

Private Sub RefreshFormulaTabProjects() 'OK
    Dim myTable As ListObject
    Dim currentIndex As Long
    Dim e As Range
    Dim C_IDPROJ As String
    Dim C_IDSUBPROJ As String
    Dim COL_DEADLINE As Integer
    
    C_IDPROJ = GetAdrressColTable("IDProj") 'D " & C_IDPROJ & "
    C_IDSUBPROJ = GetAdrressColTable("IDSubProj") 'E " & C_IDSUBPROJ & "
    COL_DEADLINE = GetIdColTable("Deadline")
    
    Set myTable = ShProjetClassique.ListObjects("TabProjet")
    If myTable.ListRows.count > 0 Then
        myTable.ListColumns("IDProj").DataBodyRange.Formula = "=IF(ISNUMBER($" & C_IDPROJ & "13),IF([@[Is Project]] <>"""",$" & C_IDPROJ & "13+1,$" & C_IDPROJ & "13),1)"
        'myTable.ListColumns("IsJalon").DataBodyRange.Formula = "=IF(AND([@IDSubProj] > $E14), ""X"", """")"
        myTable.ListColumns("IDSubProj").DataBodyRange.Formula = "=IFERROR(IF(AND([@[Is Project]]= """", [@IsJalon]= """"),$" & C_IDSUBPROJ & "13+1, IF([@IsJalon]= ""X"",$" & C_IDSUBPROJ & "13, 0)),1)"
        For Each e In myTable.ListColumns("Is Project").DataBodyRange
            If e = "X" Then
                'e.Offset(, 7).Formula = "=IF(MAXIFS([Deadline],[Is Project],"""",[IDProj],[@IDProj],[IsJalon],"""") = 0,"""",MAXIFS([Deadline],[Is Project],"""",[IDProj],[@IDProj],[IsJalon],""""))"
                Cells(e.Row, COL_DEADLINE).Formula = "=IF(GetDeadline([@IDProj]) = 0, """", GetDeadline([@IDProj]))"
            Else
                Cells(e.Row, COL_DEADLINE).Formula = "=IF([@Début]="""","""",WORKDAY(IF(WEEKDAY([@Début],1)=7, [@Début]+2,IF(WEEKDAY([@Début],1)=1,[@Début]+1,[@Début])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),))"
            End If
        Next e
    End If
End Sub

Public Sub UpdateMonths() 'OK
    Dim e As Range
    Dim count As Long
    Dim oldValue As Long
    Dim currentMonth As Long
    
    With ShProjetClassique.ListObjects("TabProjet")
        count = .ListColumns("C0").Index + .HeaderRowRange.Column - 1
    End With
    oldValue = 0
    Application.Calculation = xlCalculationManual
    ShProjetClassique.Range("GanttHeaderMonthsWeeks").ClearContents
    ShProjetClassique.Range("GanttHeaderYears").ClearContents
    With ShProjetClassique
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
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub AddProject() 'OK
    Dim myTable As ListObject
    Dim NewRow As ListRow
    
    StartCode
    Set myTable = ShProjetClassique.ListObjects("TabProjet")
    
    With myTable
        If ActiveCell.ListObject Is Nothing Or .ListRows.count = 0 Then
            Set NewRow = .ListRows.Add
        ElseIf ActiveCell.ListObject.Name = myTable.Name Then
            Set NewRow = .ListRows.Add(ActiveCell.Row - .DataBodyRange.Row + 1)
        End If
        If Application.Caller = "BtnAddProject" Then
            .ListColumns("Is Project").DataBodyRange.Rows(NewRow.Index) = "X"
            .ListColumns("Début").DataBodyRange.Rows(NewRow.Index).Formula = "=IF(MinStartDate([@IDProj]) = 0, """", MinStartDate([@IDProj]))"
            .ListColumns("Charge en J").DataBodyRange.Rows(NewRow.Index).Formula = "=IFERROR(SUMIFS([Charge en J],[Is Project],"""",[IDProj],[@IDProj]), """")"
        Else
            .ListColumns("Début").DataBodyRange.Rows(NewRow.Index).Formula = ""
            .ListColumns("Charge en J").DataBodyRange.Rows(NewRow.Index).Formula = ""
            '.ListColumns("Deadline").DataBodyRange.Rows(NewRow.Index).Formula = "=IF(OR([@IsJalon] = ""X"",[@Début]=""""),"""",WORKDAY(IF(WEEKDAY([@Début],1)=7,[@Début]+2,IF(WEEKDAY([@Début],1)=1,[@Début]+1,[@Début])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),))"
        End If
        If Application.Caller = "BtnAddJalon" Then .ListColumns("IsJalon").DataBodyRange.Rows(NewRow.Index) = "X"
        .ListColumns("Deadline").DataBodyRange.Rows(NewRow.Index).Formula = "=IF(GetDeadline([@IDProj]) = 0, """", GetDeadline([@IDProj]))"
        .DataBodyRange.Rows.AutoFit
        Call RefreshFormulaTabProjects
        Call RefreshCondFormat
    End With
    StopCode
End Sub

Public Sub DeleteProject() 'OK
    Dim myTable As ListObject
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    
    StartCode
    Application.Calculation = xlCalculationManual
    Set myTable = ShProjetClassique.ListObjects("TabProjet")
    
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

Public Sub ConvertToProject() 'OK
    Dim myTable As ListObject
    Dim currentIndex As Long
    
    StartCode
    Set myTable = ShProjetClassique.ListObjects("TabProjet")
    
    Application.Calculation = xlCalculationManual
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
        MsgBox ("La sélection n'est pas valide")
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        currentIndex = ActiveCell.Row - myTable.DataBodyRange.Row + 1
        With myTable.ListColumns("Is Project").DataBodyRange
            If .Rows(currentIndex) = "X" Then
                .Rows(currentIndex) = ""
                myTable.ListColumns("Début").DataBodyRange.Rows(currentIndex).Formula = ""
                myTable.ListColumns("Charge en J").DataBodyRange.Rows(currentIndex).Formula = ""
                myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IF([@Début]="""","""",WORKDAY(IF(WEEKDAY([@Début],1)=7,[@Début]+2,IF(WEEKDAY([@Début],1)=1,[@Début]+1,[@Début])),roundup([@[Charge en J]],0)-IF([@[Charge en J]] > 0, 1, 0),))"
            Else
                .Rows(currentIndex) = "X"
                myTable.ListColumns("Début").DataBodyRange.Rows(currentIndex).Formula = "=IF(MinStartDate([@IDProj]) = 0, """", MinStartDate([@IDProj]))"
                myTable.ListColumns("Charge en J").DataBodyRange.Rows(currentIndex).Formula = "=IFERROR(SUMIFS([Charge en J],[Is Project],"""",[IDProj],[@IDProj]), """")"
                myTable.ListColumns("Deadline").DataBodyRange.Rows(currentIndex).Formula = "=IF(GetDeadline([@IDProj]) = 0, """", GetDeadline([@IDProj]))"
            End If
        End With
        myTable.DataBodyRange.Rows.AutoFit
        Call RefreshFormulaTabProjects
    End If
    StopCode
    Application.Calculation = xlCalculationAutomatic
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
Public Sub RefreshCondFormat() 'OK
    If ShProjetClassique.ListObjects("TabProjet").ListRows.count > 0 Then
        ShProjetClassique.Cells.FormatConditions.Delete
        Call RefreshCondFormProjectJalon
        Call RefreshCondFormListGantt
        Call RefreshCondFormToday
        Call UpdateMonths
        Call RefreshCondFormatJalon
    End If
End Sub

Sub RefreshCondFormatJalon()
    With ShProjetClassique.ListObjects("TabProjet").ListColumns("C0").DataBodyRange.Resize(, 56)
            .FormatConditions.AddIconSetCondition
            .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1)
            .ReverseOrder = False
            .ShowIconOnly = True
            .IconSet = ActiveWorkbook.IconSets(xl3Flags)
        End With
        With .FormatConditions(1).IconCriteria(2)
            .Type = xlConditionValueNumber
            .value = 0
            .Operator = 7
            .Icon = xlIconNoCellIcon
        End With
        With .FormatConditions(1).IconCriteria(3)
            .Type = xlConditionValueNumber
            .value = 1
            .Operator = 7
        End With
    End With
End Sub

Sub RefreshCondFormListGantt() 'OK
    Dim TableList As ListObject
    Dim e As Range
    Dim myFormula As String
    
    Dim C_C0 As String
    Dim C_DEBUT As String
    Dim C_IS_PROJECT As String
    Dim C_DEADLINE As String
    Dim C_IS_JALON As String
  
    C_C0 = GetAdrressColTable("C0") 'K " & C_C0 & "
    C_DEBUT = GetAdrressColTable("Début") 'H " & C_DEBUT & "
    C_IS_PROJECT = GetAdrressColTable("Is Project") 'C " & C_IS_PROJECT & "
    C_DEADLINE = GetAdrressColTable("Deadline") 'J " & C_DEADLINE & "
    C_IS_JALON = GetAdrressColTable("IsJalon") 'F " & C_IS_JALON & "
    

    For Each e In ShSettings.ListObjects("TabColor").ListColumns("Color").DataBodyRange
        'FROM V0.8
        If e = "Phase" Then
            myFormula = "=" & GetOr & "(" & GetAnd & "(" & C_C0 & "$8>=$" & C_DEBUT & "14" & GetSep _
            & "$" & C_IS_PROJECT & "14 = ""X""" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep _
            & " " & C_C0 & "$8<=$" & C_DEADLINE & "14)" & GetSep & " " & GetAnd & "(" & GetYear _
            & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8)" & GetSep _
            & "$" & C_IS_PROJECT & "14 = ""X""" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) = " & GetIsoWeek & "($" & C_DEBUT & "14)))"
        Else
            myFormula = "=" & GetOr & "(" & GetAnd & "(" & C_C0 & "$8>=$" & C_DEBUT & "14" & GetSep _
            & "$" & C_IS_JALON & "14 <> ""X""" & GetSep & " $" & C_DEBUT & "14 <> """"" & GetSep _
            & " " & C_C0 & "$8<=$" & C_DEADLINE & "14)" & GetSep & " " & GetAnd & "(" & "$" & C_IS_JALON & "14 <> ""X""" & GetSep _
            & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8)" & GetSep _
            & GetIsoWeek & "(" & C_C0 & "$8) = " & GetIsoWeek & "($" & C_DEBUT & "14)))"
        End If
        
        'CURRENT VERSION 1.0
        If e = "test1" Then
            myFormula = "=" & GetAnd & "($" & C_DEBUT & "14 <> """"" & GetSep & "$" & C_IS_PROJECT & "14 = ""X""" & GetSep _
            & GetIsoWeek & "(" & C_C0 & "$8) >= " & GetIsoWeek & "($" & C_DEBUT & "14)" & GetSep _
            & GetIsoWeek & "(" & C_C0 & "$8) <= " & GetIsoWeek & "($" & C_DEADLINE & "14)" & GetSep _
            & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8))"
        ElseIf e = "test2" Then
            myFormula = "=" & GetAnd & "($" & C_DEBUT & "14 <> """"" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) >=" & GetIsoWeek _
            & "($" & C_DEBUT & "14)" & GetSep & GetIsoWeek & "(" & C_C0 & "$8) <=" & GetIsoWeek & "($" & C_DEADLINE & "14)" & GetSep _
            & GetYear & "($" & C_DEBUT & "14) = " & GetYear & "(" & C_C0 & "$8))"
        End If
        With ShProjetClassique.ListObjects("TabProjet").ListColumns("C0").DataBodyRange.Resize(, 56)
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
            .FormatConditions(1).StopIfTrue = False
        End With
        Application.CutCopyMode = False
    Next e
End Sub

Sub RefreshCondFormProjectJalon() 'OK
    Dim e As Range
    Dim myFormula As String
    Dim COL_IS_PROJECT As String
    Dim COL_IS_JALON As String
    
    COL_IS_PROJECT = "$" & GetAdrressColTable("Is Project") & 14 'C
    COL_IS_JALON = "$" & GetAdrressColTable("IsJalon") & 14 'F
    
    Set e = ShSettings.Range("IsProject")
    For Each e In ShSettings.ListObjects("TabProjJalon").ListColumns("Color").DataBodyRange
        If e = "Project" Then
            myFormula = "=" & COL_IS_PROJECT & "<>"""""
        ElseIf e = "Jalon" Then
            myFormula = "=" & COL_IS_JALON & "<>"""""
        End If
        With ShProjetClassique.ListObjects("TabProjet").ListColumns("ID").DataBodyRange.Resize(, 9)
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

Sub RefreshCondFormToday() 'OK
    Dim e As Range
    Dim myFormula As String
    Dim COL_C0 As String
    Dim COL_C1 As String
    
    COL_C0 = GetAdrressColTable("C0") 'K
    COL_C1 = GetAdrressColTable("C1") 'L

    myFormula = "=" & GetAnd & "(" & GetIsoWeek & "(" & GetToday & "())>=" & COL_C0 & "$9" & GetSep & "" & GetIsoWeek & "(" & GetToday & "())<" & COL_C1 & "$9)"

    Set e = ShSettings.Range("TodayFormat")
    With ShProjetClassique.ListObjects("TabProjet").ListColumns("C0").DataBodyRange.Resize(, 56)
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


    'ShGantt SHEET-----------------------------------------------
    With ShProjetClassique
        trad = vbNullString
        For Each e In Union(.Range("H1:H2"), .Range("N1:N2"), .Range("AK1:AK3"), .Range("G4"), .Range("H4"), .Range("B11:J11"), .Range("G12"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabProjet").ListColumns("WBS - Activité / tâche").DataBodyRange
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
    
    CELL_KEY = ShProjetClassique.Range("Langue")
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








