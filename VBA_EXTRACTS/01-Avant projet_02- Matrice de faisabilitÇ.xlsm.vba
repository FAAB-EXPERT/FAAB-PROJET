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

'------------------------------------------------------------
'Version 1.4 : 19 / 3 / 2023
'By Samir Ouarrak
'------------------------------------------------------------


Private Sub Workbook_Open()
    Call UpdateChartsFillColor
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



' --- Module: Dashboard.bas ---
Attribute VB_Name = "Dashboard"
Option Explicit

Public Sub UpdateChartsFillColor()
    StartCode
    ShDashboard.Unprotect
    Call RefreshDataBackend
    Call UpdateChartColor("ChartSyntheseProjects")
    Call BuildBackendChart
    Call RefreshCondFormList
    ShDashboard.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    StopCode
End Sub

Public Sub BuildBackendChart()
    Dim myChart As Chart
    Dim mySeries As Series
    Dim myTable As ListObject
    Dim e As Range
    Dim count As Double
    Dim SheetName As String
    
    SheetName = "=" & ShSettings.Name & "!"
    Set myChart = ShDashboard.ChartObjects("ChartBackEnd").Chart
    Set myTable = ShSettings.ListObjects("TabSettEvalProjet")
    count = 0
    While myChart.FullSeriesCollection.count > 0
        myChart.FullSeriesCollection(1).Delete
    Wend
    For Each e In myTable.ListColumns("L�gende").DataBodyRange
        Set mySeries = myChart.SeriesCollection.NewSeries
        count = count + 0.5
        With mySeries
            .Name = SheetName & e.Address
            .Values = SheetName & e.Offset(, 3).Address
        End With
        With mySeries.Format.Fill
             .Visible = msoTrue
             .TwoColorGradient msoGradientVertical, 1
             .GradientStops(1).Color.RGB = e.Font.Color
             .GradientStops(1).Transparency = 0.85
             .GradientStops(1).Color.Brightness = 0.1
             .GradientStops(2).Color.RGB = e.Font.Color
             .GradientStops(2).Transparency = 0.6 - count * 0.05
         End With
    Next e
    myChart.Axes(xlValue).MaximumScale = ShBackend.Range("MaxEvaluation")
    myChart.Axes(xlValue).MinimumScale = 0
End Sub

Sub UpdateChartColor(MyChartName As String)
    Dim Chrt As Chart
    Dim TabSettEvalProjet As ListObject
    Dim i As Integer
    Dim categoryName As Variant
    Dim Position As Long
    Dim ToFind As Long
    Dim myTable As ListObject
    
    Set TabSettEvalProjet = ShSettings.ListObjects("TabSettEvalProjet")
    Set myTable = ShDashboard.ListObjects("TabSynthese")
    Set Chrt = ShDashboard.ChartObjects(MyChartName).Chart
    If Chrt Is Nothing Then Exit Sub
    If Chrt.SeriesCollection(1) Is Nothing Then Exit Sub
    categoryName = Chrt.Axes(xlCategory).CategoryNames
    For i = 1 To UBound(categoryName)
        On Error Resume Next
        ToFind = Application.Match(categoryName(i), myTable.ListColumns("Volet").DataBodyRange, 0)
        If Not IsError(Position) Then
            Position = myTable.ListColumns("Position").DataBodyRange.Rows(ToFind).value
            With Chrt.SeriesCollection(1).Points(i).Format.Fill
                .Visible = msoTrue
                .TwoColorGradient msoGradientHorizontal, 1
                .GradientStops(1).Color.RGB = TabSettEvalProjet.ListColumns("L�gende").DataBodyRange.Rows(Position).Font.Color
                .GradientStops(1).Color.Brightness = 0.25
                .GradientStops(2).Color.RGB = TabSettEvalProjet.ListColumns("L�gende").DataBodyRange.Rows(Position).Font.Color
            End With
        End If
    Next i
    Chrt.Axes(xlValue).MaximumScale = ShBackend.Range("MaxEvaluation")
    Chrt.Axes(xlValue).MinimumScale = 0
End Sub

Sub UpdateEvalutionSynthese()
    Dim TabBackendProjects As ListObject
    Dim TabSettEvalProjet As ListObject
    Dim e As Variant
    Dim l As Variant
    
    Set TabBackendProjects = ShBackend.ListObjects("TabBackendProjects")
    Set TabSettEvalProjet = ShSettings.ListObjects("TabSettEvalProjet")
    
    For Each e In TabBackendProjects.ListColumns("L�gende").DataBodyRange
        For Each l In TabSettEvalProjet.ListColumns("L�gende").DataBodyRange
            If e = l Then
                With ShDashboard.Shapes("ShrpResult" & e.Offset(0, -2)).Fill
                    .TwoColorGradient msoGradientHorizontal, 1
                    .GradientStops(1).Color.RGB = l.Font.Color
                    .GradientStops(1).Color.Brightness = 0.25
                    .GradientStops(2).Color.RGB = l.Font.Color
                End With
                With ShDashboard.Shapes("ShrpTitle" & e.Offset(0, -2)).Fill
                    .TwoColorGradient msoGradientVertical, 1
                    .GradientStops(1).Color.RGB = l.Font.Color
                    .GradientStops(1).Transparency = 0
                    .GradientStops(1).Color.Brightness = 0.3
                    .GradientStops(2).Color.RGB = l.Font.Color
                    .GradientStops(2).Transparency = 0.4
                End With
                With ShDashboard.Shapes("ShrpEval" & e.Offset(0, -2)).Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = l.Font.Color
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = 0.7
                    .Transparency = 0
                    .Solid
                End With
                With ShDashboard.Shapes("ShrpNote" & e.Offset(0, -2)).Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = l.Font.Color
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = 0.5
                    .Transparency = 0
                    .Solid
                End With
            End If
        Next l
    Next e
End Sub


Sub RefreshCondFormList()
    Dim TableList As ListObject
    Dim e As Range
    Dim myFormula As String
    Dim myTab As ListObject
    
    Dim COL_LEGENDE As String
    
    Set myTab = ShDashboard.ListObjects("TabSynthese")
    COL_LEGENDE = IIf(ShSettings.Range("Langue") = 2, "L�gende", "L�gende")
    
    myFormula = "=" & Mid(myTab.ListColumns(COL_LEGENDE).DataBodyRange(1).Address, 1, 2) & myTab.DataBodyRange.Row & "="
    ShDashboard.Cells.FormatConditions.Delete
    Set TableList = ShSettings.ListObjects("TabSettEvalProjet")

    For Each e In TableList.ListColumns(2).DataBodyRange
        With myTab.DataBodyRange
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

Sub RefreshDataBackend()
    Dim tabData As ListObject
    Dim countTables As Long
    Dim currentRows As Long
    Dim CurrTab As ListObject
    Dim SumPonderation As Long
    Dim SumCoeffPonderation As Long
    Dim ColCoeffPond As String
    Dim i As Long

    Set tabData = ShDashboard.ListObjects("TabSynthese")
    countTables = ShMatrice.ListObjects.count
    ColCoeffPond = IIf(ShSettings.Range("Langue") = 1, "Weighting Coefficient", "Coefficient de pond�ration")
    
    If Not tabData.DataBodyRange Is Nothing Then
        currentRows = tabData.ListRows.count
        
        If currentRows > countTables Then
            tabData.DataBodyRange.Offset(countTables, 0).Resize(currentRows - countTables).Rows.Delete
        ElseIf currentRows < countTables Then
            Do While tabData.ListRows.count < countTables
                tabData.ListRows.Add
            Loop
        End If
    Else
        Do While tabData.ListRows.count < countTables
            tabData.ListRows.Add
        Loop
    End If
    If countTables = 0 Then
        tabData.ListRows.Add
    End If
    i = 1
    For Each CurrTab In ShMatrice.ListObjects
        tabData.ListColumns("Volet").DataBodyRange(i) = CurrTab.HeaderRowRange(1).Offset(-1)
        SumPonderation = Application.WorksheetFunction.Sum(CurrTab.ListColumns("Pond�ration").DataBodyRange)
        SumCoeffPonderation = Application.WorksheetFunction.Sum(CurrTab.ListColumns(ColCoeffPond).DataBodyRange)
        If SumCoeffPonderation > 0 Then
            tabData.ListColumns("Moyenne").DataBodyRange(i) = SumPonderation / SumCoeffPonderation
        Else
            tabData.ListColumns("Moyenne").DataBodyRange(i) = 0
        End If
        i = i + 1
    Next CurrTab
End Sub





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

Private Sub Worksheet_activate()
    Call UpdateChartsFillColor
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
        Case "BtnDashboard": ShDashboard.Activate
        Case "BtnMatrice": ShMatrice.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Public Function GetTableName() As String
    Dim cell As Range
    Dim lo As ListObject
    
    GetTableName = vbNullString
    Set cell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    
    On Error Resume Next
    Set lo = cell.ListObject
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        GetTableName = lo.Name
    End If
End Function

Public Sub AddRow()
    Dim tableName As String
    StartCode
    tableName = GetTableName
    If Not tableName = vbNullString Then
        Call AddRowHelper(tableName)
    End If
    StopCode
End Sub

Public Sub DeleteRow()
    Dim tableName As String
    
    StartCode
    tableName = GetTableName
    If Not tableName = vbNullString Then
        Call DeleteRowHelper(tableName)
    End If
    StopCode
End Sub

Public Sub DeleteRowHelper(tabName As String)
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim SelectedRange As Range
    
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
        If myTable.ListRows.count = 0 Then myTable.ListRows.Add
    End If
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub AddRowHelper(tabName As String)
    Dim myTable As ListObject
    Dim MyCell As Range
    
    Set myTable = ActiveSheet.ListObjects(tabName)
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Or ActiveCell.Row < myTable.DataBodyRange.Row Then
        myTable.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
    End If
End Sub

Public Function FindValueInRange(lookupValue As Double, tableRange As Range) As Long
    Dim statusRange As Range
    Dim minRange As Range
    Dim maxRange As Range
    Dim i As Long

    Set statusRange = tableRange.Columns(3)
    Set minRange = tableRange.Columns(1)
    Set maxRange = tableRange.Columns(4)
    
    For i = 1 To tableRange.Rows.count
        If lookupValue > IIf(minRange.Cells(i).value = "-", -10000, minRange.Cells(i).value) And lookupValue <= IIf(maxRange.Cells(i).value = "" Or maxRange.Cells(i).value = "-", 10000, maxRange.Cells(i).value) Then
            FindValueInRange = i
            Exit Function
        End If
    Next i
End Function

Public Sub StartCode()
    If ActiveSheet.Name = ShMatrice.Name Then ShMatrice.Unprotect
    Application.EnableEvents = False
    Application.ScreenUpdating = False

End Sub

Public Sub StopCode()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If ActiveSheet.Name = ShMatrice.Name Then ShMatrice.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Application.Calculation = xlCalculationAutomatic
End Sub




' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit
Public Sub RefreshCondFormat()
    StartCode
    ShMatrice.Unprotect
    ShMatrice.Cells.FormatConditions.Delete
    Call RefreshCondFormListChoixEval
    Call RefreshCondFormCoeffPond
    ShMatrice.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    StopCode
End Sub

Sub LogInAdmin()
    Dim Password As String
    If Application.Caller = "BtnLogInAdmin" Then
        Password = InputBox("Veuillez taper le mot de passe pour d�verouiller l'interface Admin")
        If Password = "BGDP" Then
            ShAdmin.Visible = xlSheetVisible
            ShAdmin.Activate
        Else
            MsgBox ("Mot de passe incorect")
        End If
    ElseIf Application.Caller = "BtnLogOutAdmin" Then
        ShAdmin.Visible = xlSheetVeryHidden
        ShSettings.Activate
    End If
End Sub

Private Sub RefreshCondFormListChoixEval()
    Dim TableList As ListObject
    Dim e As Range
    Dim UnionSections As Range
    Dim i As Integer
    
    If ShMatrice.ListObjects.count = 0 Then Exit Sub
    
    With ShMatrice
        Set UnionSections = .ListObjects(1).ListColumns("Evaluation").DataBodyRange
        For i = 2 To .ListObjects.count
            Set UnionSections = Union(UnionSections, .ListObjects(i).ListColumns("Evaluation").DataBodyRange)
        Next
    End With
    Set TableList = ShSettings.ListObjects("TabSettLegende")

    For Each e In TableList.ListColumns("Choix").DataBodyRange
        With UnionSections
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$D18=""" & e.value & """"
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

Sub RefreshCondFormCoeffPond()
    Dim e As Range
    Dim UnionSections As Range
    Dim CoeffPond As String
    Dim i As Integer
    
    CoeffPond = IIf(ShSettings.Range("Langue") = 2, "Coefficient de pond�ration", "Weighting Coefficient")
    
    If ShMatrice.ListObjects.count = 0 Then Exit Sub
    
    With ShMatrice
        Set UnionSections = .ListObjects(1).ListColumns(CoeffPond).DataBodyRange
        For i = 2 To .ListObjects.count
            Set UnionSections = Union(UnionSections, .ListObjects(i).ListColumns(CoeffPond).DataBodyRange)
        Next
    End With

    With UnionSections
        .FormatConditions.AddDatabar
        .FormatConditions(.FormatConditions.count).ShowValue = True
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueFormula, newvalue:=0  'xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueFormula, newvalue:=WorksheetFunction.Max(ShSettings.ListObjects("TabSettCoeffPond").DataBodyRange)  'xlConditionValueAutomaticMax
        End With
        With .FormatConditions(1).BarColor
            .Color = 2668287
            .TintAndShade = 0
        End With
        .FormatConditions(1).BarFillType = xlDataBarFillGradient
        .FormatConditions(1).Direction = xlContext
        .FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        .FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        .FormatConditions(1).NegativeBarFormat.BorderColorType = xlDataBarColor
        With .FormatConditions(1).BarBorder.Color
            .Color = 2668287
            .TintAndShade = 0
        End With
        .FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With .FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With .FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With .FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
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
    
    ShMatrice.Unprotect
    ShDashboard.Unprotect
    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("K1:K5"), .Range("D6"), .Range("K6"), .Range("D8:L9"), .Range("D11:L14"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShMatrice
        trad = vbNullString
        For Each e In Union(.Range("D1:D15"), .Range("C16:F200"), .Range("C6"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
    End With
    
    With ShDashboard
        trad = vbNullString
        For Each e In Union(.Range("G1:G5"), .Range("HEADER_SYNTH"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
    End With

    With ShBackend
        trad = vbNullString
        For Each e In .Range("ModelTab")
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
    End With


    For Each sht In ActiveWorkbook.Sheets
        trad = vbNullString
        For Each sh In sht.Shapes
            If sh.Type = 1 Or (sh.Type = 28 And Left(sh.Name, 3) = "Rem") Then
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
    ShMatrice.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ShDashboard.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ShSettings.Activate
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







' --- Module: Matrice.bas ---
Attribute VB_Name = "Matrice"
Option Explicit

Sub AddVolet()
    Dim tabModelRange As Range
    Dim lastRow As Long
    Dim VoletName As String
    Dim shp As Shape
    Dim uniqueName As String
    Dim currentTime As String
    
    StartCode
    Set tabModelRange = ShBackend.Range("ModelTab")
    
    lastRow = ShMatrice.Cells(Rows.count, 3).End(xlUp).Row
    
    If lastRow < 16 Then lastRow = 14
    
    VoletName = InputBox("Entrez le nom du volet :", "Ajout de volet")
    
    If VoletName <> "" Then
        
        currentTime = Format(Now, "hhmmss")
    
        For Each shp In ShBackend.Shapes
            uniqueName = Left(shp.Name, 6) & "_" & currentTime
            shp.Name = uniqueName
        Next shp
        
        tabModelRange.Copy ShMatrice.Range("C" & lastRow + 2)
        
        ShMatrice.Range("C" & lastRow + 2).value = VoletName
        
        MsgBox "Volet " & VoletName & " ajout� avec succ�s", vbInformation
    End If
    StopCode
    Call RefreshCondFormat
End Sub

Sub DeleteVolet()
    Dim tableName As String
    Dim voletRange As Range
    Dim shp As Shape
    Dim rep As VbMsgBoxResult
    Dim VoletName As String
    
    StartCode
    tableName = GetTableName
    If tableName = vbNullString Then Exit Sub
    Set voletRange = ShMatrice.ListObjects(tableName).Range
    Set voletRange = voletRange.Resize(voletRange.Rows.count + 2).Offset(-1, 0)
    VoletName = voletRange(1)

    rep = MsgBox("Confirmez-vous la suppression de ce volet ?", vbYesNo + vbQuestion, "Confirmation")
    
    
    If rep = vbYes Then
        For Each shp In ShMatrice.Shapes
            On Error Resume Next
            If Left(shp.Name, 4) <> "Drop" Then
                If Not Intersect(shp.TopLeftCell, voletRange) Is Nothing Then
                    shp.Delete
                End If
            End If
        Next shp
        
        voletRange.Delete Shift:=xlUp
        MsgBox "Volet " & VoletName & " supprim� avec succ�s"
    End If
    StopCode
End Sub


