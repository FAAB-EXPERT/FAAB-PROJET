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
    'Call RefreshCondFormat
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



' --- Module: ShRegPP.cls ---
Attribute VB_Name = "ShRegPP"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim currentIndex As Long
'    Dim tabPP As ListObject
'    Dim e As Range
'
'    If Target.count > 100 Then Exit Sub
'    StartCode
'    Set tabPP = ShRegPP.ListObjects("TabPP")
'    For Each e In Target
'        currentIndex = e.Row - tabPP.HeaderRowRange.Row
'        With tabPP
'            If Not Intersect(e, .ListColumns("Intérêt").DataBodyRange.Resize(, 2)) Is Nothing Then
'                Select Case .ListColumns("Intérêt").DataBodyRange(currentIndex) & .ListColumns("Pouvoir").DataBodyRange(currentIndex)
'                    Case "FaibleFaible"
'                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = "Effort Minimal"
'                    Case "FaibleÉlevé"
'                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = "Garder Satisfait"
'                    Case "ÉlevéFaible"
'                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = "Garder Informé"
'                    Case "ÉlevéÉlevé"
'                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = "Acteur Clé"
'                End Select
'            ElseIf Not Intersect(e, tabPP.ListColumns("Stratégie").DataBodyRange) Is Nothing Then
'                Select Case .ListColumns("Stratégie").DataBodyRange(currentIndex)
'                    Case "Effort Minimal"
'                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = "Faible"
'                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = "Faible"
'                    Case "Garder Satisfait"
'                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = "Faible"
'                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = "Élevé"
'                    Case "Garder Informé"
'                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = "Élevé"
'                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = "Faible"
'                    Case "Acteur Clé"
'                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = "Élevé"
'                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = "Faible"
'                End Select
'            End If
'        End With
'    Next e
'    StopCode
End Sub



' --- Module: Helper.bas ---
Attribute VB_Name = "Helper"
Option Explicit

Public Sub NavBar()
    Select Case Application.Caller
        Case "BtnGuide": ShGuide.Activate
        Case "BtnPP": ShRegPP.Activate
        Case "BtnMatrice": ShMatrice.Activate
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
    Set myTable = ActiveSheet.ListObjects("TabPP")
    
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
    If myTable.ListRows.count > 0 Then
        myTable.DataBodyRange.Rows.AutoFit
    End If
    Application.Calculation = xlCalculationAutomatic
    StopCode
End Sub

Public Sub AddRow()
    Dim myTable As ListObject
    Dim MyCell As Range
    
    StartCode
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    Set myTable = ActiveSheet.ListObjects("TabPP")
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Then
        myTable.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
    End If
    myTable.DataBodyRange.Rows.AutoFit
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


' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit

'Public Sub RefreshCondFormat()
'    ShRegPP.Cells.FormatConditions.Delete
'    Call RefreshCondFormListStrategie
'End Sub


' --- Module: PP.bas ---
Attribute VB_Name = "PP"
Option Explicit

'Public Sub RefreshCondFormListStrategie()
'    Dim tabStrategie As ListObject
'    Dim e As Range
'
'    Set tabStrategie = ShSettings.ListObjects("TabStratégie")
'
'    For Each e In tabStrategie.ListColumns("Stratégie").DataBodyRange
'        With ShRegPP.ListObjects("TabPP").DataBodyRange
'            .FormatConditions.Add Type:=xlExpression, Formula1:="=$H14=""" & e.Value & """"
'            .FormatConditions(.FormatConditions.count).SetFirstPriority
'            With .FormatConditions(1).Font
'                On Error Resume Next
'                .Color = e.Font.Color
'                If e.Font.ThemeColor > 0 Then .ThemeColor = e.Font.ThemeColor
'                .Bold = e.Font.Bold
'                .Italic = e.Font.Italic
'                .TintAndShade = e.Font.TintAndShade
'            End With
'            With .FormatConditions(1).Interior
'                .PatternColorIndex = e.Interior.PatternColorIndex
'                If e.Interior.ThemeColor > 0 Then
'                    .ThemeColor = e.Interior.ThemeColor
'                    .Pattern = e.Interior.Pattern
'                    .PatternColorIndex = e.Interior.PatternColorIndex
'                Else
'                    .Color = e.Interior.Color
'                End If
'                .TintAndShade = e.Interior.TintAndShade
'            End With
'            .FormatConditions(1).StopIfTrue = False
'        End With
'        Application.CutCopyMode = False
'    Next e
'End Sub



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

Sub Refresh_Connection()
    Dim objConnection As Variant
    Dim bBackground As Variant
    
   ' For Each objConnection In ThisWorkbook.Connections
        Set objConnection = ActiveWorkbook.Connections("Query - TabPP")
        'Get current background-refresh value
        bBackground = objConnection.OLEDBConnection.BackgroundQuery
        
        'Temporarily disable background-refresh
        objConnection.OLEDBConnection.BackgroundQuery = False
        
        'Refresh this connection
        objConnection.Refresh

        'Set background-refresh value back to original value
        objConnection.OLEDBConnection.BackgroundQuery = bBackground
   ' Next
    
End Sub

Private Sub Worksheet_Activate()
    StartCode

  '  Refresh_Connection
    copyData
    'ShMatrice.Columns(5).ColumnWidth = 30.44
    'ShMatrice.Columns(6).ColumnWidth = 30.44
    SetDiagram
    StopCode
End Sub

Private Sub copyData()
    Dim count As Long

    count = ShRegPP.ListObjects("TabPP").ListColumns(1).DataBodyRange.count
    ShMatrice.Range("E18:F200").ClearContents
    ShMatrice.Range("E18").Resize(count, 1).Value = ShRegPP.ListObjects("TabPP").ListColumns(1).DataBodyRange.Value
    ShMatrice.Range("F18").Resize(count, 1).Value = ShRegPP.ListObjects("TabPP").ListColumns(2).DataBodyRange.Value

End Sub


' --- Module: Matrice.bas ---
Attribute VB_Name = "Matrice"
Option Explicit

Private Sub ResetDiagram()
    Dim e As Shape
    
    For Each e In ShMatrice.Shapes
        If Left(e.Name, 4) = "MAT_" Then
            e.Delete
        End If
    Next e
End Sub

Public Sub SetDiagram()
    Dim tabPP As ListObject
    Dim e As Range
    Dim sh As Shape
    Dim x As Long
    Dim y As Long
    Dim i As Integer
    Dim tabX(1 To 9) As Long
    Dim tabY(1 To 9) As Long
    Dim tabCount(1 To 9) As Long
    Set tabPP = ShRegPP.ListObjects("TabPP")
    Const DM As Long = 11
    
    
    ResetDiagram
    
    With ShMatrice
        tabX(1) = .Range("G13").Left + .Range("G13").WIDTH / 2 - DM
        tabY(1) = .Range("E13").Top + .Range("E13").Height / 3 - DM
        
        tabX(2) = .Range("F13").Left + .Range("F13").WIDTH / 2 - DM
        tabY(2) = .Range("E13").Top + .Range("E13").Height / 3 - DM
        
        tabX(3) = .Range("E13").Left + .Range("E13").WIDTH / 2 - DM
        tabY(3) = .Range("E13").Top + .Range("E13").Height / 3 - DM
        
        tabX(4) = .Range("G13").Left + .Range("G13").WIDTH / 2 - DM
        tabY(4) = .Range("E14").Top + .Range("E14").Height / 3 - DM
        
        tabX(5) = .Range("F13").Left + .Range("F13").WIDTH / 2 - DM
        tabY(5) = .Range("E14").Top + .Range("E14").Height / 3 - DM
        
        tabX(6) = .Range("E13").Left + .Range("E13").WIDTH / 2 - DM
        tabY(6) = .Range("E14").Top + .Range("E14").Height / 3 - DM
        
        tabX(7) = .Range("G13").Left + .Range("G13").WIDTH / 2 - DM
        tabY(7) = .Range("E15").Top + .Range("E15").Height / 3 - DM
        
        tabX(8) = .Range("F13").Left + .Range("F13").WIDTH / 2 - DM
        tabY(8) = .Range("E15").Top + .Range("E15").Height / 3 - DM
        
        tabX(9) = .Range("E13").Left + .Range("E13").WIDTH / 2 - DM
        tabY(9) = .Range("E15").Top + .Range("E15").Height / 3 - DM
        
        For Each e In tabPP.ListColumns("Acteur").DataBodyRange
            If (e.Offset(, 2) = 5 Or e.Offset(, 2) = 4) And (e.Offset(, 3) = 5 Or e.Offset(, 3) = 4) Then
                x = tabX(1)
                y = tabY(1)
                tabCount(1) = tabCount(1) + 1
                tabX(1) = IIf(tabCount(1) Mod 2, tabCount(1) Mod 7, -tabCount(1) Mod 7) * DM * 2 + tabX(1)
                If tabCount(1) = 7 Then
                    tabY(1) = DM * 2 + tabY(1)
                    tabX(1) = .Range("G13").Left + .Range("G13").WIDTH / 2 - DM
                End If
                
            ElseIf (e.Offset(, 2) = 5 Or e.Offset(, 2) = 4) And (e.Offset(, 3) = 3 Or e.Offset(, 3) = 2) Then
                x = tabX(2)
                y = tabY(2)
                tabCount(2) = tabCount(2) + 1
                tabX(2) = IIf(tabCount(2) Mod 2, tabCount(2) Mod 7, -tabCount(2) Mod 7) * DM * 2 + tabX(2)
                If tabCount(2) = 7 Then
                    tabY(2) = DM * 2 + tabY(1)
                    tabX(2) = .Range("F13").Left + .Range("F13").WIDTH / 2 - DM
                End If
        'FAIL
            ElseIf (e.Offset(, 2) = 5 Or e.Offset(, 2) = 4) And (e.Offset(, 3) = 1 Or e.Offset(, 3) = 0) Then
                x = tabX(3)
                y = tabY(3)
                tabCount(3) = tabCount(3) + 1
                tabX(3) = IIf(tabCount(3) Mod 2, tabCount(3) Mod 7, -tabCount(3) Mod 7) * DM * 2 + tabX(3)
                If tabCount(3) = 7 Then
                    tabY(3) = DM * 2 + tabY(3)
                    tabX(3) = .Range("E13").Left + .Range("E13").WIDTH / 2 - DM
                End If
                
            ElseIf (e.Offset(, 2) = 3 Or e.Offset(, 2) = 2) And (e.Offset(, 3) = 5 Or e.Offset(, 3) = 4) Then
                x = tabX(4)
                y = tabY(4)
                tabCount(4) = tabCount(4) + 1
                tabX(4) = IIf(tabCount(4) Mod 2, tabCount(4) Mod 7, -tabCount(4) Mod 7) * DM * 2 + tabX(4)
                If tabCount(4) = 7 Then
                    tabY(4) = DM * 2 + tabY(4)
                    tabX(4) = .Range("G13").Left + .Range("G13").WIDTH / 2 - DM
                End If
                
            ElseIf (e.Offset(, 2) = 3 Or e.Offset(, 2) = 2) And (e.Offset(, 3) = 3 Or e.Offset(, 3) = 2) Then
                x = tabX(5)
                y = tabY(5)
                tabCount(5) = tabCount(5) + 1
                tabX(5) = IIf(tabCount(5) Mod 2, tabCount(5) Mod 7, -tabCount(5) Mod 7) * DM * 2 + tabX(5)
                If tabCount(5) = 7 Then
                    tabY(5) = DM * 2 + tabY(5)
                    tabX(5) = .Range("F13").Left + .Range("F13").WIDTH / 2 - DM
                End If
            ElseIf (e.Offset(, 2) = 3 Or e.Offset(, 2) = 2) And (e.Offset(, 3) = 1 Or e.Offset(, 3) = 0) Then
                x = tabX(6)
                y = tabY(6)
                tabCount(6) = tabCount(6) + 1
                tabX(6) = IIf(tabCount(6) Mod 2, tabCount(6) Mod 7, -tabCount(6) Mod 7) * DM * 2 + tabX(6)
                If tabCount(6) = 7 Then
                    tabY(6) = DM * 2 + tabY(6)
                    tabX(6) = .Range("E13").Left + .Range("E13").WIDTH / 2 - DM
                End If
            ElseIf (e.Offset(, 2) = 1 Or e.Offset(, 2) = 0) And (e.Offset(, 3) = 5 Or e.Offset(, 3) = 4) Then
                x = tabX(7)
                y = tabY(7)
                tabCount(7) = tabCount(7) + 1
                tabX(7) = IIf(tabCount(7) Mod 2, tabCount(7) Mod 7, -tabCount(7) Mod 7) * DM * 2 + tabX(7)
                If tabCount(7) = 7 Then
                    tabY(7) = DM * 2 + tabY(7)
                    tabX(7) = .Range("G13").Left + .Range("G13").WIDTH / 2 - DM
                End If
            ElseIf (e.Offset(, 2) = 1 Or e.Offset(, 2) = 0) And (e.Offset(, 3) = 3 Or e.Offset(, 3) = 2) Then
                x = tabX(8)
                y = tabY(8)
                tabCount(8) = tabCount(8) + 1
                tabX(8) = IIf(tabCount(8) Mod 2, tabCount(8) Mod 7, -tabCount(8) Mod 7) * DM * 2 + tabX(8)
                If tabCount(8) = 7 Then
                    tabY(8) = DM * 2 + tabY(8)
                    tabX(8) = .Range("F13").Left + .Range("F13").WIDTH / 2 - DM
                End If
            ElseIf (e.Offset(, 2) = 1 Or e.Offset(, 2) = 0) And (e.Offset(, 3) = 1 Or e.Offset(, 3) = 0) Then
                x = tabX(9)
                y = tabY(9)
                tabCount(9) = tabCount(9) + 1
                tabX(9) = IIf(tabCount(9) Mod 2, tabCount(9) Mod 7, -tabCount(9) Mod 7) * DM * 2 + tabX(9)
                If tabCount(9) = 7 Then
                    tabY(9) = DM * 2 + tabY(9)
                    tabX(9) = .Range("E13").Left + .Range("E13").WIDTH / 2 - DM
                End If
            End If
            Set sh = ShMatrice.Shapes.AddShape(msoShapeOval, x, y, 22, 22)
            sh.Name = "MAT_" & e.Offset(, -1)
            sh.TextFrame2.TextRange.Characters.Text = e.Offset(, -1)
            With sh.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
                .FirstLineIndent = 0
                .Alignment = msoAlignLeft
            End With
            With sh.TextFrame2
                .TextRange.Characters(1, 1).Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
            With sh.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorAccent2
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0
                .Solid
            End With
            sh.Line.Visible = msoFalse
            sh.TextFrame2.WordWrap = msoFalse
        Next e
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



