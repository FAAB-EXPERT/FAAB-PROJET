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

Private Sub Workbook_Open()
        Application.ErrorCheckingOptions.BackgroundChecking = False
End Sub

Private Sub Workbook_Deactivate()
        Application.ErrorCheckingOptions.BackgroundChecking = True
End Sub


' --- Module: ShDemandes.cls ---
Attribute VB_Name = "ShDemandes"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim oldVal As String
    Dim newVal As String
    Dim myTable As ListObject
    
    Set myTable = ShDemandes.ListObjects("TabDemandes")
    
    If Target.count > 1 Then GoTo exitHandler
        

    If Not Intersect(Target, myTable.ListColumns("Impact").DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        newVal = Target.value
        Application.Undo
        oldVal = Target.value
        Target.value = newVal
        If oldVal <> "" And newVal <> "" Then
            If InStr(1, oldVal, newVal) = 0 Then
                Target.value = oldVal & ", " & newVal
            Else
                Target.value = oldVal
            End If
        End If
    End If
    
exitHandler:
    Application.EnableEvents = True
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
        Case "BtnDemandes": ShDemandes.Activate
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


' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit

Public Sub RefreshCondFormat()
    StartCode
    ShDemandes.Cells.FormatConditions.Delete
    Call RefreshCondFormListStatut

    StopCode
End Sub




' --- Module: Demandes.bas ---
Attribute VB_Name = "Demandes"
Option Explicit

Sub RefreshCondFormListStatut()
    Dim TableList As ListObject
    Dim e As Range
    Dim myFormula As String
    Set TableList = ShSettings.ListObjects("TabStatut")

    myFormula = "=" & Mid(ShDemandes.ListObjects("TabDemandes").ListColumns("Status").DataBodyRange(1).Address, 1, 2) & "12="
    For Each e In TableList.ListColumns("Statut").DataBodyRange
        With ShDemandes.ListObjects("TabDemandes").ListColumns("Status").DataBodyRange
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


Public Sub AddColumn()
    Dim TabDemandes As ListObject
    Dim currentCol As Long
    
    StartCode
    Set TabDemandes = ShDemandes.ListObjects("TabDemandes")
    If Not Intersect(ActiveCell, TabDemandes.HeaderRowRange.Offset(-1)) Is Nothing Then
        currentCol = ActiveCell.Column - TabDemandes.DataBodyRange.Column + 1
        ShDemandes.Cells(ShDemandes.Range("HeaderRange").Row, currentCol + 1).Insert Shift:=xlToRight
        TabDemandes.ListColumns.Add currentCol
        ShDemandes.Columns(currentCol + 1).EntireColumn.Validation.Delete
    ElseIf Intersect(ActiveCell, TabDemandes.DataBodyRange) Is Nothing Then
        TabDemandes.ListColumns.Add
        Cells(ShDemandes.Range("HeaderRange").Row, TabDemandes.ListColumns.count + 1).Insert Shift:=xlToRight
        With Cells(ShDemandes.Range("HeaderRange").Row, TabDemandes.ListColumns.count + 1)
            .Offset(, -1).Copy
            .PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End With
        ShDemandes.Columns(TabDemandes.ListColumns.count + 1).EntireColumn.Validation.Delete
    Else
        MsgBox ("La selection n'est pas valide !")
    End If
    StopCode
    RefreshCondFormat
End Sub

Public Sub DeleteColumn()
    Dim TabDemandes As ListObject
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim currentCol As Long
    Dim SelectedRange As Range
    
    StartCode
    Set TabDemandes = ShDemandes.ListObjects("TabDemandes")
    
    Set SelectedRange = Intersect(Selection, Selection.Rows(1), TabDemandes.HeaderRowRange.Offset(-1))
    If Not SelectedRange Is Nothing And TabDemandes.ListColumns.count > 1 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de " & SelectedRange.count & " colonne(s) ?", vbOKCancel)
        If MsgAnswer = vbOK Then
            For count = SelectedRange.count To 1 Step -1
                currentCol = SelectedRange(count).Column - TabDemandes.DataBodyRange.Column + 1
                If TabDemandes.ListColumns.count > 2 Then
                    If Not Intersect(SelectedRange(1, count), ShDemandes.Range("HeaderRange")) Is Nothing Then
                        ShDemandes.Cells(ShDemandes.Range("HeaderRange").Row, currentCol + 1).Delete Shift:=xlToLeft
                        TabDemandes.ListColumns(currentCol).Delete
                    End If
                End If
            Next count
        End If
    Else
        MsgBox ("La selection n'est pas valide!")
    End If
    StopCode
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
        For Each e In Union(.Range("E1:N5"), .Range("C6"), .Range("J6"), .Range("C9"), .Range("E9:F9"), .Range("H9"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabStatut").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabPriority").ListColumns("Niveaux de Priorité").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabImpact").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShDemandes
        trad = vbNullString
        For Each e In Union(.Range("E1:M5"), .Range("C10:L10"), .Range("C6"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabDemandes").ListColumns("Status").DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        For Each e In .ListObjects("TabDemandes").ListColumns("Impact").DataBodyRange
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





