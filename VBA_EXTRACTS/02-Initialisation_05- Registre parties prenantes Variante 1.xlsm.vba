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
    Dim currentIndex As Long
    Dim tabPP As ListObject
    Dim e As Range
    
    If Target.count > 100 Then Exit Sub
    StartCode
    Set tabPP = ShRegPP.ListObjects("TabPP")
    For Each e In Target
        currentIndex = e.Row - tabPP.HeaderRowRange.Row
        With tabPP
            If Not Intersect(e, .ListColumns("Intérêt").DataBodyRange.Resize(, 2)) Is Nothing Then
                Select Case .ListColumns("Intérêt").DataBodyRange(currentIndex) & .ListColumns("Pouvoir").DataBodyRange(currentIndex)
                    Case ShLangue.Range("H3") & ShLangue.Range("H3") '"FaibleFaible"
                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = ShSettings.Range("CE17") ' "Effort Minimal"
                    Case ShLangue.Range("H3") & ShLangue.Range("H2") '"FaibleÉlevé"
                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = ShSettings.Range("CE18") '"Garder Satisfait"
                    Case ShLangue.Range("H2") & ShLangue.Range("H3") '"ÉlevéFaible"
                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = ShSettings.Range("CE16") '"Garder Informé"
                    Case ShLangue.Range("H2") & ShLangue.Range("H2") '"ÉlevéÉlevé"
                        .ListColumns("Stratégie").DataBodyRange(currentIndex) = ShSettings.Range("CE15") '"Acteur Clé"
                End Select
            ElseIf Not Intersect(e, tabPP.ListColumns("Stratégie").DataBodyRange) Is Nothing Then
                Select Case .ListColumns("Stratégie").DataBodyRange(currentIndex)
                    Case ShSettings.Range("CE17") '"Effort Minimal"
                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = ShLangue.Range("H3") '"Faible"
                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = ShLangue.Range("H3") '"Faible"
                    Case ShSettings.Range("CE18") '"Garder Satisfait"
                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = ShLangue.Range("H3") '"Faible"
                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = ShLangue.Range("H2") '"Élevé"
                    Case ShSettings.Range("CE16") '"Garder Informé"
                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = ShLangue.Range("H2") '"Élevé"
                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = ShLangue.Range("H3") '"Faible"
                    Case ShSettings.Range("CE15") '"Acteur Clé"
                        .ListColumns("Intérêt").DataBodyRange(currentIndex) = ShLangue.Range("H2") '"Élevé"
                        .ListColumns("Pouvoir").DataBodyRange(currentIndex) = ShLangue.Range("H3") '"Faible"
                End Select
            End If
        End With
    Next e
    StopCode
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


Public Sub AddRow()
    If Application.Caller = "BtnAddType" Then
        Call AddRowHelper("TabType")
    ElseIf Application.Caller = "btnAddPP" Then
        Call AddRowHelper("TabPP")
    End If
End Sub

Public Sub DeleteRow()
    If Application.Caller = "BtnDeleteType" Then
        Call DeleteRowHelper("TabType")
    ElseIf Application.Caller = "BtnDeletePP" Then
        Call DeleteRowHelper("TabPP")
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
'    myTable.DataBodyRange.Rows.AutoFit
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
'    myTable.DataBodyRange.Rows.AutoFit
'    StopCode
'End Sub

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
    ShRegPP.Cells.FormatConditions.Delete
    Call RefreshCondFormListStrategie
End Sub


' --- Module: PP.bas ---
Attribute VB_Name = "PP"
Option Explicit

Public Sub RefreshCondFormListStrategie()
    Dim tabStrategie As ListObject
    Dim e As Range
 
    Set tabStrategie = ShSettings.ListObjects("TabStratégie")

    For Each e In tabStrategie.ListColumns("Stratégie").DataBodyRange
        With ShRegPP.ListObjects("TabPP").DataBodyRange
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$H14=""" & e.value & """"
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
    Call GetMatrice
    StopCode
End Sub



' --- Module: Matrice.bas ---
Attribute VB_Name = "Matrice"
Option Explicit

Public Sub GetMatrice()
    Dim tabPP As ListObject
    Dim e As Range
    Dim currentIndex As Long
    Dim shpName As String
    
    Set tabPP = ShRegPP.ListObjects("TabPP")
    ShMatrice.Shapes("MatFaibleFaible").TextFrame2.TextRange.Characters.Text = vbNullString
    ShMatrice.Shapes("MatFaibleÉlevé").TextFrame2.TextRange.Characters.Text = vbNullString
    ShMatrice.Shapes("MatÉlevéFaible").TextFrame2.TextRange.Characters.Text = vbNullString
    ShMatrice.Shapes("MatÉlevéÉlevé").TextFrame2.TextRange.Characters.Text = vbNullString
    If tabPP.DataBodyRange.count > 0 Then
        For Each e In tabPP.ListColumns("Stratégie").DataBodyRange
            With tabPP
                currentIndex = e.Row - .HeaderRowRange.Row
                Select Case e
                    Case ShSettings.Range("CE17") ' "Effort Minimal"
                        shpName = "MatFaibleFaible"
                    Case ShSettings.Range("CE18") '"Garder Satisfait"
                        shpName = "MatFaibleÉlevé"
                    Case ShSettings.Range("CE16") '"Garder Informé"
                        shpName = "MatÉlevéFaible"
                    Case ShSettings.Range("CE15") '"Acteur Clé"
                        shpName = "MatÉlevéÉlevé"
                    Case Else
                        shpName = vbNullString
                End Select
                If shpName <> vbNullString Then
                    With ShMatrice.Shapes(shpName).TextFrame2.TextRange.Characters
                        .Text = .Text & IIf(Len(.Text) > 0, vbNewLine, "") & _
                        tabPP.ListColumns("ID").DataBodyRange(currentIndex) & " - " & tabPP.ListColumns("Nom").DataBodyRange(currentIndex)
                    End With
                End If
            End With
        Next e
    End If
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
        For Each e In Union(.Range("E1:L5"), .Range("C6"), .Range("L6"), .Range("C15:C16"), .Range("CE15:CE18"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShMatrice
        trad = vbNullString
        For Each e In Union(.Range("E1:K6"), .Range("C6"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShRegPP
        trad = vbNullString
        For Each e In Union(.Range("E1:K5"), .Range("C6"), .Range("D14:K25"), .Range("B12:L12"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

    With ShLangue
        trad = vbNullString
        For Each e In .Range("H2:J3")
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








