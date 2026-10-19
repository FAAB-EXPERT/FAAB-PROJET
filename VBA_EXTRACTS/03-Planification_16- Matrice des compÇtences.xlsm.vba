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
Option Explicit



' --- Module: ShMatTachesComp.cls ---
Attribute VB_Name = "ShMatTachesComp"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler ' Gestion des erreurs
    Dim tabTacheComp As ListObject
    Dim e As Range

    ' Vérification de l'objet ShMatTachesComp
    If ShMatTachesComp Is Nothing Then Exit Sub

    StartCode
    Set tabTacheComp = ShMatTachesComp.ListObjects("TabTachesCompetences")
    
    ' Vérification de l'objet tabTacheComp
    If tabTacheComp Is Nothing Then Exit Sub

    For Each e In Target
        ' Vérifie si la cellule modifiée est dans l'en-tête de la table
        If Not Intersect(e, tabTacheComp.HeaderRowRange) Is Nothing Then
            Call UpdateTableRessComp((e.Column - 3) * 2, e.value)
        End If
        
        ' Vérifie si la cellule modifiée est dans la plage "Filter_Fonction"
        If Not Intersect(e, Range("Filter_Fonction")) Is Nothing Then
            Call Apply_Filter
        End If
    Next e

    StopCode
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbExclamation, "Erreur"
    StopCode
End Sub

Private Sub Apply_Filter()
    On Error Resume Next
    Dim filterValue As Variant
    
    If ShMatTachesComp Is Nothing Then Exit Sub
    
    If ShMatTachesComp.ListObjects("TabTachesCompetences").AutoFilter Is Nothing = False Then
        ShMatTachesComp.ListObjects("TabTachesCompetences").AutoFilter.ShowAllData
    End If
    
    filterValue = ShMatTachesComp.Range("Filter_Fonction").value
    ShMatTachesComp.ListObjects("TabTachesCompetences").Range.AutoFilter Field:=2, _
        Criteria1:=filterValue
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

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim loTask As ListObject
    Dim wsTaches As Worksheet
    Dim loTachesComp As ListObject
    Dim loComp As ListObject
    Dim oldValue As String, newValue As String
    Dim colComp As Long, rowComp As Long
    Dim wsRess As Worksheet
    Dim loRessComp As ListObject
    Dim i As Long, j As Long
    Dim cell As Range

    Set loComp = Me.ListObjects("TabComp")
    Set loTask = Me.ListObjects("TabTask")
    Set wsTaches = ShMatTachesComp
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")
    Set wsTaches = ShMatTachesComp
    Set wsRess = ShMatRessComp
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")
    Set loRessComp = wsRess.ListObjects("TabRessComp")

    ' Vérifie si la modification concerne la colonne "Compétence" de TabComp
    If Not Intersect(Target, loComp.ListColumns(1).DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Intersect(Target, loComp.ListColumns(1).DataBodyRange)
            oldValue = cell.value ' Nouveau nom
            newValue = cell.value

            ' Trouver l'ancien nom (avant modification)
            Application.Undo
            oldValue = cell.value
            cell.value = newValue

            ' Mise à jour dans TabTachesCompetences (en-tête)
            For i = 1 To loTachesComp.ListColumns.count
                If loTachesComp.HeaderRowRange.Cells(1, i).value = oldValue Then
                    loTachesComp.HeaderRowRange.Cells(1, i).value = newValue
                End If
            Next i

            ' Mise à jour dans TabRessComp (en-tête)
            For i = 1 To loRessComp.ListColumns.count
                If loRessComp.HeaderRowRange.Cells(1, i).value = oldValue Then
                    loRessComp.HeaderRowRange.Cells(1, i).value = newValue
                End If
            Next i
        Next cell
        Application.EnableEvents = True
    End If
    If Not Intersect(Target, loTask.ListColumns(1).DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Intersect(Target, loTask.ListColumns(1).DataBodyRange)
            newValue = cell.value
            Application.Undo
            oldValue = cell.value
            cell.value = newValue

            ' Met à jour toutes les lignes correspondantes dans TabTachesCompetences
            For i = 1 To loTachesComp.ListRows.count
                If loTachesComp.DataBodyRange.Cells(i, 1).value = oldValue Then
                    loTachesComp.DataBodyRange.Cells(i, 1).value = newValue
                End If
            Next i
        Next cell
        Application.EnableEvents = True
    End If
End Sub


' --- Module: Settings.bas ---
Attribute VB_Name = "Settings"
Option Explicit

Public Sub AddNewTaskToCompetences()
    Dim wsSettings As Worksheet, wsTaches As Worksheet
    Dim loTask As ListObject, loFonction As ListObject, loTachesComp As ListObject
    Dim funcArr As Variant
    Dim i As Long, nbCols As Long
    Dim newRow As ListRow
    Dim taskName As String
    Dim selRow As Long, insertPos As Long
    Dim selectedCell As Range
    Dim filterWasOn As Boolean

    StartCode
    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set loTask = wsSettings.ListObjects("TabTask")
    Set loFonction = wsSettings.ListObjects("TabFonction")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")

    ' Demander le nom de la tâche
    taskName = InputBox("Saisir le nom de la nouvelle tâche :", "Nouvelle tâche")
    If Trim(taskName) = "" Then Exit Sub

    ' Déterminer la position d'insertion dans TabTask
    Set selectedCell = Application.Selection

    If Not loTask.DataBodyRange Is Nothing Then
        If Not Intersect(selectedCell, loTask.DataBodyRange) Is Nothing Then
            selRow = selectedCell.Row - loTask.DataBodyRange.Rows(1).Row + 1
            insertPos = selRow
        Else
            insertPos = loTask.ListRows.count + 1 ' Ajout à la fin
        End If
    Else
        insertPos = 1 ' Table vide, on ajoute en première ligne
    End If

    ' Insérer la tâche dans TabTask
    loTask.ListRows.Add insertPos
    loTask.DataBodyRange.Rows(insertPos).Cells(1, 1).value = taskName

    ' Récupérer toutes les fonctions existantes
    funcArr = loFonction.DataBodyRange.value
    nbCols = loTachesComp.ListColumns.count

    ' Supprimer le filtre si actif
    filterWasOn = loTachesComp.AutoFilter.FilterMode
    If filterWasOn Then loTachesComp.AutoFilter.ShowAllData

    ' Pour chaque fonction, insérer la tâche à la même position dans TabTachesCompetences
    For i = 1 To UBound(funcArr, 1)
        Dim pos As Long, found As Boolean
        found = False
        For pos = 1 To loTachesComp.ListRows.count
            If loTachesComp.DataBodyRange.Cells(pos, 2).value = funcArr(i, 1) Then
                found = True
                Exit For
            End If
        Next pos
        If found Then
            loTachesComp.ListRows.Add insertPos
            loTachesComp.DataBodyRange.Rows(insertPos).Cells(1, 1).value = taskName
            loTachesComp.DataBodyRange.Rows(insertPos).Cells(1, 2).value = funcArr(i, 1)
            If nbCols > 2 Then
                loTachesComp.DataBodyRange.Rows(insertPos).Cells(1, 3).Resize(1, nbCols - 2).value = ""
            End If
        Else
            loTachesComp.ListRows.Add
            loTachesComp.DataBodyRange.Rows(loTachesComp.ListRows.count).Cells(1, 1).value = taskName
            loTachesComp.DataBodyRange.Rows(loTachesComp.ListRows.count).Cells(1, 2).value = funcArr(i, 1)
'            If nbCols > 2 Then
'                loTachesComp.DataBodyRange.Rows(loTachesComp.ListRows.count).Cells(1, 3).Resize(1, nbCols - 2).value = ""
'            End If
        End If
    Next i

    loTachesComp.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")

    StopCode
    MsgBox "Nouvelle tâche ajoutée.", vbInformation
End Sub

Public Sub SuppTaskCompetences()
    Dim wsSettings As Worksheet, wsTaches As Worksheet
    Dim loTask As ListObject, loTachesComp As ListObject
    Dim selectedRange As Range, cell As Range
    Dim taskNames As Collection
    Dim filterWasOn As Boolean
    Dim i As Long, idx As Long

    StartCode
    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set loTask = wsSettings.ListObjects("TabTask")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")

    Set selectedRange = Application.Intersect(Selection, loTask.DataBodyRange.Columns(1))
    If selectedRange Is Nothing Then
        MsgBox "Sélectionnez une ou plusieurs cellules de tâches à supprimer.", vbExclamation
        Exit Sub
    End If

    Set taskNames = New Collection
    For Each cell In selectedRange
        On Error Resume Next
        taskNames.Add cell.value, CStr(cell.value)
        On Error GoTo 0
    Next cell

    If taskNames.count = 0 Then Exit Sub

    If MsgBox("Supprimer " & taskNames.count & " tâche(s)?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' Supprimer le filtre si actif
    filterWasOn = loTachesComp.AutoFilter.FilterMode
    If filterWasOn Then loTachesComp.AutoFilter.ShowAllData

    ' Supprimer dans TabTask (parcours à l'envers)
    For idx = loTask.ListRows.count To 1 Step -1
        For i = 1 To taskNames.count
            If loTask.DataBodyRange.Cells(idx, 1).value = taskNames(i) Then
                loTask.ListRows(idx).Delete
                Exit For
            End If
        Next i
    Next idx

    ' Supprimer dans TabTachesCompetences (parcours à l'envers)
    For idx = loTachesComp.ListRows.count To 1 Step -1
        For i = 1 To taskNames.count
            If loTachesComp.DataBodyRange.Cells(idx, 1).value = taskNames(i) Then
                loTachesComp.ListRows(idx).Delete
                Exit For
            End If
        Next i
    Next idx
    loTachesComp.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")

    StopCode
    MsgBox "Tâche(s) supprimée(s)", vbInformation
End Sub

Public Sub multiNewTasksToCompetences()
    Dim wsSettings As Worksheet, wsTaches As Worksheet
    Dim loTask As ListObject, loFonction As ListObject, loTachesComp As ListObject
    Dim funcArr As Variant, taskArr As Variant
    Dim i As Long, j As Long, nbCols As Long
    Dim lastTaskRow As Long, firstNewRow As Long, lastNewRow As Long
    Dim newRow As ListRow

    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set loTask = wsSettings.ListObjects("TabTask")
    Set loFonction = wsSettings.ListObjects("TabFonction")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")

    ' Détecter la plage des nouvelles tâches collées (sélection)
    If Application.Selection Is Nothing Then
        MsgBox "Sélectionnez les nouvelles tâches ajoutées dans TabTask.", vbExclamation
        Exit Sub
    End If

    If Intersect(Application.Selection, loTask.DataBodyRange.Columns(1)) Is Nothing Then
        MsgBox "Sélectionnez les nouvelles tâches dans la colonne des tâches de TabTask.", vbExclamation
        Exit Sub
    End If

    firstNewRow = Application.Selection.Row - loTask.DataBodyRange.Rows(1).Row + 1
    lastNewRow = firstNewRow + Application.Selection.Rows.count - 1

    ' Récupérer toutes les fonctions existantes
    funcArr = loFonction.DataBodyRange.value
    nbCols = loTachesComp.ListColumns.count

    ' Pour chaque nouvelle tâche sélectionnée
    For i = firstNewRow To lastNewRow
        Dim taskName As String
        taskName = loTask.DataBodyRange.Cells(i, 1).value
        If Trim(taskName) <> "" Then
            ' Pour chaque fonction, ajouter la tâche dans TabTachesCompetences
            For j = 1 To UBound(funcArr, 1)
                Set newRow = loTachesComp.ListRows.Add
                newRow.Range(1, 1).value = taskName
                newRow.Range(1, 2).value = funcArr(j, 1)
                If nbCols > 2 Then
                    newRow.Range(1, 3).Resize(1, nbCols - 2).value = ""
                End If
            Next j
        End If
    Next i
    loTachesComp.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")
    MsgBox "Tâches ajoutées à TabTachesCompetences pour toutes les fonctions.", vbInformation
End Sub


Public Sub OldAddNewFunction()
    Dim wsSettings As Worksheet, wsTaches As Worksheet
    Dim loFonction As ListObject, loTask As ListObject, loTachesComp As ListObject
    Dim funcName As String
    Dim taskArr As Variant
    Dim nbTasks As Long, nbCols As Long
    Dim i As Long, newRow As ListRow
    Dim filterWasOn As Boolean

    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set loFonction = wsSettings.ListObjects("TabFonction")
    Set loTask = wsSettings.ListObjects("TabTask")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")

    funcName = InputBox("Nom de la nouvelle fonction à ajouter :", "Ajouter une fonction")
    If Trim(funcName) = "" Then Exit Sub

    ' Supprimer le filtre si actif
    filterWasOn = loTachesComp.AutoFilter.FilterMode
    If filterWasOn Then loTachesComp.AutoFilter.ShowAllData

    ' Ajouter la fonction à TabFonction
    loFonction.ListRows.Add
    loFonction.DataBodyRange.Cells(loFonction.ListRows.count, 1).value = funcName

    ' Ajouter pour chaque tâche dans TabTachesCompetences
    taskArr = loTask.DataBodyRange.value
    nbTasks = UBound(taskArr, 1)
    nbCols = loTachesComp.ListColumns.count

    For i = 1 To nbTasks
        Set newRow = loTachesComp.ListRows.Add
        newRow.Range(1, 1).value = taskArr(i, 1)      ' Libellé tâche
        newRow.Range(1, 2).value = funcName           ' Fonction
'        If nbCols > 2 Then
'            newRow.Range(1, 3).Resize(1, nbCols - 2).value = ""
'        End If
    Next i

    MsgBox "Fonction ajoutée", vbInformation
End Sub

Public Sub OldDeleteFunction()
    Dim wsSettings As Worksheet, wsTaches As Worksheet
    Dim loFonction As ListObject, loTachesComp As ListObject
    Dim selectedCell As Range
    Dim funcName As String
    Dim i As Long
    Dim filterWasOn As Boolean
    
    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set loFonction = wsSettings.ListObjects("TabFonction")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")

    ' Supprimer le filtre si actif
    filterWasOn = loTachesComp.AutoFilter.FilterMode
    If filterWasOn Then loTachesComp.AutoFilter.ShowAllData

    Set selectedCell = Application.Selection
    If Intersect(selectedCell, loFonction.DataBodyRange.Columns(1)) Is Nothing Then
        MsgBox "Sélectionnez une cellule de la fonction à supprimer dans TabFonction.", vbExclamation
        Exit Sub
    End If

    funcName = selectedCell.value
    If MsgBox("Supprimer la fonction '" & funcName & "' dans TabFonction et TabTachesCompetences ?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' Supprimer la fonction dans TabFonction
    loFonction.DataBodyRange.Rows(selectedCell.Row - loFonction.DataBodyRange.Rows(1).Row + 1).Delete

    ' Supprimer toutes les lignes associées dans TabTachesCompetences
    For i = loTachesComp.ListRows.count To 1 Step -1
        If loTachesComp.DataBodyRange.Cells(i, 2).value = funcName Then
            loTachesComp.ListRows(i).Delete
        End If
    Next i
    loTachesComp.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")
    MsgBox "Fonction supprimée.", vbInformation
End Sub

Public Sub AddNewFunction()
    Dim wsSettings As Worksheet, wsTaches As Worksheet, wsRess As Worksheet
    Dim loFonction As ListObject, loTask As ListObject, loTachesComp As ListObject, loRessComp As ListObject, loRessCompAttendu As ListObject
    Dim funcName As String
    Dim taskArr As Variant
    Dim nbTasks As Long, nbCols As Long, nbColsRess As Long
    Dim i As Long, newRow As ListRow
    Dim filterWasOn As Boolean

    StartCode
    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set wsRess = ShMatRessComp
    Set loFonction = wsSettings.ListObjects("TabFonction")
    Set loTask = wsSettings.ListObjects("TabTask")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")
    Set loRessComp = wsRess.ListObjects("TabRessComp")
    Set loRessCompAttendu = wsRess.ListObjects("TabRessCompAttendu")

    funcName = InputBox("Nom de la nouvelle fonction à ajouter :", "Ajouter une fonction")
    If Trim(funcName) = "" Then Exit Sub
    ' Supprimer le filtre si actif
    filterWasOn = loTachesComp.AutoFilter.FilterMode
    If filterWasOn Then loTachesComp.AutoFilter.ShowAllData
    ' Ajouter la fonction à TabFonction
    loFonction.ListRows.Add
    loFonction.DataBodyRange.Cells(loFonction.ListRows.count, 1).value = funcName

    ' Ajouter pour chaque tâche dans TabTachesCompetences
    taskArr = loTask.DataBodyRange.value
    nbTasks = UBound(taskArr, 1)
    nbCols = loTachesComp.ListColumns.count

    For i = 1 To nbTasks
        Set newRow = loTachesComp.ListRows.Add
        newRow.Range(1, 1).value = taskArr(i, 1)      ' Libellé tâche
        newRow.Range(1, 2).value = funcName           ' Fonction
        If nbCols > 2 Then
            newRow.Range(1, 3).Resize(1, nbCols - 2).value = ""
        End If
    Next i

    ' Ajouter la fonction à TabRessComp (en dernière ligne)
    nbColsRess = loRessComp.ListColumns.count
    Set newRow = loRessComp.ListRows.Add
    newRow.Range(1, 2).value = funcName
    
    nbColsRess = loRessCompAttendu.ListColumns.count
    Set newRow = loRessCompAttendu.ListRows.Add
    newRow.Range(1, 2).value = funcName
    
'    If nbColsRess > 1 Then
'        newRow.Range(1, 2).Resize(1, nbColsRess - 1).value = ""
'    End If
    loTachesComp.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")
    StopCode
    MsgBox "Fonction ajoutée", vbInformation
End Sub

Public Sub DeleteFunction()
    Dim wsSettings As Worksheet, wsTaches As Worksheet, wsRess As Worksheet
    Dim loFonction As ListObject, loTachesComp As ListObject, loRessComp As ListObject, loRessCompAttendu As ListObject
    Dim selectedCell As Range
    Dim funcName As String
    Dim i As Long
    Dim filterWasOn As Boolean
    
    StartCode
    Set wsSettings = ShSettings
    Set wsTaches = ShMatTachesComp
    Set wsRess = ShMatRessComp
    Set loFonction = wsSettings.ListObjects("TabFonction")
    Set loTachesComp = wsTaches.ListObjects("TabTachesCompetences")
    Set loRessComp = wsRess.ListObjects("TabRessComp")
    Set loRessCompAttendu = wsRess.ListObjects("TabRessCompAttendu")

    Set selectedCell = Application.Selection
    If Intersect(selectedCell, loFonction.DataBodyRange.Columns(1)) Is Nothing Then
        MsgBox "Sélectionnez une cellule de la fonction à supprimer dans TabFonction.", vbExclamation
        Exit Sub
    End If
    ' Supprimer le filtre si actif
    filterWasOn = loTachesComp.AutoFilter.FilterMode
    If filterWasOn Then loTachesComp.AutoFilter.ShowAllData
    funcName = selectedCell.value
    If MsgBox("Supprimer la fonction '" & funcName & "' ?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' Supprimer la fonction dans TabFonction
    loFonction.DataBodyRange.Rows(selectedCell.Row - loFonction.DataBodyRange.Rows(1).Row + 1).Delete

    ' Supprimer toutes les lignes associées dans TabTachesCompetences
    For i = loTachesComp.ListRows.count To 1 Step -1
        If loTachesComp.DataBodyRange.Cells(i, 2).value = funcName Then
            loTachesComp.ListRows(i).Delete
        End If
    Next i

    ' Supprimer la fonction dans TabRessComp
    For i = loRessComp.ListRows.count To 1 Step -1
        If loRessComp.DataBodyRange.Cells(i, 2).value = funcName Then
            loRessComp.ListRows(i).Delete
        End If
    Next i
    ' Supprimer la fonction dans TabRessCompAttendu
    For i = loRessCompAttendu.ListRows.count To 1 Step -1
        If loRessCompAttendu.DataBodyRange.Cells(i, 2).value = funcName Then
            loRessCompAttendu.ListRows(i).Delete
        End If
    Next i
    loTachesComp.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")
    StopCode
    MsgBox "Fonction supprimée", vbInformation
End Sub


' --- Module: ShMatRessComp.cls ---
Attribute VB_Name = "ShMatRessComp"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



' --- Module: MatTachesComp.bas ---
Attribute VB_Name = "MatTachesComp"
Option Explicit

Public Sub AddCompetence()
    Dim tabTacheComp As ListObject
    Dim currentCol As Long
    Dim compName As String
    Dim e As Range
    Dim data As Range
    Dim tabRessComp As ListObject
    Dim tabRessCompAttendu As ListObject
    
    StartCode
    Set tabTacheComp = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set tabRessComp = ShMatRessComp.ListObjects("TabRessComp")
    Set tabRessCompAttendu = ShMatRessComp.ListObjects("TabRessCompAttendu")
    compName = InputBox("Merci de taper le nom de la compétence à ajouter")
    If compName = vbNullString Then Exit Sub
    For Each e In tabTacheComp.HeaderRowRange
        If compName = e Then
            MsgBox ("La compétence existe dans la matrice")
            Exit Sub
        End If
    Next e
    If Not ActiveCell.ListObject Is Nothing Then
        currentCol = ActiveCell.Row - ShSettings.ListObjects("TabComp").DataBodyRange.Row + 3 'ActiveCell.Column - tabTacheComp.HeaderRowRange.Column + 1
        If currentCol > 2 And currentCol <= tabTacheComp.ListColumns.count Then
            tabTacheComp.ListColumns.Add (currentCol)
            tabTacheComp.HeaderRowRange(, currentCol) = compName
        Else
            tabTacheComp.ListColumns.Add
            tabTacheComp.HeaderRowRange(, tabTacheComp.HeaderRowRange.count) = compName
            currentCol = tabTacheComp.ListColumns.count
        End If
    Else
        tabTacheComp.ListColumns.Add
        tabTacheComp.HeaderRowRange(, tabTacheComp.HeaderRowRange.count) = compName
        currentCol = tabTacheComp.ListColumns.count
    End If
    Call AddCompetenceToMatCompRessources(currentCol, compName, tabRessComp)
    Call AddCompetenceToMatCompRessources(currentCol, compName, tabRessCompAttendu)
    Call AddCompetenceToCompTab(currentCol - 2, compName)
    If tabTacheComp.ListColumns.count > 2 Then
        Set data = tabTacheComp.ListColumns(3).DataBodyRange.Resize(, tabTacheComp.ListColumns.count - 1)
        Call AddDataValidation(data)
        Call SetCondFormatColorScale(data)
    End If
    StopCode
End Sub

Public Sub DeleteCompetence()
    Dim tabTacheComp As ListObject, tabComp As ListObject
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim currentCol As Long
    Dim selectedRange As Range
    
    StartCode
    Set tabTacheComp = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set tabComp = ShSettings.ListObjects("TabComp")
    
    Set selectedRange = Intersect(Selection, Selection.Rows(1))
    If tabComp.DataBodyRange.count > 0 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de " & selectedRange.count & " compétence(s) ?", vbOKCancel)
        If MsgAnswer = vbOK Then
            For count = selectedRange.count To 1 Step -1
                currentCol = selectedRange(count).Row - tabComp.HeaderRowRange.Row
                'If tabTacheComp.ListColumns.count > 2 Then
                    'If Not Intersect(SelectedRange(1, count), tabTacheComp.HeaderRowRange) Is Nothing Then
                        ShMatRessComp.ListObjects("TabRessComp").ListColumns(currentCol + 2).Delete
                        ShMatRessComp.ListObjects("TabRessCompAttendu").ListColumns(currentCol + 2).Delete
                        tabTacheComp.ListColumns(currentCol + 2).Delete
                        ShSettings.ListObjects("TabComp").ListRows(currentCol).Delete
                    'End If
                'End If
            Next count
        End If
    End If
    StopCode
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
        Case "BtnMatTacheComp": ShMatTachesComp.Activate
        Case "BtnMatRessComp": ShMatRessComp.Activate
        Case "BtnMatTacheRess": ShComparaison.Activate
        Case "BtnSettings": ShSettings.Activate
    End Select
End Sub

Public Sub DeleteRow()
    Dim myTable As ListObject
    Dim MyCell As Range
    Dim MsgAnswer As VbMsgBoxResult
    Dim count As Long
    Dim selectedRange As Range
    
    StartCode
    Application.Calculation = xlCalculationManual
    Set MyCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(2, 0)
    Set myTable = ActiveSheet.ListObjects(MyCell.ListObject.Name)
    
    Set selectedRange = Intersect(Selection, Selection.Columns(1))
    If myTable.ListRows.count > 0 Then
        MsgAnswer = MsgBox("Confirmez-vous la suppression de " & selectedRange.count & " ligne(s) ?", vbOKCancel)
        If MsgAnswer = vbOK Then
            For count = selectedRange.count To 1 Step -1
                If myTable.ListRows.count > 0 Then
                    If Not Intersect(selectedRange(count, 1), myTable.DataBodyRange) Is Nothing Then
                           myTable.ListRows(selectedRange(count).Row - myTable.DataBodyRange.Row + 1).Delete
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
    
    If ActiveCell.ListObject Is Nothing Or myTable.ListRows.count = 0 Or ActiveCell.Row - myTable.DataBodyRange.Row + 1 < 1 Then
        myTable.ListRows.Add
    ElseIf ActiveCell.ListObject.Name = myTable.Name Then
        myTable.ListRows.Add ActiveCell.Row - myTable.DataBodyRange.Row + 1
    End If
    StopCode
End Sub

Sub SetCondFormatColorScale(data As Range)
    data.FormatConditions.Delete
    data.HorizontalAlignment = xlCenter
    With data
        .FormatConditions.AddColorScale ColorScaleType:=2
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).ColorScaleCriteria(1)
            .Type = xlConditionValueLowestValue
            .FormatColor.Color = 10285055
            .FormatColor.TintAndShade = 0
        End With
        With .FormatConditions(1).ColorScaleCriteria(2)
            .Type = xlConditionValueHighestValue
            .FormatColor.Color = 8109667
            .FormatColor.TintAndShade = 0
        End With

    End With
End Sub

Public Sub AddDataValidation(data As Range)
    Dim formula As String
    Dim rng As Range
    
    'If data.ListObject.Name = "TabTachesCompetences" Then
        formula = "=INDIRECT(" & ShSettings.Name & "!$D$8 & ""[Niveau]"")"
    'Else
    '    formula = "=INDIRECT(D$14 & ""[Niveau]"")"
    'End If
    With data.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub



' --- Module: MatTachesRess.bas ---
Attribute VB_Name = "MatTachesRess"
Option Explicit

Private Function SumMinRanges(R1 As Range, R2 As Range) As Long
    Dim i As Integer
    Dim j As Integer
    
    SumMinRanges = 0
    j = 1
    For i = 1 To R1.count Step 1
        SumMinRanges = SumMinRanges + WorksheetFunction.Min(IIf(R1(i) = vbNullString, 0, R1(i)), IIf(R2(j) = vbNullString, 0, R2(j)))
        j = j + 2
    Next i
End Function

Private Sub FillMatriceTachesRessources()
    Dim tabTachComp As ListObject
    Dim tabRessComp As ListObject
    Dim tabTachRess As ListObject
    Dim e As Range
    Dim currentRow As Long
    Dim currentCol As Long
    
    Set tabTachComp = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set tabRessComp = ShMatRessComp.ListObjects("TabRessComp")
    Set tabTachRess = ShMatTachesRess.ListObjects("TabTachesRessources")
    If tabTachComp.ListRows.count > 0 And tabTachComp.ListColumns.count > 2 And tabRessComp.ListRows.count > 0 Then
        For Each e In tabTachRess.DataBodyRange
            If e.Column > 3 Then
                currentRow = e.Row - tabTachRess.HeaderRowRange.Row
                currentCol = e.Column - tabTachRess.HeaderRowRange.Column - 1
                e = SumMinRanges(tabTachComp.ListRows(currentRow).Range.Resize(1, tabTachComp.ListColumns.count - 2).Offset(, 2), tabRessComp.ListRows(currentCol).Range.Resize(1, tabRessComp.ListColumns.count - 1).Offset(, 1))
            End If
        Next e
    End If
End Sub

Private Sub ResetMatriceTachesRessources()
    Dim tabTachComp As ListObject
    Dim tabRessComp As ListObject
    Dim tabTachRess As ListObject
    
    Set tabTachComp = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set tabRessComp = ShMatRessComp.ListObjects("TabRessComp")
    Set tabTachRess = ShMatTachesRess.ListObjects("TabTachesRessources")
    
    tabTachRess.DataBodyRange.ClearContents
    If tabTachRess.ListColumns.count > 3 Then
        tabTachRess.ListColumns(3).Range.Resize(, tabTachRess.ListColumns.count - 2).Delete
    End If
    
    If tabRessComp.ListRows.count > 0 Then
        tabTachRess.Resize tabTachRess.Range.Resize(, tabRessComp.ListRows.count + 2)
        tabTachRess.HeaderRowRange(1, 3).Resize(1, tabRessComp.ListRows.count) = Application.Transpose(tabRessComp.ListColumns(1).DataBodyRange)
    End If
    
    If tabTachComp.ListRows.count > 0 Then
        tabTachRess.Resize tabTachRess.Range.Resize(tabTachComp.ListRows.count + 1)
        tabTachComp.ListColumns(1).DataBodyRange.Resize(, 2).Copy
        tabTachRess.Range(2, 1).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        Range("D11").Activate
    End If
    
    If tabTachRess.ListRows.count > 0 And tabTachRess.ListColumns.count > 2 Then
        Call SetCondFormatColorScale(tabTachRess.DataBodyRange.Resize(tabTachRess.ListRows.count, tabTachRess.ListColumns.count - 2).Offset(, 2))
    End If
End Sub

Public Sub GetMatriceTachesRessources()
    StartCode
    Application.Calculation = xlCalculationManual
    Call ResetMatriceTachesRessources
    Call FillMatriceTachesRessources
    Application.Calculation = xlCalculationAutomatic
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



' --- Module: Comparaison.bas ---
Attribute VB_Name = "Comparaison"
Option Explicit

Public Sub clearData()
    With ShComparaison
        .Range("B16:CR200").Clear
        .Range("C13:CR13").ClearFormats
        .Range("C14:CR15").Clear
    End With


End Sub

Public Sub getData()
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim destRow As Long
    Dim r As Range

    Set tbl = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set col = tbl.ListColumns(1)
    destRow = 16

    For Each r In col.DataBodyRange.Rows
        If Not r.EntireRow.Hidden Then
            ShComparaison.Cells(destRow, "B").value = r.value
            destRow = destRow + 1
        End If
    Next r

End Sub




Public Sub fillHeader()
    On Error GoTo ErrorHandler ' Gestion des erreurs
    Dim countFunction As Integer
    Dim startRow As Integer
    Dim startCol As Integer
    Dim headerData As Variant
    Dim i As Integer
    Dim j As Integer
    Dim lastRow As Long
    
    startRow = 14
    startCol = 3
    ' Vérifie que la table TabFonction existe
    If Not ShSettings Is Nothing And Not ShSettings.ListObjects("TabFonction") Is Nothing Then
        countFunction = ShSettings.ListObjects("TabFonction").DataBodyRange.count
    Else
        MsgBox "La table TabFonction est introuvable.", vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Charge les données de HeaderComparaison dans un tableau
    headerData = ShBackEnd.Range("HeaderComparaison").value
    
    ' Remplit les en-têtes en utilisant le tableau HeaderComparaison
    For i = 0 To countFunction - 1
        ' Fusionne les cellules de la première colonne pour chaque fonction
        With ShComparaison.Range(ShComparaison.Cells(startRow, startCol + i * 3), ShComparaison.Cells(startRow, startCol + i * 3 + 2))
            .Merge
            .value = ShSettings.ListObjects("TabFonction").DataBodyRange(i + 1).value
            .HorizontalAlignment = xlCenter ' Centrage horizontal
            .VerticalAlignment = xlCenter   ' Centrage vertical
        End With
        
        ' Remplit les autres colonnes avec les données de HeaderComparaison
        For j = 1 To 3
            ShComparaison.Cells(startRow + 1, startCol + i * 3 + j - 1).value = headerData(2, j) ' Ligne 2 du tableau
        Next j
    Next i
    
    lastRow = ShComparaison.Cells(ShComparaison.Rows.count, "B").End(xlUp).Row
    ShBackEnd.Range("FormatCellCmp").Copy
    ShComparaison.Range(ShComparaison.Cells(startRow + 2, startCol - 1), ShComparaison.Cells(lastRow, startCol + countFunction * 3 - 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ShBackEnd.Range("HeaderComparaisonFormat").Copy
    ShComparaison.Range(ShComparaison.Cells(startRow - 1, startCol), ShComparaison.Cells(startRow + 1, startCol + countFunction * 3 - 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    

    
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite dans fillHeader : " & Err.Description, vbExclamation, "Erreur"
End Sub



Public Sub mainComparaison()
    StartCode
    clearData
    getData
    fillHeader
    Range("C16").Activate
    RemplirTableauComparaison
    RemplirPlanAccompagnement
    StopCode
End Sub

Sub RemplirTableauComparaisonOld()
    Dim loTaches As ListObject, loRess As ListObject, loComp As ListObject, loFonction As ListObject
    Dim iFunc As Long, iComp As Long
    Dim nbFonctions As Long, nbComps As Long
    Dim coeffs() As Double
    Dim niveauAttendu As Double, niveauReel As Double
    Dim sommeCoeff As Double
    Dim startCell As Range
    Dim destRow As Long
    Dim r As Range
    Dim fonctionNom As String
    Dim iTask As Long

    'StartCode
    Set loTaches = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set loRess = ShMatRessComp.ListObjects("TabRessComp")
    Set loComp = ShSettings.ListObjects("TabComp")
    Set loFonction = ShSettings.ListObjects("TabFonction")
    Set startCell = ShComparaison.Range("C16")

    nbFonctions = loFonction.ListRows.count
    nbComps = loComp.ListRows.count

    ' Récupérer les coefficients dans un tableau
    ReDim coeffs(1 To nbComps)
    For iComp = 1 To nbComps
        coeffs(iComp) = loComp.DataBodyRange(iComp, 2).value
    Next iComp

    destRow = 0
    For iFunc = 1 To nbFonctions
        fonctionNom = loFonction.DataBodyRange(iFunc, 1).value

        ' Filtrer TabTachesCompetences sur la fonction en cours
        loTaches.Range.AutoFilter Field:=3, Criteria1:=fonctionNom ' 2 = colonne "Fonction"

        destRow = 0 ' <-- Réinitialisation ici

        For Each r In loTaches.DataBodyRange.Rows
            If Not r.EntireRow.Hidden Then
                ' Libellé tâche
                'startCell.Offset(destRow, -1).value = r.Cells(2).value

                ' Calcul du niveau attendu (moyenne pondérée)
                niveauAttendu = 0
                sommeCoeff = 0
                For iComp = 1 To nbComps
                    niveauAttendu = niveauAttendu + r.Cells(iComp + 3).value * coeffs(iComp)
                    sommeCoeff = sommeCoeff + coeffs(iComp)
                Next iComp
                If sommeCoeff > 0 Then niveauAttendu = niveauAttendu / sommeCoeff

                ' Niveau réel (somme des niveaux réels pour la fonction)
                niveauReel = 0
                For iComp = 1 To nbComps
                    niveauReel = niveauReel + loRess.DataBodyRange(iFunc, iComp + 2).value
                Next iComp
                niveauReel = niveauReel / nbComps

                ' Remplir les colonnes dans la feuille destination
                Dim c As Long
                c = (iFunc - 1) * 3
                startCell.Offset(destRow, c).value = Round(niveauAttendu, 2)
                startCell.Offset(destRow, c + 1).value = Round(niveauReel, 2)
                startCell.Offset(destRow, c + 2).value = Round(niveauAttendu - niveauReel, 2)

                destRow = destRow + 1
            End If
        Next r
    Next iFunc
    'StopCode
     loTaches.Range.AutoFilter Field:=2, Criteria1:=ShMatTachesComp.Range("Filter_Fonction")
    'loTaches.AutoFilter.ShowAllData ' Retire le filtre à la fin
End Sub

Sub RemplirTableauComparaison()
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual

    Dim loTaches As ListObject, loRess As ListObject, loComp As ListObject, loFonction As ListObject
    Dim iFunc As Long, iComp As Long, nbFonctions As Long, nbComps As Long
    Dim coeffs() As Double, niveauAttendu As Double, niveauReel As Double, sommeCoeff As Double
    Dim startCell As Range, fonctionNom As String
    Dim tachesArr, ressArr, compArr, foncArr, resultArr()
    Dim iRow As Long, destRow As Long, c As Long

    Set loTaches = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set loRess = ShMatRessComp.ListObjects("TabRessComp")
    Set loComp = ShSettings.ListObjects("TabComp")
    Set loFonction = ShSettings.ListObjects("TabFonction")
    Set startCell = ShComparaison.Range("C16")

    tachesArr = loTaches.DataBodyRange.value
    ressArr = loRess.DataBodyRange.value
    compArr = loComp.DataBodyRange.value
    foncArr = loFonction.DataBodyRange.value

    nbFonctions = UBound(foncArr, 1)
    nbComps = UBound(compArr, 1)

    ' Récupérer les coefficients dans un tableau
    ReDim coeffs(1 To nbComps)
    For iComp = 1 To nbComps
        coeffs(iComp) = compArr(iComp, 2)
    Next iComp

    ' Préparer le tableau résultat (lignes = nb tâches visibles max, colonnes = nbFonctions*3)
    ReDim resultArr(1 To UBound(tachesArr, 1), 1 To nbFonctions * 3)

    For iFunc = 1 To nbFonctions
        fonctionNom = foncArr(iFunc, 1)
        destRow = 0
        For iRow = 1 To UBound(tachesArr, 1)
            If tachesArr(iRow, 2) = fonctionNom Then ' Colonne 3 = "Fonction"
                ' Calcul du niveau attendu
                niveauAttendu = 0
                sommeCoeff = 0
                For iComp = 1 To nbComps
                    niveauAttendu = niveauAttendu + tachesArr(iRow, iComp + 2) * coeffs(iComp)
                    sommeCoeff = sommeCoeff + coeffs(iComp)
                Next iComp
                If sommeCoeff > 0 Then niveauAttendu = niveauAttendu / sommeCoeff

                ' Niveau réel
                niveauReel = 0
                For iComp = 1 To nbComps
                    niveauReel = niveauReel + ressArr(iFunc, iComp + 2)
                Next iComp
                niveauReel = niveauReel / nbComps

                ' Remplir le tableau résultat
                c = (iFunc - 1) * 3 + 1
                resultArr(destRow + 1, c) = Round(niveauAttendu, 2)
                resultArr(destRow + 1, c + 1) = Round(niveauReel, 2)
                resultArr(destRow + 1, c + 2) = Round(niveauAttendu - niveauReel, 2)
                destRow = destRow + 1
            End If
        Next iRow
    Next iFunc

    ' Écrire le résultat d'un coup
    startCell.Resize(UBound(resultArr, 1), UBound(resultArr, 2)).value = resultArr

'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    Application.Calculation = xlCalculationAutomatic
End Sub


' --- Module: MatRessComp.bas ---
Attribute VB_Name = "MatRessComp"
Option Explicit

Public Sub AddCompetenceToMatCompRessources(ByVal Position As Integer, compName As String, tabRessComp As ListObject)
    'Dim tabRessComp As ListObject
    'Set tabRessComp = ShMatRessComp.ListObjects("TabRessComp")
    Dim e As Range
    Dim currentCol As Long
    

    tabRessComp.ListColumns.Add Position:=Position
    tabRessComp.HeaderRowRange(Position).value = compName
    Call FormatTabRessComp(tabRessComp)
End Sub


Public Sub AddCompetenceToCompTab(ByVal Position As Integer, compName As String)
    Dim tabComp As ListObject
    Dim e As Range
    Dim currentCol As Long
    
    Set tabComp = ShSettings.ListObjects("TabComp")
    tabComp.ListRows.Add Position:=Position
    tabComp.DataBodyRange(Position, 1).value = compName
End Sub

Public Sub UpdateTableRessComp(id As Long, newName As String)
    Dim tabRessComp As ListObject
    Dim currentCol As Long
    
    Set tabRessComp = ShMatRessComp.ListObjects("TabRessComp")
    currentCol = id + tabRessComp.HeaderRowRange.Column - 1
    With tabRessComp
        .HeaderRowRange(, id) = newName & " M"
        .HeaderRowRange(, id + 1) = newName & " I"
        ShMatRessComp.Cells(.HeaderRowRange.Row - 2, currentCol) = newName
    End With
End Sub

Public Sub FormatTabRessComp(tabRessComp As ListObject)
    'Dim tabRessComp As ListObject
    Dim rangeToFormat As Range
    Dim data As Range

    'Set tabRessComp = ShMatRessComp.ListObjects("TabRessComp")
    Set rangeToFormat = Intersect(tabRessComp.HeaderRowRange, tabRessComp.HeaderRowRange.Offset(, 1))
    Set rangeToFormat = rangeToFormat.Offset(-1)
    If tabRessComp.ListColumns.count > 2 Then
        Set data = tabRessComp.ListColumns(3).DataBodyRange.Resize(, tabRessComp.ListColumns.count - 1)
        Call AddDataValidation(data)
        Call SetCondFormatColorScale(data)
    End If
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
    Dim lobj As ListObject
    Dim sh As Shape
    Dim originalText As String
    Dim sht As Worksheet
    StartCode
    Set dict = CreateLangDict
    

    With ShMatTachesComp
        trad = vbNullString
        For Each e In Union(.Range("D1:Z5"), .Range("B6"), .Range("D12"), .Range("B13:AG13"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabTachesCompetences").ListColumns(2).DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShMatRessComp
        trad = vbNullString
        For Each e In Union(.Range("D1:AD5"), .Range("C12"), .Range("B14"), .Range("B6"), .Range("C13:BE15"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    
    With ShMatTachesRess
        trad = vbNullString
        For Each e In Union(.Range("E1:U5"), .Range("C10:L10"), .Range("B6"), .Range("B10:C10"), .Range("D9"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each e In .ListObjects("TabTachesRessources").ListColumns(2).DataBodyRange
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With


    With ShSettings
        trad = vbNullString
        For Each e In Union(.Range("E1:K5"), .Range("C6"), .Range("D9:H9"), .Range("E11:H25"), .Range("D8"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
        trad = vbNullString
        For Each lobj In .ListObjects
            trad = dict(lobj.Name)
            If trad <> vbNullString Then
                lobj.Name = trad
            End If
        Next lobj
    End With

    With ShBackEnd
        trad = vbNullString
        For Each e In .Range("A2:B2")
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With
    With ShGuide
        trad = vbNullString
        For Each e In Union(.Range("E1:K5"), .Range("C6"))
            trad = dict(e.value)
            If e <> vbNullString And trad <> vbNullString Then
                e = trad
            End If
        Next e
    End With

'
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






' --- Module: ShComparaison.cls ---
Attribute VB_Name = "ShComparaison"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



Private Sub Worksheet_Activate()
    mainComparaison
End Sub



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
    Call refreshAll
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler ' Gestion des erreurs

    Dim e As Range

    StartCode
    
    For Each e In Target
        If Not Intersect(e, ShDashboard.Range("ProfilName")) Is Nothing Then
            Call refreshRadarRess
            Call RadarCombine
        ElseIf Not Intersect(e, ShDashboard.Range("P_Fonction")) Is Nothing Then
            Call refreshRadarFonction
            Call RadarCombine
        End If
    Next e

    StopCode
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbExclamation, "Erreur"
    StopCode
End Sub




' --- Module: ShPlanAcc.cls ---
Attribute VB_Name = "ShPlanAcc"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
    mainComparaison
End Sub



' --- Module: BackendRess.bas ---
Attribute VB_Name = "BackendRess"
Option Explicit

Sub SauvegarderDonneesAccompagnement(ressourceNom As String)
    Dim wsDest As Worksheet, wsData As Worksheet
    Dim i As Long, lastRow As Long
    Dim competence As String
    
    Set wsDest = ShComparaisonV2
    Set wsData = ShBackendRess
    
    If ressourceNom = "" Then Exit Sub
    
    ' Supprime les anciennes données de cette ressource
    Dim iLast As Long
    
    iLast = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row
    Dim iRow As Long
    For iRow = iLast To 2 Step -1
        If wsData.Cells(iRow, 1).value = ressourceNom Then
            wsData.Rows(iRow).Delete
        End If
    Next iRow
    
    ' Sauvegarde les nouvelles données
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row + 1
    
    For i = 16 To wsDest.Cells(wsDest.Rows.count, "B").End(xlUp).Row
        competence = wsDest.Cells(i, 2).value ' colonne B = compétence (à adapter si besoin)
        If competence <> "" Then
            wsData.Cells(lastRow, 1).value = ressourceNom
            wsData.Cells(lastRow, 2).value = competence
            wsData.Cells(lastRow, 3).value = wsDest.Cells(i, 6).value ' Plan
            wsData.Cells(lastRow, 4).value = wsDest.Cells(i, 7).value ' Responsable
            wsData.Cells(lastRow, 5).value = wsDest.Cells(i, 8).value ' Date cible
            lastRow = lastRow + 1
        End If
    Next i
End Sub

Sub ChargerDonneesAccompagnement()
    Dim wsDest As Worksheet, wsData As Worksheet
    Dim ressourceNom As String
    Dim i As Long, j As Long
    Dim dict As Object, competence As String
    
    Set wsDest = ShComparaisonV2
    Set wsData = ShBackendRess
    Set dict = CreateObject("Scripting.Dictionary")
    
    ressourceNom = wsDest.Range("RESSOURCE_SELECTED").value
    If ressourceNom = "" Then Exit Sub
    
    ' Charger les données de la ressource en mémoire
    Dim iLast As Long
    iLast = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row
    For i = 2 To iLast
        If wsData.Cells(i, 1).value = ressourceNom Then
            dict(wsData.Cells(i, 2).value) = Array(wsData.Cells(i, 3).value, _
                                                   wsData.Cells(i, 4).value, _
                                                   wsData.Cells(i, 5).value)
        End If
    Next i
    
    ' Injecter dans la feuille
    For j = 16 To wsDest.Cells(wsDest.Rows.count, "B").End(xlUp).Row
        competence = wsDest.Cells(j, 2).value
        If dict.exists(competence) Then
            wsDest.Cells(j, 6).value = dict(competence)(0)
            wsDest.Cells(j, 7).value = dict(competence)(1)
            wsDest.Cells(j, 8).value = dict(competence)(2)
        Else
            wsDest.Cells(j, 6).ClearContents
            wsDest.Cells(j, 7).ClearContents
            wsDest.Cells(j, 8).ClearContents
        End If
    Next j
End Sub




' --- Module: PlanAcc.bas ---
Attribute VB_Name = "PlanAcc"
Option Explicit


Public Sub RemplirPlanAccompagnement()
    Dim wsComp As Worksheet, wsPlan As Worksheet
    Dim loPlan As ListObject
    Dim lastRowComp As Long, lastColComp As Long
    Dim i As Long, j As Long, colEcart As Long
    Dim tache As String, fonction As String, ecart As Variant
    Dim existe As Boolean
    Dim planArr As Variant
    Dim colEtat As Long

    Set wsComp = ShComparaison
    Set wsPlan = ShPlanAcc
    Set loPlan = wsPlan.ListObjects("TabPlanAcc")

    'StartCode
    ' Détecter la dernière ligne et colonne utilisée dans Comparaison
    lastRowComp = wsComp.Cells(wsComp.Rows.count, 2).End(xlUp).Row
    lastColComp = wsComp.Cells(15, wsComp.Columns.count).End(xlToLeft).Column

    ' Trouver la colonne "Etat" dans TabPlanAcc
    colEtat = loPlan.ListColumns("Etat").Index

    ' Parcours des lignes de tâches (à partir de la première ligne de données)
    For i = 16 To lastRowComp
        tache = wsComp.Cells(i, 2).value ' Colonne C = Tâche
        If tache = "" Then Exit For

        ' Parcours des blocs de fonctions (3 colonnes par fonction)
        For j = 3 To lastColComp Step 3
            fonction = wsComp.Cells(14, j).value ' Ligne d'en-tête = Fonction
            colEcart = j + 2 ' Ecart est la 3ème colonne du bloc
            ecart = wsComp.Cells(i, colEcart).value

            If IsNumeric(ecart) Then
                existe = False
                If loPlan.ListRows.count > 0 Then
                    planArr = loPlan.DataBodyRange.value
                    Dim k As Long
                    For k = 1 To UBound(planArr, 1)
                        If planArr(k, 1) = tache And planArr(k, 3) = fonction Then
                            ' Mettre à jour l'écart et l'état
                            loPlan.DataBodyRange.Cells(k, 2).value = ecart
                            If ecart <= 0 Then
                                loPlan.DataBodyRange.Cells(k, colEtat).value = "Done"
                            Else
                                loPlan.DataBodyRange.Cells(k, colEtat).value = ""
                            End If
                            existe = True
                            Exit For
                        End If
                    Next k
                End If

                ' Ajouter la ligne si elle n'existe pas déjà et si l'écart est strictement positif
                If Not existe And ecart > 0 Then
                    With loPlan.ListRows.Add
                        .Range(1, 1).value = tache
                        .Range(1, 2).value = ecart
                        .Range(1, 3).value = fonction
                        .Range(1, colEtat).value = ""
                    End With
                End If
            End If
        Next j
    Next i

    'StopCode
    'MsgBox "Plan d'accompagnement mis à jour.", vbInformation
End Sub



' --- Module: helperCompetence.bas ---
Attribute VB_Name = "helperCompetence"
Option Explicit

Sub AjouterCompetence()
    Dim loComp As ListObject, loTaches As ListObject, loRess As ListObject
    Dim nomComp As String
    Dim selCell As Range, insertIndex As Long, colInsert As Long
    Dim lastColTaches As Long, lastColRess As Long
    Dim data As Range
    Dim i As Long

    StartCode

    Set loTaches = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set loRess = ShMatRessComp.ListObjects("TabRessComp")
    Set loComp = ShSettings.ListObjects("TabComp")

    Set selCell = Selection

    nomComp = InputBox("Nom de la nouvelle compétence à ajouter :", "Ajouter une compétence")
    If nomComp = "" Then Exit Sub

    ' Déterminer où insérer dans TabComp
    If Not Intersect(selCell, loComp.DataBodyRange) Is Nothing Then
        insertIndex = selCell.Row - loComp.DataBodyRange.Row + 1
        loComp.ListRows.Add insertIndex
        loComp.DataBodyRange(insertIndex, 1).value = nomComp
        loComp.DataBodyRange(insertIndex, 2).value = 1 ' Valeur par défaut du coefficient
    Else
        loComp.ListRows.Add
        loComp.DataBodyRange(loComp.ListRows.count, 1).value = nomComp
        loComp.DataBodyRange(loComp.ListRows.count, 2).value = 1
        insertIndex = loComp.ListRows.count
    End If

    ' Ajouter la colonne dans TabTachesCompetences (à partir de la colonne 4)
    colInsert = 4 + insertIndex - 1
    loTaches.ListColumns.Add colInsert
    loTaches.HeaderRowRange(colInsert).value = nomComp
'    For i = 1 To loTaches.ListRows.count
'        loTaches.DataBodyRange(i, colInsert).value = 0
'    Next i

    ' Ajouter la colonne dans TabRessComp (hors table, à droite du header, offset -2)
    lastColRess = ShMatRessComp.Cells(13, ShMatRessComp.Columns.count).End(xlToLeft).Column
    ShMatRessComp.Cells(13, lastColRess + 1).value = nomComp
    Call FormatTabRessComp
'    For i = 15 To ShMatRessComp.Cells(ShMatRessComp.Rows.count, 1).End(xlUp).Row
'        ShMatRessComp.Cells(i, lastColRess + 1).value = 0
'    Next i
    If loTaches.ListColumns.count > 3 Then
        Set data = loTaches.ListColumns(3).DataBodyRange.Resize(, loTaches.ListColumns.count - 2)
        Call AddDataValidation(data)
        Call SetCondFormatColorScale(data)
    End If
    
    'MsgBox "Compétence ajoutée avec succès !", vbInformation
    StopCode
End Sub


Sub SupprimerCompetence()
    Dim loComp As ListObject, loTaches As ListObject, loRess As ListObject
    Dim selCell As Range, delIndex As Long, colDel As Long
    Dim nomComp As String
    Dim lastColRess As Long, i As Long

    StartCode

    Set loTaches = ShMatTachesComp.ListObjects("TabTachesCompetences")
    Set loRess = ShMatRessComp.ListObjects("TabRessComp")
    Set loComp = ShSettings.ListObjects("TabComp")

    Set selCell = Selection

    ' Déterminer la compétence à supprimer
    If Not Intersect(selCell, loComp.DataBodyRange) Is Nothing Then
        delIndex = selCell.Row - loComp.DataBodyRange.Row + 1
        nomComp = loComp.DataBodyRange(delIndex, 1).value
        loComp.ListRows(delIndex).Delete
    Else
        MsgBox "Sélectionnez une compétence à supprimer dans TabComp.", vbExclamation
        Exit Sub
    End If

    ' Supprimer la colonne dans TabTachesCompetences (à partir de la colonne 4)
    For colDel = 4 To loTaches.ListColumns.count
        If loTaches.HeaderRowRange(colDel).value = nomComp Then
            loTaches.ListColumns(colDel).Delete
            Exit For
        End If
    Next colDel

    ' Supprimer la colonne dans TabRessComp (hors table, à droite du header, offset -2)
    lastColRess = ShMatRessComp.Cells(14, ShMatRessComp.Columns.count).End(xlToLeft).Column
    For colDel = 1 To lastColRess
        If ShMatRessComp.Cells(14, colDel).value = nomComp Then
            ShMatRessComp.Columns(colDel).Delete
            Exit For
        End If
    Next colDel

    MsgBox "Compétence supprimée avec succès !", vbInformation
End Sub


' --- Module: Dashboard.bas ---
Attribute VB_Name = "Dashboard"
Option Explicit

Public Sub refreshAll()
    Call refreshRadarRess
    Call refreshRadarFonction
    Call RadarCombine
End Sub
Public Sub refreshRadarFonction()
    Dim N_Fonction As Long
    Dim chartObj As ChartObject
    Dim seriesColl As Object
    Dim nbComp As Long
    Dim firstCol As String, lastCol As String

    On Error GoTo ErrHandler

    N_Fonction = ShDashboard.Range("N_Fonction").value

    ' Récupère le nombre de compétences depuis TabComp
    nbComp = ShSettings.ListObjects("TabComp").ListRows.count

    ' La première colonne des valeurs est D (colonne 4)
    firstCol = "D"
    ' Calcule la lettre de la dernière colonne à partir du nombre de compétences
    lastCol = Split(Cells(1, 4 + nbComp - 1).Address(True, False), "$")(0)

    Set chartObj = ShDashboard.ChartObjects("ChartFonction")
    Set seriesColl = chartObj.Chart.FullSeriesCollection

    If seriesColl.count = 0 Then
        MsgBox "Aucune série trouvée dans le graphique.", vbExclamation
        Exit Sub
    End If

    With seriesColl(1)
        .Name = ShDashboard.Range("P_Fonction").value
        .Values = "='" & ShBackEnd.Name & "'!$" & firstCol & "$" & 14 & ":$" & lastCol & "$" & 14
        .XValues = "='" & ShMatTachesComp.Name & "'!$" & firstCol & "$" & 13 & ":$" & lastCol & "$" & 13

    End With

    Exit Sub
ErrHandler:
    MsgBox "Erreur dans refreshRadarRess : " & Err.Description, vbExclamation
End Sub


Public Sub refreshRadarRess()
    Dim Nprofil As Long
    Dim chartObj As ChartObject
    Dim seriesColl As Object
    Dim nbComp As Long
    Dim firstCol As String, lastCol As String

    On Error GoTo ErrHandler

    Nprofil = ShDashboard.Range("N_PROFIL").value

    ' Récupère le nombre de compétences depuis TabComp
    nbComp = ShSettings.ListObjects("TabComp").ListRows.count

    ' La première colonne des valeurs est D (colonne 4)
    firstCol = "D"
    ' Calcule la lettre de la dernière colonne à partir du nombre de compétences
    lastCol = Split(Cells(1, 4 + nbComp - 1).Address(True, False), "$")(0)

    Set chartObj = ShDashboard.ChartObjects("Chart 7")
    Set seriesColl = chartObj.Chart.FullSeriesCollection

    If seriesColl.count = 0 Then
        MsgBox "Aucune série trouvée dans le graphique.", vbExclamation
        Exit Sub
    End If

    With seriesColl(1)
        .Name = "'" & ShMatRessComp.Name & "'!$B$" & Nprofil
        .Values = "='" & ShMatRessComp.Name & "'!$" & firstCol & "$" & Nprofil & ":$" & lastCol & "$" & Nprofil
        .XValues = "='" & ShMatRessComp.Name & "'!$" & firstCol & "$" & 14 & ":$" & lastCol & "$" & 14
    End With

    Exit Sub
ErrHandler:
    MsgBox "Erreur dans refreshRadarRess : " & Err.Description, vbExclamation
End Sub


Public Sub RadarCombine()
    Dim chartObj As ChartObject
    Dim chartRadar As Chart
    Dim nbComp As Long
    Dim firstCol As String, lastCol As String
    Dim N_Fonction As Long, Nprofil As Long
    Dim wsDash As Worksheet

    Set wsDash = ShDashboard

    ' Récupère le nombre de compétences depuis TabComp
    nbComp = ShSettings.ListObjects("TabComp").ListRows.count

    ' La première colonne des valeurs est D (colonne 4)
    firstCol = "D"
    lastCol = Split(Cells(1, 4 + nbComp - 1).Address(True, False), "$")(0)

    N_Fonction = wsDash.Range("N_Fonction").value
    Nprofil = wsDash.Range("N_PROFIL").value

    ' Crée ou récupère le graphique RadarComb
    On Error Resume Next
    Set chartObj = wsDash.ChartObjects("RadarComb")
    On Error GoTo 0

    If chartObj Is Nothing Then
        Set chartObj = wsDash.ChartObjects.Add(Left:=300, Top:=50, Width:=350, Height:=250)
        chartObj.Name = "RadarComb"
    End If

    Set chartRadar = chartObj.Chart
    chartRadar.ChartType = xlRadar

    ' Efface les anciennes séries
    Do While chartRadar.SeriesCollection.count > 0
        chartRadar.SeriesCollection(1).Delete
    Loop

    ' Ajoute la série Fonction
    chartRadar.SeriesCollection.NewSeries
    With chartRadar.SeriesCollection(1)
        .Name = "Fonction"
        .Values = "='" & ShBackEnd.Name & "'!$" & firstCol & "$" & 14 & ":$" & lastCol & "$" & 14
        .XValues = "='" & ShMatTachesComp.Name & "'!$" & firstCol & "$" & 13 & ":$" & lastCol & "$" & 13
    End With

    ' Ajoute la série Ressource
    chartRadar.SeriesCollection.NewSeries
    With chartRadar.SeriesCollection(2)
        .Name = "Ressource"
        .Values = "='" & ShMatRessComp.Name & "'!$" & firstCol & "$" & Nprofil & ":$" & lastCol & "$" & Nprofil
        .XValues = "='" & ShMatRessComp.Name & "'!$" & firstCol & "$" & 14 & ":$" & lastCol & "$" & 14
    End With

    chartRadar.HasLegend = True
    chartRadar.Legend.Position = xlLegendPositionBottom
End Sub


' --- Module: ShComparaisonV2.cls ---
Attribute VB_Name = "ShComparaisonV2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private PrevRessource As String

' --- On mémorise la valeur courante quand l'utilisateur sélectionne la cellule ---
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("RESSOURCE_SELECTED")) Is Nothing Then
        PrevRessource = Me.Range("RESSOURCE_SELECTED").value
    End If
End Sub

' --- Quand la valeur change (après sélection dans la liste) ---
'Private Sub Worksheet_Change(ByVal Target As Range)
'    If Not Intersect(Target, Me.Range("RESSOURCE_SELECTED")) Is Nothing Then
'        Application.EnableEvents = False
'        ' Sauvegarde sous le nom précédent (si existant)
'        If PrevRessource <> "" Then
'            SauvegarderDonneesAccompagnement PrevRessource
'        End If
'        ' Remplir le tableau pour la nouvelle ressource
'        RemplirTableauComparaisonV2
'        ' Charger les données d'accompagnement (si déjà sauvegardées) pour la nouvelle ressource
'        ChargerDonneesAccompagnement
'        ' Mettre à jour PrevRessource pour la prochaine fois
'        PrevRessource = Me.Range("RESSOURCE_SELECTED").value
'        Application.EnableEvents = True
'    End If
'End Sub


Private Sub Worksheet_Activate()
    mainComparaisonV2 (PrevRessource)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim e As Range

    StartCode
    For Each e In Target
        If Not Intersect(e, Range("RESSOURCE_SELECTED")) Is Nothing Then
            Call mainComparaisonV2(PrevRessource)
            PrevRessource = Me.Range("RESSOURCE_SELECTED").value
        End If
    Next e
    
    StopCode
End Sub


' --- Module: ComparaisonV2.bas ---
Attribute VB_Name = "ComparaisonV2"
Option Explicit

Public Sub mainComparaisonV2(PrevRessource As String)
    StartCode
    If PrevRessource <> "" Then
        Call SauvegarderDonneesAccompagnement(PrevRessource)
    End If
    clearDataV2
    getDataV2
    setFormatTableauComparaisonV2
    Range("RESSOURCE_SELECTED").Select
    RemplirTableauComparaisonV2
    ' Recharger les infos d’accompagnement
    Call ChargerDonneesAccompagnement
    StopCode
End Sub

Public Sub setFormatTableauComparaisonV2()
    Dim lastRow As Long
    
    With ShComparaisonV2
        lastRow = .Cells(.Rows.count, "B").End(xlUp).Row
        ShBackEnd.Range("FormatCellCmp").Copy
        .Range(.Cells(16, 2), .Cells(lastRow, 5)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        ShBackEnd.Range("FormatCellLibre").Copy
        .Range(.Cells(16, 6), .Cells(lastRow, 8)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    End With
    Application.CutCopyMode = False
End Sub


Public Sub clearDataV2()
    With ShComparaisonV2
        .Range("B16:M200").Clear
        '.Range("C14:CR15").Clear
    End With
End Sub

Public Sub getDataV2()
    Dim destRow As Long

    destRow = 16
    ShSettings.ListObjects("TabComp").ListColumns(1).DataBodyRange.Copy ShComparaisonV2.Cells(destRow, "B")
End Sub

Sub RemplirTableauComparaisonV2()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim loReel As ListObject
    Dim loAttendu As ListObject
    Dim ressourceNom As String
    Dim iRowReel As Long, iRowAttendu As Long
    Dim iCol As Long
    Dim competence As String
    Dim valReel As Variant, valAttendu As Variant
    Dim diff As Variant
    Dim ligneDest As Long
    
    ' Feuilles
    Set wsSource = ShMatRessComp
    Set wsDest = ShComparaisonV2
    
    ' Tables
    Set loReel = wsSource.ListObjects("TabRessComp")
    Set loAttendu = wsSource.ListObjects("TabRessCompAttendu")
    
    ' Nom de la ressource sélectionnée
    ressourceNom = wsDest.Range("RESSOURCE_SELECTED").value
    
    If ressourceNom = "" Then
        MsgBox "Veuillez sélectionner une ressource dans la liste déroulante.", vbExclamation
        Exit Sub
    End If
    
    ' Recherche de la ligne correspondante dans chaque table
    On Error Resume Next
    iRowReel = Application.Match(ressourceNom, loReel.ListColumns(1).DataBodyRange, 0)
    iRowAttendu = Application.Match(ressourceNom, loAttendu.ListColumns(1).DataBodyRange, 0)
    On Error GoTo 0
    
    If IsError(iRowReel) Or IsError(iRowAttendu) Or iRowReel = 0 Or iRowAttendu = 0 Then
        MsgBox "La ressource '" & ressourceNom & "' n'existe pas dans les tables.", vbCritical
        Exit Sub
    End If
    
    ' Nettoyer la zone de destination
    'wsDest.Range("A2:D100").ClearContents
    
    ' Titres
    'wsDest.Range("A1:D1").value = Array("Compétence", "Niveau attendu", "Niveau réel", "Écart")
    
    ' Parcourir les colonnes de compétences (à partir de la 3e colonne)
    ligneDest = 16
    For iCol = 3 To loReel.ListColumns.count
        competence = loReel.ListColumns(iCol).Name
        
        valReel = loReel.DataBodyRange.Cells(iRowReel, iCol).value
        valAttendu = loAttendu.DataBodyRange.Cells(iRowAttendu, iCol).value
        
        If IsNumeric(valReel) And IsNumeric(valAttendu) Then
            diff = valReel - valAttendu
        Else
            diff = ""
        End If
        
        'wsDest.Cells(ligneDest, 1).value = competence
        wsDest.Cells(ligneDest, 3).value = valAttendu
        wsDest.Cells(ligneDest, 4).value = valReel
        wsDest.Cells(ligneDest, 5).value = diff
        
        ligneDest = ligneDest + 1
    Next iCol
    
    'wsDest.Columns("A:D").AutoFit
    
   ' MsgBox "Comparaison terminée pour : " & ressourceNom & " ?", vbInformation

End Sub




' --- Module: ShBackendRess.cls ---
Attribute VB_Name = "ShBackendRess"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



