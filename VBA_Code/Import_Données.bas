Attribute VB_Name = "Import_Données"
' ============================================
' Module : Import_Données
' Description : Importe les fichiers INSEE (6, 7, 8) dans la feuille Données_Brutes
' Date : 2025
' ============================================

Option Explicit

' Macro principale : Importer les données INSEE
Sub Importer_Donnees_INSEE()
    
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileDialog As fileDialog
    Dim lastRow As Long
    Dim startTime As Double
    
    ' Démarrer le chronomètre
    startTime = Timer
    
    ' Désactiver les mises à jour d'écran (pour accélérer)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Définir la feuille cible
    Set ws = ThisWorkbook.Sheets("Données_Brutes")
    
    ' Effacer les anciennes données (garde les en-têtes)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 2 Then
        ws.Range("A3:I" & lastRow).ClearContents
    End If
    
    ' Ouvrir la boîte de dialogue pour sélectionner les fichiers
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Sélectionner les fichiers INSEE (6, 7, 8)"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx; *.xls; *.csv"
        .AllowMultiSelect = True ' Permet de sélectionner plusieurs fichiers
        
        If .Show = -1 Then ' Si l'utilisateur clique sur OK
            
            Dim i As Integer
            Dim selectedFile As Variant
            
            ' Boucle sur chaque fichier sélectionné
            For i = 1 To .SelectedItems.Count
                selectedFile = .SelectedItems(i)
                
                ' Appeler la fonction d'importation
                Call Importer_Fichier(selectedFile, ws)
                
            Next i
            
            ' Message de succès
            MsgBox "? Importation terminée en " & Format(Timer - startTime, "0.00") & " secondes." & vbCrLf & _
                   "?? Nombre de lignes importées : " & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 2, _
                   vbInformation, "MORTEX - Import réussi"
            
        Else
            MsgBox "? Aucun fichier sélectionné.", vbExclamation, "MORTEX - Import annulé"
        End If
        
    End With
    
    ' Réactiver les mises à jour
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Logger l'action
    Call Logger_Action("Importation données INSEE", "Succès", Timer - startTime)
    
End Sub

' ============================================
' Fonction : Importer un fichier spécifique
' ============================================
Sub Importer_Fichier(filePath As String, ws As Worksheet)
    
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim dataRange As Range
    
    On Error GoTo ErrorHandler
    
    ' Ouvrir le fichier source
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1) ' Première feuille du fichier
    
    ' Trouver la dernière ligne avec données
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    ' Vérifier s'il y a des données
    If lastRowSource < 2 Then
        MsgBox "Le fichier " & wbSource.Name & " ne contient pas de données.", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    
    ' Copier les données (ignore la première ligne si c'est un en-tête)
    Set dataRange = wsSource.Range("A2:I" & lastRowSource) ' Ajuste selon ta structure
    
    ' Trouver la dernière ligne dans la feuille cible
    lastRowTarget = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Coller les données
    dataRange.Copy Destination:=ws.Range("A" & lastRowTarget)
    
    ' Ajouter la source dans la colonne I
    ws.Range("I" & lastRowTarget & ":I" & lastRowTarget + dataRange.Rows.Count - 1).Value = wbSource.Name
    
    ' Fermer le fichier source
    wbSource.Close False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'importation du fichier : " & filePath & vbCrLf & Err.Description, vbCritical
    If Not wbSource Is Nothing Then wbSource.Close False
    
End Sub

' ============================================
' Fonction : Logger les actions (optionnel)
' ============================================
Sub Logger_Action(action As String, statut As String, duree As Double)
    
    Dim wsLog As Worksheet
    Dim lastRow As Long
    
    Set wsLog = ThisWorkbook.Sheets("Logs")
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    wsLog.Cells(lastRow, 1).Value = Now ' Date et heure
    wsLog.Cells(lastRow, 2).Value = action
    wsLog.Cells(lastRow, 3).Value = statut
    wsLog.Cells(lastRow, 4).Value = Format(duree, "0.00") & " sec"
    
End Sub


