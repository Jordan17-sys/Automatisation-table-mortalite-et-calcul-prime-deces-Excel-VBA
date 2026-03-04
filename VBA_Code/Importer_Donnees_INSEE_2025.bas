Attribute VB_Name = "Importer_Donnees_INSEE_2025"
Sub Importer_Donnees_INSEE_2025()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim startTime As Double
    Dim i As Long
    Dim destRow As Long
    
    ' Démarrer le chronomètre
    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Sélectionner le fichier source
    filePath = Application.GetOpenFilename("Fichiers Excel (*.xlsx; *.xls), *.xlsx; *.xls", , "Selectionne le FICHIER 7")
    
    If filePath = "Faux" Or filePath = "False" Then
        MsgBox "Aucun fichier selectionne.", vbExclamation
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Ouvrir le fichier
    On Error GoTo ErreurOuverture
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    
    ' Chercher la feuille 2025
    On Error Resume Next
    Set wsSource = wbSource.Sheets("2025")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "La feuille 2025 est introuvable!", vbCritical
        wbSource.Close False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Préparer la feuille de destination
    Set wsDest = ThisWorkbook.Sheets("Données_Brutes")
    wsDest.Range("A2:J200").ClearContents
    
    ' Importer uniquement les âges 0 à 109
    destRow = 2
    
    For i = 6 To 115
        If IsNumeric(wsSource.Cells(i, 1).Value) Then
            wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, 1).Value
            
            ' ENSEMBLE
            wsDest.Cells(destRow, 2).Value = wsSource.Cells(i, 3).Value / 100000
            wsDest.Cells(destRow, 3).Value = wsSource.Cells(i, 2).Value
            wsDest.Cells(destRow, 4).Value = wsSource.Cells(i, 4).Value
            
            ' FEMMES
            wsDest.Cells(destRow, 5).Value = wsSource.Cells(i, 6).Value / 100000
            wsDest.Cells(destRow, 6).Value = wsSource.Cells(i, 5).Value
            wsDest.Cells(destRow, 7).Value = wsSource.Cells(i, 7).Value
            
            ' HOMMES
            wsDest.Cells(destRow, 8).Value = wsSource.Cells(i, 9).Value / 100000
            wsDest.Cells(destRow, 9).Value = wsSource.Cells(i, 8).Value
            wsDest.Cells(destRow, 10).Value = wsSource.Cells(i, 10).Value
            
            destRow = destRow + 1
        End If
    Next i
    
    ' Mise en forme
    With wsDest
        .Range("B2:B" & destRow - 1).NumberFormat = "0.00000"
        .Range("E2:E" & destRow - 1).NumberFormat = "0.00000"
        .Range("H2:H" & destRow - 1).NumberFormat = "0.00000"
        
        .Range("C2:C" & destRow - 1).NumberFormat = "#,##0"
        .Range("F2:F" & destRow - 1).NumberFormat = "#,##0"
        .Range("I2:I" & destRow - 1).NumberFormat = "#,##0"
        
        .Range("D2:D" & destRow - 1).NumberFormat = "0.00"
        .Range("G2:G" & destRow - 1).NumberFormat = "0.00"
        .Range("J2:J" & destRow - 1).NumberFormat = "0.00"
    End With
    
    ' Fermer le fichier source
    wbSource.Close False
    
    ' Réactiver l'affichage
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Message de succès SIMPLIFIÉ
    MsgBox "Import reussi en " & Format(Timer - startTime, "0.00") & " secondes!" & vbCrLf & _
           "Lignes importees : " & (destRow - 2) & " (ages 0 a 109)" & vbCrLf & vbCrLf & _
           "Verifications :" & vbCrLf & _
           "Age 0 : qx=" & Format(wsDest.Cells(2, 2).Value, "0.00000") & _
           " / lx=" & Format(wsDest.Cells(2, 3).Value, "#,##0") & _
           " / ex=" & Format(wsDest.Cells(2, 4).Value, "0.00"), _
           vbInformation, "MORTEX - Import termine"
    
    Exit Sub
    
ErreurOuverture:
    MsgBox "Erreur : " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


