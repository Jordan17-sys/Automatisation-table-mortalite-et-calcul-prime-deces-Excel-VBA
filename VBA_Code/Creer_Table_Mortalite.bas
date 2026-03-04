Attribute VB_Name = "Creer_Table_Mortalite"
Sub Creer_Table_Mortalite()
    ' ======================================
    ' MACRO : Créer la structure de la table de mortalité
    ' ======================================
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Table_Mortalité")
    
    ' Effacer les anciennes données
    ws.Cells.Clear
    
    ' ======================================
    ' 1. TITRE ET EN-TÊTES
    ' ======================================
    With ws
        ' Titre principal
        .Range("A1:H1").Merge
        .Range("A1").Value = "TABLE DE MORTALITE - FRANCE METROPOLITAINE 2025"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)
        
        ' En-têtes de colonnes
        .Range("A2").Value = "Age"
        .Range("B2").Value = "qx"
        .Range("C2").Value = "px"
        .Range("D2").Value = "lx"
        .Range("E2").Value = "dx"
        .Range("F2").Value = "Lx"
        .Range("G2").Value = "Tx"
        .Range("H2").Value = "ex"
        
        ' Mise en forme des en-têtes
        .Range("A2:H2").Font.Bold = True
        .Range("A2:H2").HorizontalAlignment = xlCenter
        .Range("A2:H2").Interior.Color = RGB(217, 225, 242)
        .Range("A2:H2").Borders.LineStyle = xlContinuous
        
        ' Largeur des colonnes
        .Columns("A:A").ColumnWidth = 8
        .Columns("B:H").ColumnWidth = 12
        
        ' Figer les volets
        .Range("A3").Select
        ActiveWindow.FreezePanes = True
    End With
    
    MsgBox "Structure créée avec succès !" & vbCrLf & _
           "Prochaine étape : Remplir les formules", vbInformation, "MORTEX"
End Sub


