Attribute VB_Name = "Creation_Prime_Pure"
Sub CreerFeuillePrimePure()
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Prime_Pure")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Prime_Pure"
    Else
        ws.Cells.Clear
    End If
    
    Application.ScreenUpdating = False
    
    ' -------------------------------------------------------
    ' ZONE 1 : EN-TÊTE ET PARAMÈTRES
    ' -------------------------------------------------------
    
    ' Titre principal
    With ws.Range("A1:D1")
        .Merge
        .Value = "CALCULATEUR DE PRIME PURE - ASSURANCE DÉCÈS"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102) ' Bleu foncé
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ' Section Paramètres
    ws.Range("A3").Value = "PARAMÈTRES DE CALCUL"
    With ws.Range("A3:D3")
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) ' Jaune clair
        .HorizontalAlignment = xlLeft
    End With
    
    ' Labels des paramètres
    ws.Range("A5").Value = "Âge du souscripteur :"
    ws.Range("A6").Value = "Capital assuré (€) :"
    ws.Range("A7").Value = "Taux d'intérêt technique (%) :"
    ws.Range("A8").Value = "Durée du contrat (années) :"
    
    ' Valeurs par défaut
    ws.Range("B5").Value = 30
    ws.Range("B6").Value = 100000
    ws.Range("B7").Value = 0.02
    ws.Range("B8").Value = 30
    
    ' Unités
    ws.Range("C5").Value = "ans"
    ws.Range("C6").Value = "€"
    ws.Range("C7").Value = "%"
    ws.Range("C8").Value = "ans"
    
    ' Format des cellules de paramètres
    ws.Range("B5,B8").NumberFormat = "0"
    ws.Range("B6").NumberFormat = "#,##0 €"
    ws.Range("B7").NumberFormat = "0.00%"
    
    ' Mise en forme des paramètres
    With ws.Range("A5:C8")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.Range("B5:B8").Interior.Color = RGB(217, 225, 242) ' Bleu très clair
    ws.Range("A5:A8").Font.Bold = True
    
    ' Section Résultats
    ws.Range("A10").Value = "RÉSULTATS"
    With ws.Range("A10:D10")
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218) ' Vert clair
        .HorizontalAlignment = xlLeft
    End With
    
    ' Labels des résultats
    ws.Range("A12").Value = "Prime pure annuelle :"
    ws.Range("A13").Value = "Prime commerciale (+25%) :"
    ws.Range("A14").Value = "Coût total sur la durée :"
    
    ' Cellules de résultats (vides pour l'instant)
    ws.Range("B12:B14").NumberFormat = "#,##0.00 €"
    ws.Range("B12:B14").Interior.Color = RGB(255, 255, 204) ' Jaune très clair
    ws.Range("A12:A14").Font.Bold = True
    
    With ws.Range("A12:C14")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Unités
    ws.Range("C12:C14").Value = "€"
    
    ' -------------------------------------------------------
    ' ZONE 2 : TABLEAU PAR ÂGE
    ' -------------------------------------------------------
    
    ws.Range("F3").Value = "ÉVOLUTION DES PRIMES PAR ÂGE"
    With ws.Range("F3:I3")
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .HorizontalAlignment = xlLeft
    End With
    
    ' En-têtes du tableau
    ws.Range("F5").Value = "Âge"
    ws.Range("G5").Value = "Prime pure (€/an)"
    ws.Range("H5").Value = "Prime comm. (€/an)"
    ws.Range("I5").Value = "Coût total (€)"
    
    With ws.Range("F5:I5")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' -------------------------------------------------------
    ' ZONE 3 : GRAPHIQUE (Placeholder)
    ' -------------------------------------------------------
    
    ws.Range("K3").Value = "GRAPHIQUE : PRIME EN FONCTION DE L'ÂGE"
    With ws.Range("K3:O3")
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
        .HorizontalAlignment = xlLeft
    End With
    
    ' Message placeholder
    ws.Range("K5:O10").Merge
    ws.Range("K5").Value = "Le graphique sera généré automatiquement" & vbCrLf & _
                           "après le calcul du tableau des primes"
    ws.Range("K5").HorizontalAlignment = xlCenter
    ws.Range("K5").VerticalAlignment = xlCenter
    ws.Range("K5").WrapText = True
    ws.Range("K5").Interior.Color = RGB(242, 242, 242)
    
    ' -------------------------------------------------------
    ' AJUSTEMENTS FINAUX
    ' -------------------------------------------------------
    
    ' Largeur des colonnes
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 5
    ws.Columns("F:I").ColumnWidth = 18
    ws.Columns("K:O").ColumnWidth = 15
    
    ' Masquer les colonnes de calcul intermédiaire (Q:V)
    ws.Columns("Q:V").Hidden = True
    
    Application.ScreenUpdating = True
    
    MsgBox "Feuille Prime_Pure créée avec succès !" & vbCrLf & vbCrLf & _
           "Prochaine étape : Ajouter les boutons et les macros de calcul", _
           vbInformation, "Création terminée"
    
End Sub


Sub AjouterBoutons()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Prime_Pure")
    
    ' Supprimer les anciens boutons s'ils existent
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp
    
    ' -------------------------------------------------------
    ' BOUTON 1 : CALCULER LA PRIME (Zone Paramètres)
    ' -------------------------------------------------------
    
    Dim btn1 As Button
    Set btn1 = ws.Buttons.Add(ws.Range("A16").Left, ws.Range("A16").Top, 200, 30)
    
    With btn1
        .Caption = "CALCULER LA PRIME"
        .OnAction = "CalculerPrimePure"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' -------------------------------------------------------
    ' BOUTON 2 : GÉNÉRER TABLEAU COMPLET (Zone Tableau)
    ' -------------------------------------------------------
    
    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(ws.Range("F7").Left, ws.Range("F7").Top, 220, 30)
    
    With btn2
        .Caption = "GÉNÉRER TABLEAU COMPLET"
        .OnAction = "GenererTableauPrimes"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' -------------------------------------------------------
    ' BOUTON 3 : GÉNÉRER GRAPHIQUE (Zone Graphique)
    ' -------------------------------------------------------
    
    Dim btn3 As Button
    Set btn3 = ws.Buttons.Add(ws.Range("K12").Left, ws.Range("K12").Top, 200, 30)
    
    With btn3
        .Caption = "GÉNÉRER GRAPHIQUE"
        .OnAction = "GenererGraphiquePrimes"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    MsgBox "Boutons ajoutés avec succès !" & vbCrLf & vbCrLf & _
           "3 boutons créés :" & vbCrLf & _
           "Calculer la prime (pour 1 âge)" & vbCrLf & _
           "Générer tableau (pour tous les âges)" & vbCrLf & _
           "Générer graphique", _
           vbInformation, "Boutons créés"
    
End Sub
Function CalculerPrimeUnique(age As Integer, capital As Double, taux As Double, duree As Integer) As Double
    Dim wsTable As Worksheet: Set wsTable = ThisWorkbook.Sheets("Table_Mortalité")
    Dim v As Double: v = 1 / (1 + taux)

    Dim m0 As Variant, lx0 As Double
    m0 = Application.Match(age, wsTable.Range("A:A"), 0)
    If IsError(m0) Then CalculerPrimeUnique = 0: Exit Function
    lx0 = CDbl(wsTable.Cells(m0, 4).Value)

    Dim t As Integer, ageActuel As Integer, m As Variant, dx As Double
    Dim somme As Double: somme = 0

    For t = 0 To duree - 1
        ageActuel = age + t
        m = Application.Match(ageActuel, wsTable.Range("A:A"), 0)
        If IsError(m) Then CalculerPrimeUnique = 0: Exit Function
        dx = CDbl(wsTable.Cells(m, 5).Value)

        somme = somme + (dx / lx0) * (v ^ (t + 0.5))
    Next t

    CalculerPrimeUnique = capital * somme
End Function

Function CalculerPrime(age As Integer, capital As Double, taux As Double, duree As Integer) As Double

    If age < 18 Or age > 100 Then
        CalculerPrime = 0
        Exit Function
    End If

    Dim wsTable As Worksheet
    Set wsTable = ThisWorkbook.Sheets("Table_Mortalité")

    Dim v As Double: v = 1 / (1 + taux)

    Dim numerateur As Double, denominateur As Double
    numerateur = 0: denominateur = 0

    Dim x As Integer, ageActuel As Integer
    Dim m As Variant
    Dim lx As Double, dx As Double

    For x = 0 To duree - 1
        ageActuel = age + x
        If ageActuel > 110 Then Exit For

        m = Application.Match(ageActuel, wsTable.Range("A:A"), 0)
        If IsError(m) Then
            CalculerPrime = 0
            Exit Function
        End If

        ' D = lx ; E = dx (d’après ta table)
        lx = CDbl(wsTable.Cells(m, 4).Value)
        dx = CDbl(wsTable.Cells(m, 5).Value)

        numerateur = numerateur + dx * (v ^ (x + 0.5))
        denominateur = denominateur + lx * (v ^ x)
    Next x

    If denominateur > 0 Then
        CalculerPrime = capital * (numerateur / denominateur)
    Else
        CalculerPrime = 0
    End If

End Function

Sub CalculerPrimePure()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Prime_Pure")
    
    ' Récupération des paramètres
    Dim age As Integer, capital As Double, taux As Double, duree As Integer
    
    On Error Resume Next
    age = ws.Range("B5").Value
    capital = ws.Range("B6").Value
    taux = ws.Range("B7").Value
    duree = ws.Range("B8").Value
    On Error GoTo 0
    
    ' Validations
    If age < 18 Or age > 80 Then
        MsgBox "L'âge doit être entre 18 et 80 ans", vbExclamation, "Erreur de saisie"
        ws.Range("B5").Select
        Exit Sub
    End If
    
    If capital <= 0 Or capital > 10000000 Then
        MsgBox "Le capital doit être entre 1 € et 10 000 000 €", vbExclamation, "Erreur de saisie"
        ws.Range("B6").Select
        Exit Sub
    End If
    
    If taux < 0 Or taux > 0.1 Then
        MsgBox "Le taux doit être entre 0% et 10%", vbExclamation, "Erreur de saisie"
        ws.Range("B7").Select
        Exit Sub
    End If
    
    If duree < 1 Or duree > 50 Then
        MsgBox "La durée doit être entre 1 et 50 ans", vbExclamation, "Erreur de saisie"
        ws.Range("B8").Select
        Exit Sub
    End If
    
    If age + duree > 110 Then
        MsgBox "?? La durée dépasse la limite de la table de mortalité" & vbCrLf & _
               "(Âge final : " & age + duree & " ans)", vbExclamation, "Erreur"
        ws.Range("B8").Select
        Exit Sub
    End If
    
    ' Calcul de la prime
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim primePure As Double
    primePure = CalculerPrime(age, capital, taux, duree)
    
    ' Affichage des résultats
    ws.Range("B12").Value = primePure
    ws.Range("B13").Value = primePure * 1.25
    ws.Range("B14").Value = primePure * 1.25 * duree
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Message de confirmation
    MsgBox "Calcul terminé !" & vbCrLf & vbCrLf & _
           "Détails :" & vbCrLf & _
           "• Âge : " & age & " ans" & vbCrLf & _
           "• Capital : " & Format(capital, "#,##0 €") & vbCrLf & _
           "• Durée : " & duree & " ans" & vbCrLf & vbCrLf & _
           "Prime pure annuelle : " & Format(primePure, "#,##0.00 €") & vbCrLf & _
           "Prime commerciale : " & Format(primePure * 1.25, "#,##0.00 €") & vbCrLf & _
           "Coût total : " & Format(primePure * 1.25 * duree, "#,##0.00 €"), _
           vbInformation, "Résultat du calcul"
    
End Sub

