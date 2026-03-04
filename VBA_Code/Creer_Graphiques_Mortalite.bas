Attribute VB_Name = "Creer_Graphiques_Mortalite"
Sub Creer_Graphiques_Mortalite()
    ' ======================================
    ' MACRO : Créer les 4 graphiques principaux
    ' (Version corrigée - 1 seule courbe par graphe)
    ' ======================================
    
    Dim ws As Worksheet
    Dim wsGraph As Worksheet
    Dim lastRow As Long
    Dim chart1 As ChartObject, chart2 As ChartObject
    Dim chart3 As ChartObject, chart4 As ChartObject
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Table_Mortalité")
    
    ' Vérifier si la feuille Graphiques existe
    On Error Resume Next
    Set wsGraph = ThisWorkbook.Sheets("Graphiques")
    On Error GoTo 0
    
    If Not wsGraph Is Nothing Then
        If wsGraph.ChartObjects.Count > 0 Then
            wsGraph.ChartObjects.Delete
        End If
        wsGraph.Cells.Clear
    Else
        Set wsGraph = ThisWorkbook.Sheets.Add(After:=ws)
        wsGraph.Name = "Graphiques"
        wsGraph.Tab.Color = RGB(255, 192, 0)
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    ' ======================================
    ' GRAPHIQUE 1 : COURBE DE SURVIE (lx)
    ' ======================================
    Set chart1 = wsGraph.ChartObjects.Add(Left:=10, Top:=10, Width:=500, Height:=300)
    With chart1.Chart
        .ChartType = xlLine
        ' Utiliser A:D pour inclure les deux colonnes côte à côte
        .SetSourceData Source:=ws.Range("A1:D" & lastRow)
        .HasTitle = True
        .ChartTitle.Text = "Courbe de survie (lx) - France 2025"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Age"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Survivants (lx)"
        
        ' Supprimer toutes les séries SAUF la 4ème (colonne D = lx)
        For i = .SeriesCollection.Count To 1 Step -1
            If i <> 4 Then .SeriesCollection(i).Delete
        Next i
        
        .SeriesCollection(1).Name = "Survivants (lx)"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(1).Format.Line.Weight = 2.5
    End With
    
    ' ======================================
    ' GRAPHIQUE 2 : PROBABILITÉ DE DÉCÈS (qx)
    ' ======================================
    Set chart2 = wsGraph.ChartObjects.Add(Left:=520, Top:=10, Width:=500, Height:=300)
    With chart2.Chart
        .ChartType = xlLine
        .SetSourceData Source:=ws.Range("A1:B" & lastRow)
        .HasTitle = True
        .ChartTitle.Text = "Probabilité de décès (qx) - France 2025"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Age"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Probabilité (qx)"
        
        ' Supprimer toutes les séries SAUF la 2ème (colonne B = qx)
        For i = .SeriesCollection.Count To 1 Step -1
            If i <> 2 Then .SeriesCollection(i).Delete
        Next i
        
        .SeriesCollection(1).Name = "Probabilité qx"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(1).Format.Line.Weight = 2.5
        
        ' Échelle logarithmique
        On Error Resume Next
        .Axes(xlValue).ScaleType = xlScaleLogarithmic
        On Error GoTo 0
    End With
    
    ' ======================================
    ' GRAPHIQUE 3 : ESPÉRANCE DE VIE (ex)
    ' ======================================
    Set chart3 = wsGraph.ChartObjects.Add(Left:=10, Top:=320, Width:=500, Height:=300)
    With chart3.Chart
        .ChartType = xlLine
        .SetSourceData Source:=ws.Range("A1:H" & lastRow)
        .HasTitle = True
        .ChartTitle.Text = "Espérance de vie résiduelle (ex) - France 2025"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Age"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Espérance de vie (années)"
        
        ' Supprimer toutes les séries SAUF la 8ème (colonne H = ex)
        For i = .SeriesCollection.Count To 1 Step -1
            If i <> 8 Then .SeriesCollection(i).Delete
        Next i
        
        .SeriesCollection(1).Name = "Espérance ex"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(1).Format.Line.Weight = 2.5
    End With
    
    ' ======================================
    ' GRAPHIQUE 4 : NOMBRE DE DÉCÈS (dx)
    ' ======================================
    Set chart4 = wsGraph.ChartObjects.Add(Left:=520, Top:=320, Width:=500, Height:=300)
    With chart4.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("A1:E" & lastRow)
        .HasTitle = True
        .ChartTitle.Text = "Nombre de décès par âge (dx) - France 2025"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Age"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Nombre de décès"
        
        ' Supprimer toutes les séries SAUF la 5ème (colonne E = dx)
        For i = .SeriesCollection.Count To 1 Step -1
            If i <> 5 Then .SeriesCollection(i).Delete
        Next i
        
        .SeriesCollection(1).Name = "Décès dx"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
    End With
    
    wsGraph.Activate
    Application.ScreenUpdating = True
    
    MsgBox "4 graphiques créés avec succès !" & vbCrLf & vbCrLf & _
           "1 seule courbe par graphique", _
           vbInformation, "MORTEX"
End Sub


