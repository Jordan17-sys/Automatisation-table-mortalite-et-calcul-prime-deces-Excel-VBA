Attribute VB_Name = "Remplir_Formules_Table_M"
Sub Remplir_Formules_Table_M()
    ' ======================================
    ' MACRO : Remplir automatiquement les formules
    ' ======================================
    
    Dim ws As Worksheet
    Dim wsBrutes As Worksheet
    Dim lastRowBrutes As Long
    Dim lastRowTable As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Table_Mortalité")
    Set wsBrutes = ThisWorkbook.Sheets("Données_Brutes")
    
    ' Trouver la dernière ligne de données brutes
    lastRowBrutes = wsBrutes.Cells(wsBrutes.Rows.Count, 1).End(xlUp).Row
    lastRowTable = lastRowBrutes - 2 + 3  ' Conversion : ligne brutes ? ligne table
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ======================================
    ' REMPLIR LES FORMULES LIGNE PAR LIGNE
    ' ======================================
    For i = 3 To lastRowTable
        With ws
            ' Age
            .Cells(i, 1).Formula = "='Données_Brutes'!A" & (i - 1)
            
            ' qx (probabilité de décès)
            .Cells(i, 2).Formula = "='Données_Brutes'!B" & (i - 1)
            
            ' px (probabilité de survie)
            .Cells(i, 3).Formula = "=1-B" & i
            
            ' lx (nombre de vivants)
            .Cells(i, 4).Formula = "='Données_Brutes'!C" & (i - 1)
            
            ' dx (nombre de décès)
            .Cells(i, 5).Formula = "=D" & i & "*B" & i
            
            ' Lx (années vécues entre x et x+1)
            If i < lastRowTable Then
                .Cells(i, 6).Formula = "=(D" & i & "+D" & (i + 1) & ")/2"
            Else
                ' Dernière ligne : Lx = lx * ex (car pas de ligne suivante)
                .Cells(i, 6).Formula = "=D" & i & "*0.5"  ' Approximation
            End If
            
            ' Tx (années totales à vivre) - CORRECTION ICI
            .Cells(i, 7).Formula = "=SUM(F" & i & ":F$" & lastRowTable & ")"
            
            ' ex (espérance de vie)
            .Cells(i, 8).Formula = "=IF(D" & i & ">0,G" & i & "/D" & i & ",0)"
        End With
    Next i
    
    ' ======================================
    ' MISE EN FORME
    ' ======================================
    With ws
        .Range("B3:C" & lastRowTable).NumberFormat = "0.00000"  ' qx, px
        .Range("D3:F" & lastRowTable).NumberFormat = "#,##0"     ' lx, dx, Lx
        .Range("G3:G" & lastRowTable).NumberFormat = "#,##0"     ' Tx
        .Range("H3:H" & lastRowTable).NumberFormat = "0.00"      ' ex
        
        .Range("A3:H" & lastRowTable).Borders.LineStyle = xlContinuous
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Formules remplies avec succes !" & vbCrLf & _
           "Lignes calculees : " & (lastRowTable - 2) & vbCrLf & vbCrLf & _
           "Verifications :" & vbCrLf & _
           "Age 0 : lx=" & Format(ws.Cells(3, 4).Value, "#,##0") & _
           " / ex=" & Format(ws.Cells(3, 8).Value, "0.00"), _
           vbInformation, "MORTEX"
End Sub


