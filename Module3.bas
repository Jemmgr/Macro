Attribute VB_Name = "Module3"
Function RemoveAccents(text As String) As String
    Dim accentedChars As String
    Dim unaccentedChars As String
    Dim i As Integer
    
    ' La formule change les lettres avec des accents par des lettres simples, pour rechercher les mois de mani�re simplifi�e
    accentedChars = "��������������������������������"
    unaccentedChars = "AEIOUAEUUOAIUECaeiouaeuuoaiuec"
    
    For i = 1 To Len(accentedChars)
        text = Replace(text, Mid(accentedChars, i, 1), Mid(unaccentedChars, i, 1))
    Next i
    
    RemoveAccents = text
End Function

Sub CrearTablaFiltrosrecientefr()
    Dim wsPersonal As Worksheet
    Dim wsFiltros As Worksheet
    Dim wsBasseDeD As Worksheet
    Dim lastRowPersonal As Long
    Dim lastRowFiltros As Long
    Dim i As Long
    Dim mes As String
    Dim joursDeTravail As Long
    Dim fechaDebut As Date
    Dim fechaFin As Date
    Dim anne As Long
    Dim moiss As String
    Dim wbPersonal As Workbook
    Dim Nom_complet_fichier_de_traitement As String
    Dim Nom_court_fichier_de_traitement As String
    Dim formulaStr As String
    
    ' Demander � l'utilisateur de s�lectionner le mois
    mes = InputBox("S'il vous pla�t, entrez le mois (en formato MMM):", "S�lection du mois")
    
    If mes = "" Then Exit Sub ' Quitter si l'utilisateur annule
    
    ' Convertir le mois saisi en majuscules et enlever les accents
    mes = RemoveAccents(UCase(mes))
    
    ' D�finir la feuille "BASSE DE D" dans le livre actuel
    Set wsBasseDeD = ActiveWorkbook.Sheets("BASSE DE D")
    
    ' Rechercher le mois dans la feuille "BASSE DE D" et obtenir les donn�es correspondantes
    Dim rngMes As Range
    Dim cell As Range
    Dim found As Boolean
    found = False
    
    For Each cell In wsBasseDeD.Columns("B").Cells
        If InStr(1, RemoveAccents(UCase(cell.Value)), mes) > 0 Then
            Set rngMes = cell
                found = True
            Exit For
        End If
    Next cell
    
    If Not found Then
        MsgBox "Le mois saisi n'a pas �t� trouv� dans le tableau.", vbExclamation
        Exit Sub
    End If
    
    ' V�rifier si la feuille "FILTRES" existe d�j� et demander une confirmation pour la remplacer
    On Error Resume Next
    Set wsFiltros = ThisWorkbook.Sheets("FILTRES")
    On Error GoTo 0
    If Not wsFiltros Is Nothing Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Voulez-vous remplacer la table de FILTRES existante ?", vbYesNo + vbQuestion, "Confirmer le remplacement")
        If answer = vbYes Then
            Application.DisplayAlerts = False ' D�sactiver les alertes pour supprimer la feuille
            wsFiltros.Delete ' Supprimer la feuille existante
            Application.DisplayAlerts = True
        Else
            Exit Sub ' Quitter si l'utilisateur d�cide de ne pas remplacer le tableau
        End If
    End If
    
    ' Demander � l'utilisateur de s�lectionner le fichier de donn�es
    Nom_complet_fichier_de_traitement = Application.GetOpenFilename(, , "Veuillez s�lectionner le fichier contenant la feuille de PRC")
    If Nom_complet_fichier_de_traitement = "False" Then Exit Sub ' Quitter si l'utilisateur annule
    
    ' Ouvrir le fichier en utilisant le chemin complet
    Set wbPersonal = Workbooks.Open(Nom_complet_fichier_de_traitement)
    Set wsPersonal = wbPersonal.Sheets(1) ' Assumer que les donn�es sont dans la premi�re feuille
    
    ' Cr�er la nouvelle feuille appel�e "FILTRES"
    Set wsFiltros = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsFiltros.Name = "FILTRES"
    
    ' Stocker l'information dans des variables
    anne = rngMes.Offset(0, -1).Value ' Colonne A
    moiss = CStr(rngMes.Value) ' Colonne B
    fechaDebut = rngMes.Offset(0, 1).Value ' Colonne C
    fechaFin = rngMes.Offset(0, 2).Value ' Columna D
    joursDeTravail = rngMes.Offset(0, 3).Value ' Colonne E
    
    ' Trouver la derni�re ligne avec des donn�es dans la feuille s�lectionn�e
    lastRowPersonal = wsPersonal.Cells(wsPersonal.Rows.Count, "K").End(xlUp).Row
    
    ' Titres de colonne et attribution de valeurs communes
    With wsFiltros
        .Range("A2:H4").Merge
        .Range("A2:H4").Value = "AFFECTATIONS AUTOS " & moiss & " " & anne & ""
        .Range("A2:H4").HorizontalAlignment = xlCenter
        .Range("A2:H4").VerticalAlignment = xlCenter
        .Range("A2:H4").Font.Size = 22
        .Range("A5:H5").Font.Size = 11
        .Range("A2:H4, A5, C5, G5").Font.Bold = True
        
        .Range("A8:H1000").HorizontalAlignment = xlLeft
        
        .Range("A7").Value = "AGENCE"
        .Range("B7").Value = "NOM"
        .Range("C7").Value = "AFFECTATION"
        .Range("D7").Value = "POURCENTAGE"
        .Range("E7").Value = "Nb jours"
        .Range("F7").Value = "Rattachement Agence"
        .Range("G7").Value = "Prenom"
        .Range("H7").Value = "Libell�"
        
        .Range("G5").Value = "Jours de travail:"
        .Range("A5").Value = "Date debut:"
        .Range("C5").Value = "Date fin:"
        .Range("H5").Value = joursDeTravail
        .Range("B5").Value = fechaDebut
        .Range("D5").Value = fechaFin
        
        ' Appliquer un format aux titres
        .Range("A7:H7").Font.Bold = True
        .Range("A7:H7").HorizontalAlignment = xlCenter
    End With
    
    ' Trouver la derni�re ligne avec des donn�es dans la colonne BC
    lastRow = wsPersonal.Cells(wsPersonal.Rows.Count, "BC").End(xlUp).Row
    
    ' Parcourir chaque cellule de la colonne BC
    For i = 1 To lastRow
        cellValue = wsPersonal.Cells(i, "BC").Value
        If IsNumeric(cellValue) Then
            ' Si la valeur est num�rique, l'arrondir � un chiffre apr�s la virgule et supprimer les z�ros inutiles
            cellValue = Round(cellValue, 1)
            If Int(cellValue) = cellValue Then
                ' Si le nombre arrondi est un entier, supprimer la d�cimale
                wsPersonal.Cells(i, "BC").Value = Int(cellValue)
            Else
                ' Si le nombre arrondi n'est pas un entier, l'afficher avec une d�cimale
                wsPersonal.Cells(i, "BC").Value = cellValue
            End If
            
            ' Supprimer le z�ro des cellules vides
            If wsPersonal.Cells(i, "BC").Value = 0 Then
                wsPersonal.Cells(i, "BC").ClearContents
            End If
        End If
    Next i
    
    ' Cr�er un objet Dictionary pour stocker des noms uniques
    Dim nombresUnicos As Object
    Set nombresUnicos = CreateObject("Scripting.Dictionary")
    
    ' Copier les noms des travailleurs et les codes op�rationnels dans FILTRES
    For i = 2 To lastRowPersonal ' Commencer � partir de la deuxi�me ligne en supposant que la premi�re est un en-t�te
        ' Obtenir le nom du travailleur et le code op�rationnel
        Dim nombre As String
        Dim codigo As String
        Dim prenom As String
        nombre = wsPersonal.Cells(i, "K").Value
        codigo = wsPersonal.Cells(i, "AN").Value
        prenom = wsPersonal.Cells(i, "L").Value
        
        ' V�rifier si le nom est d�j� dans le tableau FILTRES
        If Not nombresUnicos.Exists(codigo & nombre) Then
            ' Ajouter le nom au Dictionary
            nombresUnicos(codigo & nombre) = True
            ' Trouver la prochaine ligne vide dans FILTRES
            lastRowFiltros = wsFiltros.Cells(wsFiltros.Rows.Count, "A").End(xlUp).Row + 1
            ' Copier les informations des colonnes suppl�mentaires de la feuille de donn�es
            With wsFiltros
                ' Supprimer le texte ind�sirable "(e)" avant d'attribuer les valeurs � la colonne I
                Dim valorI As Variant
                valorI = Replace(wsPersonal.Cells(i, "I").Value, "(e)", "")
                .Cells(lastRowFiltros, 1).Value = valorI
                .Cells(lastRowFiltros, 2).Value = nombre
                .Cells(lastRowFiltros, 3).Value = codigo
                ' D�finir la formule dans la colonne D
                .Cells(lastRowFiltros, 4).Formula = "=IFERROR(SUMIFS('[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!$BC:$BC, '[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!K:K, """ & nombre & """, '[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!L:L, """ & prenom & """, '[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!AN:AN, """ & codigo & """)/SUMIFS('[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!$BC:$BC, '[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!K:K, """ & nombre & """, '[" & wbPersonal.Name & "]" & wsPersonal.Name & "'!L:L, """ & prenom & """), """")"
                ' D�finir la formule dans la colonne E
                .Cells(lastRowFiltros, 5).Formula = "=IFERROR(ROUND(D" & lastRowFiltros & "*" & joursDeTravail & ", 0), 0)"
                .Cells(lastRowFiltros, 7).Value = prenom ' Colonne L (Prenom)
                .Cells(lastRowFiltros, 8).Value = wsPersonal.Cells(i, "R").Value ' Colonne R (Libell�)
            End With
        End If
    Next i
    
    ' Trier le tableau par la colonne NOM
    With wsFiltros.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsFiltros.Range("B8:B" & lastRowFiltros), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsFiltros.Range("A7:H" & lastRowFiltros)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Initialiser les variables pour le suivi des sections
    Dim inicioSeccion As Long
    Dim finSeccion As Long
    Dim sumaDias As Double
    
    ' It�rer sur les lignes pour ajuster les sommes dans chaque section
    inicioSeccion = 8 ' Commencer � partir de la ligne 8 o� commencent les donn�es
    For i = 8 To lastRowFiltros
        If wsFiltros.Cells(i, 1).Value <> wsFiltros.Cells(i + 1, 1).Value Or i = lastRowFiltros Then
            ' Si la section change ou si c'est la derni�re ligne, calculer la somme des jours dans la section
            finSeccion = i
            ' Calculer la somme des jours dans la section
            sumaDias = Application.WorksheetFunction.Sum(wsFiltros.Range("E" & inicioSeccion & ":E" & finSeccion))
            
            ' Probl�me en janvier
            If sumaDias = 0 Then
                sumaDias = 1
            End If
            
            ' Ajuster la somme des jours � la quantit� de jours dans la section
            Dim proporcionDias As Double
            proporcionDias = joursDeTravail / sumaDias
            
            ' Appliquer l'ajustement � la colonne E dans la section
            For j = inicioSeccion To finSeccion
                wsFiltros.Cells(j, 5).Value = Round(wsFiltros.Cells(j, 5).Value * proporcionDias, 0)
            Next j
            
            ' Recalculer la somme des jours apr�s l'ajustement
            sumaDias = Application.WorksheetFunction.Sum(wsFiltros.Range("E" & inicioSeccion & ":E" & finSeccion))
            
            ' Si la somme de la section d�passe joursDeTravail, soustraire l'exc�dent du dernier jour de la section
            If sumaDias > joursDeTravail Then
                wsFiltros.Cells(finSeccion, 5).Value = wsFiltros.Cells(finSeccion, 5).Value - (sumaDias - joursDeTravail)
            End If
            
            ' Si la somme de la section est inf�rieure � joursDeTravail, ajouter la diff�rence au dernier jour de la section
            If sumaDias < joursDeTravail Then
                wsFiltros.Cells(finSeccion, 5).Value = wsFiltros.Cells(finSeccion, 5).Value + (joursDeTravail - sumaDias)
            End If
            
        ' Mettre � jour le d�but de la section suivante
        inicioSeccion = i + 1
        End If
    Next i
    
    ' Copier et coller les valeurs des colonnes D et E � la fin du tableau
    wsFiltros.Range("D8:E" & lastRowFiltros).Value = wsFiltros.Range("D8:E" & lastRowFiltros).Value
    
    ' Format de pourcentage dans la colonne POURCENTAGE
    wsFiltros.Range("D8:D" & lastRowFiltros).NumberFormat = "0.0%"
    
    ' Appliquer un style � la table
    wsFiltros.ListObjects.Add(xlSrcRange, wsFiltros.Range("A7:H" & lastRowFiltros), , xlYes).TableStyle = "TableStyleMedium9"
    
    ' Ex�cuter la macro MultiplicarPourcentage
    MultiplicarPorcentaje
    
    ' Enregistrer le nouveau document et copier la feuille FILTRES dans le nouveau classeur
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add ' Cr�er un nouveau classeur
    ThisWorkbook.Sheets("FILTRES").Copy Before:=newWorkbook.Sheets(1) ' Copier la feuille FILTRES dans le nouveau classeur
    
    ' Permettre � l'utilisateur de s�lectionner le dossier o� enregistrer le fichier
    Dim selectedFolder As FileDialog
    Set selectedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    selectedFolder.Title = "S�lectionnez un dossier pour enregistrer le fichier"
    
    If selectedFolder.Show = -1 Then
        Dim folderPath As String
        folderPath = selectedFolder.SelectedItems(1)
        
        ' D�finir le nom du fichier avec le chemin complet
        Dim fileName As String
        fileName = folderPath & "\Imp-VLJOUR-" & Format(fechaFin, "yyyymmdd") & ".xlsm"
        
        ' Enregistrer le nouveau classeur dans le dossier s�lectionn�
        Application.DisplayAlerts = False
        newWorkbook.SaveAs fileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled ' Format de fichier compatible avec les macros
        Application.DisplayAlerts = True
        
        ' Supprimer la deuxi�me feuille du nouveau classeur
        Application.DisplayAlerts = False
        newWorkbook.Sheets(2).Delete
        Application.DisplayAlerts = True
        
        ' Fermer le classeur sans enregistrer les modifications
        newWorkbook.Close SaveChanges:=False
        
        ' Ouvrir le nouveau document enregistr�
        Workbooks.Open fileName
        MsgBox "La table a �t� cr��e avec succ�s dans une nouvelle feuille et le document a �t� sauvegard� sous le nom : " & fileName & " et ouvert.", vbInformation
    Else
        MsgBox "Aucun dossier n'a �t� s�lectionn�. Le processus a �t� annul�.", vbExclamation
    End If
    
    ' Fermer le fichier de donn�es sans enregistrer les modifications
    wbPersonal.Close SaveChanges:=False
End Sub

Sub MultiplicarPorcentaje()
    Dim Nom_complet_fichier_de_traitement As String
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim wsFiltres As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim currentEmployer As String
    Dim totalPercentage As Double
    Dim scaleFactor As Double
    Dim totalAdjusted As Double
    Dim adjustment As Integer
    Dim nextRowFiltres As Long
    Dim sortRange As Range
    Dim newValue As Double
    Dim nom As String
    Dim prenom As String
    Dim code As String
    Dim i As Long
    Dim found As Boolean
    
    ' S�lectionner le fichier
    Nom_complet_fichier_de_traitement = Application.GetOpenFilename(, , "Veuillez s�lectionner le fichier de pointage")
    
    If Nom_complet_fichier_de_traitement = "False" Then
        MsgBox "Aucun fichier n'a �t� s�lectionn�."
        Exit Sub
    End If
    
    ' Ouvrir le fichier en utilisant le chemin complet
    Set wbSource = Workbooks.Open(Nom_complet_fichier_de_traitement)
    
    ' D�finir la feuille de travail "repart" comme la premi�re feuille du fichier s�lectionn�
    Set ws = wbSource.Sheets(1)
    
    ' D�finir la feuille de travail "FILTRES" dans le livre actuel
    Set wsFiltres = ThisWorkbook.Sheets("FILTRES")
    
    ' Valeur prise de H5 de la feuille FILTRES
    newValue = wsFiltres.Range("H5").Value
    
    ' D�sactiver le calcul automatique pour acc�l�rer le processus
    Application.Calculation = xlCalculationManual
    
    ' Trouver la derni�re ligne avec des donn�es dans la colonne A de la feuille
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialiser la plage de d�but pour chaque section
    currentRow = 2
    totalPercentage = 0
    totalAdjusted = 0
    
    ' Parcourir chaque ligne dans la colonne A de la feuille
    For i = 2 To lastRow
        ' V�rifier si le nom et le pr�nom de l'employeur ont chang�
        If ws.Cells(i, 1).Value <> currentEmployer Or ws.Cells(i, 2).Value <> prenom Then
            ' S'il s'agit d'un nouvel employeur ou d'un nouveau pr�nom, calculer le facteur d'�chelle pour la section pr�c�dente
            If totalPercentage > 0 Then
                scaleFactor = newValue / totalPercentage ' Le nouveau montant est utilis�
                ' Multiplier chaque pourcentage par le facteur d'�chelle et arrondir
                For j = currentRow To i - 1
                    ws.Cells(j, 5).Value = Round(ws.Cells(j, 4).Value * scaleFactor, 0)
                    totalAdjusted = totalAdjusted + ws.Cells(j, 5).Value
                Next j
                ' Ajuster la derni�re valeur pour que la somme soit �gale au nouveau montant
                adjustment = newValue - totalAdjusted
                ws.Cells(i - 1, 5).Value = ws.Cells(i - 1, 5).Value + adjustment
            End If
            ' Mettre � jour le nom et le pr�nom de l'employeur et r�initialiser les compteurs
            currentEmployer = ws.Cells(i, 1).Value
            prenom = ws.Cells(i, 2).Value
            currentRow = i
            totalPercentage = 0
            totalAdjusted = 0
        End If
        ' Additionner les pourcentages pour cette section
        totalPercentage = totalPercentage + ws.Cells(i, 4).Value
    Next i
    
    ' Calculer le facteur d'�chelle pour la derni�re section
    
    If totalPercentage > 0 Then
        scaleFactor = newValue / totalPercentage ' Le nouveau montant est utilis�
        ' Multiplie chaque pourcentage par le facteur d'�chelle et arrondis
        For j = currentRow To lastRow
            ws.Cells(j, 5).Value = Round(ws.Cells(j, 4).Value * scaleFactor, 0)
            totalAdjusted = totalAdjusted + ws.Cells(j, 5).Value
        Next j
        ' Ajuste la derni�re valeur pour que la somme soit �gale � la nouvelle valeur
        adjustment = newValue - totalAdjusted
        ws.Cells(lastRow, 5).Value = ws.Cells(lastRow, 5).Value + adjustment
    End If
    
    ' Supprimer les lignes dans FILTRES qui correspondent au nom et pr�nom dans repart
    For i = lastRow To 2 Step -1
        nom = ws.Cells(i, 1).Value
        prenom = ws.Cells(i, 2).Value
        For j = wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row To 8 Step -1
            If wsFiltres.Cells(j, 2).Value = nom And wsFiltres.Cells(j, 7).Value = prenom Then
                wsFiltres.Rows(j).EntireRow.Delete
            End If
        Next j
    Next i
    
    ' Trouve la prochaine ligne disponible dans la feuille "FILTRES"
    nextRowFiltres = wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row + 1
    
    ' Copie l'information de la table "repart" dans la feuille "FILTRES"
    For i = 2 To lastRow
        ' Obtiens le nom, le pr�nom et le code de la ligne actuelle dans "repart"
        nom = UCase(ws.Cells(i, 1).Value)
        prenom = UCase(ws.Cells(i, 2).Value)
        code = ws.Cells(i, 3).Value
        
        ' V�rifie si le nom, le pr�nom et le code existent d�j� dans la table "FILTRES"
        found = False
        For j = 8 To nextRowFiltres
            If UCase(wsFiltres.Cells(j, 2).Value) = nom And UCase(wsFiltres.Cells(j, 7).Value) = prenom And wsFiltres.Cells(j, 3).Value = code Then
                ' Si trouv�, supprime les informations suppl�mentaires et remplace les valeurs par celles de "repart"
                wsFiltres.Cells(j, 1).Value = "" ' Supprime l'information dans la colonne A
                wsFiltres.Cells(j, 8).Value = "" ' Supprime l'information dans la colonne H
                wsFiltres.Cells(j, 2).Value = ws.Cells(i, 1).Value ' Copie le nom
                wsFiltres.Cells(j, 3).Value = ws.Cells(i, 3).Value ' Copie le code
                wsFiltres.Cells(j, 4).Value = ws.Cells(i, 4).Value ' Copie le pourcentage
                wsFiltres.Cells(j, 5).Value = ws.Cells(i, 5).Value ' Copie la valeur calcul�e
                wsFiltres.Cells(j, 7).Value = ws.Cells(i, 2).Value ' Copie le prenom
                found = True
                Exit For
            End If
        Next j
        
        ' Si le nom, le pr�nom et le code ne sont pas trouv�s, copiez la ligne vers "FILTRES"
        If Not found Then
            wsFiltres.Cells(nextRowFiltres, 1).Value = "" ' Supprime l'information dans la colonne A
            wsFiltres.Cells(nextRowFiltres, 8).Value = "" ' Supprime l'information dans la colonne H
            wsFiltres.Cells(nextRowFiltres, 2).Value = ws.Cells(i, 1).Value ' Copie le nom
            wsFiltres.Cells(nextRowFiltres, 3).Value = ws.Cells(i, 3).Value ' Copie le code
            wsFiltres.Cells(nextRowFiltres, 4).Value = ws.Cells(i, 4).Value ' Copie le pourcentage
            wsFiltres.Cells(nextRowFiltres, 5).Value = ws.Cells(i, 5).Value ' Copie la valeur calcul�e
            wsFiltres.Cells(nextRowFiltres, 7).Value = ws.Cells(i, 2).Value ' Copie le prenom
            nextRowFiltres = nextRowFiltres + 1
        End If
    Next i
    
    ' D�finis le plage � trier
    Set sortRange = wsFiltres.Range("A7:I" & nextRowFiltres - 1)
    
    ' Trie les donn�es par nom de famille dans la colonne B
    With wsFiltres.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsFiltres.Range("B8:B" & nextRowFiltres - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes ' La premi�re ligne est une ligne d'en-t�te
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Supprime les lignes o� la valeur dans la colonne D est z�ro
    For i = nextRowFiltres - 1 To 8 Step -1
        If wsFiltres.Cells(i, 5).Value = 0 Then
            wsFiltres.Rows(i).EntireRow.Delete
        End If
    Next i
    
    ' Recalculer les pourcentages pour que chaque section totalise 100 %
    Dim currentNom As String
    Dim currentPrenom As String
    Dim sectionStart As Long
    Dim sectionEnd As Long
    sectionStart = 8 ' Commencez � partir de la ligne 8 o� les donn�es commencent
    
    For i = 8 To wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row
        If wsFiltres.Cells(i, 2).Value <> currentNom Or wsFiltres.Cells(i, 7).Value <> currentPrenom Then
            ' Si la section change, recalculer les pourcentages
            sectionEnd = i - 1
            ' Additionner les valeurs actuelles de la colonne D dans la section
            totalPercentage = Application.WorksheetFunction.Sum(wsFiltres.Range("D" & sectionStart & ":D" & sectionEnd))
            
            ' Recalculer les pourcentages pour que la somme soit 100%
            If totalPercentage <> 0 Then
                For j = sectionStart To sectionEnd
                    wsFiltres.Cells(j, 4).Value = wsFiltres.Cells(j, 4).Value / totalPercentage
                Next j
            End If
            
            ' Mettre � jour le d�but de la section suivante
            sectionStart = i
            currentNom = wsFiltres.Cells(i, 2).Value
            currentPrenom = wsFiltres.Cells(i, 7).Value
        End If
    Next i
    
    ' Assurer le recalcul de la derni�re section
    sectionEnd = wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row
    totalPercentage = Application.WorksheetFunction.Sum(wsFiltres.Range("D" & sectionStart & ":D" & sectionEnd))
    
    If totalPercentage <> 0 Then
        For j = sectionStart To sectionEnd
            wsFiltres.Cells(j, 4).Value = wsFiltres.Cells(j, 4).Value / totalPercentage
        Next j
    End If
    
    ' Convertir toutes les cellules des colonnes B et G en majuscules dans la feuille "FILTRES"
    Dim lastFiltresRow As Long
    lastFiltresRow = wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row
    
    ' Parcourir la colonne B
    For i = 8 To lastFiltresRow
        wsFiltres.Cells(i, 2).Value = UCase(wsFiltres.Cells(i, 2).Value)
    Next i
    
    ' Parcourir la colonne G
    For i = 8 To lastFiltresRow
        wsFiltres.Cells(i, 7).Value = UCase(wsFiltres.Cells(i, 7).Value)
    Next i
    
    ' Fermer le fichier source sans enregistrer les modifications
    wbSource.Close SaveChanges:=False
    
    ' R�active le calcul automatique
    Application.Calculation = xlCalculationAutomatic
    
End Sub

