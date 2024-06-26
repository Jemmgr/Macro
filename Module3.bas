Attribute VB_Name = "Macro"

Sub CreerTableFiltres()
    Dim wsPersonnel As Worksheet
    Dim wsFiltres As Worksheet
    Dim wsBasseDeD As Worksheet
    Dim lastRowPersonnel As Long
    Dim lastRowFiltres As Long
    Dim i As Long
    Dim mois As String
    Dim joursDeTravail As Long
    Dim dateDebut As Date
    Dim dateFin As Date
    Dim anne As Long
    Dim moiss As String
    Dim wbPersonnel As Workbook
    Dim Nom_complet_fichier_de_traitement As String
    Dim Nom_court_fichier_de_traitement As String
    Dim formulaStr As String
    Dim wb As Workbook
    Dim fileExtension As String
    Dim tempWorkbook As Workbook
    Dim tempFilePath As String
    Dim uniqueID As String
    
    ' Référencer le livre de travail actuel explicitement
    Set wb = ThisWorkbook
    
    ' Demander à l'utilisateur de sélectionner le mois
    mois = InputBox("S'il vous plaît, entrez le mois (en formato MMM):", "Sélection du mois")
    ' Quitter si l'utilisateur annule
    If mois = "" Then Exit Sub
    
    ' Convertir le mois saisi en majuscules et enlever les accents
    mois = RemoveAccents(UCase(mois))
    
    ' Définir la feuille "BASSE DE D" dans le livre actuel
    Set wsBasseDeD = wb.Sheets("BASSE DE D")
    
    ' Rechercher le mois dans la feuille "BASSE DE D" et obtenir les données correspondantes
    Dim rngmois As Range
    Dim cell As Range
    Dim found As Boolean
    found = False
    
    For Each cell In wsBasseDeD.Columns("B").Cells
        If InStr(1, RemoveAccents(UCase(cell.Value)), mois) > 0 Then
            Set rngmois = cell
                found = True
            Exit For
        End If
    Next cell
    
    If Not found Then
        MsgBox "Le mois saisi n'a pas été trouvé dans le tableau.", vbExclamation
        Exit Sub
    End If
    
    ' Vérifier si la feuille "FILTRES" existe déjà et demander une confirmation pour la remplacer
    On Error Resume Next
    Set wsFiltres = wb.Sheets("FILTRES")
    On Error GoTo 0
    
    If Not wsFiltres Is Nothing Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Voulez-vous remplacer la table de FILTRES existante ?", vbYesNo + vbQuestion, "Confirmer le remplacement")
        If answer = vbYes Then
            Application.DisplayAlerts = False ' Désactiver les alertes pour supprimer la feuille
            wsFiltres.Delete ' Supprimer la feuille existante
            Application.DisplayAlerts = True
        Else
            Exit Sub ' Quitter si l'utilisateur décide de ne pas remplacer le tableau
        End If
    End If
    
    ' Créer la nouvelle feuille appelée "FILTRES"
    On Error Resume Next
        ' Ajouter une nouvelle feuille après la dernier feiulle existente
        Set wsFiltres = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    On Error GoTo 0
    
    If wsFiltres Is Nothing Then
        Exit Sub
    End If
    
    ' Assurer que le nom de la feuille est unique
    Dim uniqueName As String
    uniqueName = "FILTRES"
    i = 1
    Do While SheetExists(uniqueName)
        uniqueName = "FILTRES" & i
        i = i + 1
    Loop
    wsFiltres.Name = uniqueName
    
    ' Demander à l'utilisateur de sélectionner le fichier de données
    Nom_complet_fichier_de_traitement = Application.GetOpenFilename(, , "Veuillez sélectionner le fichier contenant la feuille de PRC")
    If Nom_complet_fichier_de_traitement = "False" Then Exit Sub ' Quitter si l'utilisateur annule
    
    ' Vérifier l'extension du fichier
    fileExtension = Right(Nom_complet_fichier_de_traitement, Len(Nom_complet_fichier_de_traitement) - InStrRev(Nom_complet_fichier_de_traitement, "."))
    
    ' Générer un identifiant unique basé sur la date et l'heure actuelles
    uniqueID = Format(Now, "yyyymmddHHMMSS")
    
    ' S'il s'agit d'un fichier CSV, convertissez-le en XLSX et séparez les colonnes
    If fileExtension = "csv" Or fileExtension = "CSV" Then
        ' Ouvrir le fichier CSV
        Workbooks.OpenText fileName:=Nom_complet_fichier_de_traitement, Origin:=xlMSDOS, startRow:=1, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False, Space:=False, Other:=False, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), TrailingMinusNumbers:=True
            
        ' Référence du fichier ouvert
        Set tempWorkbook = ActiveWorkbook
        
        ' Enregistrez le fichier CSV au format XLSX dans un emplacement temporaire avec un nom unique
        tempFilePath = Environ("TEMP") & "\temp_converted_file_" & uniqueID & ".xlsx"
        tempWorkbook.SaveAs fileName:=tempFilePath, FileFormat:=xlOpenXMLWorkbook
        tempWorkbook.Close SaveChanges:=False
        
        ' Ouvrir le fichier converti
        Set wbPersonnel = Workbooks.Open(tempFilePath)
    Else
        ' Ouvrir le fichier directement s'il ne s'agit pas d'un CSV
        Set wbPersonnel = Workbooks.Open(Nom_complet_fichier_de_traitement)
    End If
    
    Set wsPersonnel = wbPersonnel.Sheets(1) ' Assumer que les données sont dans la première feuille
    
    ' Stocker l'information dans des variables
    anne = rngmois.Offset(0, -1).Value ' Colonne A
    moiss = CStr(rngmois.Value) ' Colonne B
    dateDebut = rngmois.Offset(0, 1).Value ' Colonne C
    dateFin = rngmois.Offset(0, 2).Value ' Columna D
    joursDeTravail = rngmois.Offset(0, 3).Value ' Colonne E
    
    ' Trouver la dernière ligne avec des données dans la feuille sélectionnée
    lastRowPersonnel = wsPersonnel.Cells(wsPersonnel.Rows.Count, "K").End(xlUp).Row
    
    ' Titres de colonne et attribution de valeurs communes
    With wsFiltres
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
        .Range("H7").Value = "Libellé"
        .Range("G5").Value = "Jours de travail:"
        .Range("A5").Value = "Date debut:"
        .Range("C5").Value = "Date fin:"
        .Range("H5").Value = joursDeTravail
        .Range("B5").Value = dateDebut
        .Range("D5").Value = dateFin
        
        ' Appliquer un format aux titres
        .Range("A7:H7").Font.Bold = True
        .Range("A7:H7").HorizontalAlignment = xlCenter
    End With
    
    ' Trouver la dernière ligne avec des données dans la colonne BC
    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "BC").End(xlUp).Row
    
    ' Parcourir chaque cellule de la colonne BC
    For i = 1 To lastRow
        cellValue = wsPersonnel.Cells(i, "BC").Value
        If IsNumeric(cellValue) Then
            ' Si la valeur est numérique, l'arrondir à un chiffre après la virgule et supprimer les zéros inutiles
            cellValue = Round(cellValue, 1)
            If Int(cellValue) = cellValue Then
                ' Si le nombre arrondi est un entier, supprimer la décimale
                wsPersonnel.Cells(i, "BC").Value = Int(cellValue)
            Else
                ' Si le nombre arrondi n'est pas un entier, l'afficher avec une décimale
                wsPersonnel.Cells(i, "BC").Value = cellValue
            End If
            
            ' Supprimer le zéro des cellules vides
            If wsPersonnel.Cells(i, "BC").Value = 0 Then
                wsPersonnel.Cells(i, "BC").ClearContents
            End If
        End If
    Next i
    
    ' Créer un objet Dictionary pour stocker des noms uniques
    Dim nomsUniques As Object
    Set nomsUniques = CreateObject("Scripting.Dictionary")
    
    ' Copier les noms des travailleurs et les codes opérationnels dans FILTRES
    For i = 2 To lastRowPersonnel ' Commencer à partir de la deuxième ligne en supposant que la première est un en-tête
        ' Obtenir le nom du travailleur et le code opérationnel
        Dim nombre As String
        Dim code As String
        Dim prenom As String
        nombre = wsPersonnel.Cells(i, "K").Value
        code = wsPersonnel.Cells(i, "AN").Value
        prenom = wsPersonnel.Cells(i, "L").Value
        
        ' Vérifier si le nom est déjà dans le tableau FILTRES
        If Not nomsUniques.Exists(code & nombre) Then
            ' Ajouter le nom au Dictionary
            nomsUniques(code & nombre) = True
            ' Trouver la prochaine ligne vide dans FILTRES
            lastRowFiltres = wsFiltres.Cells(wsFiltres.Rows.Count, "A").End(xlUp).Row + 1
            ' Copier les informations des colonnes supplémentaires de la feuille de données
            With wsFiltres
                ' Supprimer le texte indésirable "(e)" avant d'attribuer les valeurs à la colonne I
                Dim valorI As Variant
                valorI = Replace(wsPersonnel.Cells(i, "I").Value, "(e)", "")
                .Cells(lastRowFiltres, 1).Value = valorI
                .Cells(lastRowFiltres, 2).Value = nombre
                .Cells(lastRowFiltres, 3).Value = code
                ' Définir la formule dans la colonne D
                .Cells(lastRowFiltres, 4).Formula = "=IFERROR(SUMIFS('[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!$BC:$BC, '[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!K:K, """ & nombre & """, '[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!L:L, """ & prenom & """, '[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!AN:AN, """ & code & """)/SUMIFS('[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!$BC:$BC, '[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!K:K, """ & nombre & """, '[" & wbPersonnel.Name & "]" & wsPersonnel.Name & "'!L:L, """ & prenom & """), """")"
                ' Définir la formule dans la colonne E
                .Cells(lastRowFiltres, 5).Formula = "=IFERROR(ROUND(D" & lastRowFiltres & "*" & joursDeTravail & ", 0), 0)"
                .Cells(lastRowFiltres, 7).Value = prenom ' Colonne L (Prenom)
                .Cells(lastRowFiltres, 8).Value = wsPersonnel.Cells(i, "R").Value ' Colonne R (Libellé)
            End With
        End If
    Next i
    
    ' Vérifier que lastRowFiltres est supérieur ou égal à 8
    If lastRowFiltres >= 8 Then
        ' Trier le tableau par la colonne NOM
        With wsFiltres.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsFiltres.Range("B8:B" & lastRowFiltres), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wsFiltres.Range("A7:H" & lastRowFiltres)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    ' Ajuster les sommes des sections
    AjustarSommes wsFiltres, joursDeTravail
    
    ' Copier et coller les valeurs des colonnes D et E à la fin du tableau
    wsFiltres.Range("D8:E" & lastRowFiltres).Value = wsFiltres.Range("D8:E" & lastRowFiltres).Value
    
    ' Format de pourcentage dans la colonne POURCENTAGE
    wsFiltres.Range("D8:D" & lastRowFiltres).NumberFormat = "0.0%"
    
    ' Appliquer un style à la table
    wsFiltres.ListObjects.Add(xlSrcRange, wsFiltres.Range("A7:H" & lastRowFiltres), , xlYes).TableStyle = "TableStyleMedium9"
    
    ' Exécuter la macro MultiplierPourcentage
    MultiplierPourcentage wb
    
    ' Enregistrer le nouveau document et copier la feuille FILTRES dans le nouveau classeur
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add ' Créer un nouveau classeur
    wsFiltres.Copy Before:=newWorkbook.Sheets(1) ' Copier la feuille FILTRES dans le nouveau classeur
    
    ' Supprimer la feuille vide initiale du nouveau classeur
    Dim ws As Worksheet
    For Each ws In newWorkbook.Sheets
        If ws.Name <> "FILTRES" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    ' Permettre à l'utilisateur de sélectionner le dossier où enregistrer le fichier
    Dim selectedFolder As FileDialog
    Set selectedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    selectedFolder.Title = "Sélectionnez un dossier pour enregistrer le fichier"
    
    If selectedFolder.Show = -1 Then
        Dim folderPath As String
        folderPath = selectedFolder.SelectedItems(1)
        
        ' Définir le nom du fichier avec le chemin complet
        Dim fileName As String
        fileName = folderPath & "\Imp-VLJOUR-" & Format(dateFin, "yyyymmdd") & ".xlsx"
        
        ' Enregistrer le nouveau classeur dans le dossier sélectionné
        Application.DisplayAlerts = False
        newWorkbook.SaveAs fileName, FileFormat:=xlOpenXMLWorkbook ' Format de fichier compatible avec les macros
        Application.DisplayAlerts = True
        
        ' Fermer le classeur sans enregistrer les modifications
        newWorkbook.Close SaveChanges:=False
        
        ' Ouvrir le nouveau document enregistré
        Workbooks.Open fileName
        MsgBox "La table a été créée avec succès dans une nouvelle feuille et le document a été sauvegardé sous le nom : " & fileName & " et ouvert.", vbInformation
    Else
        MsgBox "Aucun dossier n'a été sélectionné. Le processus a été annulé.", vbExclamation
    End If
    
    ' Fermer le fichier de données sans enregistrer les modifications
    wbPersonnel.Close SaveChanges:=False
End Sub

Sub AjustarSommes(wsFiltres As Worksheet, joursDeTravail As Long)
    Dim debutSection As Long
    Dim finSection As Long
    Dim sommeJours As Double
    Dim i As Long, j As Long
    Dim difference As Long
    
    debutSection = 8 ' Commencer à partir de la ligne 8 où commencent les données
    
    For i = 8 To wsFiltres.Cells(wsFiltres.Rows.Count, "A").End(xlUp).Row
        If wsFiltres.Cells(i, 1).Value <> wsFiltres.Cells(i + 1, 1).Value Or i = wsFiltres.Cells(wsFiltres.Rows.Count, "A").End(xlUp).Row Then
            finSection = i
            ' Calculer la somme des jours dans la section
            sommeJours = Application.WorksheetFunction.Sum(wsFiltres.Range("E" & debutSection & ":E" & finSection))
            
            ' Assurez-vous que sommeJours n'est pas zéro
            If sommeJours = 0 Then
                sommeJours = 1
            End If
            
            ' Ajuster la somme des jours à la quantité de jours dans la section
            Dim proportionJours As Double
            proportionJours = joursDeTravail / sommeJours
            
            ' Appliquer l'ajustement à la colonne E dans la section
            For j = debutSection To finSection
                wsFiltres.Cells(j, 5).Value = Round(wsFiltres.Cells(j, 5).Value * proportionJours, 0)
            Next j
            
            ' Recalculer la somme des jours après l'ajustement
            sommeJours = Application.WorksheetFunction.Sum(wsFiltres.Range("E" & debutSection & ":E" & finSection))
            
            ' Ajuster la somme des jours pour qu'elle corresponde à joursDeTravail
            difference = joursDeTravail - sommeJours
            
            If difference <> 0 Then
                If difference > 0 Then
                    ' Ajouter des jours si la différence est positive
                    For j = finSection To debutSection Step -1
                        If difference = 0 Then Exit For
                        wsFiltres.Cells(j, 5).Value = wsFiltres.Cells(j, 5).Value + 1
                        difference = difference - 1
                    Next j
                Else
                    ' Retirer des jours si la différence est négative
                    For j = finSection To debutSection Step -1
                        If difference = 0 Then Exit For
                        If wsFiltres.Cells(j, 5).Value > 0 Then
                            If wsFiltres.Cells(j, 5).Value + difference >= 0 Then
                                wsFiltres.Cells(j, 5).Value = wsFiltres.Cells(j, 5).Value + difference
                                difference = 0
                            Else
                                difference = difference + wsFiltres.Cells(j, 5).Value
                                wsFiltres.Cells(j, 5).Value = 0
                            End If
                        End If
                    Next j
                End If
            End If
            
            ' Mettre à jour le début de la section suivante
            debutSection = i + 1
        End If
    Next i
End Sub

Sub MultiplierPourcentage(wb As Workbook)
    Dim Nom_complet_fichier_de_traitement As String
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim wsFiltres As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim currentRow As Long
    Dim currentEmployer As String
    Dim totalPourcentage As Double
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
    Dim cell As Range
    
    ' Sélectionner le fichier
    Nom_complet_fichier_de_traitement = Application.GetOpenFilename(, , "Veuillez sélectionner le fichier de pointage")
    
    If Nom_complet_fichier_de_traitement = "False" Then
        MsgBox "Aucun fichier n'a été sélectionné."
        Exit Sub
    End If
    
    ' Ouvrir le fichier en utilisant le chemin complet
    Set wbSource = Workbooks.Open(Nom_complet_fichier_de_traitement)
    
    ' Définir la feuille de travail "repart" comme la première feuille du fichier sélectionné
    Set ws = wbSource.Sheets(1)
    
    ' Définir la feuille de travail "FILTRES" dans le livre actuel
    Set wsFiltres = wb.Sheets("FILTRES")
    
    ' Valeur prise de H5 de la feuille FILTRES
    newValue = wsFiltres.Range("H5").Value
    
    ' Désactiver le calcul automatique pour accélérer le processus
    Application.Calculation = xlCalculationManual
    
    ' Trouver la dernière ligne avec des données dans la colonne A de la feuille
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Recherchez la cellule contenant "NOM" dans la colonne A et obtenez la ligne de départ
    startRow = 0
    For Each cell In ws.Columns("A").Cells
        If cell.Value = "NOM" Then
            startRow = cell.Row + 1
            Exit For
        End If
    Next cell
    
    ' Si "NOM" n'est pas trouvé, commencez par la deuxième ligne
    If startRow = 0 Then
        startRow = 2
    End If
    
    ' Initialiser la plage de début pour chaque section
    currentRow = startRow
    totalPourcentage = 0
    totalAdjusted = 0
    
    ' Parcourir chaque ligne dans la colonne A de la feuille
    For i = startRow To lastRow
        ' Vérifier si le nom et le prénom de l'employeur ont changé
        If ws.Cells(i, 1).Value <> currentEmployer Or ws.Cells(i, 2).Value <> prenom Then
            ' S'il s'agit d'un nouvel employeur ou d'un nouveau prénom, calculer le facteur d'échelle pour la section précédente
            If totalPourcentage > 0 Then
                scaleFactor = newValue / totalPourcentage ' Le nouveau montant est utilisé
                ' Multiplier chaque pourcentage par le facteur d'échelle et arrondir
                For j = currentRow To i - 1
                    ws.Cells(j, 5).Value = Round(ws.Cells(j, 4).Value * scaleFactor, 0)
                    totalAdjusted = totalAdjusted + ws.Cells(j, 5).Value
                Next j
                ' Ajuster la dernière valeur pour que la somme soit égale au nouveau montant
                adjustment = newValue - totalAdjusted
                ws.Cells(i - 1, 5).Value = ws.Cells(i - 1, 5).Value + adjustment
            End If
            ' Mettre à jour le nom et le prénom de l'employeur et réinitialiser les compteurs
            currentEmployer = ws.Cells(i, 1).Value
            prenom = ws.Cells(i, 2).Value
            currentRow = i
            totalPourcentage = 0
            totalAdjusted = 0
        End If
        ' Additionner les pourcentages pour cette section
        totalPourcentage = totalPourcentage + ws.Cells(i, 4).Value
    Next i
    
    ' Calculer le facteur d'échelle pour la dernière section
    If totalPourcentage > 0 Then
        scaleFactor = newValue / totalPourcentage ' Le nouveau montant est utilisé
        ' Multiplie chaque pourcentage par le facteur d'échelle et arrondis
        For j = currentRow To lastRow
            ws.Cells(j, 5).Value = Round(ws.Cells(j, 4).Value * scaleFactor, 0)
            totalAdjusted = totalAdjusted + ws.Cells(j, 5).Value
        Next j
        ' Ajuste la dernière valeur pour que la somme soit égale à la nouvelle valeur
        adjustment = newValue - totalAdjusted
        ws.Cells(lastRow, 5).Value = ws.Cells(lastRow, 5).Value + adjustment
    End If
    
    ' Supprimer les lignes dans FILTRES qui correspondent au nom et prénom dans repart
    For i = lastRow To startRow Step -1
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
    For i = startRow To lastRow
        ' Obtiens le nom, le prénom et le code de la ligne actuelle dans "repart"
        nom = UCase(ws.Cells(i, 1).Value)
        prenom = UCase(ws.Cells(i, 2).Value)
        code = ws.Cells(i, 3).Value
        
        ' Vérifie si le nom, le prénom et le code existent déjà dans la table "FILTRES"
        found = False
        For j = 8 To nextRowFiltres
            If UCase(wsFiltres.Cells(j, 2).Value) = nom And UCase(wsFiltres.Cells(j, 7).Value) = prenom And wsFiltres.Cells(j, 3).Value = code Then
                ' Si trouvé, supprime les informations supplémentaires et remplace les valeurs par celles de "repart"
                wsFiltres.Cells(j, 1).Value = "" ' Supprime l'information dans la colonne A
                wsFiltres.Cells(j, 8).Value = "" ' Supprime l'information dans la colonne H
                wsFiltres.Cells(j, 2).Value = ws.Cells(i, 1).Value ' Copie le nom
                wsFiltres.Cells(j, 3).Value = ws.Cells(i, 3).Value ' Copie le code
                wsFiltres.Cells(j, 4).Value = ws.Cells(i, 4).Value ' Copie le pourcentage
                wsFiltres.Cells(j, 5).Value = ws.Cells(i, 5).Value ' Copie la valeur calculée
                wsFiltres.Cells(j, 7).Value = ws.Cells(i, 2).Value ' Copie le prenom
                found = True
                Exit For
            End If
        Next j
        
        ' Si le nom, le prénom et le code ne sont pas trouvés, copiez la ligne vers "FILTRES"
        If Not found Then
            wsFiltres.Cells(nextRowFiltres, 1).Value = "" ' Supprime l'information dans la colonne A
            wsFiltres.Cells(nextRowFiltres, 8).Value = "" ' Supprime l'information dans la colonne H
            wsFiltres.Cells(nextRowFiltres, 2).Value = ws.Cells(i, 1).Value ' Copie le nom
            wsFiltres.Cells(nextRowFiltres, 3).Value = ws.Cells(i, 3).Value ' Copie le code
            wsFiltres.Cells(nextRowFiltres, 4).Value = ws.Cells(i, 4).Value ' Copie le pourcentage
            wsFiltres.Cells(nextRowFiltres, 5).Value = ws.Cells(i, 5).Value ' Copie la valeur calculée
            wsFiltres.Cells(nextRowFiltres, 7).Value = ws.Cells(i, 2).Value ' Copie le prenom
            nextRowFiltres = nextRowFiltres + 1
        End If
    Next i
    
    ' Définis le plage à trier
    Set sortRange = wsFiltres.Range("A7:I" & nextRowFiltres - 1)
    
    ' Trie les données par nom de famille dans la colonne B
    With wsFiltres.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsFiltres.Range("B8:B" & nextRowFiltres - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes ' La première ligne est une ligne d'en-tête
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Supprime les lignes où la valeur dans la colonne D est zéro
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
    sectionStart = 8 ' Commencez à partir de la ligne 8 où les données commencent
    
    For i = 8 To wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row
        If wsFiltres.Cells(i, 2).Value <> currentNom Or wsFiltres.Cells(i, 7).Value <> currentPrenom Then
            ' Si la section change, recalculer les pourcentages
            sectionEnd = i - 1
            ' Additionner les valeurs actuelles de la colonne D dans la section
            totalPourcentage = Application.WorksheetFunction.Sum(wsFiltres.Range("D" & sectionStart & ":D" & sectionEnd))
            
            ' Recalculer les pourcentages pour que la somme soit 100%
            If totalPourcentage <> 0 Then
                For j = sectionStart To sectionEnd
                    wsFiltres.Cells(j, 4).Value = wsFiltres.Cells(j, 4).Value / totalPourcentage
                Next j
            End If
            
            ' Mettre à jour le début de la section suivante
            sectionStart = i
            currentNom = wsFiltres.Cells(i, 2).Value
            currentPrenom = wsFiltres.Cells(i, 7).Value
        End If
    Next i
    
    ' Assurer le recalcul de la dernière section
    sectionEnd = wsFiltres.Cells(wsFiltres.Rows.Count, "B").End(xlUp).Row
    totalPourcentage = Application.WorksheetFunction.Sum(wsFiltres.Range("D" & sectionStart & ":D" & sectionEnd))
    
    If totalPourcentage <> 0 Then
        For j = sectionStart To sectionEnd
            wsFiltres.Cells(j, 4).Value = wsFiltres.Cells(j, 4).Value / totalPourcentage
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
    
    ' Réactive le calcul automatique
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Function RemoveAccents(text As String) As String
    Dim accentedChars As String
    Dim unaccentedChars As String
    Dim i As Integer
    
    ' La formule change les lettres avec des accents par des lettres simples, pour rechercher les mois de manière simplifiée
    accentedChars = "ÁÉÍÓÚÀÈÙÛÔÂÎÏÜËÇáéíóúàèùûôâîïüëç"
    unaccentedChars = "AEIOUAEUUOAIUECaeiouaeuuoaiuec"
    
    For i = 1 To Len(accentedChars)
        text = Replace(text, Mid(accentedChars, i, 1), Mid(unaccentedChars, i, 1))
    Next i
    
    RemoveAccents = text
End Function

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
    
End Function
