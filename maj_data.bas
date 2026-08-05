Sub refresh_data_dia()
	Application.ScreenUpdating = False
    Dim DictPrincipal As Object
    Set DictPrincipal = CreateObject("Scripting.Dictionary")

    ' récupération du path actuelle de la macro
    Dim CheminDossierPrincipal As String
    Dim ClasseurPrincipal As Workbook
    Set ClasseurPrincipal = ThisWorkbook
    CheminDossierPrincipal = ClasseurPrincipal.Path
    
    DictPrincipal.Add "path_principal", CheminDossierPrincipal
	
    ' validation de l'existence du dossier "data_dia"
    Const DirectoryDataDia As String = "data_dia"
    Dim CheminDossierDataDia As String
    CheminDossierDataDia = get_path_directory(CheminDossierPrincipal, DirectoryDataDia)
	
	' récupération chemin d'accès fichier DIA
	Dim DataDiaFile() As String
	Dim CheminDataDia As String
	DictPrincipal.Add "data_dia", CreateObject("Scripting.Dictionary")
	DictPrincipal("data_dia").Add "path", CheminDossierDataDia
	DataDiaFile = GetFilesList(DictPrincipal("data_dia")("path"))
	CheminDataDia = GetSingleElement(DataDiaFile)
	DictPrincipal("data_dia").Add "path_data_dia_file", CheminDataDia
	
	' récupération des nom de colonne du fichier DIA
	Dim HeadersDiaFile() As String
	HeadersDiaFile = GetColumnNamesFromPath(DictPrincipal("data_dia")("path_data_dia_file"))
	
	'MsgBox Join(HeadersDiaFile, ", ")
	
    ' validation de l'existence du dossier "templates"
    Const DirectoryTemplates As String = "templates"
    Dim CheminDossierTemplates As String
    CheminDossierTemplates = get_path_directory(CheminDossierPrincipal, DirectoryTemplates)
    
    DictPrincipal.Add "path_templates", CreateObject("Scripting.Dictionary")
    DictPrincipal("path_templates").Add "path", CheminDossierTemplates
    DictPrincipal("path_templates").Add "directory_in_templates", CreateObject("Scripting.Dictionary")
    
	' création du dictionnaire contenant les sous dossiers de "templates"
    Dim CheminDirectoryTemplates() As String
    CheminDirectoryTemplates = GetSubFoldersList(CheminDossierTemplates)
    
    Set DictPrincipal = AddFoldersToDictionary(DictPrincipal, CheminDirectoryTemplates)
    
	' récupération des fichiers dans chacun des dossiers du dossier "templates"
	Dim Cle As Variant
	Dim CheminFilesTemplates() As String
	Dim HeadersFilesList() As String
	Dim HeadersDiaList() As String

	For Each Cle In DictPrincipal("path_templates")("directory_in_templates").Keys
		Application.ScreenUpdating = False
		CheminFilesTemplates = GetFilesList(DictPrincipal("path_templates")("directory_in_templates")(Cle)("path"))
		Set DictPrincipal("path_templates")("directory_in_templates")(Cle)("files_in_path") = AddFilesToDictionary(DictPrincipal("path_templates")("directory_in_templates")(Cle)("files_in_path"), CheminFilesTemplates)
		
		' récupération des nom de colonne du fichier commençant par "DIA_" dans les différents templates ne commencant pas par "suivi_"
		HeadersFilesList = GetHeadersList(CheminFilesTemplates)
		HeadersDiaList = ExtractDiaValues(HeadersFilesList)
		
		' validation toutes les colonnes issues de DIA dans les templates sont biens présentes dans le fichier extrait de dia
		ContientTousLesElements HeadersDiaList, HeadersDiaFile
		Application.ScreenUpdating = True
	Next Cle

	' on remonte d'un niveau dans notre arborescence
	Dim CheminDirectorySuivieProd As String
	CheminDirectorySuivieProd = RemonterUnNiveau(CheminDossierPrincipal)
	
	' validation de l'existence du dossier "collaborateurs"
    Const DirectoryCollaborateurs As String = "collaborateurs"
	Dim DirectoryCollaborateursexist As Boolean
	DirectoryCollaborateursexist = DossierExiste(CheminDirectorySuivieProd & DirectoryCollaborateurs)	
	
	' si dossier "collaborateurs" est inexistant on crée le dossier et l'arborescence depuis les données DIA des DA, DC, DS
	If Not DirectoryCollaborateursexist Then
	
		' création du dossier collaborateurs
		CreerDossier CheminDirectorySuivieProd, DirectoryCollaborateurs

		On Error Resume Next
		Set DataDiaFileWork = Workbooks.Open(Filename:=DictPrincipal("data_dia")("path_data_dia_file"), ReadOnly:=True)
		On Error GoTo 0
		
		' on récupère tous les nom de DA, DC, DS du fichier dia
		Dim CellsLabel As Range
		Dim PlageLabel As Range
		Dim DirectoriesToCreate() As String
		Dim CleanDirectory As String
		For Each Cle In DictPrincipal("path_templates")("directory_in_templates").Keys
			' Création des dossier des DA, DC, DS
			CleanDirectory = GetCheminEnfant(CheminDirectorySuivieProd, DirectoryCollaborateurs)
			CreerDossier CleanDirectory, Cle
			
			Set CellsLabel = letter_colonne(Cle, DataDiaFileWork)
			Set PlageLabel = GetPlageSousCellule(CellsLabel)
			DirectoriesToCreate = GetValeursUniques(PlageLabel)
			
			' Création des dossiers des collaborateurs
			Dim DirectoryForEachCollaborateurType As String
			For Each DirectoryToCreate In DirectoriesToCreate
				DirectoryForEachCollaborateurType = GetCheminEnfant(CleanDirectory, Cle)
				CreerDossier DirectoryForEachCollaborateurType, DirectoryToCreate
				
				' intégration des fichier depuis le dossier template pour chaque catégorie de collaborateur
				DirectoryForEachCollaborateurType = GetCheminEnfant(DirectoryForEachCollaborateurType, DirectoryToCreate)
				CopierTousLesFichiers DictPrincipal("path_templates")("directory_in_templates")(Cle)("path"), DirectoryForEachCollaborateurType
			Next DirectoryToCreate
		Next Cle

	End If
	
	
	
    Dim message As String
    message = DictionnaryToString(DictPrincipal)
    
    MsgBox message
	
	
	Application.ScreenUpdating = True
	
	{
	path_principal :"",
	path_templates : {
		path: "path_principal + templates"
		directory_in_templates:{		
			da :{
				name:"DA",
				path : "path_principal + templates + DA",
				files_in_path: {
					TVA : {
						name: "TVA.xlsx",
						path_file : "path_principal + templates + DA + TVA.xlsx",
						headers_file : []
						}
					}
				}
			dc:{
				name:"DC"
				path : "path_principal + templates + DC"
				},
			ds:{
				name:"DS"
				path : "path_principal + templates + DS"
				}
			}
		}
	}

    
    
    
    ' au passage on dépose les bonnes matrices dans les bon dossiers
    ' sinon validation de l'existence des dossiers "DC" et "DA"
    ' on récupère les données des fichiers ne commencant pas par "suivi_" pour centraliser les données dans un unique fichier
    ' validation de l'existence du dossier "archives" dans le dossier "responsables"
    ' validation de l'existence du dossier "temp" dans le dossier "archives"
    ' on crée notre fichier unique dans ce dossier "temp"
    ' on supprime le dossier "collaborateurs"
    ' on utilise la fonction de création de l'arborescence qui reprend les données DIA
    ' depuis notre fichier unique dans "temp" on affecte les données client aux différents DA, DC
    ' pour les fichiers de suivit des DC et DS il suffira d'appuyer sur actualiser pour récupérer les données pas besoin de les transférer
    ' en cas de dossier client non affecté on lévera une erreur
    ' validation de l'existence du dossier "archives" dans le dossier "archive"
    ' on transfère le fichier unique depuis le dossier "temp" vers archives en l'horodatant
    ' on affiche un message récap des traitements, dossier DA, DC, DS créés ou supprimés
End Sub

' =========================================================================
' LES FONCTIONS
' =========================================================================

    Function get_path_directory(path_directory As String, path_directory_concat As String) As String
        Dim path_to_return As String
        Dim FSO As Object
        
        If Right(path_directory, 1) = "\" Then
            path_directory = Left(path_directory, Len(path_directory) - 1)
        End If
        
        path_to_return = path_directory & "\" & path_directory_concat & "\"
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If Not FSO.FolderExists(path_to_return) Then
            MsgBox "Erreur : Le dossier '" & path_to_return & "' est introuvable à l'emplacement :" & path_directory, vbCritical, "Dossier Manquant"
            End
        End If
        get_path_directory = path_to_return
    End Function
    
    
    Function GetPathFile(CheminDossier As String, FileToCheck As String) As String
        Dim CheminComplet As String
        
        CheminComplet = CheminDossier & FileToCheck
        If Dir(CheminComplet) <> "" Then
            GetPathFile = CheminComplet
        Else
            MsgBox "Erreur : Le fichier '" & FileToCheck & "' est introuvable à l'emplacement :" & CheminDossier, vbCritical, "Fichier Manquant"
            End
        End If
    End Function
    
    
    Function GetSubFoldersList(CheminDossier As String) As String()
        Dim FSO As Object
        Dim DossierParent As Object
        Dim SousDossier As Object
        Dim ListeDossiers() As String
        Dim i As Long
        
        ' Initialisation de l'objet FileSystemObject
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        Set DossierParent = FSO.GetFolder(CheminDossier)
        
        ' Vérifie s'il y a au moins un sous-dossier
        If DossierParent.SubFolders.Count = 0 Then
            GetSubFoldersList = Split("", "") ' Retourne un tableau vide
            Exit Function
        End If
        
        ' Dimensionne le tableau de sortie
        ReDim ListeDossiers(0 To DossierParent.SubFolders.Count - 1)
        
        ' Parcourt les sous-dossiers et extrait leur chemin complet
        i = 0
        For Each SousDossier In DossierParent.SubFolders
            ListeDossiers(i) = SousDossier.Path
            ' Remarque : Utilisez SousDossier.Name pour récupérer uniquement le nom du dossier
            i = i + 1
        Next SousDossier
        
        GetSubFoldersList = ListeDossiers
    End Function
    
    
    Function AddFoldersToDictionary(DictDestination As Object, ListeDossiers() As String) As Object
        Dim FSO As Object
        Dim i As Long
        Dim NomDossier As String
        Dim DictDirInTemplates As Object
        
        ' Vérifie que le tableau n'est pas vide
        If (Not ListeDossiers) = -1 Then
            End
        End If
        
        Set DictDirInTemplates = DictDestination("path_templates")("directory_in_templates")
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        For i = LBound(ListeDossiers) To UBound(ListeDossiers)
            If ListeDossiers(i) <> "" Then
                ' Extraction automatique du dernier dossier du chemin (ex: "C:\A\B\Templates" -> "Templates")
                NomDossier = FSO.GetFolder(ListeDossiers(i)).Name
                
                ' Ajoute au dictionnaire s'il n'existe pas déjà pour éviter les doublons
                If Not DictDirInTemplates.Exists(NomDossier) Then
                    DictDirInTemplates.Add NomDossier, CreateObject("Scripting.Dictionary")
                    DictDirInTemplates(NomDossier).Add "path", ListeDossiers(i)
					DictDirInTemplates(NomDossier).Add "name", NomDossier
					DictDirInTemplates(NomDossier).Add "files_in_path", CreateObject("Scripting.Dictionary")
                End If
            End If
        Next i
        
        Set AddFoldersToDictionary = DictDestination
    End Function
        
		
	Function GetFilesList(ByVal CheminDossier As String) As String()
		Dim FSO As Object
		Dim DossierSource As Object
		Dim Fichier As Object
		Dim ListeFichiers() As String
		Dim i As Long
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		Set DossierSource = FSO.GetFolder(CheminDossier)
		
		' Vérifie s'il y a des fichiers dans le dossier
		If DossierSource.Files.Count = 0 Then
			ReDim ListeFichiers(0 To -1)
			GetFilesList = ListeFichiers
			Exit Function
		End If
		
		' Redimensionne le tableau au nombre exact de fichiers
		ReDim ListeFichiers(0 To DossierSource.Files.Count - 1)
		
		i = 0
		For Each Fichier In DossierSource.Files
			ListeFichiers(i) = Fichier.Path
			i = i + 1
		Next Fichier
		
		GetFilesList = ListeFichiers
	End Function
		
		
	Function AddFilesToDictionary(ByVal DictDestination As Object, ByRef ListeFiles() As String) As Object
		Dim FSO As Object
		Dim i As Long
		Dim NomFichier As String
		Dim NomSansExtension As String
		
		' 1. Sortie propre si le tableau est vide (évite d'utiliser End)
		If (Not ListeFiles) = -1 Then
			Set AddFilesToDictionary = DictDestination
			Exit Function
		End If
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		For i = LBound(ListeFiles) To UBound(ListeFiles)
			If ListeFiles(i) <> "" Then
				' 2. Correction de la variable passée à GetBaseName
				NomFichier = FSO.GetFileName(ListeFiles(i))
				NomSansExtension = FSO.GetBaseName(ListeFiles(i))
				
				' 3. Utilisation constante de NomSansExtension et de DictDestination
				If Not DictDestination.Exists(NomSansExtension) Then
					DictDestination.Add NomSansExtension, CreateObject("Scripting.Dictionary")
					DictDestination(NomSansExtension).Add "path_file", ListeFiles(i)
					DictDestination(NomSansExtension).Add "name", NomFichier
					DictDestination(NomSansExtension).Add "headers_file", New Collection
				End If
			End If
		Next i
		
		Set AddFilesToDictionary = DictDestination
	End Function
	
	
	Function GetHeadersList(CheminFichierList() As String) As String()
		Dim WkCible As Workbook
		Dim DerniereCol As Long
		Dim i As Long, c As Long
		Dim Titre As String
		Dim PathFile As String
		
		Dim ListeEnTetes As Collection
		Set ListeEnTetes = New Collection
		
		Application.ScreenUpdating = False
		
		' Sécurité : tableau d'entrée vide
		If (Not CheminFichierList) = -1 Then
			GetHeadersList = Split("")
			Exit Function
		End If
		
		' 1. Parcours des fichiers (index 'i')
		For i = LBound(CheminFichierList) To UBound(CheminFichierList)
			PathFile = CheminFichierList(i)
			
			If Trim(PathFile) <> "" Then
				On Error Resume Next
				Set WkCible = Workbooks.Open(Filename:=PathFile, ReadOnly:=True)
				On Error GoTo 0
				
				If WkCible Is Nothing Then
					MsgBox "Erreur : Impossible d'ouvrir le fichier " & PathFile, vbCritical
				Else
					With WkCible.Sheets(1)
						DerniereCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
						
						' 2. Parcours des colonnes (index 'c' séparé pour ne pas corrompre 'i')
						For c = 1 To DerniereCol
							Titre = CStr(.Cells(2, c).MergeArea.Cells(1, 1).Value)
							
							If .Cells(3, c).Value <> "" And .Cells(3, c).Value <> Titre Then
								Titre = Titre & " - " & .Cells(3, c).Value
							End If
							
							' Ajout de l'élément à la suite dans la collection globale
							ListeEnTetes.Add Titre
						Next c
					End With
					
					WkCible.Close SaveChanges:=False
					Set WkCible = Nothing
				End If
			End If
			
			Application.ScreenUpdating = True
		Next i
		
		' 3. Conversion de la Collection en tableau String()
		If ListeEnTetes.Count = 0 Then
			GetHeadersList = Split("")
		Else
			Dim Entetes() As String
			ReDim Entetes(0 To ListeEnTetes.Count - 1)
			
			For i = 1 To ListeEnTetes.Count
				Entetes(i - 1) = ListeEnTetes(i)
			Next i
			
			GetHeadersList = Entetes
		End If
	End Function


	Function ExtractDiaValues(HeadersList() As String) As String()
		Dim i As Long
		Dim Valeur As String
		Dim CleExtraite As String
		Dim DictUnique As Object
		
		Set DictUnique = CreateObject("Scripting.Dictionary")
		' Rendre la comparaison des clés insensible à la casse (ex: "Code" et "CODE" seront traités comme doublons)
		DictUnique.CompareMode = 1 
		
		' Sécurité : vérification si le tableau d'entrée est vide
		If (Not HeadersList) = -1 Then
			ExtractDiaValues = Split("")
			Exit Function
		End If
		
		For i = LBound(HeadersList) To UBound(HeadersList)
			Valeur = Trim(HeadersList(i))
			
			' Vérifie si la chaîne commence par "DIA_"
			If InStr(1, Valeur, "DIA_", vbTextCompare) = 1 Then
				CleExtraite = Mid(Valeur, 5)
				
				' L'affectation par clé évite les doublons : si la clé existe déjà, elle ne sera pas dupliquée
				If Not DictUnique.Exists(CleExtraite) Then
					DictUnique.Add CleExtraite, CleExtraite
				End If
			End If
		Next i
		
		' Conversion des clés du dictionnaire en tableau String()
		If DictUnique.Count = 0 Then
			ExtractDiaValues = Split("")
		Else
			Dim Resultat() As String
			Dim KeysArray As Variant
			
			KeysArray = DictUnique.Keys
			ReDim Resultat(0 To DictUnique.Count - 1)
			
			For i = 0 To DictUnique.Count - 1
				Resultat(i) = CStr(KeysArray(i))
			Next i
			
			ExtractDiaValues = Resultat
		End If
	End Function
	

	Function GetSingleElement(ListeString() As String) As String
		' 1. Vérifie si le tableau est alloué / non vide
		If (Not ListeString) = -1 Then
			MsgBox "Erreur : Le dossier 'data_dia' est vide (non initialisé).", vbExclamation
			End
			Exit Function
		End If

		' 2. Vérifie s'il y a exactement un seul élément (UBound - LBound + 1 = 1)
		If (UBound(ListeString) - LBound(ListeString) + 1) = 1 Then
			' Retourne la valeur unique (en nettoyant les espaces superflus si besoin)
			GetSingleElement = Trim(ListeString(LBound(ListeString)))
		Else
			MsgBox "Erreur : Le dossier 'data_dia' contient " & (UBound(ListeString) - LBound(ListeString) + 1) & " éléments au lieu d'un seul.", vbExclamation
			End
		End If
	End Function
	
	
	Function GetColumnNamesFromPath(ByVal PathFile As String) As String()
		Dim WkSource As Workbook
		Dim DerniereCol As Long
		Dim c As Long
		Dim Entetes() As String
		
		' Desactivation des rafraîchissements visuels et alertes
		Application.DisplayAlerts = False
		
		' Ouverture sécurisée du classeur en arrière-plan (lecture seule)
		On Error Resume Next
		Set WkSource = Workbooks.Open(Filename:=PathFile, ReadOnly:=True)
		On Error GoTo 0
		
		If WkSource Is Nothing Then
			MsgBox "Erreur : Impossible d'ouvrir le fichier " & PathFile, vbCritical
			Application.DisplayAlerts = True
			GetColumnNamesFromPath = Split("")
			Exit Function
		End If
		
		' Extraction des en-têtes de la ligne 1 de la première feuille
		With WkSource.Sheets(1)
			' Détermine la dernière colonne renseignée sur la ligne 1
			DerniereCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
			
			' Si la ligne 1 est totalement vide (la dernière cellule renvoyée est A1 mais vide)
			If DerniereCol = 1 And Trim(.Cells(1, 1).Value) = "" Then
				GetColumnNamesFromPath = Split("")
			Else
				ReDim Entetes(0 To DerniereCol - 1)
				
				For c = 1 To DerniereCol
					' Récupère la valeur de la cellule (en gérant le cas des cellules fusionnées si nécessaire)
					Entetes(c - 1) = CStr(.Cells(1, c).MergeArea.Cells(1, 1).Value)
				Next c
				
				GetColumnNamesFromPath = Entetes
			End If
		End With
		
		' Fermeture sans enregistrer et réactivation des options d'affichage
		WkSource.Close SaveChanges:=False
		Set WkSource = Nothing
		
		Application.DisplayAlerts = True
	End Function
	
	
	Function RemonterUnNiveau(ByVal CheminDossier As String) As String
		Dim FSO As Object
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		' GetParentFolderName extrait automatiquement le dossier parent
		RemonterUnNiveau = FSO.GetParentFolderName(CheminDossier)
	End Function
	
	
	Function DossierExiste(ByVal CheminDossier As String) As Boolean
		Dim FSO As Object
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		' S'assure que le chemin se termine par un anti-slash
		If Right(CheminDossier, 1) <> "\" Then CheminDossier = CheminDossier & "\"
		
		' Vérifie la présence du dossier "collaborateurs"
		DossierExiste = FSO.FolderExists(CheminDossier & "collaborateurs")
	End Function
	
	
	Function letter_colonne(ByVal label_to_find As String, ByVal WkSource As Workbook, Optional ByVal NomFeuille As String = "") As Range
		Application.ScreenUpdating = False
		Dim sheet_to_use As Worksheet
		Dim CelluleTrouvee As Range
		
		' Si aucun nom de feuille n'est précisé, on prend la 1ère feuille par défaut
		If NomFeuille = "" Then
			Set sheet_to_use = WkSource.Sheets(1)
		Else
			Set sheet_to_use = WkSource.Sheets(NomFeuille)
		End If
		
		' Recherche dans la ligne 1
		Set CelluleTrouvee = sheet_to_use.Rows(1).Find(What:=label_to_find, LookIn:=xlValues, LookAt:=xlWhole)
		
		If CelluleTrouvee Is Nothing Then
			MsgBox "Erreur critique : Le label '" & label_to_find & "' est introuvable dans l'onglet '" & sheet_to_use.Name & "' du fichier '" & WkSource.Name & "'.", vbCritical
			End
		End If
		Application.ScreenUpdating = True
		Set letter_colonne = CelluleTrouvee
	End Function
	
	
	Function GetPlageSousCellule(ByVal CelluleLabel As Range, Optional ByVal LigneDepartDonnees As Long = 2) As Range
		Application.ScreenUpdating = False
		Dim WS As Worksheet
		Dim NumCol As Long
		Dim DerniereLigne As Long

		' 1. Vérification que la cellule transmise n'est pas Nothing
		If CelluleLabel Is Nothing Then
			MsgBox "Erreur : La cellule fournie est invalide (Nothing).", vbCritical
			End
		End If

		' 2. Récupération de la feuille parent et du numéro de colonne
		Set WS = CelluleLabel.Worksheet
		NumCol = CelluleLabel.Column

		' 3. Détermination de la dernière ligne renseignée dans cette colonne
		DerniereLigne = WS.Cells(WS.Rows.Count, NumCol).End(xlUp).Row

		' 4. Vérification qu'il y a des données sous l'en-tête
		If DerniereLigne < LigneDepartDonnees Then
			MsgBox "Aucune donnée à traiter sous le label '" & CelluleLabel.Value & _
				   "' dans la feuille '" & WS.Name & "' (" & WS.Parent.Name & ").", vbExclamation
			Exit Function
		End If
		Application.ScreenUpdating = True
		' 5. Renvoi de la plage de données (ex: de A2 jusqu'à A100)
		Set GetPlageSousCellule = WS.Range(WS.Cells(LigneDepartDonnees, NumCol), WS.Cells(DerniereLigne, NumCol))
	End Function
	
	
	Function GetValeursUniques(ByVal PlageDonnees As Range, Optional ByVal IgnorerVides As Boolean = True) As String()
		Application.ScreenUpdating = False
		Dim DictUnique As Object
		Dim Cellule As Range
		Dim Valeur As String
		Dim i As Long
		
		' Sécurité : Si la plage est vide ou invalide
		If PlageDonnees Is Nothing Then
			GetValeursUniques = Split("")
			Exit Function
		End If
		
		' Création du dictionnaire
		Set DictUnique = CreateObject("Scripting.Dictionary")
		
		' CompareMode = 1 (vbTextCompare) rend la comparaison insensible à la casse ("ABC" = "abc")
		' Si vous voulez que la casse compte, mettez CompareMode = 0
		DictUnique.CompareMode = 1
		
		' Parcours de toutes les cellules de la plage
		For Each Cellule In PlageDonnees.Cells
			Valeur = Trim(CStr(Cellule.Value))
			
			' Gestion des cellules vides selon l'option passée
			If Not (IgnorerVides And Valeur = "") Then
				' Si la clé n'existe pas encore dans le dictionnaire, on l'ajoute
				If Not DictUnique.Exists(Valeur) Then
					DictUnique.Add Valeur, Valeur
				End If
			End If
		Next Cellule
		
		Application.ScreenUpdating = True
		
		' Conversion des clés du dictionnaire vers un tableau String()
		If DictUnique.Count = 0 Then
			GetValeursUniques = Split("")
		Else
			Dim Resultat() As String
			Dim Cles As Variant
			
			Cles = DictUnique.Keys
			ReDim Resultat(0 To DictUnique.Count - 1)
			
			For i = 0 To DictUnique.Count - 1
				Resultat(i) = CStr(Cles(i))
			Next i
			
			GetValeursUniques = Resultat
		End If
	End Function
	
	
	Function GetCheminEnfant(ByVal CheminParent As String, ByVal NomEnfant As String) As String
		Dim FSO As Object
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		' S'assure de nettoyer les séparateurs et combine les chemins proprement
		GetCheminEnfant = FSO.BuildPath(CheminParent, NomEnfant)
	End Function

' =========================================================================
' LES FONCTIONS
' =========================================================================



' =========================================================================
' DEBUG DICT
' =========================================================================


Function DictionnaryToString(ByVal Dico As Object, Optional ByVal Niveau As Long = 0) As String
    Dim Cle As Variant
    Dim Indentation As String
    Dim Texte As String
    
    If Dico Is Nothing Then Exit Function
    
    ' Crée les espaces pour l'indentation hiérarchique
    Indentation = String(Niveau * 4, " ")
    
    For Each Cle In Dico.Keys
        If IsObject(Dico(Cle)) Then
            If TypeName(Dico(Cle)) = "Dictionary" Then
                Texte = Texte & Indentation & "[" & Cle & "] :" & vbCrLf
                ' Appel récursif pour les sous-dictionnaires
                Texte = Texte & DictionnaryToString(Dico(Cle), Niveau + 1)
            Else
                Texte = Texte & Indentation & Cle & " : [Objet " & TypeName(Dico(Cle)) & "]" & vbCrLf
            End If
        Else
            Texte = Texte & Indentation & Cle & " : " & Dico(Cle) & vbCrLf
        End If
    Next Cle
    
    DictionnaryToString = Texte
End Function






' =========================================================================
' LES SUB
' =========================================================================

	Sub ContientTousLesElements(ByRef ListeATester As Variant, ByRef ListeReference As Variant)
		Dim DictRef As Object
		Dim i As Long
		Dim ElementManquant As Boolean
		
		ElementManquant = True
		
		Set DictRef = CreateObject("Scripting.Dictionary")
		' Ignorer la casse lors des comparaisons de texte ("ABC" = "abc")
		DictRef.CompareMode = 1 'vbTextCompare
		
		' 1. Chargement de la liste de référence dans le dictionnaire
		For i = LBound(ListeReference) To UBound(ListeReference)
			If Not DictRef.Exists(ListeReference(i)) Then
				DictRef.Add ListeReference(i), True
			End If
		Next i
		
		' 2. Vérification de chaque valeur de la liste à tester
		For i = LBound(ListeATester) To UBound(ListeATester)
			If Not DictRef.Exists(ListeATester(i)) Then
				ElementManquant = False
				MsgBox "La colonne '"& ListeATester(i)&"' est manquante dans le fichier extrait de Dia."
			End If
		Next i
		
		If Not ElementManquant Then
			End
        End If

	End Sub
	
	
	Sub CreerDossier(ByVal PathDirectorySource As String, ByVal DirectoryToCreate As String)
		Dim FSO As Object
		Dim CheminDossier As String
		
		If Right(PathDirectorySource, 1) = "\" Then
            CheminDossier = Left(PathDirectorySource, Len(PathDirectorySource) - 1)
        End If
        
        CheminDossier = PathDirectorySource & "\" & DirectoryToCreate
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		' Vérifie si le dossier n'existe pas déjà avant de le créer
		If Not FSO.FolderExists(CheminDossier) Then
			FSO.CreateFolder CheminDossier
		Else
			MsgBox "Le dossier '" & DirectoryToCreate & "' existe déjà.", vbExclamation
			End
		End If
	End Sub
	
	
	Sub CopierTousLesFichiers(ByVal DossierSource As String, ByVal DossierDestination As String)
		Dim FSO As Object
		Dim MasqueSource As String

		Set FSO = CreateObject("Scripting.FileSystemObject")

		' 3. Construction du filtre (ex: "C:\DossierA\*.*" ou "C:\DossierA\*")
		MasqueSource = FSO.BuildPath(DossierSource, "*")

		' 4. Copie de TOUS les fichiers en un seul bloc (Overwrite:=True écrase les fichiers existants)
		On Error Resume Next
		FSO.CopyFile MasqueSource, DossierDestination & "\", True
		
		If Err.Number <> 0 Then
			MsgBox "Erreur lors de la copie : " & Err.Description, vbExclamation
			End
		End If
		On Error GoTo 0
	End Sub

' =========================================================================
' LES SUB
' =========================================================================

