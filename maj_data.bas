Sub refresh_data_dia()
	Dim AncienCalcul As XlCalculation

	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.DisplayAlerts = False

	AncienCalcul = Application.Calculation
	Application.Calculation = xlCalculationManual
	
	
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
		
		' ajout des noms de colonnes de chaque fichiers dans le "DictPrincipal" sous la clée "headers_file"
		AjouterEntetesAuDictionnaire DictPrincipal, CheminFilesTemplates, Cle, "headers_file"

		Application.ScreenUpdating = True
	Next Cle
	
	' on récupère un dictionnaire contenant les filtres à appliquer dans chaque template pour la récupération des données
	Dim DictFiltresTemplates As Object
	Set DictFiltresTemplates = GetDictionnaireFiltresTemplates(DictPrincipal)

	' on remonte d'un niveau dans notre arborescence
	Dim CheminDirectorySuivieProd As String
	CheminDirectorySuivieProd = RemonterUnNiveau(CheminDossierPrincipal)
	
	' validation de l'existence du dossier "collaborateurs"
    Const DirectoryCollaborateurs As String = "\collaborateurs"
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
			Dim FilesTemplate() As String
			Dim NameFilesTemplate() As String
			For Each DirectoryToCreate In DirectoriesToCreate
				DirectoryForEachCollaborateurType = GetCheminEnfant(CleanDirectory, Cle)
				CreerDossier DirectoryForEachCollaborateurType, DirectoryToCreate
				
				' intégration des fichier depuis le dossier template pour chaque catégorie de collaborateur
				DirectoryForEachCollaborateurType = GetCheminEnfant(DirectoryForEachCollaborateurType, DirectoryToCreate)
				CopierTousLesFichiers DictPrincipal("path_templates")("directory_in_templates")(Cle)("path"), DirectoryForEachCollaborateurType
				
				' DirectoryToCreate = nom collab'
				' Cle = Data Analyste
				' DirectoryForEachCollaborateurType = chemin complet pour récupérer fichiers templates
				' NameFilesTemplate= le nom des fichiers sans extension pour les utiliser comme clés dans notre dictionnaire (liste)
				FilesTemplate = GetFilesList(DirectoryForEachCollaborateurType)
				NameFilesTemplate = GetNameFiles(FilesTemplate)
				
				' on injecte les bonnes données dans les bon template en fonction du type de colaborateur, du collaborateur et des filtres de matrice
				InjectDataInTemplates Cle, DirectoryToCreate, DirectoryForEachCollaborateurType, NameFilesTemplate, DictPrincipal, DataDiaFileWork, DictFiltresTemplates
				
			Next DirectoryToCreate
		Next Cle

	End If
	
	' Si le dossier "collaborateur" existe déjà alors on met à jour 
	If DirectoryCollaborateursexist Then
		' On récupère l'intégralité des données des fichiers remplit par les collaborateurs
		CentraliserDonneesCollaborateurs DictPrincipal
		
		' On les centralise dans un fichier sur 3 onglets dans le dosser archive
		' on regénère l'arborescence des dossiers COLLABORATEUR
		' on injecte les données DIA
		' on injecte les données précédemment remplies selon le "code du dossier"
		' attention pour la tva, bien le coupler avec le mois
		' on horodate le fichier
	End If
	
    ' AfficherDictionnaire DictPrincipal
	
	
	If Not DataDiaFileWork Is Nothing Then
        DataDiaFileWork.Close SaveChanges:=False
        Set DataDiaFileWork = Nothing
    End If
	
	Application.Calculation = AncienCalcul
	Application.DisplayAlerts = True
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	
	MsgBox "Traitement terminé."
	
	
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
	
	
	Function GetDictionnaireFiltresTemplates(ByVal DictPrincipal As Object) As Object
		Dim DictFiltres As Object
		Dim DictFiltreTemplate As Object
		
		Dim TypeCollab As Variant
		Dim NomTemplate As Variant
		
		Dim WkTemplate As Workbook
		Dim WsFiltre As Worksheet
		
		Dim PathTemplate As String
		Dim ColonneDia As String
		Dim ValeurAGarder As String
		
		Set DictFiltres = CreateObject("Scripting.Dictionary")
		DictFiltres.CompareMode = vbTextCompare
		
		' =========================================================
		' PARCOURS DES TYPES DE COLLABORATEURS
		' DA / DC / DS / ...
		' =========================================================
		
		For Each TypeCollab In _
			DictPrincipal("path_templates") _
			("directory_in_templates").Keys
			
			
			' =====================================================
			' PARCOURS DES TEMPLATES DU TYPE DE COLLABORATEUR
			' =====================================================
			
			For Each NomTemplate In _
				DictPrincipal("path_templates") _
				("directory_in_templates")(TypeCollab) _
				("files_in_path").Keys
				
				PathTemplate = _
					DictPrincipal("path_templates") _
					("directory_in_templates")(TypeCollab) _
					("files_in_path")(NomTemplate) _
					("path_file")
				
				' =================================================
				' OUVERTURE DU TEMPLATE
				' =================================================
				
				Set WkTemplate = Workbooks.Open( _
					Filename:=PathTemplate, _
					ReadOnly:=True)
				
				' =================================================
				' RECHERCHE DE L'ONGLET "filtre"
				' =================================================
				
				Set WsFiltre = Nothing
				
				On Error Resume Next
				Set WsFiltre = WkTemplate.Worksheets("filtre")
				On Error GoTo 0
				
				' =================================================
				' SI L'ONGLET EXISTE
				' =================================================
				
				If Not WsFiltre Is Nothing Then
						
					' =============================================
					' RÉCUPÉRATION DU FILTRE
					'
					' A2 = colonne DIA
					' B2 = valeur à conserver
					' =============================================
					
					ColonneDia = Trim(CStr(WsFiltre.Range("A2").Value))
					ValeurAGarder = Trim(CStr(WsFiltre.Range("B2").Value))
					
					' =============================================
					' CONTRÔLE DU PARAMÉTRAGE
					' =============================================
					
					If ColonneDia = "" Then	
						MsgBox _
							"Le template '" & NomTemplate & _
							"' possède un onglet 'filtre'," & vbCrLf & _
							"mais aucune colonne DIA n'est indiquée en A2.", _
							vbCritical
						
						WkTemplate.Close SaveChanges:=False
						End
					End If
					
					If ValeurAGarder = "" Then
						MsgBox _
							"Le template '" & NomTemplate & _
							"' possède un onglet 'filtre'," & vbCrLf & _
							"mais aucune valeur à garder n'est indiquée en B2.", _
							vbCritical
						WkTemplate.Close SaveChanges:=False
						End
					End If
					
					' =============================================
					' CRÉATION DU SOUS-DICTIONNAIRE
					' =============================================
					
					Set DictFiltreTemplate = _
						CreateObject("Scripting.Dictionary")
					
					DictFiltreTemplate.CompareMode = vbTextCompare
					
					DictFiltreTemplate.Add _
						"colonne_dia", _
						ColonneDia
					
					DictFiltreTemplate.Add _
						"valeur_a_garder", _
						ValeurAGarder
					
					' =============================================
					' AJOUT AU DICTIONNAIRE PRINCIPAL
					'
					' NomTemplate correspond au nom sans extension
					' ex :
					' TVA
					' impot_IS
					' =============================================
					
					If Not DictFiltres.Exists(CStr(NomTemplate)) Then
						
						DictFiltres.Add _
							CStr(NomTemplate), _
							DictFiltreTemplate
						
					Else
						
						' Si le même template existe dans plusieurs
						' catégories DA/DC/DS, on contrôle que son
						' paramétrage est identique.
						
						If _
							StrComp( _
								CStr(DictFiltres(NomTemplate)("colonne_dia")), _
								ColonneDia, _
								vbTextCompare _
							) <> 0 _
						Or _
							StrComp( _
								CStr(DictFiltres(NomTemplate)("valeur_a_garder")), _
								ValeurAGarder, _
								vbTextCompare _
							) <> 0 _
						Then
							
							MsgBox _
								"Le template '" & NomTemplate & _
								"' existe plusieurs fois avec des filtres différents." & _
								vbCrLf & vbCrLf & _
								"Type collaborateur : " & TypeCollab, _
								vbCritical
							
							WkTemplate.Close SaveChanges:=False
							End
							
						End If
						
					End If
					
				End If
				
				' =================================================
				' FERMETURE DU TEMPLATE
				' =================================================
				
				WkTemplate.Close SaveChanges:=False
				
				Set WsFiltre = Nothing
				Set WkTemplate = Nothing
				
			Next NomTemplate
			
		Next TypeCollab
		
		Set GetDictionnaireFiltresTemplates = DictFiltres
	End Function
	
	
	Function GetHeadersList(CheminFichierList() As String) As String()
		Dim WkCible As Workbook
		Dim DerniereCol As Long
		Dim i As Long, c As Long
		Dim Titre As String
		Dim PathFile As String
		
		Dim ListeEnTetes As Collection
		Dim Entetes() As String
		
		Set ListeEnTetes = New Collection
		
		' =========================================================
		' SÉCURITÉ : TABLEAU D'ENTRÉE VIDE
		' =========================================================
		
		If (Not CheminFichierList) = -1 Then
			GetHeadersList = Split("")
			Exit Function
		End If
		 
		' =========================================================
		' PARCOURS DES FICHIERS
		' =========================================================
		
		For i = LBound(CheminFichierList) To UBound(CheminFichierList)
			
			PathFile = CheminFichierList(i)
			
			If Trim(PathFile) <> "" Then
				
				Set WkCible = Nothing
				
				On Error Resume Next
				Set WkCible = Workbooks.Open( _
					Filename:=PathFile, _
					ReadOnly:=True)
				On Error GoTo 0

				If WkCible Is Nothing Then
					MsgBox _
						"Erreur : Impossible d'ouvrir le fichier " & _
						PathFile, _
						vbCritical
				Else
					With WkCible.Sheets(1)
						DerniereCol = _
							.Cells(2, .Columns.Count) _
							.End(xlToLeft).Column

						' =========================================
						' PARCOURS DES COLONNES
						' =========================================
						For c = 1 To DerniereCol
							
							Titre = Trim(CStr( _
								.Cells(2, c) _
								.MergeArea.Cells(1, 1).Value _
							))

							' Sous-titre éventuel ligne 3
							If .Cells(3, c).Value <> "" _
							   And .Cells(3, c).Value <> Titre Then
								
								Titre = Titre & " - " & _
										Trim(CStr(.Cells(3, c).Value))   
							End If
							
							' =====================================
							' ON CONSERVE UNIQUEMENT LES DIA_
							' ET ON IGNORE LES VALEURS VIDES
							' =====================================
							
							If Titre <> "" Then
								If UCase(Left(Titre, 4)) = "DIA_" Then
									ListeEnTetes.Add Titre
								End If
							End If
						Next c
					End With
					
					WkCible.Close SaveChanges:=False
					Set WkCible = Nothing
				End If
			End If
		Next i
		
		' =========================================================
		' CONVERSION COLLECTION -> TABLEAU STRING()
		' =========================================================
		
		If ListeEnTetes.Count = 0 Then
			GetHeadersList = Split("")
		Else
			ReDim Entetes(0 To ListeEnTetes.Count - 1)
			
			For i = 1 To ListeEnTetes.Count
				
				Entetes(i - 1) = CStr(ListeEnTetes(i))
				
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
		WkSource.Close
		
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
		
		' Vérifie la présence du dossier "collaborateurs"
		DossierExiste = FSO.FolderExists(CheminDossier)
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


	Function GetNameFiles(ByRef LstFiles() As String) As String()
		Dim FSO As Object
		Dim Resultat() As String
		Dim i As Long

		Set FSO = CreateObject("Scripting.FileSystemObject")
		ReDim Resultat(LBound(LstFiles) To UBound(LstFiles))

		For i = LBound(LstFiles) To UBound(LstFiles)
			' Extraction du nom sans chemin ni extension
			Resultat(i) = FSO.GetBaseName(LstFiles(i))
		Next i

		GetNameFiles = Resultat
	End Function
	
	
	Function GetColumnNumberByHeader( _
		ByVal Ws As Worksheet, _
		ByVal HeaderName As String, _
		Optional ByVal HeaderRow As Long = 1 _
	) As Long

		Dim CelluleTrouvee As Range
		
		Set CelluleTrouvee = Ws.Rows(HeaderRow).Find( _
			What:=HeaderName, _
			After:=Ws.Cells(HeaderRow, Ws.Columns.Count), _
			LookIn:=xlValues, _
			LookAt:=xlWhole, _
			SearchOrder:=xlByColumns, _
			SearchDirection:=xlNext, _
			MatchCase:=False)
		
		
		If CelluleTrouvee Is Nothing Then
			
			GetColumnNumberByHeader = 0
			
		Else
			
			GetColumnNumberByHeader = CelluleTrouvee.Column
			
		End If

	End Function
	
	
	Function GetTemplateColumnNumber( _
		ByVal Ws As Worksheet, _
		ByVal HeaderToFind As String _
	) As Long

		Dim DerniereCol As Long
		Dim c As Long
		
		Dim Titre As String
		
		DerniereCol = Ws.Cells(2, Ws.Columns.Count) _
						  .End(xlToLeft).Column
		
		
		For c = 1 To DerniereCol
			
			' ==========================================
			' TITRE PRINCIPAL LIGNE 2
			' Gestion des cellules fusionnées
			' ==========================================
			
			Titre = CStr( _
				Ws.Cells(2, c) _
				  .MergeArea.Cells(1, 1).Value _
			)
			
			
			' ==========================================
			' SOUS-TITRE ÉVENTUEL LIGNE 3
			' ==========================================
			
			If Ws.Cells(3, c).Value <> "" _
			   And Ws.Cells(3, c).Value <> Titre Then
				
				Titre = Titre & " - " & _
						CStr(Ws.Cells(3, c).Value)
				
			End If
			
			
			' ==========================================
			' COMPARAISON
			' ==========================================
			
			If StrComp( _
				Trim(Titre), _
				Trim(HeaderToFind), _
				vbTextCompare _
			) = 0 Then
				
				GetTemplateColumnNumber = c
				Exit Function
				
			End If
			
		Next c
		
		
		GetTemplateColumnNumber = 0

	End Function
	
	
	Function DossierDoitEtreInjecte( _
		ByVal NomTemplate As String, _
		ByVal WsSource As Worksheet, _
		ByVal LigneSource As Long, _
		ByVal DictFiltresTemplates As Object _
	) As Boolean

		Dim NomColonneDia As String
		Dim ValeurAGarder As String
		Dim ValeurSource As String
		
		Dim ColFiltre As Long
		
		
		' =========================================================
		' PAR DÉFAUT :
		' LE DOSSIER EST ACCEPTÉ
		' =========================================================
		
		DossierDoitEtreInjecte = True
		
		
		' =========================================================
		' SI LE TEMPLATE N'A PAS DE FILTRE
		' =========================================================
		
		If Not DictFiltresTemplates.Exists(NomTemplate) Then
			Exit Function
		End If
		
		
		' =========================================================
		' RÉCUPÉRATION DU FILTRE
		' =========================================================
		
		NomColonneDia = _
			CStr( _
				DictFiltresTemplates(NomTemplate) _
				("colonne_dia") _
			)
		
		
		ValeurAGarder = _
			CStr( _
				DictFiltresTemplates(NomTemplate) _
				("valeur_a_garder") _
			)
		
		
		' =========================================================
		' RECHERCHE DE LA COLONNE DANS DIA
		' =========================================================
		
		ColFiltre = GetColumnNumberByHeader( _
			WsSource, _
			NomColonneDia, _
			1)
		
		
		If ColFiltre = 0 Then
			
			MsgBox _
				"La colonne utilisée pour le filtre est introuvable dans DIA :" & _
				vbCrLf & vbCrLf & _
				NomColonneDia & vbCrLf & vbCrLf & _
				"Template : " & NomTemplate, _
				vbCritical
			
			End
			
		End If
		
		
		' =========================================================
		' VALEUR DU DOSSIER
		' =========================================================
		
		ValeurSource = Trim(CStr( _
			WsSource.Cells( _
				LigneSource, _
				ColFiltre _
			).Value _
		))
		
		
		' =========================================================
		' COMPARAISON
		' =========================================================
		
		If StrComp( _
			ValeurSource, _
			Trim(ValeurAGarder), _
			vbTextCompare _
		) <> 0 Then
			
			DossierDoitEtreInjecte = False
			
		End If

	End Function
	
	
	Function GetMoisTVA( _
		ByVal TypeTVA As String, _
		ByVal IndexPeriode As Long _
	) As String

		Select Case UCase(Trim(TypeTVA))
			
			Case UCase("CA3 Mensuelle")
				
				Select Case IndexPeriode
					Case 1: GetMoisTVA = "janvier"
					Case 2: GetMoisTVA = "février"
					Case 3: GetMoisTVA = "mars"
					Case 4: GetMoisTVA = "avril"
					Case 5: GetMoisTVA = "mai"
					Case 6: GetMoisTVA = "juin"
					Case 7: GetMoisTVA = "juillet"
					Case 8: GetMoisTVA = "août"
					Case 9: GetMoisTVA = "septembre"
					Case 10: GetMoisTVA = "octobre"
					Case 11: GetMoisTVA = "novembre"
					Case 12: GetMoisTVA = "décembre"
				End Select
			
			
			Case UCase("CA3 Trimestrielle")
				
				Select Case IndexPeriode
					Case 1: GetMoisTVA = "mars"
					Case 2: GetMoisTVA = "juin"
					Case 3: GetMoisTVA = "septembre"
					Case 4: GetMoisTVA = "décembre"
				End Select
			
			
			Case Else
				
				GetMoisTVA = ""
				
		End Select

	End Function
	

	Function FeuilleExiste(ByVal Wk As Workbook, ByVal NomFeuille As String) As Boolean
		Dim Ws As Worksheet
		
		Set Ws = Nothing
		
		On Error Resume Next
		Set Ws = Wk.Worksheets(NomFeuille)
		On Error GoTo 0
		
		FeuilleExiste = Not Ws Is Nothing

	End Function
	
	
	Function DerniereLigneUtilisee(ByVal Ws As Worksheet) As Long
		Dim Cellule As Range
		
		Set Cellule = Ws.Cells.Find( _
			What:="*", _
			After:=Ws.Cells(1, 1), _
			LookAt:=xlPart, _
			LookIn:=xlFormulas, _
			SearchOrder:=xlByRows, _
			SearchDirection:=xlPrevious)
		
		If Cellule Is Nothing Then
			DerniereLigneUtilisee = 0
		Else	
			DerniereLigneUtilisee = Cellule.Row		
		End If
	End Function


	Function DerniereColonneUtilisee(ByVal Ws As Worksheet) As Long
		Dim Cellule As Range
		
		Set Cellule = Ws.Cells.Find( _
			What:="*", _
			After:=Ws.Cells(1, 1), _
			LookAt:=xlPart, _
			LookIn:=xlFormulas, _
			SearchOrder:=xlByColumns, _
			SearchDirection:=xlPrevious)
		
		If Cellule Is Nothing Then
			DerniereColonneUtilisee = 0
		Else
			DerniereColonneUtilisee = Cellule.Column
		End If

	End Function
	
' =========================================================================
' LES FONCTIONS
' =========================================================================



' =========================================================================
' DEBUG DICT
' =========================================================================


	Sub AfficherDictionnaire(ByVal Dict As Object, Optional ByVal Niveau As Long = 0)
        Dim Cle As Variant
        Dim Indentation As String
        Indentation = String(Niveau * 4, " ") ' Décale le texte selon le niveau d'imbrication
    
        If Dict Is Nothing Then Exit Sub
    
        For Each Cle In Dict.Keys
            If IsObject(Dict(Cle)) Then
                Debug.Print Indentation & "[" & Cle & "] (Sous-dictionnaire/Objet)"
                ' Appel récursif pour explorer le sous-dictionnaire
                On Error Resume Next
                AfficherDictionnaire Dict(Cle), Niveau + 1
                On Error GoTo 0
            ElseIf IsArray(Dict(Cle)) Then
                Dim ArrTemp As Variant
                ArrTemp = Dict(Cle)
                
                ' Affiche : Clé : [Elément1, Elément2, Elément3]
                Debug.Print Indentation & Cle & " : [" & Join(ArrTemp, ", ") & "]"
            Else
                Debug.Print Indentation & Cle & " : " & Dict(Cle)
            End If
        Next Cle
    End Sub





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
	
	
	Sub AjouterEntetesAuDictionnaire(ByRef Dict As Object, _
									  ByVal CheminFichierList As Variant, _
									  ByVal Cle As String, _
									  ByVal CleDestination As String)
		Dim WkCible As Workbook
		Dim DictSousEntetes As Object
		Dim ListeEnTetes As Collection
		Dim Entetes() As String
		Dim DerniereCol As Long
		Dim i As Long, c As Long
		Dim Titre As String
		Dim PathFile As String
		Dim FSO As Object
		Dim NomSansExtension As String
		
		Set FSO = CreateObject("Scripting.FileSystemObject")

		Application.ScreenUpdating = False
		
		' 2. Parcours des fichiers
		For i = LBound(CheminFichierList) To UBound(CheminFichierList)
			PathFile = CheminFichierList(i)
			NomSansExtension = FSO.GetBaseName(PathFile)
			Set DictSousEntetes = Dict("path_templates")("directory_in_templates")(Cle)("files_in_path")(NomSansExtension)
			
			If Trim(PathFile) <> "" Then
				Set WkCible = Workbooks.Open(Filename:=PathFile, ReadOnly:=True)
				
				If Not WkCible Is Nothing Then
					Set ListeEnTetes = New Collection
					
					With WkCible.Sheets(1)
						DerniereCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
						
						' Parcours des colonnes
						For c = 1 To DerniereCol
							Titre = CStr(.Cells(2, c).MergeArea.Cells(1, 1).Value)
							
							If .Cells(3, c).Value <> "" And .Cells(3, c).Value <> Titre Then
								Titre = Titre & " - " & .Cells(3, c).Value
							End If
							
							ListeEnTetes.Add Titre
						Next c
					End With
					
					WkCible.Close SaveChanges:=False
					Set WkCible = Nothing
					
					' Conversion de la Collection en tableau String() pour ce fichier, on stock uniquement les headers relatif au fichier DIA_data
					If ListeEnTetes.Count > 0 Then
						ReDim Entetes(0 To ListeEnTetes.Count - 1)
						For c = 1 To ListeEnTetes.Count
							If Left(ListeEnTetes(c), 4) = "DIA_" Then
								Entetes(c - 1) = ListeEnTetes(c)
							End If
						Next c
						
						' Enregistrement du tableau sous le chemin du fichier dans le sous-dictionnaire
						DictSousEntetes(CleDestination) = Entetes
					End If
				End If
			End If
		Next i

		Application.ScreenUpdating = True

	End Sub
	
	
	Sub InjectDataInTemplates( _
		ByVal TypeCollab As String, _
		ByVal NameCollab As String, _
		ByVal PathCollab As String, _
		ByRef NameFilesTemplate() As String, _
		ByVal DictPrincipal As Object, _
		ByVal WkSource As Workbook, _
		ByVal DictFiltresTemplates As Object)

		Dim i As Long
		Dim HeadersToInject As Variant
		
		Dim WkDestination As Workbook
		Dim WsSource As Worksheet
		Dim WsDestination As Worksheet
		
		Dim PathFileDestination As String
		Dim NameFileWithExtension As String
		
		Dim ColFiltreSource As Long
		Dim DerniereLigneSource As Long
		
		Dim LigneSource As Long
		Dim LigneDestination As Long
		
		Dim FSO As Object
		
		Application.ScreenUpdating = False
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set WsSource = WkSource.Sheets(1)

		' =========================================================
		' RECHERCHE DE LA COLONNE DA / DC / DS DANS LE FICHIER DIA
		' =========================================================
		
		ColFiltreSource = GetColumnNumberByHeader( _
			WsSource, _
			TypeCollab, _
			1)
		
		If ColFiltreSource = 0 Then
			
			MsgBox _
				"La colonne '" & TypeCollab & _
				"' est introuvable dans le fichier DIA.", _
				vbCritical
			
			End
			
		End If
		
		DerniereLigneSource = WsSource.Cells( _
			WsSource.Rows.Count, _
			ColFiltreSource _
		).End(xlUp).Row

		' =========================================================
		' PARCOURS DES FICHIERS DU COLLABORATEUR
		' =========================================================
		
		For i = LBound(NameFilesTemplate) To UBound(NameFilesTemplate)
			' =====================================================
			' ON IGNORE LES FICHIERS SUIVI_*
			' =====================================================
			
			If Not (UCase(Left(NameFilesTemplate(i), 6)) = "SUIVI_") Then
				' =================================================
				' RÉCUPÉRATION DES HEADERS DIA_ DU TEMPLATE
				' =================================================
				
				HeadersToInject = _
					DictPrincipal("path_templates") _
					("directory_in_templates")(TypeCollab) _
					("files_in_path")(NameFilesTemplate(i)) _
					("headers_file")
				
				' =================================================
				' NOM RÉEL DU FICHIER AVEC EXTENSION
				' =================================================
				
				NameFileWithExtension = _
					DictPrincipal("path_templates") _
					("directory_in_templates")(TypeCollab) _
					("files_in_path")(NameFilesTemplate(i)) _
					("name")
							
				' =================================================
				' CONSTRUCTION DU CHEMIN DU FICHIER COLLABORATEUR
				' =================================================
				
				PathFileDestination = _
					FSO.BuildPath( _
						PathCollab, _
						NameFileWithExtension _
					)
							
				' =================================================
				' OUVERTURE DU FICHIER
				' =================================================
				
				Set WkDestination = Workbooks.Open( _
					Filename:=PathFileDestination, _
					ReadOnly:=False)
							
				Set WsDestination = WkDestination.Sheets(1)
							
				' =================================================
				' PREMIÈRE LIGNE D'INJECTION
				'
				' Headers lignes 2 / 3
				' Données à partir de la ligne 4
				' =================================================
				
				LigneDestination = 4
							
				' =================================================
				' PARCOURS DES LIGNES DU FICHIER DIA
				' =================================================
				
				For LigneSource = 2 To DerniereLigneSource
									
					' =================================================
					' FILTRE COLLABORATEUR
					' =================================================
					
					If Trim(CStr( _
						WsSource.Cells( _
							LigneSource, _
							ColFiltreSource _
						).Value _
					)) = Trim(NameCollab) Then
												
						' =================================================
						' FILTRE SPÉCIFIQUE AU TEMPLATE
						'
						' Si aucun filtre n'existe pour le template :
						' DossierDoitEtreInjecte = True
						' =================================================
						
						If DossierDoitEtreInjecte( _
							NameFilesTemplate(i), _
							WsSource, _
							LigneSource, _
							DictFiltresTemplates _
						) Then
														
							' =============================================
							' CHOIX DU TRAITEMENT
							' =============================================
							
							If LCase(Trim(NameFilesTemplate(i))) = "tva" Then
																
								' =========================================
								' TRAITEMENT SPÉCIFIQUE TVA
								'
								' Mensuelle :
								' 12 lignes
								'
								' Trimestrielle :
								' 4 lignes
								'
								' Autres :
								' 1 ligne
								' =========================================
								
								InjectDataTVA _
									WsSource, _
									WsDestination, _
									LigneSource, _
									LigneDestination, _
									HeadersToInject, _
									WkSource
								
							Else
								
								' =========================================
								' TRAITEMENT STANDARD
								' =========================================
								
								InjectDataStandard _
									WsSource, _
									WsDestination, _
									LigneSource, _
									LigneDestination, _
									HeadersToInject, _
									WkSource
								
								
							End If
	
						End If
	
					End If
					
				Next LigneSource
				
				' =================================================
				' ENREGISTREMENT DU FICHIER COLLABORATEUR
				' =================================================
				
				WkDestination.Close SaveChanges:=True
								
				Set WsDestination = Nothing
				Set WkDestination = Nothing
					
			End If
			
		Next i
		
		Application.ScreenUpdating = True

	End Sub


	Sub InjectDataTVA( _
		ByVal WsSource As Worksheet, _
		ByVal WsDestination As Worksheet, _
		ByVal LigneSource As Long, _
		ByRef LigneDestination As Long, _
		ByVal HeadersToInject As Variant, _
		ByVal WkSource As Workbook)

		Dim j As Long
		Dim IndexDuplication As Long
		Dim NbDuplications As Long
		
		Dim HeaderTemplate As String
		Dim HeaderDia As String
		
		Dim ColSource As Long
		Dim ColDestination As Long
		Dim ColTypeTVA As Long
		Dim ColMois As Long
		
		Dim TypeTVA As String
		Dim MoisAAffecter As String
		
		' =========================================================
		' RECHERCHE TYPE DE TVA
		' =========================================================
		
		ColTypeTVA = GetColumnNumberByHeader( _
			WsSource, _
			"Type de TVA", _
			1)
		
		If ColTypeTVA = 0 Then
			
			MsgBox _
				"La colonne 'Type de TVA' est introuvable dans DIA.", _
				vbCritical
			
			End
			
		End If
		
		TypeTVA = Trim(CStr( _
			WsSource.Cells( _
				LigneSource, _
				ColTypeTVA _
			).Value _
		))
		
		' =========================================================
		' RECHERCHE COLONNE MOIS DANS LE TEMPLATE TVA
		' =========================================================
		
		ColMois = GetTemplateColumnNumber( _
			WsDestination, _
			"mois")
		
		If ColMois = 0 Then
			
			MsgBox _
				"La colonne 'mois' est introuvable dans le template TVA.", _
				vbCritical
			
			End
			
		End If
		
		' =========================================================
		' NOMBRE DE LIGNES À CRÉER
		' =========================================================
		
		Select Case UCase(TypeTVA)
			
			Case UCase("CA3 Mensuelle")
				NbDuplications = 12
			
			Case UCase("CA3 Trimestrielle")
				NbDuplications = 4
			
			Case Else
				NbDuplications = 1
				
		End Select
		
		' =========================================================
		' CRÉATION DES LIGNES
		' =========================================================
		
		For IndexDuplication = 1 To NbDuplications
			
			' =====================================================
			' INJECTION DES COLONNES DIA_
			' =====================================================
			
			For j = LBound(HeadersToInject) To UBound(HeadersToInject)
				
				HeaderTemplate = Trim(CStr(HeadersToInject(j)))
				
				If HeaderTemplate <> "" Then
					
					If UCase(Left(HeaderTemplate, 4)) = "DIA_" Then
						
						HeaderDia = Mid(HeaderTemplate, 5)
						
						ColSource = GetColumnNumberByHeader( _
							WsSource, _
							HeaderDia, _
							1)
						
						If ColSource = 0 Then
							
							MsgBox _
								"Colonne DIA introuvable :" & vbCrLf & _
								HeaderDia & vbCrLf & vbCrLf & _
								"Fichier : " & WkSource.Name, _
								vbCritical
							
							End
							
						End If
						
						ColDestination = GetTemplateColumnNumber( _
							WsDestination, _
							HeaderTemplate)
						
						If ColDestination = 0 Then
							
							MsgBox _
								"Colonne template introuvable :" & vbCrLf & _
								HeaderTemplate & vbCrLf & vbCrLf & _
								"Fichier : " & WsDestination.Parent.Name, _
								vbCritical
							
							End
							
						End If
						
						WsDestination.Cells( _
							LigneDestination, _
							ColDestination _
						).Value = _
							WsSource.Cells( _
								LigneSource, _
								ColSource _
							).Value
						
					End If
					
				End If
				
			Next j
			
			' =====================================================
			' MOIS ASSOCIÉ
			' =====================================================
			
			MoisAAffecter = GetMoisTVA( _
				TypeTVA, _
				IndexDuplication)
			
			WsDestination.Cells( _
				LigneDestination, _
				ColMois _
			).Value = MoisAAffecter
			
			LigneDestination = LigneDestination + 1
			
		Next IndexDuplication

	End Sub
	
	
	Sub InjectDataStandard( _
		ByVal WsSource As Worksheet, _
		ByVal WsDestination As Worksheet, _
		ByVal LigneSource As Long, _
		ByRef LigneDestination As Long, _
		ByVal HeadersToInject As Variant, _
		ByVal WkSource As Workbook)

		Dim j As Long
		
		Dim HeaderTemplate As String
		Dim HeaderDia As String
		
		Dim ColSource As Long
		Dim ColDestination As Long
		
		For j = LBound(HeadersToInject) To UBound(HeadersToInject)
			
			HeaderTemplate = Trim(CStr(HeadersToInject(j)))
			
			If HeaderTemplate <> "" Then
				
				If UCase(Left(HeaderTemplate, 4)) = "DIA_" Then
					
					HeaderDia = Mid(HeaderTemplate, 5)
					
					ColSource = GetColumnNumberByHeader( _
						WsSource, _
						HeaderDia, _
						1)
					
					If ColSource = 0 Then
						
						MsgBox _
							"Colonne DIA introuvable :" & vbCrLf & _
							HeaderDia & vbCrLf & vbCrLf & _
							"Fichier : " & WkSource.Name, _
							vbCritical
						
						End
						
					End If
					
					ColDestination = GetTemplateColumnNumber( _
						WsDestination, _
						HeaderTemplate)
					
					If ColDestination = 0 Then
						
						MsgBox _
							"Colonne template introuvable :" & vbCrLf & _
							HeaderTemplate & vbCrLf & vbCrLf & _
							"Fichier : " & WsDestination.Parent.Name, _
							vbCritical
						
						End
						
					End If
					
					WsDestination.Cells( _
						LigneDestination, _
						ColDestination _
					).Value = _
						WsSource.Cells( _
							LigneSource, _
							ColSource _
						).Value
					
				End If
				
			End If
			
		Next j
		
		LigneDestination = LigneDestination + 1

	End Sub
	
' =========================================================================
' LES SUB
' =========================================================================

	Sub CentraliserDonneesCollaborateurs( _
		ByVal DictPrincipal As Object)

		Dim FSO As Object
		
		Dim PathPrincipal As String
		Dim PathRacine As String
		Dim PathCollaborateurs As String
		Dim PathArchives As String
		Dim PathFichierCentral As String
		
		Dim WkCentral As Workbook
		Dim WkSource As Workbook
		
		Dim WsSource As Worksheet
		Dim WsDestination As Worksheet
		
		Dim DossierType As Object
		Dim DossierCollaborateur As Object
		Dim Fichier As Object
		
		Dim NomFichierSansExtension As String
		Dim NomOngletDestination As String
		
		Dim DerniereLigneSource As Long
		Dim DerniereColSource As Long
		Dim LigneDestination As Long
		
		Dim DictPremiereSource As Object
		
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		
		' =========================================================
		' CHEMINS
		' =========================================================
		
		PathPrincipal = DictPrincipal("path_principal")
		
		PathRacine = _
			FSO.GetParentFolderName(PathPrincipal)
		
		PathCollaborateurs = _
			FSO.BuildPath( _
				PathRacine, _
				"collaborateurs" _
			)
		
		PathArchives = _
			FSO.BuildPath( _
				PathPrincipal, _
				"archives" _
			)
		
		' =========================================================
		' CRÉATION DU DOSSIER ARCHIVES SI NÉCESSAIRE
		' =========================================================
		
		If Not FSO.FolderExists(PathArchives) Then
			
			FSO.CreateFolder PathArchives
			
		End If
		
		
		' =========================================================
		' FICHIER CENTRAL
		' =========================================================
		
		PathFichierCentral = _
			FSO.BuildPath( _
				PathArchives, _
				"Centralisation.xlsx" _
			)
		
		
		' =========================================================
		' OUVERTURE / CRÉATION DU FICHIER CENTRAL
		' =========================================================
		
		If FSO.FileExists(PathFichierCentral) Then
			
			Set WkCentral = Workbooks.Open( _
				Filename:=PathFichierCentral, _
				ReadOnly:=False _
			)
			
		Else
			
			Set WkCentral = Workbooks.Add
			
			
			' On garde temporairement une seule feuille.
			Do While WkCentral.Worksheets.Count > 1
				
				WkCentral.Worksheets( _
					WkCentral.Worksheets.Count _
				).Delete
				
			Loop
			
			WkCentral.Worksheets(1).Name = "__temp__"
			
			WkCentral.SaveAs _
				Filename:=PathFichierCentral, _
				FileFormat:=xlOpenXMLWorkbook
			
		End If
		
		' =========================================================
		' NETTOYAGE DES FEUILLES EXISTANTES
		'
		' On repart de zéro à chaque centralisation
		' =========================================================
		
		Application.DisplayAlerts = False
		
		Do While WkCentral.Worksheets.Count > 1
			
			WkCentral.Worksheets( _
				WkCentral.Worksheets.Count _
			).Delete
			
		Loop
		
		WkCentral.Worksheets(1).Cells.Clear
		WkCentral.Worksheets(1).Name = "__temp__"
		
		Application.DisplayAlerts = True
		
		' =========================================================
		' DICTIONNAIRE
		'
		' Clé = nom du template
		' Valeur = True si aucune source encore injectée
		' =========================================================
		
		Set DictPremiereSource = _
			CreateObject("Scripting.Dictionary")
		
		DictPremiereSource.CompareMode = vbTextCompare
		
		' =========================================================
		' PARCOURS DES TYPES DE COLLABORATEURS
		' =========================================================
		
		For Each DossierType In _
			FSO.GetFolder(PathCollaborateurs).SubFolders
			
			' =====================================================
			' PARCOURS DES COLLABORATEURS
			' =====================================================
			
			For Each DossierCollaborateur In _
				DossierType.SubFolders
				
				' =================================================
				' PARCOURS DE TOUS LES FICHIERS
				' =================================================
				
				For Each Fichier In _
					DossierCollaborateur.Files
					
					
					NomFichierSansExtension = _
						FSO.GetBaseName(Fichier.Name)
					
					' =================================================
					' ON IGNORE LES FICHIERS COMMENÇANT PAR suivi_
					' =================================================
					
					If UCase(Left( _
						Trim(NomFichierSansExtension), _
						6 _
					)) <> "SUIVI_" Then
					
						' =============================================
						' LE NOM DU FICHIER DEVIENT LE NOM DE L'ONGLET
						' =============================================
						
						NomOngletDestination = NomFichierSansExtension
						
						' =============================================
						' CRÉATION DE L'ONGLET SI NÉCESSAIRE
						' =============================================
						
						If Not FeuilleExiste( _
							WkCentral, _
							NomOngletDestination _
						) Then
							
							
							Set WsDestination = _
								WkCentral.Worksheets.Add( _
									After:=WkCentral.Worksheets( _
										WkCentral.Worksheets.Count _
									) _
								)
							
							WsDestination.Name = _
								NomOngletDestination
							
							
							DictPremiereSource.Add _
								NomOngletDestination, _
								True
							
						Else
							
							Set WsDestination = _
								WkCentral.Worksheets( _
									NomOngletDestination _
								)
							
							
							If Not DictPremiereSource.Exists( _
								NomOngletDestination _
							) Then
								
								DictPremiereSource.Add _
									NomOngletDestination, _
									True
								
							End If
							
						End If

						' =============================================
						' OUVERTURE DU FICHIER SOURCE
						' =============================================
						
						Set WkSource = Nothing
						
						On Error Resume Next
						
						Set WkSource = Workbooks.Open( _
							Filename:=Fichier.Path, _
							ReadOnly:=True _
						)
						
						On Error GoTo 0
						
						If WkSource Is Nothing Then
							
							MsgBox _
								"Impossible d'ouvrir le fichier :" & _
								vbCrLf & vbCrLf & _
								Fichier.Path, _
								vbExclamation
							
						Else
							
							' =========================================
							' PREMIER ONGLET
							' =========================================
							
							Set WsSource = _
								WkSource.Worksheets(1)
							
							DerniereLigneSource = _
								DerniereLigneUtilisee( _
									WsSource _
								)
							
							DerniereColSource = _
								DerniereColonneUtilisee( _
									WsSource _
								)
							
							If DerniereLigneSource > 0 _
							   And DerniereColSource > 0 Then

								' =====================================
								' PREMIÈRE SOURCE POUR CE TEMPLATE
								'
								' COPIE DE TOUT :
								' lignes 1 à dernière ligne
								' =====================================
								
								If DictPremiereSource( _
									NomOngletDestination _
								) = True Then
									
									
									WsDestination.Cells( _
										1, _
										1 _
									).Resize( _
										DerniereLigneSource, _
										DerniereColSource _
									).Value = _
									
									WsSource.Cells( _
										1, _
										1 _
									).Resize( _
										DerniereLigneSource, _
										DerniereColSource _
									).Value
									
									
									DictPremiereSource( _
										NomOngletDestination _
									) = False
									
								Else
									
									' =================================
									' SOURCES SUIVANTES
									'
									' DONNÉES À PARTIR DE LA LIGNE 4
									' =================================
									
									If DerniereLigneSource >= 4 Then
											
										LigneDestination = _
											DerniereLigneUtilisee( _
												WsDestination _
											) + 1
										
										
										WsDestination.Cells( _
											LigneDestination, _
											1 _
										).Resize( _
											DerniereLigneSource - 3, _
											DerniereColSource _
										).Value = _
										
										WsSource.Cells( _
											4, _
											1 _
										).Resize( _
											DerniereLigneSource - 3, _
											DerniereColSource _
										).Value
												
									End If
	
								End If

							End If							
							
							WkSource.Close SaveChanges:=False
							
							Set WsSource = Nothing
							Set WkSource = Nothing						
							
						End If
											
					End If
									
				Next Fichier
							
			Next DossierCollaborateur
				
		Next DossierType
		
		' =========================================================
		' SUPPRESSION DE LA FEUILLE TEMPORAIRE
		' =========================================================
		
		If FeuilleExiste(WkCentral, "__temp__") Then
			
			' On ne peut pas supprimer la dernière feuille d'un classeur
			If WkCentral.Worksheets.Count > 1 Then
				
				Application.DisplayAlerts = False
				
				WkCentral.Worksheets("__temp__").Delete
				
				Application.DisplayAlerts = True
				
			End If
			
		End If
		
		' =========================================================
		' ENREGISTREMENT
		' =========================================================
		
		WkCentral.Save
		
		WkCentral.Close SaveChanges:=False
		
		Set WkCentral = Nothing

	End Sub
