Sub refresh_data_dia()
	Dim DictPrincipal As Object
	Set DictPrincipal = CreateObject("Scripting.Dictionary")

    ' récupération du path actuelle de la macro
	Dim CheminDossierPrincipal As String
	Dim ClasseurPrincipal As Workbook
	Set ClasseurPrincipal = ThisWorkbook
    CheminDossierPrincipal = ClasseurPrincipal.Path
	
	DictPrincipal.Add "path_principal", CheminDossierPrincipal
	
	' validation de l'existence du dossier "templates"
	Const DirectoryTemplates As String = "templates"
	Dim CheminDossierTemplates As String
	CheminDossierTemplates = get_path_directory(CheminDossierPrincipal, DirectoryTemplates)
	
	DictPrincipal.Add "path_templates", CreateObject("Scripting.Dictionary")
	DictPrincipal("path_templates").Add "path", CheminDossierTemplates
	DictPrincipal("path_templates").Add "directory_in_templates", CreateObject("Scripting.Dictionary")
	
	Dim CheminDirectoryTemplates() As String
	CheminDirectoryTemplates = GetSubFoldersList(CheminDossierTemplates)
	
	AddFoldersToDictionary(DictPrincipal, CheminDirectoryTemplates)
	
	MsgBox Join(CheminDirectoryTemplates, ", ")
	
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
	
    ' validation de l'existence du dossier "templates"
	Const DirectoryTemplates As String = "templates"
	Dim CheminDossierTemplates As String
	CheminDossierTemplates = get_path_directory(CheminDossierPrincipal, DirectoryTemplates)

    ' validation de l'existence des dossiers "DA", "DC", "DS"
	Const DirectorytemplatesDa As String = "DA"
	Dim CheminDossierTemplatesDa As String
	CheminDossierTemplatesDa = get_path_directory(CheminDossierTemplates, DirectorytemplatesDa)
	
	Const DirectorytemplatesDc As String = "DC"
	Dim CheminDossierTemplatesDc As String
	CheminDossierTemplatesDc = get_path_directory(CheminDossierTemplates, DirectorytemplatesDc)
	
	Const DirectorytemplatesDs As String = "DS"
	Dim CheminDossierTemplatesDs As String
	CheminDossierTemplatesDs = get_path_directory(CheminDossierTemplates, DirectorytemplatesDs)
			
    ' validation de la présence d'un fichier excel dans le dossier "DA" = "TVA.xlsx"
	Const TvaTemplates As String = "TVA.xlsx"
	Dim CheminTvaTemplates As String
	CheminTvaTemplates = GetPathFile(CheminDossierTemplatesDa, TvaTemplates)

    ' validation de la présence des fichiers excel dans le dossier "DC" = "accompte_IS.xlsx", "liasse.xlsx", "suivi_dc.xlsx"
	Const AcompteIsTemplates As String = "acompte_IS.xlsx"
	Dim CheminAcompteIsTemplates As String
	CheminAcompteIsTemplates = GetPathFile(CheminDossierTemplatesDc, AcompteIsTemplates)
	
	Const LiasseTemplates As String = "liasse.xlsx"
	Dim CheminLiasseTemplates As String
	CheminLiasseTemplates = GetPathFile(CheminDossierTemplatesDc, LiasseTemplates)
	
	Const SuiviDcTemplates As String = "suivi_dc.xlsx"
	Dim CheminSuiviDcTemplates As String
	CheminSuiviDcTemplates = GetPathFile(CheminDossierTemplatesDc, SuiviDcTemplates)

    ' validation de la présence d'un fichier excel dans le dossier "DS" = "suivi_ds.xlsx"
	Const SuiviDsTemplates As String = "suivi_ds.xlsx"
	Dim CheminSuiviDsTemplates As String
	CheminSuiviDsTemplates = GetPathFile(CheminDossierTemplatesDs, SuiviDsTemplates)
	
    ' récupération des nom de colonne du fichier commençant par "DIA_" dans les différents templates ne commencant pas par "suivi_"
	Dim LstTvaHeaders() As String
	LstTvaHeaders = GetHeadersList(CheminTvaTemplates)
	
	
    ' validation de l'existence du dossier "data_dia"
    ' validation de la présence d'un fichier excel dans le dossier "data_dia"
    ' lecture du fichier, contrôle des colonnes obligatoires dans le fichier depuis la liste récupérée des templates commencant par "DIA_"
    ' on remonte d'un niveau dans notre arborescence
    ' validation de l'existence du dossier "collaborateurs"
    ' si dossier "collaborateurs" est inexistant on crée le dossier et l'arborescence depuis les données DIA des DA, DC, DS
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
	
	
	Function AddFoldersToDictionary(DictDestination As Object, ListeDossiers() As String)
		Dim FSO As Object
		Dim i As Long
		Dim NomDossier As String
		
		' Vérifie que le tableau n'est pas vide
		If (Not ListeDossiers) = -1 Then Exit Sub
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		For i = LBound(ListeDossiers) To UBound(ListeDossiers)
			If ListeDossiers(i) <> "" Then
				' Extraction automatique du dernier dossier du chemin (ex: "C:\A\B\Templates" -> "Templates")
				NomDossier = FSO.GetFolder(ListeDossiers(i)).Name
				
				' Ajoute au dictionnaire s'il n'existe pas déjà pour éviter les doublons
				If Not DictDestination("path_templates")("directory_in_templates").Exists(NomDossier) Then
					' Option A : Ajouter directement la chaîne du chemin
					DictDestination("path_templates")("directory_in_templates").Add NomDossier, ListeDossiers(i)
					
					' Option B (Si vous voulez un sous-dictionnaire pour chaque dossier) :
					DictDestination("path_templates")("directory_in_templates").Add NomDossier, CreateObject("Scripting.Dictionary")
					DictDestination("path_templates")("directory_in_templates")(NomDossier).Add "path", ListeDossiers(i)
				End If
			End If
		Next i
		
		AddFoldersToDictionary = DictDestination
	End Function
		
	
	Function GetHeadersList(CheminFichier As String) As String()
		Dim WkCible As Workbook
		Dim DerniereCol As Long
		Dim i As Long
		Dim Entetes() As String
		Dim Titre As String
		
		Application.ScreenUpdating = False
		
		' Ouverture en lecture seule
		On Error Resume Next
		Set WkCible = Workbooks.Open(Filename:=CheminFichier, ReadOnly:=True)
		On Error GoTo 0
		
		If WkCible Is Nothing Then
			MsgBox "Erreur : Impossible d'ouvrir le fichier " & CheminFichier, vbCritical
			Exit Function
		End If
		
		With WkCible.Sheets(1)
			' Détermine la dernière colonne à partir de la ligne 2
			DerniereCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
			
			' Dimensionne le tableau de sortie de 0 à (DerniereCol - 1)
			ReDim Entetes(0 To DerniereCol - 1)
			
			For i = 1 To DerniereCol
				' Récupère la valeur de la cellule principale (ligne 2)
				Titre = .Cells(2, i).MergeArea.Cells(1, 1).Value
				
				' Si la ligne 3 contient un sous-titre distinct (ex: colonnes EDI / ECF), on le combine
				If .Cells(3, i).Value <> "" And .Cells(3, i).Value <> Titre Then
					Titre = Titre & " - " & .Cells(3, i).Value
				End If
				
				Entetes(i - 1) = Titre
			Next i
		End With
		
		' Fermeture du fichier source
		WkCible.Close SaveChanges:=False
		Application.ScreenUpdating = True

		GetHeadersList = Entetes
	End Function

' =========================================================================
' LES FONCTIONS 
' =========================================================================


' =========================================================================
' LES SUB
' =========================================================================



' =========================================================================
' LES SUB
' =========================================================================