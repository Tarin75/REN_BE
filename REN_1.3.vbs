Dim numero, nomDossier, objFSO, repertoire, repertoirePath, extensions, prefixe, excludedPrefixes, extension, file, nouveauNom
Dim fichiersOuverts, readOnlyFiles

' Ouvre une boîte de dialogue pour saisir un numéro à 6 chiffres
numero = InputBox("Entrez le dossier (6 chiffres) : ", "BE Frameries - Renommage fichiers")

' Vérifie si le numéro saisi contient exactement 6 chiffres
If Len(numero) <> 6 Then
    MsgBox "Le numéro saisi doit contenir exactement 6 chiffres.", vbExclamation, "Erreur"
    WScript.Quit 1
End If

' Crée le nom du dossier en utilisant le numéro saisi
nomDossier = "W:\06_0040\70_CRONOS\01_ANNEXES\02_RML\" & numero

' Vérifie si le dossier existe et le crée si nécessaire
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(nomDossier) Then
    objFSO.CreateFolder nomDossier
End If

' Liste des répertoires à renommer
repertoires = Array("3_Plans", "5_Documents", "6_Echanges", "1_Budget", "1_Budget\1_commandes", "4_Photos\1_avant")

' Liste des extensions de fichiers à renommer dans les répertoires 1, 4, 5 et 6
extensionsDocuments = Array(".pdf", ".docx", ".msg", ".xlsx", ".xlsm")

' Liste des extensions de fichiers à renommer dans le répertoire 3_Plans
extensionsPlans = Array(".dwg", ".dwf", ".pdf")

' Préfixe pour les fichiers renommés
prefixe = "BE_" & numero & "_PLAN"

' Fichiers à exclure du renommage
excludedPrefixes = Array("STVLT_", "BE_", "TR_", "BC_", "EX_", "BD_")

' Parcours les répertoires
For Each repertoire In repertoires
    ' Construit le chemin du sous-répertoire
    repertoirePath = nomDossier & "\" & repertoire

    ' Vérifie si le sous-répertoire existe, sinon le crée
    If Not objFSO.FolderExists(repertoirePath) Then
        objFSO.CreateFolder repertoirePath
    End If

    ' Liste des extensions de fichiers à renommer en fonction du répertoire
    If repertoire = "3_Plans" Then
        extensions = extensionsPlans
    Else
        extensions = extensionsDocuments
        ' Retire "_PLAN" du préfixe pour les répertoires ajoutés
        prefixe = "BE_" & numero & "_"
    End If

    ' Parcours les extensions et renomme les fichiers si le préfixe n'existe pas déjà
    For Each extension In extensions
        For Each file In objFSO.GetFolder(repertoirePath).Files
            ' Vérifie si le fichier est en lecture seule, vérouillé ou ouvert par une application
            If Not IsFileLocked(file.Path) Then
                If Not ArrayContains(excludedPrefixes, UCase(Left(file.Name, Len(excludedPrefixes(0))))) Then
                    If objFSO.GetExtensionName(file.Name) = Mid(extension, 2) And Not Left(file.Name, Len(prefixe)) = prefixe Then
                        ' Si le répertoire est 3_Plans et que le nom de fichier commence par "PLAN", renommer uniquement avec le préfixe
                        If repertoire = "3_Plans" And Left(file.Name, 4) = "PLAN" Then
                            nouveauNom = prefixe & Right(file.Name, Len(file.Name) - 4)
                        Else
                            ' Supprime le _ en fin de nom de fichier s'il existe
                            Dim fileNameWithoutExtension, fileExtension
                            fileNameWithoutExtension = Left(file.Name, Len(file.Name) - Len(extension))
                            fileExtension = objFSO.GetExtensionName(file.Name)
                            If Right(fileNameWithoutExtension, 1) = "_" Then
                                fileNameWithoutExtension = Left(fileNameWithoutExtension, Len(fileNameWithoutExtension) - 1)
                            End If
                            ' Construit le nouveau nom de fichier avec le préfixe
                            nouveauNom = prefixe & Replace(fileNameWithoutExtension, numero, "") & extension
                        End If
                        ' Vérifie si le fichier avec le nouveau nom existe déjà
                        If Not objFSO.FileExists(repertoirePath & "\" & nouveauNom) Then
                            ' Renomme le fichier
                            file.Name = nouveauNom
                        End If
                    End If
                End If
            Else
                ' Ajoute le nom du fichier à la liste des fichiers en lecture seule ou ouverts
                If readOnlyFiles = "" Then
                    readOnlyFiles = file.Name
                Else
                    readOnlyFiles = readOnlyFiles & vbCrLf & file.Name
                End If
            End If
        Next
    Next
Next

' Affiche une boîte de message pour les fichiers en lecture seule ou ouverts
If readOnlyFiles <> "" Then
    MsgBox "Les fichiers suivants sont en lecture seule ou ouverts par une application et n'ont pas été renommés : " & vbCrLf & readOnlyFiles, vbExclamation, "Fichiers en lecture seule ou ouverts"
End If

' Ouvre l'explorateur de fichiers dans le répertoire où se trouvent les fichiers renommés
CreateObject("WScript.Shell").Run "explorer.exe " & nomDossier, 1, True

WScript.Quit 0

' Fonction pour vérifier si un élément existe dans un tableau
Function ArrayContains(arr, item)
    Dim element
    For Each element In arr
        If UCase(element) = UCase(item) Then
            ArrayContains = True
            Exit Function
        End If
    Next
    ArrayContains = False
End Function

' Fonction pour vérifier si un fichier est en lecture seule, vérouillé ou ouvert par une application
Function IsFileLocked(file)
    Dim fso, locked
    locked = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    ' Essaye d'ouvrir le fichier en mode lecture-écriture
    Dim dummy: Set dummy = fso.OpenTextFile(file, 8, True)
    If Err.Number <> 0 Then
        locked = True
    End If
    On Error GoTo 0
    IsFileLocked = locked
End Function
