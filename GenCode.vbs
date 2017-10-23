'************************** Procédure Génération Mot de Passe**********************************

'************* Fonction Qui donne le code d'accès*****************

Sub GenererNombre(max,min,variable)
Dim code
Randomize
code = Int((max-min+1)*Rnd+min)
variable = code
'WScript.Echo code
end Sub


'*************Formater Code*********************************
Public Function FormatSpc(Quoi)
FormatSpc = String(4 - Len(Quoi), "0") & Quoi
End Function


'*******************Choisir Magasin et Nom de Vendeur (pour enregistrer dans le registre)**********************

Sub ChoixVendeurMag()
nomMagasin = InputBox("Quel magasin voulez vous ?", "Choix Magasin","Tapez le nom du magasin en majuscule...")
nomVendeur = InputBox("Entrer le nom du vendeur", "Choix Vendeur","Tapez le nom du Vendeur en majuscule...")

'Creation fichier magasin si existe pas encore
'*** Préparation de l'environnement
Const ForAppending = 8
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
FichierTXT = nomMagasin +".txt"

'*** Création du fichier texte "magasin.txt"
Set NewFichier = fso.OpenTextFile(FichierTXT, ForAppending, True)
Dim codepass
Dim MyVar
Dim Reponse 
MyVar = Now

Set Shell = CreateObject("wscript.Shell")
Set env = Shell.environment("Process")

strComputer = env.Item("Computername")


'*** Destruction des objets
Set Shell = Nothing
Set env = Nothing

Reponse = MsgBox(nomVendeur & " a t-il deja un code ?", vbYesNo + vbQuestion, "Code Existant? ")
If Reponse = VbYes Then
codepass = InputBox("Entrer le code de" & nomVendeur, "Choix Code","Tapez le code de " & nomVendeur & " ...")

Else

Call GenererNombre(0000,9999,codepass)
codepass = FormatSpc(codepass)

End If


'*** Ajout de données dans la variable "data".
data = MyVar & "|| " & strComputer & " Vendeur : " & nomVendeur & " | " & " Code : " & codepass
MsgBox("Code " & nomVendeur & " " & codepass)

'*** Ecriture des données de la variable "data" dans le fichier texte.
NewFichier.WriteLine(data)

'*** Destruction des objets
Set fso = Nothing

End Sub

Call ChoixVendeurMag()
WScript.Quit
