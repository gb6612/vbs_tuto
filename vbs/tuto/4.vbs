' 1) Procedures et Fonctions

' On utilise les procedures pour executer un certain nombre de commandes de facon repetitive

' Declaration d'une procedure
' Ici une procedure pour afficher un texte (str). On peut facilement mettre ou enlever le commentaire
' pour choisir de quelle facon afficher le texte
Sub affiche(str)
	'msgbox str,vbOKOnly,"Ivy Says"
   WScript.StdOut.WriteLine str
End Sub

affiche("1 km a pied")
affiche("ca use ca use")
affiche("2 km a pied")
affiche("ca fait mal a mes souliers")

WScript.StdOut.WriteLine
WScript.StdOut.WriteLine ("Fonctions")
' Une fonction est un type particulier de procedure, quand on a besoin de retourner une valeur

Function mon_addition(num1, num2)
      ' on retourne une valeur avec le meme nom de la fonction
      mon_addition = num1 + num2
End Function

resultat = mon_addition(5, 3)
WScript.StdOut.WriteLine resultat

