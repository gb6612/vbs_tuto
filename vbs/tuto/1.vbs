' 1) Constantes
' 2) Variables

' 1) Une constante est une valeur qu'on assigne et on ne change plus 
'    Ca peut etre des nombres, comme du texte
Const M_PI = 3.14
Const TEXTE1 = "Soleil"

WScript.StdOut.WriteLine M_PI  ' ici on affiche sa valeur

'2) Une variable est un emplacement utilisé pour contenir une valeur qui peut être modifiée 
'   pendant l'exécution du script
'Règles de déclaration des variables :
'   - Le nom de la variable doit commencer par un alphabet
'   - Les noms de variables ne peuvent pas dépasser 255 caractères
'   - Les variables ne doivent PAS contenir de point (.)
'   - Les noms de variables doivent être uniques dans le contexte déclaré

' Declaration d'une variable
Dim var0

' On donne une valeur à la variable
var0 = 123   ' ici on donne à la variable un numero
var0 = "Bonjour"  ' et ici un bout de texte

WScript.StdOut.WriteLine var0   ' Ici on affiche la valeur de la variable...

'WScript.StdOut.WriteLine ""

' On peut combiner du texte fixe avec des variables, en utilisant le charactere &
var0 = "belle"
WScript.StdOut.WriteLine "Quelle " & var0 & " journee"
var0 = "longue"
WScript.StdOut.WriteLine "Quelle " & var0 & " journee"

' Une variable peut etre un nombre aussi
Dim number1
Dim number2 
number1 = 15
number2 = 3
WScript.StdOut.WriteLine number1 * number2  ' et ici par exemple on affiche le resultat d'un calcul



