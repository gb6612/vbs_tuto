' Strings
'https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/string-functions

str0 = "Il etait une fois une BELLE journee"

WScript.StdOut.WriteLine(UCase(str0)) ' ecrit la phrase en majuscule
WScript.StdOut.WriteLine(LCase(str0)) ' ecrit la phrase en minuscule

WScript.StdOut.WriteLine("il y a " & Len(str0) & " caracteres") ' compte le nombre de caracteres

WScript.StdOut.WriteLine("les 5 premiers caracteres sont :" & Left(str0, 5)) 
WScript.StdOut.WriteLine("les 5 derniers caracteres sont :" & Right(str0, 5)) 

' retourne la position d'un texte String2 dans la phrase String1
' InStr([start,]string1,string2[,compare])
' Si rien n'est trouvé, la fct retourne 0
WScript.StdOut.WriteLine("Fonction InStr: " & InStr(str0, "une")) 

' retourne le texte à la position 'start' et d'une longueur 'length'
' Mid(string, start, [ length ])
WScript.StdOut.WriteLine("Fonction Mid: " & Mid(str0, 4, 5)) 

str0 = Replace(str0, "BELLE", "jolie")
WScript.StdOut.WriteLine("Fonction Replace: " & str0) 


' et Aussi...
  ' LTrim(texte)  pour enlever les espaces en debut de texte
  ' RTrim(texte)  pour enlever les espaces en fin de texte


' Split (string? Expression, string? Delimiter = " ", int Limit
' separe un texte en plusieurs parties, par defaut chaque espace
' voir aussi prochaine lecon...
ft = Split(str0) 
i=0
do while i <= UBound(ft)
     'Ajoute un numero de ligne...
     WScript.StdOut.WriteLine(i & " : " & ft(i))
     i=i+1
loop
