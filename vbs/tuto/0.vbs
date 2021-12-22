' Voici un commentaire
' qui est tout simplement du texte pas compilé

' Dans un terminal powershell on peut lancer ce script de la maniere suivante:
' cscript tuto0.vbs



' ecrire un message dans une fenetre popup
' Tout ce qui se trouve entre "..." est du texte (string)
msgbox "Coucou c'est moi",vbOKOnly,"Titre"

' ...
' le script se bloque tant que le message avant n'est pas fermé
' ...

WScript.StdOut.WriteLine ("*** ici on ecrit dans le terminal ***")

WScript.Echo ("on peut aussi utiliser echo")

WScript.StdOut.WriteLine "felicitation pour le premier vbscript :) "

