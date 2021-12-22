' 1) Lire une valeur donnee par l'utilisateur
' 2) Si... Alors...


' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/inputbox-function
' InputBox(prompt, [ title ], [ default ], [ xpos ], [ ypos ], [ helpfile, context ])
'
' Ici on demande qqch a l'utilisateur via une fenetre de popup
dim varStr
varStr = InputBox("Quel est ton nom ?", "Nom", "Patapouf")

If (varStr = "") Then
   ' dans ce cas l'utilisateur a cliqué sur Cancel
   msgbox "Vous voulez annuler"
Else
   msgbox "Bonjour " & varStr   ' retourne un message dans une popup
End If

' Et ici on demande a l'utilisateur directement via le terminal
WScript.StdOut.Write("Et ton animal prefere? ")
'WScript.StdIn.Read(0)
varStr = WScript.StdIn.ReadLine()
WScript.StdOut.Write(varStr & " est aussi mon animal prefere!")

' remarque : StdOut.Write ne va pas a la ligne!

' Astuce pour ecrire 2 lignes vides
WScript.StdOut.Write vbCrLf & vbCrLf   
' ou
WScript.StdOut.WriteLine
WScript.StdOut.WriteLine



' 2) Si... Alors...
WScript.StdOut.Write("Donne moi un chiffre inferieur a 100 : ")
varStr = WScript.StdIn.ReadLine()

If (varStr > 100) Then
   WScript.StdOut.WriteLine(varStr & " est plus grand que 100...bref")
elseif (varStr = 100) Then
   WScript.StdOut.WriteLine(varStr & " est egal a 100...")
else
   WScript.StdOut.WriteLine(varStr & " est mon numero prefere")
End If

