' 1) IsNumeric permet de tester si une valeur est un chiffre valable
WScript.StdOut.Write("Donne moi un chiffre : ")
num1 = WScript.StdIn.ReadLine()

If not(IsNumeric(num1)) Then  ' Si la valeur n'est PAS un chiffre valable
      WScript.StdOut.WriteLine(num1 & " n'est pas un chiffre")
ElseIf (CStr(CLng(num1))=num1) Then
      WScript.StdOut.WriteLine(num1 & " est un nombre entier")
Else
      WScript.StdOut.WriteLine(num1 & " est un nombre avec virgule")
End If


'2) SLEEP
' met en pause le script pour x ms 
for i=1 to 10
      WScript.Sleep 500  ' stop 500ms
      wscript.StdOut.Write "-"
Next
WScript.StdOut.WriteLine


' 3) random - retourne un numero aleatoire entre 0 et 1
Randomize  ' on appelle 1 seule fois
for i=1 to 10
   wscript.StdOut.WriteLine (Rnd)
Next

'
' Pour donner un numero aleatoire entier entre min et max:
'
Dim max,min
max=100
min=1
Randomize
for i=1 to 10
   wscript.StdOut.WriteLine (Int((max-min+1)*Rnd+min))
Next

'
' Pour donner une suite de characteres aleatoire
' 
Const LETTRES = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
Randomize
for i=1 to 10
   wscript.StdOut.Write (Mid(LETTRES, Int(Len(LETTRES)*Rnd+1), 1))
   ' explication:
   ' Int(Len(LETTRES)*Rnd+1) => retourne un numero aleatoire entre 1 et le nombre de characteres dans 'LETTRE'
   ' Mid(LETTRES, ..., 1) => prend 1 charactere dans 'LETTRE', à la position donnée par le numero precedant 
Next
wscript.StdOut.WriteLine
