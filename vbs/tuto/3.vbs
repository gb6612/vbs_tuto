' 1) For... Next
' 2) While... 

' Compte jusqu'a 10...(10 inclus!)
For i = 1 To 10
   WScript.StdOut.Write(i & ",")
Next
WScript.StdOut.Write vbCrLf

' Compte jusqu'a 10...mais par pas de 2
For i = 1 To 10 Step 2
   WScript.StdOut.Write(i & ",")
Next
WScript.StdOut.Write vbCrLf

' DeCompte 
For i = 10 To 0 Step -1
   WScript.StdOut.Write(i & ",")
Next
WScript.StdOut.Write vbCrLf

WScript.StdOut.Write vbCrLf

' 2) While
' Les instructions do..While seront exécutées tant que la condition est Vraie. 
' (c'est-à-dire,) La boucle doit être répétée jusqu'à ce que la condition soit False.
' Ci dessous un exemple. La variable i est assignée (i=4). Puis la boucle commence et se repete tant que (i<10).
' Il faut changer i dans la boucle sinon elle va tourner en rond à l'infini !
WScript.StdOut.WriteLine("Do While...")
i = 4
do while (i < 10)
   WScript.StdOut.Write(i & ",")
   i = i + 1
loop
WScript.StdOut.Write vbCrLf


' Les instructions do..Until seront exécutées tant que la condition est Fausse 
WScript.StdOut.WriteLine("Do Until...")
i = 4
do until (i > 10)
   WScript.StdOut.Write(i & ",")
   i = i + 1
loop
WScript.StdOut.Write vbCrLf

