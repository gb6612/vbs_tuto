' Lire et Ecrire dans un fichier texte 

Sub log(str)
   WScript.StdOut.WriteLine str
End Sub

log("read from file")
Const ForReading = 1, ForWriting = 2, ForAppending = 8


' Ceci peut etre plus rapide, pour des fichiers petits
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("list.txt",ForReading)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("list_out.html",ForWriting,true)


log("Lire un fichier entier d'un coup")
strFileText = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

objFileToWrite.WriteLine("<html>")
objFileToWrite.WriteLine("<head><style>body{font-family: sans-serif;margin: 15px;}a{line-height: 25px;}</style></head>")
objFileToWrite.WriteLine("<body>")
objFileToWrite.WriteLine("<h2>Listing all Pdfs from ssv211.</h2><h3>Please click on link to view pdf.</h3>")
objFileToWrite.WriteLine("<ul>")


ft = Split(strFileText, vbCrLf)  ' separe chaque ligne
i=0
do while i <= UBound(ft)
     ' 
     if (Len(ft(i))>0) Then
       tmpStr = "<li><a href = " & ft(i) & ">"
       
       arrNames = Split(ft(i), "\")
       intIndex = Ubound(arrNames)
       tmpFilename = arrNames(intIndex)
       tmpFilename = Replace( tmpFilename, """", "")
       tmpStr = tmpStr & tmpFilename
       
       tmpStr = tmpStr & "</a></li>"
       WScript.StdOut.WriteLine(tmpStr)
       objFileToWrite.WriteLine(tmpStr)
     End If
     i=i+1
loop



objFileToWrite.WriteLine("</ul>")
objFileToWrite.WriteLine("</body>")
objFileToWrite.WriteLine("</html>")

objFileToWrite.Close
Set objFileToWrite = Nothing



