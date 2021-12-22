' Lire et Ecrire dans un fichier texte 

Sub log(str)
   WScript.StdOut.WriteLine str
End Sub

log("read from file")
Const ForReading = 1, ForWriting = 2, ForAppending = 8

log(" Lire un fichier ligne apres ligne")
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("7_rdfile.txt",ForReading)
Dim strLine
do while not objFileToRead.AtEndOfStream
     strLine = objFileToRead.ReadLine()
     'Do something with the line
     WScript.StdOut.WriteLine(strLine)
loop
objFileToRead.Close
Set objFileToRead = Nothing


log("Lire un fichier entier d'un coup")
' Ceci peut etre plus rapide, pour des fichiers petits
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("7_rdfile.txt",ForReading)
strFileText = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing
ft = Split(strFileText, vbCrLf)  ' separe chaque ligne
i=0
do while i <= UBound(ft)
     'Ajoute un numero de ligne...
     WScript.StdOut.WriteLine(i & " : " & ft(i))
     i=i+1
loop


log("Ecrire dans un fichier")
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("7_wrfile.txt",ForWriting,true)
objFileToWrite.WriteLine(strFileText)
objFileToWrite.Close
Set objFileToWrite = Nothing


log(" Creer un directory et copier un fichier")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Const Overwrite = True
strFile = "\\server\folder\file.ext"
strFolder = "MyFolder"
If Not oFSO.FolderExists(strFolder) Then
  oFSO.CreateFolder strFolder
  oFSO.CopyFile strFile, strFolder, Overwrite
End If

