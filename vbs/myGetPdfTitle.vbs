' Run using
' > cscript <filename.vbs> > out.log

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = ".\doc"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

Dim objFile
Dim f
Dim tmpStr

Dim i
i = 0
For Each objFile in colFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "pdf" Then
        Wscript.StdOut.WriteLine i & " " & objfile.Path

        ' Open file for reading
        Set f = objFSO.OpenTextFile(objfile.Path, ForReading, false)
        
        Do Until f.AtEndOfStream
           tmpStr = f.ReadLine
'
           If foundStrMatch(tmpStr)=true Then
              tmpstr = LTrim(tmpstr)
              tmpStr = Replace(tmpStr, "<rdf:li xml:lang=""x-default"">", "") 
              tmpStr = Replace(tmpStr, "</rdf:li>", "") 
              WScript.StdOut.WriteLine tmpStr
            End If
        Loop
   
        f.Close

        i=i+1
    End If
Next

Wscript.StdOut.WriteLine i

Function foundStrMatch(tmpStr)
Dim substrToFind
substrToFind = "<rdf:li xml:lang=""x-default"">"
If InStr(tmpStr, substrToFind) > 0 Then
    foundStrMatch = true
Else
    foundStrMatch = false
End If
End Function
