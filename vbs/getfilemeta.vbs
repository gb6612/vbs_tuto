' Run using
' > cscript <filename.vbs> > out.log

Const ForReading = 1, ForWriting = 2, ForAppending = 8

myFile = "C:\Users\gabri\Documents\lettre.docx"
'myFile = "C:\mydata\vbs\list_out.html"
tmpStr = GetDetails(myFile)

mylog(tmpStr)



Sub mylog(str)   
   WScript.StdOut.WriteLine str
End Sub

Function GetDetails(fileName)
    Dim oShell 
    Dim objFolder 
    Dim szItem 
    Dim objFolderItem

    Set fso = CreateObject("Scripting.FileSystemObject") 

    Set objFile = fso.GetFile(fileName)
    tmpPathName = objFile.ParentFolder
    tmpFileName = objFile.Name
    'WScript.StdOut.WriteLine tmpPathName
    'WScript.StdOut.WriteLine tmpFileName

    Set oShell = CreateObject("Shell.Application")
    Set objFolder = oShell.NameSpace(tmpPathName)
    
    tmpStr = ""
    
    If (Not objFolder Is Nothing) Then
        Set objFolderItem = objFolder.ParseName(tmpFileName)
   
        If (Not objFolderItem Is Nothing) Then
            'szItem = objFolder.GetDetailsOf(objFolderItem, 2)
            For i = 0 To 40
                tmpStr = tmpStr & i & " : " & objFolder.GetDetailsOf(objFolderItem, i) & " : " & objFolder.GetDetailsOf(sFile, i)  & vbCrLf                
            Next
        End If
        
        Set objFolderItem = Nothing
    End If
    
    GetDetails = tmpStr 
    Set objFolder = Nothing
    Set oShell = Nothing
End Function
