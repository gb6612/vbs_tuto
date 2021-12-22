myFile = "C:\mydata\vbs\list_out.html"
'tmpStr = GetDetails(myFile)

Set fso = CreateObject("Scripting.FileSystemObject") 
mylog ("GetFileName : " & fso.GetFileName(myFile))
mylog ("GetBaseName : " & fso.GetBaseName(myFile))
mylog ("GetAbsolutePathName : " & fso.GetAbsolutePathName(myFile))
mylog ("GetExtensionName : " & fso.GetExtensionName(myFile))
mylog ("GetParentFolderName : " & fso.GetParentFolderName(myFile))

mylog("")
'On Error Resume Next

Set objFile = fso.GetFile(myFile)
Set objFolder = objFile.ParentFolder

mylog ("Name: " & objFile.Name)
mylog ("ParentFolder: " & objFile.ParentFolder)
mylog ("Path: " & objFile.Path)
mylog ("Type: " & objFile.Type)
mylog ("Parent Folder: " & objFolder.Name)
mylog ("Drive: " & objFolder.Drive)
mylog ("Parent Folder Path: " & objFolder.Path)

Sub mylog(str)
   WScript.StdOut.WriteLine str
End Sub
