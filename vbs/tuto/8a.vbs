' Strings SPLIT
'https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/string-functions

str0 = "Il etait une.fois.une BELLE journee"
WScript.StdOut.WriteLine(str0)

WScript.StdOut.WriteLine("separer avec les espaces")
ft = Split(str0) 
i=0
do while i <= UBound(ft)
     'Ajoute un numero de ligne...
     WScript.StdOut.WriteLine(i & " : " & ft(i))
     i=i+1
loop

WScript.StdOut.WriteLine
WScript.StdOut.WriteLine("separer avec les points")
ft = Split(str0,".") 
i=0
do while i <= UBound(ft)
     'Ajoute un numero de ligne...
     WScript.StdOut.WriteLine(i & " : " & ft(i))
     i=i+1
loop
