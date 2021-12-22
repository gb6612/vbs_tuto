' Lancer le script avec des parametres/arguments
' nom_du_script.vbs  argument1 argument2 argument3 ...
' dans ce cas, par exemple, executer le script comme suit:
' cscript 5.vbs hello world

' ici on compte le nombre d'arguments passés 
intCount = WScript.Arguments.Count
wscript.echo "il y a " & intCount & " parametres"

' s'il y a des arguments...
if (intCount>0) then
   for i=0 to intCount-1
      arg = WScript.Arguments.Item(i)
      wscript.echo "Le parametre " & i & " est " & arg
   Next
end if

' arg0 = WScript.Arguments.Item(0)  ' le 1er argument
' arg1 = WScript.Arguments.Item(1)  ' le 2eme argument 
' arg2 = WScript.Arguments.Item(2)  ' le 3eme argument , etc...




