'1) SLEEP
' met en pause le script pour x ms 
for i=1 to 10
    WScript.Sleep 500  ' stop 500ms
    wscript.StdOut.Write "-"
Next
WScript.StdOut.WriteLine

' Timer
' = secondes passees depuis minuit
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/timer-function
' DoEvents

Dim PauseTime, Start, Finish, TotalTime, aaa
If (MsgBox("clique Yes pour mettre en pause 5 secondes", vbYesNo)) = vbYes Then
    PauseTime = 5    ' temps de la pause
    Start = Timer    ' start time.
    Do While Timer < Start + PauseTime
        aaa = DoEvents    ' le script se met en pause, mais sans bloquer les autres processus de l'ordinateur!!!
    Loop
    Finish = Timer    ' end time.
    TotalTime = Finish - Start    ' Calcul total time.
    MsgBox "Mis en pause pour " & TotalTime & " secondes"
Else
    '
End If

' 2)
' nome della directory attuale
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory
WScript.StdOut.WriteLine strCurDir
