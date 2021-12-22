' msgbox avancé
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function

' MsgBox (prompt, [ buttons, ] [ title, ] [ helpfile, context ])
' 
' prompt : texte affiché dans la fenetre
' 
' buttons : definit les boutons et l'icone à afficher. 
'        par exemple pour avoir deux boutons OUI/NON et une icone d'erreur on ecrira :
'        vbYesNo+vbCritical
' Constant	Value	Description
' vbOKOnly	0	Display OK button only.
' vbOKCancel	1	Display OK and Cancel buttons.
' vbAbortRetryIgnore	2	Display Abort, Retry, and Ignore buttons.
' vbYesNoCancel	3	Display Yes, No, and Cancel buttons.
' vbYesNo	4	Display Yes and No buttons.
' vbRetryCancel	5	Display Retry and Cancel buttons.
' vbCritical	16	Display Critical Message icon.
' vbQuestion	32	Display Warning Query icon.
' vbExclamation	48	Display Warning Message icon.
' vbInformation	64	Display Information Message icon.
' 
dim reponse
reponse = msgbox ("voulez vous continuer ?", _
        vbYesNo+vbCritical, _
        "titre")   ' retourne un message dans une popup

' ensuite on peut tester la reponse choisie par l'utilisateur.
' Attention à tester seulement les boutons qu'on a affiché plus haut,
' par exemple si on a affiché OUI/NON, il ne faudra pas tester un OK...
' Les reponses possibles sont:
' vbOK	1	OK
' vbCancel	2	Cancel
' vbAbort	3	Abort
' vbRetry	4	Retry
' vbIgnore	5	Ignore
' vbYes	6	Yes
' vbNo	7	No


If (reponse = vbYes) Then
   msgbox "On continue", vbExclamation, "click yes"
elseif (reponse = vbNo) Then
    msgbox "On arrete ici", vbInformation, "click no"
'else
'   WScript.StdOut.WriteLine("")
End If

