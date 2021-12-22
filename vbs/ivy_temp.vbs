Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
'Set olApp=CreateObject("Outlook.Application")
	  
dim str, str1 ' used for string
dim i,j  ' General Purpose

' Finish each array cell with a space
CWeekdayArray_EN=Array("", "Sunday ", "Monday ", "Tueday ", "Wednesday ", "Thursday ", "Friday ", "Saturday ")
CWeekdayArray_FR=Array("", "Dimanche ", "Lundi ", "Mardi ", "Mercredi ", "Jeudi ", "Vendredi ", "Samedi ")

CMonthArray_EN=Array("", "January ", "February ", "March ", "April ", "May ", "June ", "July ", "August ", "September ", "October ", "November ", "December ")
CMonthArray_FR=Array("", "Janvier ", "Février ", "Mars ", "Avril ", "Mai ", "Juin ", "Juillet ", "Août ", "Septembre ", "Octobre ", "Novembre ", "Décembre ")

CGreetings_EN=Array("Good Morning ", "Good Afternoon ", "Good Evening ")
CGreetings_FR=Array("Bonjour ", "Hello ", "Bonsoir ")



Sub IvySays(str)
    Sapi.speak str 
	'msgbox str,vbOKOnly,"Ivy Says"
End Sub


Function myWeekdayName(weekday, lang)
'   if (lang="en") then
'      myWeekdayName = CWeekdayArray_EN(weekday)
'   elseif (lang="fr") then
      myWeekdayName = CWeekdayArray_FR(weekday)
'   end if
End Function

Function myMonthName(month, lang)
'   if (lang="en") then
'      myMonthName = CMonthArray_EN(month)
'   elseif (lang="fr") then
      myMonthName = CMonthArray_FR(month)
'   end if
End Function

Function RSSTopHeadline(feed, feednb)
  Set req = CreateObject("MSXML2.XMLHTTP.3.0")
  req.Open "GET", feed, False
  req.Send

  Set xml = CreateObject("Msxml2.DOMDocument")
  xml.loadXml(req.responseText)
  RSSTopHeadline = xml.getElementsByTagName("channel/item/title")(feednb).Text
End Function



Function myFormatDateTime(myDate)
   Dim str
   
   ' format to
   ' 2017-01-04 00:00
    str = Year(myDate) & "-"
		 		 
	If(Len(Month(myDate))=1) Then
        str=str & "0"
    End If
    str = str & Month(myDate) & "-"
   
	If(Len(Day(myDate))=1) Then
        str=str & "0"
    End If
    str = str & Day(myDate) & " 00:00"
      
	myFormatDateTime = str
End Function


Function myCalendarFormatDateTime(myDate)
   Dim str
   
   ' format to
   ' 2017-01-04 00:00
    str = Year(myDate) & "-"
		 		 
	If(Len(Month(myDate))=1) Then
        str=str & "0"
    End If
    str = str & Month(myDate) & "-"
   
	If(Len(Day(myDate))=1) Then
        str=str & "0"
    End If
    str = str & Day(myDate)
      
	myCalendarFormatDateTime = str
End Function


Function WeAreOnline
   dim tmp
   set oShell = WScript.CreateObject("WScript.Shell")
   ReturnCode = oShell.Run("ping -n 3 -w 1000 www.w3.org",0,true)
   
   tmp = 0
   if ReturnCode=0 then
      'msgbox  " PC is connected to internet"
	  tmp = 1
   else
       ' If unsuccessfull let's check somewhere else, in case the website was down for maintenance
       ReturnCode = oShell.Run("ping -n 3 -w 300 www.admin.ch",0,true)
       if ReturnCode=0 then
         'ok now
	     tmp = 1
	   end if
   
   end if
   WeAreOnline = tmp
End Function


' ********************************************************************
' MAIN
' ********************************************************************

IvySays("Bonjour mes heros")
IvySays("Felicitation pour avoir trouvé le code.")
IvySays("Le système est en phase d'initialisation.")
IvySays("Attendez s'il vous plait.")


Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

' Play audio
oPlayer.URL = "C:\mylib\audio\Lesser Vibes - High Tech Interface Sounds.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

IvySays("Voici votre nouveau défi.")
IvySays("Il faut décripter le message en chiffres en bas")
IvySays("a partir de l'indice écrit")
IvySays("Bonne chance")

oPlayer.URL = "C:\mylib\audio\Lesser Vibes - High Tech Interface Sounds.mp3"
'oPlayer.settings.setMode "loop", True

For i = 0 To 3
   oPlayer.controls.play 
   While oPlayer.playState <> 1 ' 1 = Stopped
     WScript.Sleep 100
   Wend
Next

' Release the audio file
oPlayer.close


