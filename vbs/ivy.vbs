Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
'Set olApp=CreateObject("Outlook.Application")
	  
dim str, str1 ' used for string
dim i,j  ' General Purpose
Dim News(4,5)  ' Max 4 sources with 5 headlines each
Dim myMeteo(8)

Const SLEEPTIMEOUT = 1800  ' Update every 30min

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

Function UpdateMeteoNews
  Set ie = CreateObject("InternetExplorer.Application")  
  ie.Navigate "http://www.meteosuisse.admin.ch/home.html?tab=report"
  While ie.Busy : WScript.Sleep 100 : Wend

  ' P(0)  : actualise le ...
  ' H3(0) : situation generale
  ' P(1)  :    ...
  ' H3(1) : today title
  ' P(2)  :    ...
  ' H3(2) : tomorrow title
  ' P(3)  :    ...
  Set collH3Tags = ie.Document.getElementsByClassName("textFCK").item(0).getElementsByTagName("h3")  
  Set collPTags  = ie.Document.getElementsByClassName("textFCK").item(0).getElementsByTagName("p")  
  
  myMeteo(0) = ie.Document.getElementsByClassName("textFCK").item(0).getElementsByTagName("h2").item(0).innertext    
  myMeteo(1) = collPTags.item(0).innertext    'actualise le ...
  myMeteo(2) = collH3Tags.item(0).innertext   'situation generale
  myMeteo(3) = collPTags.item(1).innertext    '   ...
  myMeteo(4) = "aujourd'hui " & collH3Tags.item(1).innertext   'today title
  myMeteo(5) = collPTags.item(2).innertext    '   ...
  myMeteo(6) = "demain " & collH3Tags.item(2).innertext   'tomorrow title
  myMeteo(7) = collPTags.item(3).innertext    '   ...
  'msgbox  collH3Tags.length
  'msgbox  collH3Tags.item(0).innertext
  
  ie.quit
  
End Function


Function UpdateNews
       dim i
	   dim str
	   
	   ' jeuxvideos.com
       For i = 0 To 4
         News(0,i) = RSSTopHeadline("http://www.jeuxvideo.com/rss/rss.xml", i)
       Next

	   ' 20min actualité
       For i = 0 To 4
         News(1,i) = RSSTopHeadline("http://www.20min.ch/rss/rss.tmpl?type=channel&get=17&lang=ro", i)
       Next

	   ' LeTemps economie
       For i = 0 To 4
         News(2,i) = RSSTopHeadline("https://www.letemps.ch/taxonomy/term/7/feed", i)
       Next
       '
	   ' 20min sport
       For i = 0 To 4
         News(3,i) = RSSTopHeadline("http://www.20min.ch/rss/rss.tmpl?type=channel&get=23&lang=ro", i)
       Next

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


Function IsUserAlive
   Dim tmp
   Set objShell = WScript.CreateObject("WScript.Shell")
   
   iRetVal = objShell.Popup("Il y a quelqu'un?" , SLEEPTIMEOUT, "Knock Knock", vbOKOnly + vbInformation)
   'objShell.SendKeys "% n"   ' Minimize window
   
   Select Case iRetVal
      Case vbOK
         'msgbox "You clicked OK"
		 tmp = 1
      Case -1 
         'msgbox "Popup timed out"
		 tmp = 0
   End Select
   
   'Wscript.Quit
   IsUserAlive = tmp
End Function

' ********************************************************************
' MAIN
' ********************************************************************

Dim objShell
Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"


'Get User Name
str1 = CreateObject("WScript.Network").UserName

if hour(time)<12 then
   str = CGreetings_FR(0)
elseif hour(time)<18 then
   str = CGreetings_FR(1)
else
   str = CGreetings_FR(2)
end if
str = str & str1
IvySays(str)
IvySays("Je suis a votre disposition")


do while (1)
if IsUserAlive Then

   str1=Lcase(Trim(InputBox("Que desirez vous ?","Ready to Listen")))
   
   if IsEmpty(str1) Then
       IvySays("bien")       
   
   elseif (str1="heure") then
       ' TIME
       IvySays("Il est " & hour(time) & " heures " & minute(time))

   elseif (str1="date") then
       ' DATE
       IvySays("Nous sommes " & myWeekdayName(weekday(date), "fr") & " , " & " le " & day(date) & " " & myMonthName(month(date), "fr") & " " & year(date) ) ' FR
   
   elseif (str1="nouvelles") then
       ' RSS headlines
       if WeAreOnline then
          IvySays("Je vais chercher les titres principaux.")
          IvySays("Je reviens.")
	      UpdateNews
	   else
	      IvySays("Malheureusement je n'ai pas de connexion a internet")
	      IvySays("Je vais vous donner les dernières nouvelles a ma disposition")
	   end if
	   
       IvySays("Quelques nouvelles du 20 minutes")
       For i = 0 To 4
          IvySays(News(1,i))
       Next
	   
       IvySays("Rubrique economie du temps . c h")
       For i = 0 To 2
          IvySays(News(2,i))
       Next
	   
       IvySays("Rubrique Sport ")
       For i = 0 To 2
          IvySays(News(3,i))
       Next
	   
       IvySays("Le site jeu video nous donne ceci")
       For i = 0 To 2
          IvySays(News(0,i))
       Next
       
   elseif (str1="meteo") then
       ' METEO
       if WeAreOnline then
          IvySays("Je cherche la meteo")
	      UpdateMeteoNews
	   else
	      IvySays("Malheureusement je n'ai pas de connexion a internet")
	      IvySays("Je vais vous donner les dernières nouvelles a ma disposition")
	   end if
       For Each i In myMeteo
         str = i
         IvySays(str)
       Next
       IvySays("Voila pour la meteo")
	   
   elseif (str1="messages") then
      ' E-MAILS
      IvySays("Je controle vos messages")
      Set olApp=CreateObject("Outlook.Application")
      Set olns=olApp.GetNameSpace("MAPI")
	  'olns.SendAndReceive(False)
      Set objFolder = olns.Folders("Betech").Folders("Inbox")
	  
	  ' Loop through e-mails
      Set oItems = objFolder.Items
      'Set colFilteredItems = oItems.Restrict("[Unread]=true")
      oItems.Sort "[ReceivedTime]", False	  

	  Set oItemsInDateRange = oItems.Restrict("[ReceivedTime] >= '" & myFormatDateTime(DateAdd("d", -2, Date)) & "'") 'Restrict the Items collection for the xx-day date range
      oItemsInDateRange.Sort "[ReceivedTime]"
	  
	  if oItemsInDateRange.Count=0 then
         IvySays("Il n'y a pas de messages")
	     
	  else
         For i = 1 to oItemsInDateRange.Count
            Set objMessage  = oItemsInDateRange.Item(i)
			IvySays("De " )
			IvySays(objMessage.SenderName )
			IvySays("resu le " & Day(objMessage.ReceivedTime) & " " & myMonthName(month(objMessage.ReceivedTime), "fr") & " a " & Hour(objMessage.ReceivedTime) & " heures " & minute(objMessage.ReceivedTime) )
			WScript.Sleep 500
			IvySays("Au sujet de " & objMessage.subject)
			WScript.Sleep 1000
		 Next
'            'msgbox "Subject: " & objMessage.subject
'            'msgbox "To: " & objMessage.to
'            'msgbox "SenderEmailAddress: " & objMessage.SenderEmailAddress
'            'msgbox "Sender: " & objMessage.Sender
'            'msgbox "SenderName: " & objMessage.SenderName
'            'msgbox "ReceivedTime: " & objMessage.ReceivedTime
'            'msgbox "Body: " & objMessage.body
  	     IvySays("Voila pour les messages")
		 
	   End If
       
	   
   elseif (str1="meeting") then
      ' Calendar
      IvySays("Je controle le calendrier")
      Set olApp=CreateObject("Outlook.Application")
      Set olns=olApp.GetNameSpace("MAPI")
	  'Set oCalendar = olns.GetDefaultFolder(olFolderCalendar)
	  Set oCalendar = olns.GetDefaultFolder(9)
	  	  
	  Set oItems = oCalendar.Items
      oItems.Sort "[Start]"
	  
	  if oItems.Count>0 Then
	     j=0
         For Each i In oItems
		    'msgbox i.Subject & " " & DateDiff("d", Date, i.Start) & " " & DateDiff("d", i.End, Date)
		    if (DateDiff("d", Date, i.Start)=0) and (DateDiff("d", i.End, Date)=0) Then
			   ' Only Today
                  IvySays("Aujourd'hui a " &  hour(i.Start) & " heures " & minute(i.Start) & " il y a " & i.Subject)
	              j=1
		    elseif (DateDiff("d", Date, i.Start)=0) and (DateDiff("d", i.End, Date)<0) Then
			   ' From/Start Today
                  IvySays("Aujourd'hui commence " & i.Subject)
	              j=1
		    elseif (DateDiff("d", Date, i.Start)=1) and (DateDiff("d", i.End, Date)=-1) Then
			   ' Only Tomorrow
                  IvySays("Demain a " &  hour(i.Start) & " heures " & minute(i.Start) & " il y a " & i.Subject)
	              j=1
		    elseif (DateDiff("d", Date, i.Start)=1) and (DateDiff("d", i.End, Date)<-1) Then
			   ' From/Start Tomorrow
                  IvySays("Demain commencera " & i.Subject)
	              j=1
			End If
   	     Next
		 
		 If (j=0) Then
   		    IvySays("Je ne vois aucun rendez vous pour aujourd'hui ni demain")
		 End if
		 
	  Else
   		    IvySays("Je ne vois aucun rendez vous pour aujourd'hui ni demain")
	  
	  End If
	  

	  
   elseif (str1="exit") then
       str = "Si vous n'avez plus besoin de moi, je vais dormir. A bien tot."  ' FR
       IvySays(str)
       exit do
	   
   elseif (str1="aide") or (str1="?") then
       IvySays("Voici ce que vous pouvez me demander")
       str = "heure"     & vbCrLf & _
	         "date"      & vbCrLf & _
	         "nouvelles" & vbCrLf & _
	         "meteo"     & vbCrLf & _
	         "messages"  & vbCrLf & _
	         "meeting"   & vbCrLf & _
	         "exit" 
	   msgbox str

   'elseif (str1="run ff") then
   '    Set objShell = WScript.CreateObject( "WScript.Shell" )
   '    objShell.Exec("""c:\Program Files\Mozilla Firefox\firefox.exe""")
   '    Set objShell = Nothing
	   
   else
       ' Command not valid
       IvySays("Je ne comprends pas votre demande")
	   
   end if
   
   IvySays("Autre choses ?")
   
Else
   ' User is not present
   ' Let's update the news on our own (but only if online)
   if WeAreOnline then
        UpdateNews
	    UpdateMeteoNews
   end if
   
End If   
loop
