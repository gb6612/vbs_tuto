' Dates et Heures

' La date et heure du systeme
str0 = Now
WScript.StdOut.WriteLine("Nous sommes le " & str0) 
WScript.StdOut.WriteLine
WScript.StdOut.WriteLine("Date: " & DateValue(str0)) ' ici on affiche seulement la date
WScript.StdOut.WriteLine("Heure: " & TimeValue(str0)) ' ici on affiche seulement l'heure
WScript.StdOut.WriteLine
WScript.StdOut.WriteLine("Annee: " & Year(str0)) ' annee
WScript.StdOut.WriteLine("Mois: " & Month(str0)) 
WScript.StdOut.WriteLine("Jour: " & Day(str0)) 
WScript.StdOut.WriteLine("Nom du Mois: " & MonthName(Month(str0))) ' MonthName(month, [ abbreviate ])
WScript.StdOut.WriteLine("Jour de la semaine: " & Weekday(DateValue(str0)))  ' Weekday(date, [ firstdayofweek ])
WScript.StdOut.WriteLine("Jour de la semaine: " & WeekdayName(Weekday(DateValue(str0)))) ' WeekdayName(weekday, abbreviate, firstdayofweek)
WScript.StdOut.WriteLine
WScript.StdOut.WriteLine("heure: " & Hour(str0)) 
WScript.StdOut.WriteLine("minutes: " & Minute(str0)) 
WScript.StdOut.WriteLine("secondes: " & Second(str0)) 
WScript.StdOut.WriteLine

' Pour mettre une heure dans une variable

var0 = TimeSerial(13, 45, 00) 'TimeSerial(hour, minute, second)
var1 = DateSerial(2022, 07, 04) ' DateSerial(year, month, day)
WScript.StdOut.WriteLine("Date specifiee: " & var1+var0) 


' Calculer la difference entre deux dates
' DateDiff(interval, date1, date2, [ firstdayofweek, [ firstweekofyear ]] )
' interval:
'    yyyy	Year
'    q	Quarter
'    m	Month
'    y	Day of year
'    d	Day
'    w	Weekday
'    ww	Week
'    h	Hour
'    n	Minute
'    s	Second
WScript.StdOut.WriteLine("Difference entre deux dates en jours: " & DateDiff("d", Now, var1)) 


' Pour ajouter à une date 
'DateAdd(interval, number, date)
WScript.StdOut.WriteLine("Ajouter 35 jours a partir d'aujourd'hui: " & DateAdd("d", 35, Now)) 


