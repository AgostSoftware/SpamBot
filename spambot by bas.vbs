set shell = createobject ("wscript.shell") 

strtext = inputbox ("Wat wil je spammen?") 
strtimes = inputbox ("Hoevaak moet er gespamt worden?") 
strspeed = inputbox ("Hoe snel wil je spammen? (1000 = één per seconde, 100 = 10 per seconden etc.)") 
strtimeneed = inputbox ("Hoe veel SECONDEN heb je nodig om naar je spam scherm te komen?") 

If not isnumeric (strtimes & strspeed & strtimeneed) then 
msgbox "Je hebt iets anders ingetypt dan een nummer, tijden, snelheid en/of tijd nodig. Dit sluit nu af." 
wscript.quit 
End If 
strtimeneed2 = strtimeneed * 1000 
do 
msgbox "Je hebt" & strtimeneed & "seconden om naar je spam veld te gaan voor de spam begint!." 
wscript.sleep strtimeneed2 
shell.sendkeys ("spambot by bas" & "{enter}") 
for i=0 to strtimes 
shell.sendkeys (strtext & "{enter}") 
wscript.sleep strspeed 
Next 
shell.sendkeys ("Spambot by bas." & "{enter}") 
wscript.sleep strspeed * strtimes / 10 
returnvalue=MsgBox ("Wil je opnieuw spammen met de zelfde informatie?",36) 
If returnvalue=6 Then 
Msgbox "Oké, de spambot is gedeactiveerd!." 
End If 
If returnvalue=7 Then 
msgbox "Spambot is nu uit!" 
wscript.quit 
End IF 
loop 
