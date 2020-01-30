Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
dim str

if hour(time) < 12 then
	Sapi.speak "Good Morning"
else
	if hour(time) > 12 then
		if hour(time) > 16 then
			Sapi.speak "Good evening"
		else
			Sapi.speak "Good afternoon"
		end if
	end if
end if

Sapi.speak "The current day is"
Sapi.speak date
	
Sapi.speak "The current time is"
	
if hour(time) > 12 then
	Sapi.speak hour(time)-12
else
	if hour(time) = 0 then
		Sapi.speak "12"
	else
		Sapi.speak hour(time)
	end if
end if

if minute(time) < 10 then
	Sapi.speak "o"

	if minute(time) < 1 then
		Sapi.speak "clock"
	else
		Sapi.speak minute(time)
	end if
else
	Sapi.speak minute(time)
end if
if hour(time) > 12 then
	Sapi.speak "P.M."
	else
		if hour(time) = 0 then
			if minute(time) = 0 then
				Sapi.speak "Midnight"
			else
				Sapi.speak "A.M."
			end if
		else
			if hour(time) = 12 then
				if minute(time) = 0 then
					Sapi.speak "Noon"
			else
				Sapi.speak "P.M."
			end if
		else
			Sapi.speak "A.M."
		end if
	end if
 end if
 
Sapi.speak "Checking for new messages. Please standby."
WScript.Sleep 2000

Set otl = createobject("outlook.application")
Set session = otl.getnamespace("mapi")

session.logon
	Set inbox = session.getdefaultfolder(6)
	c = 0

	For Each m In inbox.items
		If m.unread Then c = c + 1
	Next
session.logoff

s = "s"
If c = 1 Then s = ""

Sapi.speak "You have" & c & "unread message" & s
