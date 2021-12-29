Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")


dim Input

Sapi.speak "hi."
Sapi.speak "I am you Personal assistant"
Sapi.speak "Please type what you want to open"
Input=inputbox ("Please type what you want to open")


if Input = "hi" OR Input = "Hi"then
Sapi.speak "hi"
Sapi.speak "how are you?"
Input=inputbox ("How are you?")
if Input = "good"then
Sapi.speak "nice!"
else
if Input = "bad"then
Sapi.speak "sorry"
end if
end if
else
if Input = "youtube" OR Input = "Youtube"then
Sapi.speak "Opening youtube"
wshshel.run "www.youtube.com"

else
if Input = "instructables" OR Input = "Instructables" then
Sapi.speak "Opening instructables"
wshshell.run "www.instructables.com"

else
if Input = "google" OR Input = "Google" then
Sapi.speak "Opening google"
wshshell.run "www.google.com"

else
if Input = "command prompt" OR Input = "Command prompt" then
Sapi.speak "Opening command prompt"
wshshell.run "cmd"

else
if Input = "calculator" OR Input = "Calculator" then
Sapi.speak "Opening calculator"
wshshell.run "calc"

else
if Input = "notepad" OR Input = "Notepad" then
Sapi.speak "Opening notepad"
wshshell.run "notepad"
else
if Input = "hangouts" OR Input = "Hangouts" OR Input = "google hangouts" OR Input = "Google hangouts" OR Input = "Google Hangouts" then
Sapi.speak "Opening google hangouts"
wshshell.run "https://hangouts.google.com/"
else
if Input = "hangouts" OR Input = "Python" then
Sapi.speak "Opening python"
wshshell.run "idle"
else
if Input = "scratch" OR Input = "Scratch" then
Sapi.speak "Opening scratch"
wshshell.run "https://scratch.mit.edu/"
else
if Input = "calender" OR Input = "Calender" OR Input = "Google calender" OR Input = "google calender" OR Input = "Google Calender" then
Sapi.speak "Opening google calender"
wshshell.run "https://calendar.google.com/calendar/u/0/r"
else
if Input = "clock" OR Input = "Clock" then
Sapi.speak "Opening clock"
wshshell.run "https://www.clocktab.com/"
else
Sapi.speak "I don't recognize your input, Please re open the assistant app and try again"
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if