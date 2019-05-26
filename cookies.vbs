set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next 

Dim title,risposta 

title = "Hey!" 'Titolo dell'applicazione 

risposta=MsgBox("This will delete ALL cookies",vbYesNo+vbExclamation,title) 
'Visualizza una MsgBox 

if risposta=vbYes then 
'Se la risposta è si 

risposta1=MsgBox("Are you sure?",vbYesNo+vbExclamation,title)
'Visualizza esatto 

if risposta1=vbYes then fso.DeleteFile "%userprofile%\AppData\Local\Google\Chrome\User Data\Default\Cookies"
'Se la risposta è si 
end if
