set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next 

Dim title,risposta 

title = "Hey!"

risposta=MsgBox("This will delete ALL cookies",vbYesNo+vbExclamation,title) 

if risposta=vbYes then 

risposta1=MsgBox("Are you sure?",vbYesNo+vbExclamation,title)

if risposta1=vbYes then fso.DeleteFile "%userprofile%\AppData\Local\Google\Chrome\User Data\Default\Cookies"
end if
