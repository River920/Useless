dim q1, pswd, pswdinpt


q1 = msgbox("This Is The Useless Launcher Verifier, If you didn't intend to open this click 'No'", vbYesNo + vbQuestion ,"ULC")
    If q1 = vbNo Then
        WScript.Quit
    End If

pswd = 150314


Function PswdReq()
msgbox "The Password Is: " & pswd, vbInformation , "Bot Protection"
pswdinpt = InputBox("Enter The Password Shown In The Previous Prompt: ", vbQuestion , "")
    If pswdinpt = 150314 Then
        msgbox "Verified, You're Not A Bot!", vbOk + vbInformation , "Bot Protection"
    Else msgbox "Try Again.", vbOk + vbCritical
    End If
End Function

PswdReq()