Attribute VB_Name = "Module1"
Public Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Global lngTimerID As Long
Global MsgStatus As String
Global strFilename As String
Global TotalMails As Long
Global m_State As POP3States
Global showForm As Boolean
Global Play As Boolean
Global WaitFlag As Boolean

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
        Call checknewmail
End Sub

Public Sub Main()
   Dim delay As String
   
   strFilename = App.Path & "\mail.wav"
   showForm = True
   Play = True
   WaitFlag = True
   
   MsgStatus = "Mail Checker Utility for " & Form2.txtHost.Text
   AddIcon Form2.picMail, MsgStatus
   showForm = True
   
   Pause 1
   delay = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "Delay", "")
   If delay = "" Then 'nothing is in the registry
        
        SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "Delay", "20"
        delay = 20
        Form2.Show 1
   End If
   checknewmail
   lngTimerID = SetTimer(0, 0, Round(delay * 60000, 0), AddressOf TimerProc)
End Sub
Public Sub checknewmail()
    Dim NoOfEntries, arr
    DoEvents
    If CBool(IsNetConnectOnline()) = True Then
        
        
        NoOfEntries = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "NoOfEntries", "")
        If NoOfEntries <> "" And NoOfEntries <> "0" Then
            arr = GetAllKeys(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker")
            For i = LBound(arr) To UBound(arr)
                DoEvents
                m_State = POP3_Connect
                
                Form2.txtHost.Text = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(i), "ServerName", "")
                Form2.TxtUserName.Text = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(i), "UserName", "")
                Form2.TxtPassword.Text = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(i), "Password", "")
                
                MsgStatus = "Checking New mails from " & Form2.txtHost.Text
                ChangeIcon Form2.picReadingMail(3), MsgStatus
                
                'Debug.Print Form2.txtHost.Text
                Form2.Winsock1.Close
                Form2.Winsock1.LocalPort = 0
                Form2.Winsock1.Connect Form2.txtHost.Text, 110
                WaitforNextMail
                Pause 15
                
            Next
        End If
    Else
        KillTimer 0, lngTimerID
        showForm = True
        Play = False
        ChangeIcon Form2.picPause(5), MsgStatus
    End If
    
End Sub
Public Sub WaitforNextMail()
    While 1
        DoEvents
        If WaitFlag = False Then Exit Sub
    Wend
End Sub

Public Sub CheckNewMailfromOneServer(mHost As String, mUser As String, mPass As String)
    DoEvents
    m_State = POP3_Connect
    
    
    Form2.txtHost.Text = mHost
    Form2.TxtUserName.Text = mUser
    Form2.TxtPassword.Text = mPass
    
    MsgStatus = "Checking New mails from " & Form2.txtHost.Text
    ChangeIcon Form2.picReadingMail(3), MsgStatus
    
    Form2.Winsock1.Close
    Form2.Winsock1.LocalPort = 0
    Form2.Winsock1.Connect Form2.txtHost.Text, 110
    WaitforNextMail
    Pause 15

End Sub
