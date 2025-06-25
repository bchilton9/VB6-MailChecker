VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Setup POP3 account"
   ClientHeight    =   1965
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMailcommingin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picNoNewMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "Form2.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picPause 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   480
      Picture         =   "Form2.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picReadingMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   0
      Picture         =   "Form2.frx":114E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox picMail 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      Picture         =   "Form2.frx":1A18
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TxtUserName 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Server:"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "Form2.frx":1E5A
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   3855
   End
   Begin VB.Menu mnuSystray 
      Caption         =   ""
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuReadMailNow 
         Caption         =   "&Check Mail Now"
      End
      Begin VB.Menu mnutemp 
         Caption         =   "----"
      End
      Begin VB.Menu Mne_play 
         Caption         =   "&Play"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Pause 
         Caption         =   "Pa&use"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuDelay 
         Caption         =   "Set &Delay"
      End
      Begin VB.Menu Mnu_Quit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    Me.Hide
    showForm = True
End Sub

Private Sub Command1_Click()
    
    For Each c In Controls
        If TypeOf c Is TextBox Then
            If Len(c.Text) = 0 Then
                MsgBox "please Enter Required Fields"
                Exit Sub
            End If
        End If
    Next
    Me.Hide
    Dim NoOfEntries, MaxVal, arr
    NoOfEntries = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "NoOfEntries", "")
    If NoOfEntries = "" Or NoOfEntries = "0" Then
        NoOfEntries = 1
    Else
        arr = GetAllKeys(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker")
        Dim count
        count = 1
        For i = LBound(arr) To UBound(arr)
            MaxVal = arr(i)
            count = count + 1
        Next

        NoOfEntries = count
    End If
    SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "NoOfEntries", CStr(NoOfEntries)
    SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & MaxVal + 1, "ServerName", txtHost
    SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & MaxVal + 1, "UserName", TxtUserName
    SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & MaxVal + 1, "Password", TxtPassword
    
End Sub


Private Sub Form_Terminate()
    DeleteIcon picMail
    DeleteIcon picMailcommingin(2)
End Sub



Private Sub Mne_play_Click()
    If Play = True Then
        MsgBox "alredy running"
        Exit Sub
    End If
    showForm = True
    Play = True
    ChangeIcon Form2.picMail, MsgStatus
    Call Main
End Sub

Private Sub Mnu_Pause_Click()
    If Play = False Then
        MsgBox "alredy Paused"
        Exit Sub
    End If
    Winsock1.Close
    KillTimer 0, lngTimerID
    showForm = True
    Play = False
    ChangeIcon Form2.picPause(5), "Pause"
End Sub

Private Sub Mnu_Quit_Click()
    If MsgBox("Are you Sure to Quit?", vbYesNo, Form2.txtHost.Text) = vbYes Then
        KillTimer 0, lngTimerID
        Winsock1.Close
        DeleteIcon picMail
        DeleteIcon picMailcommingin(2)
        Unload Me
        End
    End If
    showForm = True
End Sub

Private Sub mnuDelay_Click()
    Dim delay As String
    delay = "ss"
    Do
        delay = InputBox("Enter Time Delay (in minutes) before checking new Mails again", "Delay", 10)
    Loop While Not IsNumeric(delay)
    
    SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "Delay", delay
    KillTimer 0, lngTimerID
    Winsock1.Close
    'Call Main
End Sub

Private Sub mnuReadMailNow_Click()
    KillTimer 0, lngTimerID
    showForm = True
    Call checknewmail
    Call Main
End Sub

Private Sub MnuSetup_Click()
    KillTimer 0, lngTimerID
    showForm = True
    Form1.Hide
    Form3.Show
End Sub

Private Sub mnutemp_Click()
    showForm = True
    HangUp
End Sub

Private Sub mnutmp_Click()
    showForm = True
    HangUp
End Sub

Private Sub picMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Select Case X
        Case trayLBUTTONUP
            If showForm = True Then
                Form1.Show
                Form1.Label1.Caption = MsgStatus
                Form1.lblServer.Caption = Form2.txtHost.Text
                sndPlaySound strFilename, 2
                Pause 15
                Form1.Hide
            End If

        Case trayRBUTTONUP
            Form1.Hide
            Form2.Hide
            showForm = False
            PopupMenu mnuSystray
            
        Case Else
    End Select
End Sub

Private Sub picMailcommingin_Click(Index As Integer)
  Select Case X
        Case trayLBUTTONUP
            If showForm = True Then
                Form1.Show
                Form1.Label1.Caption = MsgStatus
                Form1.lblServer.Caption = Form2.txtHost.Text
                sndPlaySound strFilename, 2
                Pause 15
                Form1.Hide
            End If

        Case trayRBUTTONUP
            Form1.Hide
            Form2.Hide
            showForm = False
            PopupMenu mnuSystray
            
        Case Else
    End Select

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    'Save the received data into strData variable
    On Error GoTo terror
    Winsock1.GetData strData
            

    If Left$(strData, 1) = "+" Then
        Select Case m_State
            Case POP3_Connect
                '
                'Reset the number of messages
                intMessages = 0
                '
                'Change current state of session
                m_State = POP3_USER
                '
                'Send to the server the USER command with the parameter.
                'The parameter is the name of the mail box
                'Don't forget to add vbCrLf at the end of the each command!
                Winsock1.SendData "USER " & TxtUserName & vbCrLf
                
                'Here is the end of Winsock1_DataArrival routine until the
                'next appearing of the DataArrival event. But next time this
                'section will be skipped and execution will start right after
                'the Case POP3_USER section.
            Case POP3_USER
                '
                'This part of the code runs in case of successful response to
                'the USER command.
                'Now we have to send to the server the user's password
                '
                'Change the state of the session
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & TxtPassword & vbCrLf
                
            Case POP3_PASS
                '
                'The server answered positively to the process of the
                'identification and now we can send the STAT command. As a
                'response the server is going to return the number of
                'messages in the mail box and its size in octets
                '
                ' Change the state of the session
                m_State = POP3_STAT
                '
                'Send STAT command to know how many
                'messages in the mailbox
                Winsock1.SendData "STAT" & vbCrLf
                
            Case POP3_STAT
                '
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                '(there are messages). Evidently, the first of all we have to
                'find out the first numeric value that contains in the
                'server's response
                TotalMails = 0
                TotalMails = CInt(Mid$(strData, 5, _
                              InStr(5, strData, " ") - 5))
                'If intMessages > 0 Then
                    '
                    'Oops. There is something in the mailbox!
                    'Change the session state
                    'm_State = POP3_RETR
                    '
                    'Increment the number of messages by one
                    'intCurrentMessage = intCurrentMessage + 1
                    '
                    'and we're sending to the server the RETR command in
                    'order to retrieve the first message
                    'Winsock1.SendData "RETR 1" & vbCrLf
                    
                'Else
                    'The mailbox is empty. Send the QUIT command to the
                    'server in order to close the session
                m_State = POP3_QUIT
                Winsock1.SendData "QUIT" & vbCrLf
                
                'MsgBox "You have not mail.", vbInformation
                'End If
            Case POP3_RETR
            Case POP3_QUIT
                'No matter what data we've received it's important
                'to close the connection with the mail server
                Winsock1.Close
                'Now we're calling the ListMessages routine in order to
                'fill out the ListView control with the messages we've          
                'downloaded
                
                If TotalMails > 0 Then
                    
                    MsgStatus = "You Have " & TotalMails & " Mail(s) in your Mail box"
                    ChangeIcon Form2.picMailcommingin(2), MsgStatus
                    
                    Form1.lblServer.Caption = Form2.txtHost.Text
                    Form1.Label1.Caption = MsgStatus
                    sndPlaySound strFilename, 2
                    Form1.Show
                    Pause 15
                    ChangeIcon Form2.picMail, MsgStatus
                    Form1.Hide
                Else
                    MsgStatus = "No New Mail from " & Form2.txtHost.Text
                    ChangeIcon Form2.picNoNewMail(1), MsgStatus
                End If
                WaitFlag = False
        End Select
    Else
        'As you see, there is no sophisticated error
        'handling. We just close the socket and show the server's response
        'That's all. By the way even fully featured mail applications
        'do the same.
            WaitFlag = False
            Winsock1.Close
            MsgBox "POP3 Error: " & strData, _
            vbExclamation, "POP3 Error"
            ChangeIcon Form2.picMail, "Idle"
    End If
terror:
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: #" & Number & vbCrLf & _
            Description
End Sub


