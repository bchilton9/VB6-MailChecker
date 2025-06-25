VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   20280
   ClientTop       =   15210
   ClientWidth     =   3765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OK"
      Height          =   255
      Left            =   1253
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   773
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   780
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2490
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7080
      Left            =   -360
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Label1_Click()
run = Shell("C:\Program Files\Mozilla Thunderbird\thunderbird.exe", vbNormalFocus)
Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFC0FF
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFC0FF
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&
End Sub
