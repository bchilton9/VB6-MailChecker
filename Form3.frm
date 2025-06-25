VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Remove Items"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Check mail Now"
      Height          =   375
      Left            =   960
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1200
      ItemData        =   "Form3.frx":0000
      Left            =   630
      List            =   "Form3.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add New"
      Height          =   375
      Left            =   3060
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Remove"
      Height          =   375
      Left            =   2325
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   840
      Picture         =   "Form3.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   7080
      Left            =   0
      Picture         =   "Form3.frx":14CD
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr
Private Sub Command1_Click()
    
    Dim NoOfEntries
    arr = GetAllKeys(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker")
    If List1.ListIndex >= 0 Then
        
        DeleteKey HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(List1.ListIndex)
        List1.RemoveItem List1.ListIndex
        NoOfEntries = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "NoOfEntries", "")
        SaveSettingString HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "NoOfEntries", CStr(NoOfEntries) - 1
    Else
        MsgBox "Please select any item first"
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub Command3_Click()
    Form3.Hide
    arr = GetAllKeys(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker")
    If List1.ListIndex >= 0 Then
        CheckNewMailfromOneServer CStr(GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(List1.ListIndex), "ServerName", "")), CStr(GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(List1.ListIndex), "UserName", "")), CStr(GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(List1.ListIndex), "Password", ""))
    Else
        MsgBox "Please select any item first"
    End If
End Sub

Private Sub Form_Load()
    Dim NoOfEntries, arr
    Dim mHost As String, mUserName As String, mPassword As String
    arr = GetAllKeys(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker")
    NoOfEntries = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker", "NoOfEntries", "")
    If NoOfEntries = "" Or NoOfEntries = "0" Then
        NoOfEntries = 0
    Else
        
        For i = LBound(arr) To UBound(arr)
            mHost = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(i), "ServerName", "")
            mUserName = GetSettingString(HKEY_CURRENT_USER, "InTerSoft\Applications\EmailChecker\" & arr(i), "UserName", "")
            List1.AddItem mHost
        Next
    End If
End Sub
