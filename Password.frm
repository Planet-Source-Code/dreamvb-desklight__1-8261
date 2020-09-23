VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Password.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1245
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1245
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   705
      Width           =   2910
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   4320
      TabIndex        =   0
      Top             =   0
      Width           =   4320
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password to UnLock Desktop"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Top             =   30
         Width           =   3210
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Password"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   -90
      X2              =   1530
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -30
      X2              =   1590
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Text1(0).Text = Main.PasswordKey Then
  Main.TDeskTop True
  Main.PasswordOn = False
  MsgBox "The Desktop has been Unlocked", vbInformation
  Unload Password
  Else
   MsgBox "You have entered the wrong password", vbCritical
  Main.TDeskTop False
  Unload Password
  End If
  
End Sub

Private Sub Command2_Click()
 Unload Password
 
End Sub

Private Sub Form_Load()
Main.CenterForm Password
 If Main.PasswordOn = False Then
  Command1.Enabled = False
 End If
  
End Sub

Private Sub Form_Resize()
Line1(0).X2 = Me.Width
Line1(1).X2 = Me.Width

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  Main.MoveForm Password
  End If
  
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  Main.MoveForm Password
  End If
  
End Sub
