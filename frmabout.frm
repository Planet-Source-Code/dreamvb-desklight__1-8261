VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About...."
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   420
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1425
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "By DreamVb http://www.lastsale.com/vb"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desktop Light"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   915
      TabIndex        =   0
      Top             =   210
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmabout.frx":0000
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmabout

End Sub

Private Sub Form_Load()
Main.CenterForm frmabout

End Sub
