VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Back Drop"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "...."
      Height          =   330
      Index           =   0
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   390
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   345
      TabIndex        =   0
      Text            =   "C:\Windows\Internet.bmp"
      Top             =   3285
      Width           =   3480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   300
      ScaleHeight     =   315
      ScaleWidth      =   3990
      TabIndex        =   2
      Top             =   3210
      Width           =   4050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filename"
      Height          =   1605
      Left            =   180
      TabIndex        =   3
      Top             =   2880
      Width           =   4335
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   3345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1050
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancel"
         Height          =   330
         Index           =   3
         Left            =   2250
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1050
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Set as New"
         Height          =   330
         Index           =   2
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1050
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Preview"
         Height          =   330
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1050
         Width           =   1050
      End
   End
   Begin VB.Image imgView2 
      Height          =   1425
      Left            =   1260
      Stretch         =   -1  'True
      Top             =   585
      Width           =   2085
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   75
      X2              =   75
      Y1              =   75
      Y2              =   4590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   90
      X2              =   4605
      Y1              =   75
      Y2              =   75
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   4515
      Left            =   75
      Top             =   75
      Width           =   4545
   End
   Begin VB.Image imgView1 
      Height          =   2565
      Left            =   855
      Picture         =   "Form3.frx":0000
      Top             =   195
      Width           =   2865
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
Dim MSG
 On Error Resume Next
Select Case Index
   Case 0
     txtFileName.Text = Main.OpenFile("Open Picture", "Windows Picture Files (*.All Files)", "*.bmp")
   Case 1
      imgView2.Picture = LoadPicture(txtFileName)
   Case 2
    
    If Dir(txtFileName.Text) = "" Then
      MsgBox "Can't Find file" & " " & txtFileName, vbInformation
    Else
     Main.SetBackDrop txtFileName.Text
      Command1(4).Enabled = True
    End If
  
  Case 3
   Unload Form3
  Case 4
   Unload Form3
   
End Select
If Err Then MsgBox Err.Description

End Sub

Private Sub Form_Load()
 Main.CenterForm Form3
 
End Sub
