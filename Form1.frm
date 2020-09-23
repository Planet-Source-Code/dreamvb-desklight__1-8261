VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   7245
   ClientTop       =   2955
   ClientWidth     =   2265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   2265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2265
      TabIndex        =   8
      Top             =   0
      Width           =   2265
      Begin VB.Image imgUp 
         Height          =   240
         Left            =   1710
         Picture         =   "Form1.frx":030A
         Top             =   15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1965
         Picture         =   "Form1.frx":064C
         Top             =   15
         Width           =   240
      End
      Begin VB.Image imgDown 
         Height          =   240
         Left            =   1710
         Picture         =   "Form1.frx":098E
         Top             =   15
         Width           =   240
      End
   End
   Begin VB.PictureBox Cover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3885
      Picture         =   "Form1.frx":0CD0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   3525
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   45
      X2              =   1830
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   45
      X2              =   1830
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About Desktop Light"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   7
      Top             =   3390
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Desktop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   6
      Top             =   2910
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide DeskTop Icons"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   90
      TabIndex        =   5
      Top             =   2415
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Desktop Pic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   4
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide TaskBar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   1425
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Run Program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   915
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shut Down Windows"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   435
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   6
      Left            =   30
      Top             =   3345
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   5
      Left            =   30
      Top             =   2865
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   4
      Left            =   30
      Top             =   2370
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   3
      Left            =   30
      Top             =   1890
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   30
      Top             =   1395
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   30
      Top             =   405
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   30
      Top             =   900
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   15
      X2              =   1800
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -75
      X2              =   1800
      Y1              =   345
      Y2              =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mPicWidth As Integer
Dim mPicHeight As Integer
Dim TitleOn, UpDown, TaskOnOff, DeskLock, IconOnOff As Boolean
Private TMaxHeight As Long
Private TMaxWidth As Long
Sub RemoveRover()
Dim I As Integer
 For I = 0 To 6
  Shape1(I).Visible = False
  Label1(I).ForeColor = vbBlack
  Next
  
End Sub
Sub HideTaskBarA()
 Select Case TaskOnOff
  Case True
   Label1(2).Caption = "Hide TaskBar"
    Main.HideTaskBar nShow
   TaskOnOff = False
   Case False
   Label1(2).Caption = "Show TaskBar"
   Main.HideTaskBar nHide
   TaskOnOff = True
 End Select
 
End Sub

Sub DeskLockA()
 Select Case DeskLock
  Case True
   Label1(5).Caption = "Lock Desktop"
   Password.Show
   DeskLock = False
   Case False
   Label1(5).Caption = "Unlock Desktop"
   frmPass1.Show
   DeskLock = True
 End Select
 
End Sub
Sub HideIconsA()
 Select Case IconOnOff
  Case True
   Label1(4).Caption = "Hide DeskTop Icons"
    Main.HideIcons nShow
   IconOnOff = False
   Case False
   Label1(4).Caption = "Show DeskTop Icons"
   Main.HideIcons nHide
   IconOnOff = True
 End Select
 
End Sub




Private Sub Form_Load()
 Dim OldW, OldH As Integer
 Dim TotalSize As String
 
  OldW = Screen.Width / Screen.TwipsPerPixelX
  OldH = Screen.Height / Screen.TwipsPerPixelY
   TotalSize = OldW & "x" & OldH
   
    Select Case TotalSize
      Case "800x600"
            
            With Form1
             .Left = 9615
             .Top = 4740
             .Height = 3840
             .Width = 2355
            End With
            
      Case "1024x768"
            
            With Form1
             .Left = 12900
             .Top = 7245
             .Height = 3840
             .Width = 2355
            End With
  
     Case "640x840"
     
          With Form1
           .Left = 7200
           .Top = 2910
           .Height = 3840
           .Width = 2355
         End With
End Select

 Main.DrawBar Picture1
 TitleOn = True
With Cover
 .AutoSize = True
  mPicWidth = .ScaleWidth
  mPicHeight = .ScaleHeight
  .Visible = False
End With


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RemoveRover

End Sub

Private Sub Form_Paint()
  Dim mCol As Integer
  Dim mRow As Integer
  Dim mRet As Long
  If TitleOn Then
      For mRow = 0 To TMaxHeight Step mPicHeight
               For mCol = 0 To TMaxWidth Step mPicWidth
              mRet = BitBlt(hDC, mCol, mRow, mPicWidth, mPicHeight, Cover.hDC, 0, 0, Main.SRCCOPY)
          Next
      Next
End If
End Sub

Private Sub Form_Resize()
  TMaxHeight = Height \ Screen.TwipsPerPixelY
  TMaxWidth = Width \ Screen.TwipsPerPixelX
    Line1(0).X2 = Me.Width
    Line1(1).X2 = Me.Width
    Line1(2).X2 = Me.Width
    Line1(3).X2 = Me.Width
    
End Sub

Private Sub Image1_Click()
Dim Answer
 Answer = _
  MsgBox("Are you sure you want ot quit now", _
  vbYesNo)
   If Answer = vbYes Then
   End
   Else
   End If
   
End Sub

Private Sub imgDown_Click()
imgDown.Visible = False
Form1.Height = 3840
imgUp.Visible = True

End Sub

Private Sub imgUp_Click()
imgUp.Visible = False
Form1.Height = 410
imgDown.Visible = True
End Sub

Private Sub Label1_Click(Index As Integer)
 Select Case Index
  Case 0
   Main.sndPlaySound App.Path & "\Butclick.wav", 1
   Main.SHShutDownDialog 0
  Case 1
   Main.sndPlaySound App.Path & "\Butclick.wav", 1
   Form2.Show
  Case 2
   Main.sndPlaySound App.Path & "\Butclick.wav", 1
   HideTaskBarA
  Case 3
   Main.sndPlaySound App.Path & "\Butclick.wav", 1
   Form3.Show
  Case 4
   Main.sndPlaySound App.Path & "\Butclick.wav", 1
   HideIconsA
  Case 5
   Main.sndPlaySound App.Path & "\Butclick.wav", 1
   DeskLockA
  Case 6
   frmabout.Show
 End Select
 

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
  Case 0
   Shape1(0).Visible = True
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   Shape1(3).Visible = False
   Shape1(4).Visible = False
   Shape1(5).Visible = False
   Shape1(6).Visible = False
   Label1(0).ForeColor = vbWhite
  
  Case 1
   Shape1(0).Visible = False
   Shape1(1).Visible = True
   Shape1(2).Visible = False
   Shape1(3).Visible = False
   Shape1(4).Visible = False
   Shape1(5).Visible = False
   Shape1(6).Visible = False
   Label1(1).ForeColor = vbWhite
  
  Case 2
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = True
   Shape1(3).Visible = False
   Shape1(4).Visible = False
   Shape1(5).Visible = False
   Shape1(6).Visible = False
   Label1(2).ForeColor = vbWhite
  
  Case 3
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   Shape1(3).Visible = True
   Shape1(4).Visible = False
   Shape1(5).Visible = False
   Shape1(6).Visible = False
   Label1(3).ForeColor = vbWhite
   
  Case 4
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   Shape1(3).Visible = False
   Shape1(4).Visible = True
   Shape1(5).Visible = False
   Shape1(6).Visible = False
   Label1(4).ForeColor = vbWhite
   
  Case 5
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   Shape1(3).Visible = False
   Shape1(4).Visible = False
   Shape1(5).Visible = True
   Shape1(6).Visible = False
   Label1(5).ForeColor = vbWhite
   
   
  Case 6
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   Shape1(3).Visible = False
   Shape1(4).Visible = False
   Shape1(5).Visible = False
   Shape1(6).Visible = True
   Label1(6).ForeColor = vbWhite
   
End Select

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RemoveRover
 If Button = 1 Then
  Main.MoveForm Form1
 End If
 
End Sub
