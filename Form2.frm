VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run Program"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Browse"
      Height          =   375
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1260
      Width           =   900
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Text            =   "C:\WINDOWS\NOTEPAD.EXE"
      Top             =   705
      Width           =   4410
   End
   Begin VB.PictureBox Cover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4005
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   1140
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":317A
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mPicWidth As Integer
Dim mPicHeight As Integer
Dim TitleOn As Boolean
Private TMaxHeight As Long
Private TMaxWidth As Long



Private Sub Command1_Click()
 Dim TCommand As Long
 If txtFileName.Text = "" Then
  MsgBox "Please Enter or select a filename", vbInformation
  Else
  TCommand = ShellExecute(Me.hwnd, vbNullString, txtFileName, vbNullString, "", 1)
    If Dir(txtFileName) = "" Then ' See if the file is here
       MsgBox "Filename" & " " & txtFileName & " " & "not found", vbInformation ' Will Display when now file found
       End If
   End If
   
End Sub

Private Sub Command2_Click()
 Unload Form2
 
End Sub

Private Sub Command3_Click()
txtFileName.Text = Main.OpenFile("Open Program", "Windows Picture Files (*.All Files)", "*.exe")

End Sub

Private Sub Form_Load()
 TitleOn = True
With Cover
 .AutoSize = True
  mPicWidth = .ScaleWidth
  mPicHeight = .ScaleHeight
  .Visible = False
End With
End Sub

Private Sub Form_Paint()
 Dim mCol As Integer
  Dim mRow As Integer
  Dim mRet As Long
  If TitleOn Then
      For mRow = 0 To TMaxHeight Step mPicHeight
               For mCol = 0 To TMaxWidth Step mPicWidth
              mRet = Main.BitBlt(hDC, mCol, mRow, mPicWidth, mPicHeight, Cover.hDC, 0, 0, Main.SRCCOPY)
          Next
      Next
End If

End Sub

Private Sub Form_Resize()
  TMaxHeight = Height \ Screen.TwipsPerPixelY
  TMaxWidth = Width \ Screen.TwipsPerPixelX
  
End Sub

Private Sub txtFilename_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
  Command1_Click
  End If
  
End Sub

