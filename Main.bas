Attribute VB_Name = "Main"
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public PasswordOn As Boolean
Public Const SRCCOPY = &HCC0020
Public PasswordKey As String
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const SPI_SETDESKWALLPAPER = 20

Enum TWindow
  nShow = 1
  nHide = 0
End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function OpenFile(mTitle, mFileType, mFileExt As String) As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = mFileType + Chr$(0) + mFileExt
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path
        ofn.lpstrTitle = mTitle
        ofn.flags = 0
       
        A = GetOpenFileName(ofn)
        If (A) Then
                OpenFile = Trim$(ofn.lpstrFile)
        End If
        
 End Function

Public Function TDeskTop(En As Boolean)
Dim HDesk As Long
 HDesk = FindWindow("Progman", vbNullString)
  If HDesk <> 0 Then
   EnableWindow HDesk, En
   End If
   
End Function
Public Function HideTaskBar(En As TWindow)
 Dim hTBar
  hTBar = FindWindow("Shell_traywnd", vbNullString)
   If hTBar <> 0 Then
    ShowWindow hTBar, En
  End If
    
End Function
Public Function HideIcons(En As TWindow)
 Dim nTIcon
  nTIcon = FindWindow("Progman", vbNullString)
   If nTIcon <> 0 Then
    nTIcon = ShowWindow(nTIcon, En)
End If

End Function

Public Function DrawBar(Bar As PictureBox)
 Dim X, Y
 Dim Grade
  Bar.AutoRedraw = True
  X = Bar.Width
  Y = Bar.Height
  Grade = 255
    Do Until Grade = 0
    X = X - Bar.Width / 255 * 1
    Grade = Grade - 1
    Bar.Line (0, 0)-(X, Y), RGB(Grade + 5, Grade + 5, Grade + 5), BF
    Loop
End Function

Function SetBackDrop(Picname As String)
 WallPaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, Picname, 0)
 
End Function
Function CenterForm(Frm As Form)
 Frm.Top = (Screen.Height - Frm.Height) / 2
 Frm.Left = (Screen.Width - Frm.Width) / 2
 
End Function
Function MoveForm(mHwnd As Form)
ReleaseCapture
SendMessage mHwnd.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 1

End Function
