VERSION 5.00
Begin VB.Form frmQuake 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6930
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&SOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   5
      Top             =   7635
      Width           =   1635
   End
   Begin VB.OptionButton Option4 
      Caption         =   "7.9 On Richter, Bhuj, 2001 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5745
      TabIndex        =   4
      Top             =   5685
      Width           =   3570
   End
   Begin VB.OptionButton Option3 
      Caption         =   "6.9 On Richter, Taiwan, 1999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5745
      TabIndex        =   3
      Top             =   4155
      Width           =   3885
   End
   Begin VB.OptionButton Option2 
      Caption         =   "6.5 On Richter, Uttarkashi, 1998"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5745
      TabIndex        =   2
      Top             =   2355
      Width           =   4305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "5.9 On Richter, Latur, 1994"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5775
      TabIndex        =   1
      Top             =   1545
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Experience"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10020
      TabIndex        =   0
      Top             =   7650
      Width           =   1635
   End
End
Attribute VB_Name = "frmQuake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'if u have any doubts u can mail to me tk_pramod@yahoo.com
'funny funny

Private Declare Function waveOutSetVolume Lib "winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Private Declare Function waveOutGetVolume Lib "winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
Private Declare Function SystemParametersInfo Lib "USER32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function InvalidateRect Lib "USER32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2
Dim sFileName As String
Private Const SPI_SCREENSAVERRUNNING = 97
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Function shake(T As Long, L As Long, X As Integer, Y As Integer)
'this is the the quake
X = X * 20
Y = Y * 20
Dim rtn As Long

For i = X To 0 Step -10
    If i Mod 40 = 0 Then
        rtn = sndPlaySound(sFileName, SND_ASYNC)
    End If
    For j = Y To 0 Step -20
        Me.Move L + 0, T + i
        Me.Move L + i, T + 0
        Me.Move L + 0, T - i
        Me.Move L - i, T + 0
    Next
Next

End Function

           
   
Private Sub Command1_Click()
If Not (Option1.Value Or Option2.Value Or Option3.Value Or Option4.Value) Then Exit Sub
'here set the wallpaper
a = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "c:\windows\Gujarath Relif Fund.bmp", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
a = InvalidateRect(0, 0, 1)
Dim rtn As Long
'find a temp file where to extract the wave file in res
GetTempFile "", "~rs", 0, sFileName
sFileName = sFileName + ".wav"
'Save the resource item to disk
Me.Option1.Visible = False
Me.Option2.Visible = False
Me.Option3.Visible = False
Me.Option4.Visible = False
Me.Command1.Value = False
Me.Command1.Visible = False
Me.Command2.Visible = False
Me.Refresh

Dim i As Long
Dim tmp, vol As String
vol = "9999939"
tmp = Right((Hex$(vol + 65536)), 4)
vol = CLng("&H" & tmp & tmp)
a = waveOutSetVolume(0, vol)
'here we save the wave file from .res to HDD
If SaveResItemToDisk(101, "Custom", sFileName) = 0 Then
    
End If
'we call shake function
If (Me.Option1.Value) Then shake Left, Top, 2, 2
If (Me.Option2.Value) Then shake Left, Top, 4, 4
If (Me.Option3.Value) Then shake Left, Top, 6, 6
If (Me.Option4.Value) Then shake Left, Top, 8, 8
'here we make no sound
rtn = sndPlaySound("", SND_ASYNC)
'Delete the temp file
Kill sFileName

Command2_Click

End Sub

Private Sub Command2_Click()
Dim ret As Integer
Dim pOld As Long
'set our window as normal one
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'u can access alt+ctrl+del and ctrl+escp
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
MsgBox "plz, contribute something to Gujarath Earthquake Relief Fund" & vbCrLf & vbCrLf & "Thanx to pratheesh babu, 4 this idea" & vbCrLf & "with love" & vbCrLf & "pramod kumar" & vbCrLf & "u can reach me at tk_pramod@yahoo.com"
'showing some error msg like windows GPF
If Me.Option4.Value = True Then FatalAppExit 0, "7.9 On Richter is very high This system also crashed in this quake ..."
End
End Sub

Private Sub Form_Load()

Dim Lft As Long, Tp As Long
Me.Left = 0
Me.Top = 0
'here sets the ctrl in currect place
Me.Height = Screen.Height
Me.Width = Screen.Width
Tp = (Me.Height - Me.Option1.Height) / 6
Lft = (Me.Width - Me.Option1.Width) / 2
Me.Option1.Left = Lft
Me.Option2.Left = Lft
Me.Option3.Left = Lft
Me.Option4.Left = Lft
Me.Command2.Left = Lft - Me.Command2.Width
Me.Command1.Left = Lft + Me.Option1.Width
Me.Option1.Top = 1 * Tp
Me.Option2.Top = 2 * Tp
Me.Option3.Top = 3 * Tp
Me.Option4.Top = 4 * Tp
Me.Command1.Top = 5 * Tp
Me.Command2.Top = 5 * Tp
Dim ret As Integer
Dim pOld As Long

'desable alt+ctrl+del & ctrl+escp etc.
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
'make our window topmost
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'here we capturing the desktop and make it is the picture of the form
'it will looks like the desktop window (some tricks)
Me.Picture = CaptureScreen
'here we store the picture in res file to p-icture ctrl
Me.Picture1.Picture = LoadPictureResource(102, "Custom")
'save the picture to HDD
SavePicture Picture1.Picture, "c:\windows\Gujarath Relif Fund.bmp"

End Sub
