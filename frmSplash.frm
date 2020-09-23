VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5715
   ClientLeft      =   3795
   ClientTop       =   3330
   ClientWidth     =   9435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   3690
      ScaleHeight     =   2655
      ScaleWidth      =   4440
      TabIndex        =   2
      Top             =   2580
      Width           =   4470
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   60
         Picture         =   "frmSplash.frx":08CA
         ScaleHeight     =   750
         ScaleWidth      =   840
         TabIndex        =   5
         Top             =   1830
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   1020
         X2              =   3420
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Card Creator"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   2130
         TabIndex        =   4
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "eMe"
         BeginProperty Font 
            Name            =   "Neurochrome"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   330
         TabIndex        =   3
         Top             =   60
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1230
      Picture         =   "frmSplash.frx":12F0
      ScaleHeight     =   2655
      ScaleWidth      =   4440
      TabIndex        =   6
      Top             =   2940
      Width           =   4440
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   150
      Picture         =   "frmSplash.frx":2792A
      ScaleHeight     =   2655
      ScaleWidth      =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   4470
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   6390
      Top             =   3390
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2640
      Left            =   4830
      Picture         =   "frmSplash.frx":4E4EC
      ScaleHeight     =   2640
      ScaleWidth      =   4470
      TabIndex        =   0
      Top             =   150
      Width           =   4470
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngTransparency As Long
Public p_blnSplash As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
    On Error GoTo hErr
    
  
  Dim NormalWindowStyle As Long
  Dim col As Long
  Dim ret As Long
    
    NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 50, LWA_ALPHA

    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    col = RGB(0, 0, 0)
    SetLayeredWindowAttributes Me.hwnd, col, 50, LWA_COLORKEY

'... Set Form to center of Screen
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

Exit Sub
hErr:
    Select Case MsgBox(Err.Description, vbAbortRetryIgnore + _
       vbCritical, "An Error Occured")
        Case vbAbort
            Screen.MousePointer = vbDefault
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub Timer1_Timer()
    On Error GoTo hErr
    
    If m_lngTransparency < 255 Then
        m_lngTransparency = m_lngTransparency + 5
        
        If m_lngTransparency = 255 Then
            Timer1.Interval = 5000 '2500
        End If
    
    Else
            Form1.Show
            Unload Me
    End If

Exit Sub
hErr:
    Select Case MsgBox(Err.Description, vbAbortRetryIgnore + _
       vbCritical, "An Error Occured")
        Case vbAbort
            Screen.MousePointer = vbDefault
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub
