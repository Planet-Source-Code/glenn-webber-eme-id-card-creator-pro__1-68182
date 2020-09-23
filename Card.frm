VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   ScaleHeight     =   6105
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox CardB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   4740
      TabIndex        =   9
      Top             =   3060
      Width           =   4766
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   3855
         TabIndex        =   11
         Top             =   2550
         Width           =   3855
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Signature:"
            Height          =   255
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   765
         End
         Begin VB.Line Line1 
            X1              =   30
            X2              =   3780
            Y1              =   270
            Y2              =   270
         End
      End
      Begin VB.PictureBox BARCODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1740
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   93
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   1392
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   705
         Left            =   120
         Stretch         =   -1  'True
         Top             =   150
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   705
         Left            =   120
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   1050
         TabIndex        =   16
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label3"
         Height          =   195
         Index           =   2
         Left            =   1020
         TabIndex        =   14
         Top             =   1050
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label4"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   13
         Top             =   1410
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox CardF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3060
      Left            =   0
      ScaleHeight     =   3030
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   0
      Width           =   4766
      Begin VB.Image Logo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Left            =   1770
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Image Photo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   3450
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   450
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   780
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   5
         Top             =   1140
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   4
         Top             =   1500
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   3
         Top             =   1890
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   2
         Top             =   2250
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
