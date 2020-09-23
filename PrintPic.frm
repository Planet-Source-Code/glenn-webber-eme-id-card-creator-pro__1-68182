VERSION 5.00
Begin VB.Form PrintPic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2685
      Index           =   2
      Left            =   30
      ScaleHeight     =   2655
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   5970
      Width           =   9045
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2685
      Index           =   0
      Left            =   30
      ScaleHeight     =   2655
      ScaleWidth      =   4485
      TabIndex        =   1
      Top             =   3240
      Width           =   4515
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2685
      Index           =   1
      Left            =   4560
      ScaleHeight     =   2655
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   3240
      Width           =   4515
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2685
      Left            =   4620
      Top             =   90
      Width           =   4515
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2685
      Left            =   90
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "PrintPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
