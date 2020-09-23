VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  eMe Card Creator Pro"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14880
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      Height          =   885
      Left            =   9870
      TabIndex        =   109
      Top             =   1680
      Width           =   4965
      Begin VB.OptionButton Option4 
         Caption         =   "Signature"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   112
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Barcode"
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   111
         Top             =   390
         Width           =   915
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Nill"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   110
         Top             =   390
         Width           =   555
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Back Image 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1605
      Index           =   1
      Left            =   9870
      TabIndex        =   100
      Top             =   2610
      Width           =   4965
      Begin VB.HScrollBar HScroll7 
         Height          =   225
         Left            =   2910
         TabIndex        =   104
         Top             =   510
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   225
         Left            =   2910
         TabIndex        =   103
         Top             =   990
         Width           =   1935
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Visible"
         Height          =   225
         Left            =   1770
         TabIndex        =   102
         Top             =   510
         Width           =   795
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Border"
         Height          =   195
         Left            =   1770
         TabIndex        =   101
         Top             =   990
         Value           =   1  'Checked
         Width           =   765
      End
      Begin IDCard.lvButtons_H lvButtons_H10 
         Height          =   345
         Left            =   240
         TabIndex        =   105
         Top             =   420
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H11 
         Height          =   345
         Left            =   240
         TabIndex        =   106
         Top             =   960
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "Reset"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Index           =   6
         Left            =   2910
         TabIndex        =   108
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Width"
         Height          =   255
         Index           =   7
         Left            =   2940
         TabIndex        =   107
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Back Image 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1605
      Index           =   0
      Left            =   9870
      TabIndex        =   91
      Top             =   30
      Width           =   4965
      Begin VB.HScrollBar HScroll5 
         Height          =   225
         Left            =   2910
         TabIndex        =   95
         Top             =   480
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   225
         Left            =   2910
         TabIndex        =   94
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Visible"
         Height          =   255
         Left            =   1770
         TabIndex        =   93
         Top             =   510
         Width           =   765
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Border"
         Height          =   195
         Left            =   1770
         TabIndex        =   92
         Top             =   990
         Value           =   1  'Checked
         Width           =   765
      End
      Begin IDCard.lvButtons_H lvButtons_H8 
         Height          =   345
         Left            =   240
         TabIndex        =   96
         Top             =   420
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H9 
         Height          =   345
         Left            =   240
         TabIndex        =   97
         Top             =   960
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "Reset"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Index           =   4
         Left            =   2940
         TabIndex        =   99
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Width"
         Height          =   255
         Index           =   5
         Left            =   2940
         TabIndex        =   98
         Top             =   780
         Width           =   735
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
      Left            =   9810
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   84
      Top             =   8040
      Visible         =   0   'False
      Width           =   1392
   End
   Begin VB.Frame Frame11 
      Caption         =   "Back Labels:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3525
      Left            =   9870
      TabIndex        =   61
      Top             =   4260
      Width           =   4965
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   7
         Left            =   2760
         TabIndex        =   124
         Top             =   3030
         Width           =   885
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   6
         Left            =   2760
         TabIndex        =   123
         Top             =   2670
         Width           =   885
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   5
         Left            =   2760
         TabIndex        =   122
         Top             =   2280
         Width           =   885
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   4
         Left            =   2760
         TabIndex        =   121
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   7
         Left            =   720
         MaxLength       =   20
         TabIndex        =   120
         Top             =   3000
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   720
         MaxLength       =   20
         TabIndex        =   119
         Top             =   2610
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   720
         MaxLength       =   20
         TabIndex        =   118
         Top             =   2220
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   720
         MaxLength       =   20
         TabIndex        =   117
         Top             =   1830
         Width           =   1965
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   3
         Left            =   2760
         TabIndex        =   77
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   2
         Left            =   2760
         TabIndex        =   76
         Top             =   1110
         Width           =   885
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   1
         Left            =   2760
         TabIndex        =   75
         Top             =   720
         Width           =   885
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Visible"
         Height          =   225
         Index           =   0
         Left            =   2760
         TabIndex        =   74
         Top             =   300
         Width           =   885
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   720
         MaxLength       =   20
         TabIndex        =   70
         Top             =   1440
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   720
         MaxLength       =   20
         TabIndex        =   69
         Top             =   1050
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   720
         MaxLength       =   20
         TabIndex        =   68
         Top             =   660
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   720
         MaxLength       =   20
         TabIndex        =   67
         Top             =   270
         Width           =   1965
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   0
         Left            =   3750
         TabIndex        =   78
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   1
         Left            =   3750
         TabIndex        =   79
         Top             =   630
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   2
         Left            =   3750
         TabIndex        =   80
         Top             =   1020
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   3
         Left            =   3750
         TabIndex        =   81
         Top             =   1410
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   4
         Left            =   3750
         TabIndex        =   125
         Top             =   1830
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   5
         Left            =   3750
         TabIndex        =   126
         Top             =   2220
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   6
         Left            =   3750
         TabIndex        =   127
         Top             =   2610
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H7 
         Height          =   315
         Index           =   7
         Left            =   3750
         TabIndex        =   128
         Top             =   3000
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         Caption         =   "Label 8:"
         Height          =   225
         Index           =   15
         Left            =   60
         TabIndex        =   116
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 7:"
         Height          =   225
         Index           =   14
         Left            =   60
         TabIndex        =   115
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 6:"
         Height          =   225
         Index           =   13
         Left            =   60
         TabIndex        =   114
         Top             =   2250
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 5:"
         Height          =   225
         Index           =   12
         Left            =   60
         TabIndex        =   113
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 4:"
         Height          =   225
         Index           =   11
         Left            =   60
         TabIndex        =   73
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 3:"
         Height          =   225
         Index           =   10
         Left            =   60
         TabIndex        =   72
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 2:"
         Height          =   225
         Index           =   9
         Left            =   60
         TabIndex        =   71
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 1:"
         Height          =   225
         Index           =   8
         Left            =   60
         TabIndex        =   66
         Top             =   330
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Front Labels:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3525
      Left            =   4860
      TabIndex        =   21
      Top             =   4260
      Width           =   4965
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   56
         Top             =   3030
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   55
         Top             =   2640
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   54
         Top             =   2250
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   53
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   720
         MaxLength       =   20
         TabIndex        =   49
         Top             =   3000
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   720
         MaxLength       =   20
         TabIndex        =   48
         Top             =   2610
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   720
         MaxLength       =   20
         TabIndex        =   47
         Top             =   2220
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   720
         MaxLength       =   20
         TabIndex        =   46
         Top             =   1830
         Width           =   1965
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   33
         Top             =   1470
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   32
         Top             =   1080
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   31
         Top             =   690
         Width           =   795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   30
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   720
         MaxLength       =   20
         TabIndex        =   26
         Top             =   1440
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   720
         MaxLength       =   20
         TabIndex        =   25
         Top             =   1050
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   720
         MaxLength       =   20
         TabIndex        =   24
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   720
         MaxLength       =   20
         TabIndex        =   22
         Top             =   300
         Width           =   1935
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   0
         Left            =   3750
         TabIndex        =   41
         Top             =   270
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   1
         Left            =   3750
         TabIndex        =   42
         Top             =   660
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   2
         Left            =   3750
         TabIndex        =   43
         Top             =   1050
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   3
         Left            =   3750
         TabIndex        =   44
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   4
         Left            =   3750
         TabIndex        =   57
         Top             =   1830
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   5
         Left            =   3750
         TabIndex        =   58
         Top             =   2220
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   6
         Left            =   3750
         TabIndex        =   59
         Top             =   2610
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H6 
         Height          =   315
         Index           =   7
         Left            =   3750
         TabIndex        =   60
         Top             =   3000
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "Font"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         Caption         =   "Label 8:"
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   52
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 7:"
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   51
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 6:"
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   50
         Top             =   2250
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 5:"
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   45
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 4:"
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   29
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 3:"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 2:"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label 1:"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   23
         Top             =   330
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Logo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1605
      Left            =   4860
      TabIndex        =   10
      Top             =   2610
      Width           =   4965
      Begin VB.CheckBox Check6 
         Caption         =   "Border"
         Height          =   255
         Left            =   1770
         TabIndex        =   90
         Top             =   990
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Visible"
         Height          =   195
         Left            =   1770
         TabIndex        =   89
         Top             =   510
         Width           =   765
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   225
         Left            =   2910
         TabIndex        =   13
         Top             =   510
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   225
         Left            =   2910
         TabIndex        =   12
         Top             =   1020
         Width           =   1935
      End
      Begin IDCard.lvButtons_H lvButtons_H5 
         Height          =   345
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H12 
         Height          =   345
         Left            =   240
         TabIndex        =   82
         Top             =   960
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "Reset"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Width"
         Height          =   255
         Index           =   3
         Left            =   2940
         TabIndex        =   15
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox PicResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   840
      Left            =   6690
      ScaleHeight     =   780
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   8280
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8340
      Top             =   8130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicDes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1417
      Left            =   2340
      ScaleHeight     =   1380
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   8010
      Width           =   1134
   End
   Begin VB.Frame Frame2 
      Caption         =   "Photo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2565
      Left            =   4860
      TabIndex        =   6
      Top             =   30
      Width           =   4965
      Begin VB.CheckBox Check3 
         Caption         =   "Border"
         Height          =   225
         Left            =   3660
         TabIndex        =   88
         Top             =   270
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Visible"
         Height          =   225
         Left            =   2820
         TabIndex        =   87
         Top             =   270
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   225
         Left            =   2880
         TabIndex        =   17
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   2880
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.PictureBox PicSrc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   90
         ScaleHeight     =   2145
         ScaleWidth      =   2610
         TabIndex        =   7
         Top             =   270
         Width           =   2640
      End
      Begin IDCard.lvButtons_H lvButtons_H1 
         Height          =   345
         Left            =   3900
         TabIndex        =   133
         Top             =   2100
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H4 
         Height          =   345
         Left            =   2880
         TabIndex        =   134
         Top             =   2100
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         Caption         =   "Setup"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Width"
         Height          =   255
         Index           =   1
         Left            =   2910
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Index           =   0
         Left            =   2910
         TabIndex        =   18
         Top             =   690
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Card:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7755
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   4785
      Begin VB.Frame Frame5 
         Caption         =   "Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1065
         Left            =   1170
         TabIndex        =   139
         Top             =   6600
         Width           =   2475
         Begin IDCard.lvButtons_H lvButtons_H3 
            Height          =   345
            Left            =   120
            TabIndex        =   140
            Top             =   270
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            Caption         =   "Load"
            CapAlign        =   2
            BackStyle       =   2
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin IDCard.lvButtons_H lvButtons_H2 
            Height          =   345
            Left            =   750
            TabIndex        =   141
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Caption         =   "Save"
            CapAlign        =   2
            BackStyle       =   2
            Shape           =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin IDCard.lvButtons_H lvButtons_H17 
            Height          =   345
            Left            =   1500
            TabIndex        =   142
            Top             =   270
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            Caption         =   "Delete"
            CapAlign        =   2
            BackStyle       =   2
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin IDCard.lvButtons_H lvButtons_H18 
            Height          =   345
            Left            =   420
            TabIndex        =   143
            Top             =   660
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            Caption         =   "List Users"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Templates:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   765
         Left            =   150
         TabIndex        =   135
         Top             =   5820
         Width           =   4515
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   90
            TabIndex        =   136
            Text            =   "Combo1"
            Top             =   270
            Width           =   2565
         End
         Begin IDCard.lvButtons_H lvButtons_H13 
            Height          =   345
            Left            =   2730
            TabIndex        =   137
            Top             =   240
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   609
            Caption         =   "Delete"
            CapAlign        =   2
            BackStyle       =   2
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin IDCard.lvButtons_H lvButtons_H14 
            Height          =   345
            Left            =   3480
            TabIndex        =   138
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   609
            Caption         =   "Save"
            CapAlign        =   2
            BackStyle       =   2
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
      End
      Begin VB.PictureBox CardB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2685
         Left            =   150
         ScaleHeight     =   2655
         ScaleWidth      =   4485
         TabIndex        =   5
         Top             =   3090
         Width           =   4515
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "GAW"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4110
            TabIndex        =   146
            Top             =   2460
            Width           =   345
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0800055555"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3630
            TabIndex        =   145
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CHILD LINE"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   3690
            TabIndex        =   144
            Top             =   90
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            Height          =   195
            Index           =   7
            Left            =   1770
            TabIndex        =   132
            Top             =   1410
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   195
            Index           =   6
            Left            =   1770
            TabIndex        =   131
            Top             =   1050
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   195
            Index           =   5
            Left            =   1800
            TabIndex        =   130
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label5"
            Height          =   195
            Index           =   4
            Left            =   1830
            TabIndex        =   129
            Top             =   270
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   4110
            Y1              =   2430
            Y2              =   2430
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Signature:"
            Height          =   255
            Left            =   360
            TabIndex        =   85
            Top             =   2190
            Width           =   765
         End
         Begin VB.Image Image3 
            Height          =   435
            Left            =   1620
            Top             =   2010
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label4"
            Height          =   195
            Index           =   3
            Left            =   1020
            TabIndex        =   65
            Top             =   1410
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
            Height          =   195
            Index           =   2
            Left            =   1020
            TabIndex        =   64
            Top             =   1050
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   63
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Index           =   0
            Left            =   1050
            TabIndex        =   62
            Top             =   270
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   705
            Left            =   120
            Picture         =   "Main.frx":5C12
            Stretch         =   -1  'True
            Top             =   960
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   705
            Left            =   120
            Picture         =   "Main.frx":B824
            Stretch         =   -1  'True
            Top             =   150
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VB.PictureBox CardF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2685
         Left            =   150
         Picture         =   "Main.frx":11436
         ScaleHeight     =   2655
         ScaleWidth      =   4485
         TabIndex        =   4
         Top             =   300
         Width           =   4515
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   40
            Top             =   2280
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
            TabIndex        =   39
            Top             =   1920
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
            TabIndex        =   38
            Top             =   1620
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
            TabIndex        =   37
            Top             =   1350
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
            TabIndex        =   36
            Top             =   1050
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
            TabIndex        =   35
            Top             =   720
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
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   480
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
            TabIndex        =   20
            Top             =   90
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Photo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1575
            Left            =   3030
            Picture         =   "Main.frx":6A438
            Stretch         =   -1  'True
            Top             =   150
            Width           =   1245
         End
         Begin VB.Image Logo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   795
            Left            =   1620
            Picture         =   "Main.frx":7004A
            Stretch         =   -1  'True
            Top             =   1170
            Visible         =   0   'False
            Width           =   1035
         End
      End
      Begin IDCard.lvButtons_H lvButtons_H15 
         Height          =   495
         Left            =   3780
         TabIndex        =   83
         Top             =   6960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Caption         =   "Print Card"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin IDCard.lvButtons_H lvButtons_H16 
         Height          =   525
         Left            =   180
         TabIndex        =   86
         Top             =   6930
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   926
         Caption         =   "Save Card As Image"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9030
      Top             =   8160
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   4350
      ScaleHeight     =   795
      ScaleWidth      =   885
      TabIndex        =   2
      Top             =   8400
      Width           =   945
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   915
      Left            =   3300
      ScaleHeight     =   51.584
      ScaleMode       =   0  'User
      ScaleWidth      =   48.163
      TabIndex        =   1
      Top             =   8430
      Width           =   945
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5370
      Picture         =   "Main.frx":75C5C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   8250
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Dim distX As Single
Dim distY As Single

Dim Pic As StdPicture

Dim Clicked As Boolean
Dim xOffset As Long
Dim yOffset As Long
Dim CurX As Long
Dim CurY As Long

Dim px As Integer
Dim py As Integer
Dim cl As New arisBarcode
Dim RoddersMotion As New ClassMotion

Private Sub Check1_Click(Index As Integer)

If Label3(Index).Visible = True Then
Label3(Index).Visible = False
Else
Label3(Index).Visible = True
End If

End Sub

Private Sub Check10_Click()
If Image2.BorderStyle = 1 Then
Image2.BorderStyle = 0
Else
Image2.BorderStyle = 1

End If
End Sub

Private Sub Check2_Click()
If Photo.Visible = True Then
Photo.Visible = False
Else
Photo.Visible = True
End If
End Sub

Private Sub Check3_Click()
If Photo.BorderStyle = 1 Then
Photo.BorderStyle = 0
Else
Photo.BorderStyle = 1
End If
End Sub

Private Sub Check4_Click(Index As Integer)
If Label6(Index).Visible = True Then
Label6(Index).Visible = False
Else
Label6(Index).Visible = True
End If
End Sub


Private Sub Check5_Click()
If Logo.Visible = False Then
Logo.Height = 700
Logo.Width = 700
Logo.Visible = True
Else
Logo.Visible = False
End If
End Sub

Private Sub Check6_Click()
If Logo.BorderStyle = 0 Then
Logo.BorderStyle = 1
Else
Logo.BorderStyle = 0
End If
End Sub

Private Sub Check7_Click()
If Image1.Visible = True Then
Image1.Visible = False
Else
Image1.Visible = True
End If
End Sub

Private Sub Check8_Click()
If Image1.BorderStyle = 1 Then
Image1.BorderStyle = 0
Else
Image1.BorderStyle = 1
End If
End Sub

Private Sub Check9_Click()

If Image2.Visible = True Then
Image2.Visible = False
Else
Image2.Visible = True
End If

End Sub


Private Sub Combo1_Click()

On Error Resume Next

strSQL = "SELECT * FROM Templates WHERE Name=" & Quote & Trim(Combo1.List(Combo1.ListIndex)) & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)

If gSn.Fields("PhotoShow") = True Then
Check2.Value = 1
Else
Check2.Value = 0
End If
Photo.Top = gSn.Fields("PhotoTop")
Photo.Left = gSn.Fields("PhotoLeft")
Photo.Width = gSn.Fields("PhotoW")
Photo.Height = gSn.Fields("PhotoH")
If gSn.Fields("PhotoB") = True Then
Check3.Value = 1
Else
Check3.Value = 0
End If

If gSn.Fields("LogoShow") = True Then
Check5.Value = 1
Else
Check5.Value = 0
End If
Logo.Top = gSn.Fields("LogoTop")
Logo.Left = gSn.Fields("LogoLeft")
Logo.Width = gSn.Fields("LogoW")
Logo.Height = gSn.Fields("LogoH")
If gSn.Fields("LogoB") = True Then
Check6.Value = 1
Else
Check6.Value = 0
End If

If gSn.Fields("IM1Show") = True Then
Check7.Value = 1
Else
Check7.Value = 0
End If
Image1.Top = gSn.Fields("IM1Top")
Image1.Left = gSn.Fields("IM1Left")
Image1.Width = gSn.Fields("IM1W")
Image1.Height = gSn.Fields("IM1H")
If gSn.Fields("IM1B") = True Then
Check8.Value = 1
Else
Check8.Value = 0
End If

If gSn.Fields("IM2Show") = True Then
Check9.Value = 1
Else
Check9.Value = 0
End If
Image2.Top = gSn.Fields("IM2Top")
Image2.Left = gSn.Fields("IM2Left")
Image2.Width = gSn.Fields("IM2W")
Image2.Height = gSn.Fields("IM2H")
If gSn.Fields("IM2B") = True Then
Check10.Value = 1
Else
Check10.Value = 0
End If

If gSn.Fields("BarSig") = 1 Then
Option4(0).Value = True
End If
If gSn.Fields("BarSig") = 2 Then
Option4(1).Value = True
End If
If gSn.Fields("BarSig") = 3 Then
Option4(2).Value = True
End If

If gSn.Fields("Label1") = True Then
Check1(0).Value = 1
Else
Text1(0).Text = ""
Check1(0).Value = 0
End If
If gSn.Fields("FontBold1") = True Then
Label3(0).FontBold = True
Else
Label3(0).FontBold = False
End If
If gSn.Fields("FontItalic1") = True Then
Label3(0).FontItalic = True
Else
Label3(0).FontItalic = False
End If
If gSn.Fields("FontUnderline1") = True Then
Label3(0).FontUnderline = True
Else
Label3(0).FontUnderline = False
End If
Label3(0).FontName = gSn.Fields("FontName1")
Label3(0).ForeColor = gSn.Fields("Color1")
Label3(0).FontSize = gSn.Fields("FontSize1")
Label3(0).Top = gSn.Fields("Top1")
Label3(0).Left = gSn.Fields("Left1")

If gSn.Fields("Label2") = True Then
Check1(1).Value = 1
Else
Text1(1).Text = ""
Check1(1).Value = 0
End If
If gSn.Fields("FontBold2") = True Then
Label3(1).FontBold = True
Else
Label3(1).FontBold = False
End If
If gSn.Fields("FontItalic2") = True Then
Label3(1).FontItalic = True
Else
Label3(1).FontItalic = False
End If
If gSn.Fields("FontUnderline2") = True Then
Label3(1).FontUnderline = True
Else
Label3(1).FontUnderline = False
End If
Label3(1).FontName = gSn.Fields("FontName2")
Label3(1).ForeColor = gSn.Fields("Color2")
Label3(1).FontSize = gSn.Fields("FontSize2")
Label3(1).Top = gSn.Fields("Top2")
Label3(1).Left = gSn.Fields("Left2")

If gSn.Fields("Label3") = True Then
Check1(2).Value = 1
Else
Text1(2).Text = ""
Check1(2).Value = 0
End If
If gSn.Fields("FontBold3") = True Then
Label3(2).FontBold = True
Else
Label3(2).FontBold = False
End If
If gSn.Fields("FontItalic3") = True Then
Label3(2).FontItalic = True
Else
Label3(2).FontItalic = False
End If
If gSn.Fields("FontUnderline3") = True Then
Label3(2).FontUnderline = True
Else
Label3(2).FontUnderline = False
End If
Label3(2).FontName = gSn.Fields("FontName3")
Label3(2).ForeColor = gSn.Fields("Color3")
Label3(2).FontSize = gSn.Fields("FontSize3")
Label3(2).Top = gSn.Fields("Top3")
Label3(2).Left = gSn.Fields("Left3")

If gSn.Fields("Label4") = True Then
Check1(3).Value = 1
Else
Text1(3).Text = ""
Check1(3).Value = 0
End If
If gSn.Fields("FontBold4") = True Then
Label3(3).FontBold = True
Else
Label3(3).FontBold = False
End If
If gSn.Fields("FontItalic4") = True Then
Label3(3).FontItalic = True
Else
Label3(3).FontItalic = False
End If
If gSn.Fields("FontUnderline4") = True Then
Label3(3).FontUnderline = True
Else
Label3(3).FontUnderline = False
End If
Label3(3).FontName = gSn.Fields("FontName4")
Label3(3).ForeColor = gSn.Fields("Color4")
Label3(3).FontSize = gSn.Fields("FontSize4")
Label3(3).Top = gSn.Fields("Top4")
Label3(3).Left = gSn.Fields("Left4")

If gSn.Fields("Label5") = True Then
Check1(4).Value = 1
Else
Text1(4).Text = ""
Check1(4).Value = 0
End If
If gSn.Fields("FontBold5") = True Then
Label3(4).FontBold = True
Else
Label3(4).FontBold = False
End If
If gSn.Fields("FontItalic5") = True Then
Label3(4).FontItalic = True
Else
Label3(4).FontItalic = False
End If
If gSn.Fields("FontUnderline5") = True Then
Label3(4).FontUnderline = True
Else
Label3(4).FontUnderline = False
End If
Label3(4).FontName = gSn.Fields("FontName5")
Label3(4).ForeColor = gSn.Fields("Color5")
Label3(4).FontSize = gSn.Fields("FontSize5")
Label3(4).Top = gSn.Fields("Top5")
Label3(4).Left = gSn.Fields("Left5")

If gSn.Fields("Label6") = True Then
Check1(5).Value = 1
Else
Text1(5).Text = ""
Check1(5).Value = 0
End If
If gSn.Fields("FontBold6") = True Then
Label3(5).FontBold = True
Else
Label3(5).FontBold = False
End If
If gSn.Fields("FontItalic6") = True Then
Label3(5).FontItalic = True
Else
Label3(5).FontItalic = False
End If
If gSn.Fields("FontUnderline6") = True Then
Label3(5).FontUnderline = True
Else
Label3(5).FontUnderline = False
End If
Label3(5).FontName = gSn.Fields("FontName6")
Label3(5).ForeColor = gSn.Fields("Color6")
Label3(5).FontSize = gSn.Fields("FontSize6")
Label3(5).Top = gSn.Fields("Top6")
Label3(5).Left = gSn.Fields("Left6")

If gSn.Fields("Label7") = True Then
Check1(6).Value = 1
Else
Text1(6).Text = ""
Check1(6).Value = 0
End If
If gSn.Fields("FontBold7") = True Then
Label3(6).FontBold = True
Else
Label3(6).FontBold = False
End If
If gSn.Fields("FontItalic7") = True Then
Label3(6).FontItalic = True
Else
Label3(6).FontItalic = False
End If
If gSn.Fields("FontUnderline7") = True Then
Label3(6).FontUnderline = True
Else
Label3(6).FontUnderline = False
End If
Label3(6).FontName = gSn.Fields("FontName7")
Label3(6).ForeColor = gSn.Fields("Color7")
Label3(6).FontSize = gSn.Fields("FontSize7")
Label3(6).Top = gSn.Fields("Top7")
Label3(6).Left = gSn.Fields("Left7")

If gSn.Fields("Label8") = True Then
Check1(7).Value = 1
Else
Text1(7).Text = ""
Check1(7).Value = 0
End If
If gSn.Fields("FontBold8") = True Then
Label3(7).FontBold = True
Else
Label3(7).FontBold = False
End If
If gSn.Fields("FontItalic8") = True Then
Label3(7).FontItalic = True
Else
Label3(7).FontItalic = False
End If
If gSn.Fields("FontUnderline8") = True Then
Label3(7).FontUnderline = True
Else
Label3(7).FontUnderline = False
End If
Label3(7).FontName = gSn.Fields("FontName8")
Label3(7).ForeColor = gSn.Fields("Color8")
Label3(7).FontSize = gSn.Fields("FontSize8")
Label3(7).Top = gSn.Fields("Top8")
Label3(7).Left = gSn.Fields("Left8")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If gSn.Fields("Label9") = True Then
Check4(0).Value = 1
Else
Text2(0).Text = ""
Check4(0).Value = 0
End If
If gSn.Fields("FontBold9") = True Then
Label6(0).FontBold = True
Else
Label6(0).FontBold = False
End If
If gSn.Fields("FontItalic9") = True Then
Label6(0).FontItalic = True
Else
Label6(0).FontItalic = False
End If
If gSn.Fields("FontUnderline9") = True Then
Label6(0).FontUnderline = True
Else
Label6(0).FontUnderline = False
End If
Label6(0).FontName = gSn.Fields("FontName9")
Label6(0).ForeColor = gSn.Fields("Color9")
Label6(0).FontSize = gSn.Fields("FontSize9")
Label6(0).Top = gSn.Fields("Top9")
Label6(0).Left = gSn.Fields("Left9")

If gSn.Fields("Label10") = True Then
Check4(1).Value = 1
Else
Text2(1).Text = ""
Check4(1).Value = 0
End If
If gSn.Fields("FontBold10") = True Then
Label6(1).FontBold = True
Else
Label6(1).FontBold = False
End If
If gSn.Fields("FontItalic10") = True Then
Label6(1).FontItalic = True
Else
Label6(1).FontItalic = False
End If
If gSn.Fields("FontUnderline10") = True Then
Label6(1).FontUnderline = True
Else
Label6(1).FontUnderline = False
End If
Label6(1).FontName = gSn.Fields("FontName10")
Label6(1).ForeColor = gSn.Fields("Color10")
Label6(1).FontSize = gSn.Fields("FontSize10")
Label6(1).Top = gSn.Fields("Top10")
Label6(1).Left = gSn.Fields("Left10")

If gSn.Fields("Label11") = True Then
Check4(2).Value = 1
Else
Text2(2).Text = ""
Check4(2).Value = 0
End If
If gSn.Fields("FontBold11") = True Then
Label6(2).FontBold = True
Else
Label6(2).FontBold = False
End If
If gSn.Fields("FontItalic11") = True Then
Label6(2).FontItalic = True
Else
Label6(2).FontItalic = False
End If
If gSn.Fields("FontUnderline11") = True Then
Label6(2).FontUnderline = True
Else
Label6(2).FontUnderline = False
End If
Label6(2).FontName = gSn.Fields("FontName11")
Label6(2).ForeColor = gSn.Fields("Color11")
Label6(2).FontSize = gSn.Fields("FontSize11")
Label6(2).Top = gSn.Fields("Top11")
Label6(2).Left = gSn.Fields("Left11")

If gSn.Fields("Label12") = True Then
Check4(3).Value = 1
Else
Text2(3).Text = ""
Check4(3).Value = 0
End If
If gSn.Fields("FontBold12") = True Then
Label6(3).FontBold = True
Else
Label6(3).FontBold = False
End If
If gSn.Fields("FontItalic12") = True Then
Label6(3).FontItalic = True
Else
Label6(3).FontItalic = False
End If
If gSn.Fields("FontUnderline12") = True Then
Label6(3).FontUnderline = True
Else
Label6(3).FontUnderline = False
End If
Label6(3).FontName = gSn.Fields("FontName12")
Label6(3).ForeColor = gSn.Fields("Color12")
Label6(3).FontSize = gSn.Fields("FontSize12")
Label6(3).Top = gSn.Fields("Top12")
Label6(3).Left = gSn.Fields("Left12")

If gSn.Fields("Label13") = True Then
Check4(4).Value = 1
Else
Text2(4).Text = ""
Check4(4).Value = 0
End If
If gSn.Fields("FontBold13") = True Then
Label6(4).FontBold = True
Else
Label6(4).FontBold = False
End If
If gSn.Fields("FontItalic13") = True Then
Label6(4).FontItalic = True
Else
Label6(4).FontItalic = False
End If
If gSn.Fields("FontUnderline13") = True Then
Label6(4).FontUnderline = True
Else
Label6(4).FontUnderline = False
End If
Label6(4).FontName = gSn.Fields("FontName13")
Label6(4).ForeColor = gSn.Fields("Color13")
Label6(4).FontSize = gSn.Fields("FontSize13")
Label6(4).Top = gSn.Fields("Top13")
Label6(4).Left = gSn.Fields("Left13")

If gSn.Fields("Label14") = True Then
Check4(5).Value = 1
Else
Text2(5).Text = ""
Check4(5).Value = 0
End If
If gSn.Fields("FontBold14") = True Then
Label6(5).FontBold = True
Else
Label6(5).FontBold = False
End If
If gSn.Fields("FontItalic14") = True Then
Label6(5).FontItalic = True
Else
Label6(5).FontItalic = False
End If
If gSn.Fields("FontUnderline14") = True Then
Label6(5).FontUnderline = True
Else
Label6(5).FontUnderline = False
End If
Label6(5).FontName = gSn.Fields("FontName14")
Label6(5).ForeColor = gSn.Fields("Color14")
Label6(5).FontSize = gSn.Fields("FontSize14")
Label6(5).Top = gSn.Fields("Top14")
Label6(5).Left = gSn.Fields("Left14")

If gSn.Fields("Label15") = True Then
Check4(6).Value = 1
Else
Text2(6).Text = ""
Check4(6).Value = 0
End If
If gSn.Fields("FontBold15") = True Then
Label6(6).FontBold = True
Else
Label6(6).FontBold = False
End If
If gSn.Fields("FontItalic15") = True Then
Label6(6).FontItalic = True
Else
Label6(6).FontItalic = False
End If
If gSn.Fields("FontUnderline15") = True Then
Label6(6).FontUnderline = True
Else
Label6(6).FontUnderline = False
End If
Label6(6).FontName = gSn.Fields("FontName15")
Label6(6).ForeColor = gSn.Fields("Color15")
Label6(6).FontSize = gSn.Fields("FontSize15")
Label6(6).Top = gSn.Fields("Top15")
Label6(6).Left = gSn.Fields("Left15")

If gSn.Fields("Label16") = True Then
Check4(7).Value = 1
Else
Text2(7).Text = ""
Check4(7).Value = 0
End If
If gSn.Fields("FontBold16") = True Then
Label6(7).FontBold = True
Else
Label6(7).FontBold = False
End If
If gSn.Fields("FontItalic16") = True Then
Label6(7).FontItalic = True
Else
Label6(7).FontItalic = False
End If
If gSn.Fields("FontUnderline16") = True Then
Label6(7).FontUnderline = True
Else
Label6(7).FontUnderline = False
End If
Label6(7).FontName = gSn.Fields("FontName16")
Label6(7).ForeColor = gSn.Fields("Color16")
Label6(7).FontSize = gSn.Fields("FontSize16")
Label6(7).Top = gSn.Fields("Top16")
Label6(7).Left = gSn.Fields("Left16")

gSn.Close

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub Form_Terminate()
Set c = Nothing
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub
Private Sub Frame11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub Frame12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub Frame13_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MMove Then MMove = False
End Sub

Private Sub HScroll1_Change()

HScroll1.Min = 300
HScroll1.Max = 3000

HScroll1.SmallChange = (HScroll1.Max / 20) + 1
HScroll1.LargeChange = (HScroll1.Max / 5) + 1

Photo.Height = HScroll1.Value
Photo.Height = Photo.Height

End Sub

Private Sub HScroll4_Change()
HScroll4.Min = 300
HScroll4.Max = 4700

HScroll4.SmallChange = (HScroll4.Max / 20) + 1
HScroll4.LargeChange = (HScroll4.Max / 5) + 1

Photo.Width = HScroll4.Value
Photo.Width = Photo.Width

End Sub

Private Sub HScroll5_Change()
HScroll5.Min = 300
HScroll5.Max = 3000

HScroll5.SmallChange = (HScroll5.Max / 20) + 1
HScroll5.LargeChange = (HScroll5.Max / 5) + 1

Image1.Height = HScroll5.Value
Image1.Height = Image1.Height

End Sub

Private Sub HScroll6_Change()
HScroll6.Min = 300
HScroll6.Max = 4700

HScroll6.SmallChange = (HScroll6.Max / 20) + 1
HScroll6.LargeChange = (HScroll6.Max / 5) + 1

Image1.Width = HScroll6.Value
Image1.Width = Image1.Width

End Sub

Private Sub HScroll7_Change()
HScroll7.Min = 300
HScroll7.Max = 3000

HScroll7.SmallChange = (HScroll7.Max / 20) + 1
HScroll7.LargeChange = (HScroll7.Max / 5) + 1

Image2.Height = HScroll7.Value
Image2.Height = Image2.Height

End Sub

Private Sub HScroll8_Change()
HScroll8.Min = 300
HScroll8.Max = 4700

HScroll8.SmallChange = (HScroll8.Max / 20) + 1
HScroll8.LargeChange = (HScroll8.Max / 5) + 1

Image2.Width = HScroll8.Value
Image2.Width = Image2.Width

End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
distX = x - Image1.Left
distY = y - Image1.Top
Clicked = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Clicked Then
           Image1.Left = x - distX
           Image1.Top = y - distY
           Image1.Refresh
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Clicked = False

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
distX = x - Image2.Left
distY = y - Image2.Top
Clicked = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Clicked Then
           Image2.Left = x - distX
           Image2.Top = y - distY
           Image2.Refresh
End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Clicked = False
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

distX = x - Label3(Index).Left
distY = y - Label3(Index).Top
Clicked = True

End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Clicked Then
           Label3(Index).Left = x - distX
           Label3(Index).Top = y - distY
End If

End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Clicked = False
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
distX = x - Label6(Index).Left
distY = y - Label6(Index).Top
Clicked = True
End Sub

Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Clicked Then
           Label6(Index).Left = x - distX
           Label6(Index).Top = y - distY
End If

End Sub

Private Sub Label6_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Clicked = False
End Sub

Private Sub Logo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

distX = x - Logo.Left
distY = y - Logo.Top
Clicked = True

End Sub

Private Sub Logo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Clicked Then
           Logo.Left = x - distX
           Logo.Top = y - distY
           Logo.Refresh
End If
End Sub

Private Sub Logo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Clicked = False

End Sub

Private Sub lvButtons_H1_Click()
CommonDialog1.CancelError = True
On Error GoTo Err
CommonDialog1.InitDir = Path1 & "Photos\"
CommonDialog1.Filter = "Pictures (*.jpg;*.gif;*.bmp;*.ico)|*.jpg;*.gif;*.bmp;*.ico"
CommonDialog1.ShowOpen

 
        Photo.Picture = LoadPicture(CommonDialog1.FileName)


Exit Sub
Err:
End Sub

Private Sub lvButtons_H10_Click()
CommonDialog1.CancelError = True
On Error GoTo Err
CommonDialog1.InitDir = Path1 & "Graphics\"
CommonDialog1.Filter = "Pictures (*.jpg;*.gif;*.bmp;*.ico)|*.jpg;*.gif;*.bmp;*.ico"
CommonDialog1.ShowOpen

Image2.Height = 1000
Image2.Width = 1000
        
Image2.Picture = LoadPicture(CommonDialog1.FileName)

Exit Sub
Err:
End Sub

Private Sub lvButtons_H11_Click()
Image2.Height = 700
Image2.Width = 700

Image2.Picture = Picture2.Picture

End Sub

Private Sub lvButtons_H12_Click()
Logo.Height = 700
Logo.Width = 700

Logo.Picture = Picture2.Picture
Call CenLogo

End Sub

Private Sub lvButtons_H13_Click()

If MsgBox("Would You Like To Delete " & Trim(Combo1.List(Combo1.ListIndex)) & "?", vbQuestion Or vbYesNo) = vbYes Then
strSQL = "SELECT * FROM Templates WHERE Name=" & Quote & Trim(Combo1.List(Combo1.ListIndex)) & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
gSn.Delete
gSn.Close
End If

Call gettemps

End Sub

Private Sub lvButtons_H14_Click()

On Error Resume Next

Dim MenuName

MenuName = InputBox("Enter Template Name:", "New Template")
If MenuName <> "" Then
strSQL = "SELECT * FROM Templates WHERE Name=" & Quote & MenuName & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
If gSn.RecordCount > 0 Then
gSn.Delete
End If
gSn.Close

strSQL = "SELECT * FROM Templates"
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
gSn.AddNew
gSn.Fields("Name") = MenuName
gSn.Fields("PhotoTop") = Photo.Top
gSn.Fields("PhotoLeft") = Photo.Left
gSn.Fields("PhotoW") = Photo.Width
gSn.Fields("PhotoH") = Photo.Height
If Photo.Visible = True Then
gSn.Fields("PhotoShow") = 1
Else
gSn.Fields("PhotoShow") = 0
End If
If Photo.BorderStyle = 1 Then
gSn.Fields("PhotoB") = 1
Else
gSn.Fields("PhotoB") = 0
End If

gSn.Fields("LogoTop") = Logo.Top
gSn.Fields("LogoLeft") = Logo.Left
gSn.Fields("LogoW") = Logo.Width
gSn.Fields("LogoH") = Logo.Height
If Logo.Visible = True Then
gSn.Fields("LogoShow") = 1
Else
gSn.Fields("LogoShow") = 0
End If
If Photo.BorderStyle = 1 Then
gSn.Fields("LogoB") = 1
Else
gSn.Fields("LogoB") = 0
End If

gSn.Fields("IM1Top") = Image1.Top
gSn.Fields("IM1Left") = Image1.Left
gSn.Fields("IM1W") = Image1.Width
gSn.Fields("IM1H") = Image1.Height
If Image1.Visible = True Then
gSn.Fields("IM1Show") = 1
Else
gSn.Fields("IM1Show") = 0
End If
If Image1.BorderStyle = 1 Then
gSn.Fields("IM1B") = 1
Else
gSn.Fields("IM1B") = 0
End If

gSn.Fields("IM2Top") = Image2.Top
gSn.Fields("IM2Left") = Image2.Left
gSn.Fields("IM2W") = Image2.Width
gSn.Fields("IM2H") = Image2.Height
If Image2.Visible = True Then
gSn.Fields("IM2Show") = 1
Else
gSn.Fields("IM2Show") = 0
End If
If Image2.BorderStyle = 1 Then
gSn.Fields("IM2B") = 1
Else
gSn.Fields("IM2B") = 0
End If


If Option4(0).Value = True Then
gSn.Fields("BarSig") = "1"
End If
If Option4(1).Value = True Then
gSn.Fields("BarSig") = "2"
End If
If Option4(2).Value = True Then
gSn.Fields("BarSig") = "3"
End If

If Label3(0).Visible = True Then
gSn.Fields("Label1") = 1
Else
gSn.Fields("Label1") = 0
End If
If Label3(0).FontBold = True Then
gSn.Fields("FontBold1") = 1
Else
gSn.Fields("FontBold1") = 0
End If
If Label3(0).FontItalic = True Then
gSn.Fields("FontItalic1") = 1
Else
gSn.Fields("FontItalic1") = 0
End If
If Label3(0).FontUnderline = True Then
gSn.Fields("FontUnderline1") = 1
Else
gSn.Fields("FontUnderline1") = 0
End If
gSn.Fields("FontName1") = Label3(0).FontName
gSn.Fields("Color1") = Label3(0).ForeColor
gSn.Fields("FontSize1") = Label3(0).FontSize
gSn.Fields("Top1") = Label3(0).Top
gSn.Fields("Left1") = Label3(0).Left

If Label3(1).Visible = True Then
gSn.Fields("Label2") = 1
Else
gSn.Fields("Label2") = 0
End If
If Label3(1).FontBold = True Then
gSn.Fields("FontBold2") = 1
Else
gSn.Fields("FontBold2") = 0
End If
If Label3(1).FontItalic = True Then
gSn.Fields("FontItalic2") = 1
Else
gSn.Fields("FontItalic2") = 0
End If
If Label3(1).FontUnderline = True Then
gSn.Fields("FontUnderline2") = 1
Else
gSn.Fields("FontUnderline2") = 0
End If
gSn.Fields("FontName2") = Label3(1).FontName
gSn.Fields("Color2") = Label3(1).ForeColor
gSn.Fields("FontSize2") = Label3(1).FontSize
gSn.Fields("Top2") = Label3(1).Top
gSn.Fields("Left2") = Label3(1).Left

If Label3(2).Visible = True Then
gSn.Fields("Label3") = 1
Else
gSn.Fields("Label3") = 0
End If
If Label3(2).FontBold = True Then
gSn.Fields("FontBold3") = 1
Else
gSn.Fields("FontBold3") = 0
End If
If Label3(2).FontItalic = True Then
gSn.Fields("FontItalic3") = 1
Else
gSn.Fields("FontItalic3") = 0
End If
If Label3(2).FontUnderline = True Then
gSn.Fields("FontUnderline3") = 1
Else
gSn.Fields("FontUnderline3") = 0
End If
gSn.Fields("FontName3") = Label3(2).FontName
gSn.Fields("Color3") = Label3(2).ForeColor
gSn.Fields("FontSize3") = Label3(2).FontSize
gSn.Fields("Top3") = Label3(2).Top
gSn.Fields("Left3") = Label3(2).Left

If Label3(3).Visible = True Then
gSn.Fields("Label4") = 1
Else
gSn.Fields("Label4") = 0
End If
If Label3(3).FontBold = True Then
gSn.Fields("FontBold4") = 1
Else
gSn.Fields("FontBold4") = 0
End If
If Label3(3).FontItalic = True Then
gSn.Fields("FontItalic4") = 1
Else
gSn.Fields("FontItalic4") = 0
End If
If Label3(3).FontUnderline = True Then
gSn.Fields("FontUnderline4") = 1
Else
gSn.Fields("FontUnderline4") = 0
End If
gSn.Fields("FontName4") = Label3(3).FontName
gSn.Fields("Color4") = Label3(3).ForeColor
gSn.Fields("FontSize4") = Label3(3).FontSize
gSn.Fields("Top4") = Label3(3).Top
gSn.Fields("Left4") = Label3(3).Left

If Label3(4).Visible = True Then
gSn.Fields("Label5") = 1
Else
gSn.Fields("Label5") = 0
End If
If Label3(4).FontBold = True Then
gSn.Fields("FontBold5") = 1
Else
gSn.Fields("FontBold5") = 0
End If
If Label3(4).FontItalic = True Then
gSn.Fields("FontItalic5") = 1
Else
gSn.Fields("FontItalic5") = 0
End If
If Label3(4).FontUnderline = True Then
gSn.Fields("FontUnderline5") = 1
Else
gSn.Fields("FontUnderline5") = 0
End If
gSn.Fields("FontName5") = Label3(4).FontName
gSn.Fields("Color5") = Label3(4).ForeColor
gSn.Fields("FontSize5") = Label3(4).FontSize
gSn.Fields("Top5") = Label3(4).Top
gSn.Fields("Left5") = Label3(4).Left

If Label3(5).Visible = True Then
gSn.Fields("Label6") = 1
Else
gSn.Fields("Label6") = 0
End If
If Label3(5).FontBold = True Then
gSn.Fields("FontBold6") = 1
Else
gSn.Fields("FontBold6") = 0
End If
If Label3(5).FontItalic = True Then
gSn.Fields("FontItalic6") = 1
Else
gSn.Fields("FontItalic6") = 0
End If
If Label3(5).FontUnderline = True Then
gSn.Fields("FontUnderline6") = 1
Else
gSn.Fields("FontUnderline6") = 0
End If
gSn.Fields("FontName6") = Label3(5).FontName
gSn.Fields("Color6") = Label3(5).ForeColor
gSn.Fields("FontSize6") = Label3(5).FontSize
gSn.Fields("Top6") = Label3(5).Top
gSn.Fields("Left6") = Label3(5).Left

If Label3(6).Visible = True Then
gSn.Fields("Label7") = 1
Else
gSn.Fields("Label7") = 0
End If
If Label3(6).FontBold = True Then
gSn.Fields("FontBold7") = 1
Else
gSn.Fields("FontBold7") = 0
End If
If Label3(6).FontItalic = True Then
gSn.Fields("FontItalic7") = 1
Else
gSn.Fields("FontItalic7") = 0
End If
If Label3(6).FontUnderline = True Then
gSn.Fields("FontUnderline7") = 1
Else
gSn.Fields("FontUnderline7") = 0
End If
gSn.Fields("FontName7") = Label3(6).FontName
gSn.Fields("Color7") = Label3(6).ForeColor
gSn.Fields("FontSize7") = Label3(6).FontSize
gSn.Fields("Top7") = Label3(6).Top
gSn.Fields("Left7") = Label3(6).Left

If Label3(7).Visible = True Then
gSn.Fields("Label8") = 1
Else
gSn.Fields("Label8") = 0
End If
If Label3(7).FontBold = True Then
gSn.Fields("FontBold8") = 1
Else
gSn.Fields("FontBold8") = 0
End If
If Label3(7).FontItalic = True Then
gSn.Fields("FontItalic8") = 1
Else
gSn.Fields("FontItalic8") = 0
End If
If Label3(7).FontUnderline = True Then
gSn.Fields("FontUnderline8") = 1
Else
gSn.Fields("FontUnderline8") = 0
End If
gSn.Fields("FontName8") = Label3(7).FontName
gSn.Fields("Color8") = Label3(7).ForeColor
gSn.Fields("FontSize8") = Label3(7).FontSize
gSn.Fields("Top8") = Label3(7).Top
gSn.Fields("Left8") = Label3(7).Left

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Label6(0).Visible = True Then
gSn.Fields("Label9") = 1
Else
gSn.Fields("Label9") = 0
End If
If Label6(0).FontBold = True Then
gSn.Fields("FontBold9") = 1
Else
gSn.Fields("FontBold9") = 0
End If
If Label6(0).FontItalic = True Then
gSn.Fields("FontItalic9") = 1
Else
gSn.Fields("FontItalic9") = 0
End If
If Label6(0).FontUnderline = True Then
gSn.Fields("FontUnderline9") = 1
Else
gSn.Fields("FontUnderline9") = 0
End If
gSn.Fields("FontName9") = Label6(0).FontName
gSn.Fields("Color9") = Label6(0).ForeColor
gSn.Fields("FontSize9") = Label6(0).FontSize
gSn.Fields("Top9") = Label6(0).Top
gSn.Fields("Left9") = Label6(0).Left

If Label6(1).Visible = True Then
gSn.Fields("Label10") = 1
Else
gSn.Fields("Label10") = 0
End If
If Label6(1).FontBold = True Then
gSn.Fields("FontBold10") = 1
Else
gSn.Fields("FontBold10") = 0
End If
If Label6(1).FontItalic = True Then
gSn.Fields("FontItalic10") = 1
Else
gSn.Fields("FontItalic10") = 0
End If
If Label6(1).FontUnderline = True Then
gSn.Fields("FontUnderline10") = 1
Else
gSn.Fields("FontUnderline10") = 0
End If
gSn.Fields("FontName10") = Label6(1).FontName
gSn.Fields("Color10") = Label6(1).ForeColor
gSn.Fields("FontSize10") = Label6(1).FontSize
gSn.Fields("Top10") = Label6(1).Top
gSn.Fields("Left10") = Label6(1).Left

If Label6(2).Visible = True Then
gSn.Fields("Label11") = 1
Else
gSn.Fields("Label11") = 0
End If
If Label6(2).FontBold = True Then
gSn.Fields("FontBold11") = 1
Else
gSn.Fields("FontBold11") = 0
End If
If Label6(2).FontItalic = True Then
gSn.Fields("FontItalic11") = 1
Else
gSn.Fields("FontItalic11") = 0
End If
If Label6(2).FontUnderline = True Then
gSn.Fields("FontUnderline11") = 1
Else
gSn.Fields("FontUnderline11") = 0
End If
gSn.Fields("FontName11") = Label6(2).FontName
gSn.Fields("Color11") = Label6(2).ForeColor
gSn.Fields("FontSize11") = Label6(2).FontSize
gSn.Fields("Top11") = Label6(2).Top
gSn.Fields("Left11") = Label6(2).Left

If Label6(3).Visible = True Then
gSn.Fields("Label12") = 1
Else
gSn.Fields("Label12") = 0
End If
If Label6(3).FontBold = True Then
gSn.Fields("FontBold12") = 1
Else
gSn.Fields("FontBold12") = 0
End If
If Label6(3).FontItalic = True Then
gSn.Fields("FontItalic12") = 1
Else
gSn.Fields("FontItalic12") = 0
End If
If Label6(3).FontUnderline = True Then
gSn.Fields("FontUnderline12") = 1
Else
gSn.Fields("FontUnderline12") = 0
End If
gSn.Fields("FontName12") = Label6(3).FontName
gSn.Fields("Color12") = Label6(3).ForeColor
gSn.Fields("FontSize12") = Label6(3).FontSize
gSn.Fields("Top12") = Label6(3).Top
gSn.Fields("Left12") = Label6(3).Left

If Label6(4).Visible = True Then
gSn.Fields("Label13") = 1
Else
gSn.Fields("Label13") = 0
End If
If Label6(4).FontBold = True Then
gSn.Fields("FontBold13") = 1
Else
gSn.Fields("FontBold13") = 0
End If
If Label6(4).FontItalic = True Then
gSn.Fields("FontItalic13") = 1
Else
gSn.Fields("FontItalic13") = 0
End If
If Label6(4).FontUnderline = True Then
gSn.Fields("FontUnderline13") = 1
Else
gSn.Fields("FontUnderline13") = 0
End If
gSn.Fields("FontName13") = Label6(4).FontName
gSn.Fields("Color13") = Label6(4).ForeColor
gSn.Fields("FontSize13") = Label6(4).FontSize
gSn.Fields("Top13") = Label6(4).Top
gSn.Fields("Left13") = Label6(4).Left

If Label6(5).Visible = True Then
gSn.Fields("Label14") = 1
Else
gSn.Fields("Label14") = 0
End If
If Label6(5).FontBold = True Then
gSn.Fields("FontBold14") = 1
Else
gSn.Fields("FontBold14") = 0
End If
If Label6(5).FontItalic = True Then
gSn.Fields("FontItalic14") = 1
Else
gSn.Fields("FontItalic14") = 0
End If
If Label6(5).FontUnderline = True Then
gSn.Fields("FontUnderline14") = 1
Else
gSn.Fields("FontUnderline14") = 0
End If
gSn.Fields("FontName14") = Label6(5).FontName
gSn.Fields("Color14") = Label6(5).ForeColor
gSn.Fields("FontSize14") = Label6(5).FontSize
gSn.Fields("Top14") = Label6(5).Top
gSn.Fields("Left14") = Label6(5).Left

If Label6(6).Visible = True Then
gSn.Fields("Label15") = 1
Else
gSn.Fields("Label15") = 0
End If
If Label6(6).FontBold = True Then
gSn.Fields("FontBold15") = 1
Else
gSn.Fields("FontBold15") = 0
End If
If Label6(6).FontItalic = True Then
gSn.Fields("FontItalic15") = 1
Else
gSn.Fields("FontItalic15") = 0
End If
If Label6(6).FontUnderline = True Then
gSn.Fields("FontUnderline15") = 1
Else
gSn.Fields("FontUnderline15") = 0
End If
gSn.Fields("FontName15") = Label6(6).FontName
gSn.Fields("Color15") = Label6(6).ForeColor
gSn.Fields("FontSize15") = Label6(6).FontSize
gSn.Fields("Top15") = Label6(6).Top
gSn.Fields("Left15") = Label6(6).Left

If Label6(7).Visible = True Then
gSn.Fields("Label16") = 1
Else
gSn.Fields("Label16") = 0
End If
If Label6(7).FontBold = True Then
gSn.Fields("FontBold16") = 1
Else
gSn.Fields("FontBold16") = 0
End If
If Label6(7).FontItalic = True Then
gSn.Fields("FontItalic16") = 1
Else
gSn.Fields("FontItalic16") = 0
End If
If Label6(7).FontUnderline = True Then
gSn.Fields("FontUnderline16") = 1
Else
gSn.Fields("FontUnderline16") = 0
End If
gSn.Fields("FontName16") = Label6(7).FontName
gSn.Fields("Color16") = Label6(7).ForeColor
gSn.Fields("FontSize16") = Label6(7).FontSize
gSn.Fields("Top16") = Label6(7).Top
gSn.Fields("Left16") = Label6(7).Left

gSn.Update
gSn.Close

Call gettemps

End If

End Sub

Private Sub lvButtons_H15_Click()
Me.MousePointer = 11

    Dim x As Long
    Dim y As Long
    
'... Set the scalemode of the Pictures Boxes to Pixel
'... to help speed up the Process
CardF.ScaleMode = 3
CardB.ScaleMode = 3
PrintPic.Picture1(0).ScaleMode = 3
PrintPic.Picture1(1).ScaleMode = 3

    For y = 0 To CardF.ScaleHeight
        For x = 0 To CardF.ScaleWidth
            PrintPic.Picture1(0).PSet (x, y), Blend(CardF.Point(x, y), CardF.Point(x, y))
        Next
    Next
     
     For y = 0 To CardB.ScaleHeight
        For x = 0 To CardB.ScaleWidth
             PrintPic.Picture1(1).PSet (x, y), Blend(CardB.Point(x, y), CardB.Point(x, y))
        Next
    Next
    
    CardF.ScaleMode = 1
    CardB.ScaleMode = 1
    PrintPic.Picture1(0).ScaleMode = 1
    PrintPic.Picture1(1).ScaleMode = 1
    
    '... We are Done so reset the Mouse Pointer back to the Arrow
    Me.MousePointer = 0
    
    PrintPic.Image1.Picture = PrintPic.Picture1(0).Image
    PrintPic.Image2.Picture = PrintPic.Picture1(1).Image
    'PrintPic.Show 'vbModal
    PrintPic.PrintForm
    Unload PrintPic
    
End Sub

Private Sub lvButtons_H15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H15.hwnd, "Print Card", _
          vbCrLf & "Click To Send Your ID Card To The Printer.", 0, 100
End If

End Sub

Private Sub lvButtons_H16_Click()


Me.MousePointer = 11

    Dim x As Long
    Dim y As Long
    
'... Set the scalemode of the Pictures Boxes to Pixel
'... to help speed up the Process
CardF.ScaleMode = 3
CardB.ScaleMode = 3
PrintPic.Picture1(2).ScaleMode = 3

    For y = 0 To CardF.ScaleHeight
        For x = 0 To CardF.ScaleWidth
            PrintPic.Picture1(2).PSet (x, y), Blend(CardF.Point(x, y), CardF.Point(x, y))
        Next
    Next
    
   For y = 0 To CardB.ScaleHeight
     For x = 0 To CardB.ScaleWidth
            PrintPic.Picture1(2).PSet (CardB.ScaleWidth + (10) + x, y), Blend(CardB.Point(x, y), CardB.Point(x, y))
         Next
    Next
       
        
    CardF.ScaleMode = 1
    CardB.ScaleMode = 1
    PrintPic.Picture1(2).ScaleMode = 1
    
    '... We are Done so reset the Mouse Pointer back to the Arrow
    Me.MousePointer = 0

    On Error GoTo Select_Error
    With CommonDialog1
        .InitDir = Path1 & "IDCards"
        .DefaultExt = ".bmp"
        .CancelError = True
        .DialogTitle = "Save ID Card..."
        .Filter = "Pictures File(*.bmp)|*.bmp"
        .ShowSave
    End With
           
       SavePicture PrintPic.Picture1(2).Image, Path1 & CommonDialog1.FileName
      
       MsgBox "Images has been saved", vbExclamation
     
     Unload PrintPic
   
    
    Exit Sub
Select_Error:

End Sub

Private Sub lvButtons_H17_Click()

If Len(Text1(0).Text) = 0 Then
MsgBox "You Must Enter Some Text In Label1!", vbCritical, "Error"
Text1(2).SetFocus
Exit Sub
End If

If MsgBox("Would You Like To Delete " & Text1(0).Text & "?", vbQuestion Or vbYesNo) = vbYes Then
strSQL = "SELECT * FROM Users WHERE Label1=" & Quote & Text1(0).Text & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
If gSn.RecordCount > 0 Then
gSn.Delete

For x = 0 To 7
Text1(x).Text = ""
Text2(x).Text = ""
Next

Photo.Picture = PicDes.Image
Logo.Picture = PicDes.Image
Image1.Picture = PicDes.Image
Image2.Picture = PicDes.Image

Else
MsgBox "No Record In Database!", vbExclamation, "Error"
End If
gSn.Close
End If





End Sub

Private Sub lvButtons_H17_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H17.hwnd, "Delete", _
          vbCrLf & "Click To Delete The User From The Database.", 0, 100
End If
End Sub

Private Sub lvButtons_H18_Click()


On Error Resume Next
Dim A As ListItem

Form2.ListView1.ColumnHeaders.Clear
Form2.ListView1.ColumnHeaders.Add , , "No", 1000
Form2.ListView1.ColumnHeaders.Add , , "Last Name", 2500
Form2.ListView1.ColumnHeaders.Add , , "First Name", 2500
Form2.ListView1.ListItems.Clear

strSQL = "SELECT * FROM Users order by Label1"
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
If gSn.RecordCount > 0 Then
Do While Not gSn.EOF
   
   Set A = Form2.ListView1.ListItems.Add(, , gSn.Fields("Label1"))
   A.SubItems(1) = gSn.Fields("Label2")
   A.SubItems(2) = gSn.Fields("Label3")

gSn.MoveNext
DoEvents
Loop
gSn.Close
Form2.Show vbModal
End If


End Sub

Private Sub lvButtons_H18_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H18.hwnd, "List Users", _
          vbCrLf & "Click On Me For A List Of Users In Your Database." _
          & vbCrLf & "Then Click On 'No' To Get The User's Info.", _
          0, 100
End If
End Sub

Private Sub lvButtons_H2_Click()


            Dim si As String
            Dim c As New cDIBSection
            Dim qual1

'On Error Resume Next

If Len(Text1(0).Text) = 0 Then
MsgBox "You Must Enter Some Text In Label1!", vbCritical, "Error"
Text1(0).SetFocus
Exit Sub
End If

strSQL = "SELECT * FROM Users WHERE Label1=" & Quote & Text1(0).Text & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
If gSn.RecordCount > 0 Then
gSn.Delete
End If
gSn.Close


strSQL = "SELECT * FROM Users"
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
gSn.AddNew



            si = App.Path & "\Photo.jpg"
            If FileVerify(si) Then Kill si
            c.CreateFromPicture Photo.Picture
            qual1 = 80
            If SaveJPG(c, si, qual1) Then
            ' OK!
            End If

            si = App.Path & "\Logo.jpg"
            If FileVerify(si) Then Kill si
            c.CreateFromPicture Logo.Picture
            qual1 = 80
            If SaveJPG(c, si, qual1) Then
            ' OK!
            End If

            si = App.Path & "\Image1.jpg"
            If FileVerify(si) Then Kill si
            c.CreateFromPicture Image1.Picture
            qual1 = 80
            If SaveJPG(c, si, qual1) Then
            ' OK!
            End If

            si = App.Path & "\Image2.jpg"
            If FileVerify(si) Then Kill si
            c.CreateFromPicture Image2.Picture
            qual1 = 80
 
            If SaveJPG(c, si, qual1) Then
            ' OK!
            End If

'SavePicture Photo.Picture, App.Path & "\Photo.bmp"
'SavePicture Logo.Picture, App.Path & "\Logo.bmp"
'SavePicture Image1.Picture, App.Path & "\Image1.bmp"
'SavePicture Image2.Picture, App.Path & "\Image2.bmp"


'Load Photo to Put in DB
           Open App.Path & "\Photo.jpg" For Binary As #1
           ReDim PhotoPrint(LOF(1))
           Get #1, , PhotoPrint()
           Close #1

gSn.Fields("Photo") = PhotoPrint()

'Load Logo to Put in DB
           Open App.Path & "\Logo.jpg" For Binary As #1
           ReDim PhotoPrint(LOF(1))
           Get #1, , PhotoPrint()
           Close #1

gSn.Fields("Logo") = PhotoPrint()

'Load IM1 to Put in DB
           Open App.Path & "\Image1.jpg" For Binary As #1
           ReDim PhotoPrint(LOF(1))
           Get #1, , PhotoPrint()
           Close #1

gSn.Fields("IM1") = PhotoPrint()

'Load IM2 to Put in DB
           Open App.Path & "\Image2.jpg" For Binary As #1
           ReDim PhotoPrint(LOF(1))
           Get #1, , PhotoPrint()
           Close #1

gSn.Fields("IM2") = PhotoPrint()

gSn.Fields("Label1") = Trim(Text1(0).Text)
gSn.Fields("Label2") = Trim(Text1(1).Text)
gSn.Fields("Label3") = Trim(Text1(2).Text)
gSn.Fields("Label4") = Trim(Text1(3).Text)
gSn.Fields("Label5") = Trim(Text1(4).Text)
gSn.Fields("Label6") = Trim(Text1(5).Text)
gSn.Fields("Label7") = Trim(Text1(6).Text)
gSn.Fields("Label8") = Trim(Text1(7).Text)
gSn.Fields("Label9") = Trim(Text2(0).Text)
gSn.Fields("Label10") = Trim(Text2(1).Text)
gSn.Fields("Label11") = Trim(Text2(2).Text)
gSn.Fields("Label12") = Trim(Text2(3).Text)
gSn.Fields("Label13") = Trim(Text2(4).Text)
gSn.Fields("Label14") = Trim(Text2(5).Text)
gSn.Fields("Label15") = Trim(Text2(6).Text)
gSn.Fields("Label16") = Trim(Text2(7).Text)
gSn.Update
gSn.Close

End Sub

Private Sub lvButtons_H2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H2.hwnd, "Save", _
          vbCrLf & "Click On Me To Save The User's Info In The Database.", 0, 100
End If

End Sub

Private Sub lvButtons_H3_Click()

Dim Pic As StdPicture

On Error Resume Next

If Len(Text1(0).Text) = 0 Then
MsgBox "You Must Enter Some Text In Label1!", vbCritical, "Error"
Text1(0).SetFocus
Exit Sub
End If

strSQL = "SELECT * FROM Users WHERE Label1=" & Quote & Text1(0).Text & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
If gSn.RecordCount > 0 Then

PhotoPrint() = gSn.Fields("Photo")
        Set Pic = PictureFromByteStream(PhotoPrint)
        Set Photo.Picture = Pic

PhotoPrint() = gSn.Fields("Logo")
        Set Pic = PictureFromByteStream(PhotoPrint)
        Set Logo.Picture = Pic

PhotoPrint() = gSn.Fields("IM1")
        Set Pic = PictureFromByteStream(PhotoPrint)
        Set Image1.Picture = Pic
        
 PhotoPrint() = gSn.Fields("IM2")
        Set Pic = PictureFromByteStream(PhotoPrint)
        Set Image2.Picture = Pic

Text1(0).Text = gSn.Fields("Label1")
Text1(1).Text = gSn.Fields("Label2")
Text1(2).Text = gSn.Fields("Label3")
Text1(3).Text = gSn.Fields("Label4")
Text1(4).Text = gSn.Fields("Label5")
Text1(5).Text = gSn.Fields("Label6")
Text1(6).Text = gSn.Fields("Label7")
Text1(7).Text = gSn.Fields("Label8")
Text2(0).Text = gSn.Fields("Label9")
Text2(1).Text = gSn.Fields("Label10")
Text2(2).Text = gSn.Fields("Label11")
Text2(3).Text = gSn.Fields("Label12")
Text2(4).Text = gSn.Fields("Label13")
Text2(5).Text = gSn.Fields("Label14")
Text2(6).Text = gSn.Fields("Label15")
Text2(7).Text = gSn.Fields("Label16")

Else
MsgBox "No Record In Database!", vbExclamation, "Error"
End If

gSn.Close

End Sub

Private Sub lvButtons_H3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H3.hwnd, "Load", _
          vbCrLf & "Enter Text In Label1," _
          & vbCrLf & "Then Click To Get The User's Info From The Database.", _
          0, 100
End If
End Sub

Private Sub lvButtons_H4_Click()

Call RoddersMotion.Vision_VFWFormatDialog

End Sub

Private Sub lvButtons_H6_Click(Index As Integer)

'... Sub to choose the Text Size, Color and Style
'... for the Footer

CommonDialog1.CancelError = True
On Error GoTo Err
'... Set The Text Size, Color and Style of the Text
'... to the Font Dialog Control
CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
CommonDialog1.FontName = Label3(Index).FontName
CommonDialog1.FontBold = Label3(Index).FontBold
CommonDialog1.Color = Label3(Index).ForeColor
CommonDialog1.FontSize = Label3(Index).FontSize
CommonDialog1.FontItalic = Label3(Index).FontItalic
CommonDialog1.FontUnderline = Label3(Index).FontUnderline
'... Show the Font Dialog
CommonDialog1.ShowFont

'... Set the User Chossen Font Style, Size and Color
'... to the Fotter
Label3(Index).Caption = Trim(Label3(Index).Caption)
Label3(Index).Alignment = 2

If CommonDialog1.FontBold = True Then
  Label3(Index).FontBold = True
  Else
    Label3(Index).FontBold = False
        End If

If CommonDialog1.FontItalic = True Then
  Label3(Index).FontItalic = True
  Else
    Label3(Index).FontItalic = False
        End If
        
If CommonDialog1.FontUnderline = True Then
  Label3(Index).FontUnderline = True
  Else
    Label3(Index).FontUnderline = False
        End If
Label3(Index).FontName = CommonDialog1.FontName
Label3(Index).ForeColor = CommonDialog1.Color
Label3(Index).FontSize = CommonDialog1.FontSize
'... We want to let the Label to auto Size
'... to the New Font Size
Label3(Index).AutoSize = True
'... Now we want to set it to the Width
'... ot the Picture Boxes
'Label3(Index).AutoSize = False
Label3(Index).Height = Label3(Index).Height + 85

Label3(Index).Top = Main.pic1(0).Height - Label3(Index).Height
Label3(Index).Left = 0
Label3(Index).Width = Main.pic1(0).Width
Label3(Index).Height = Label3(Index).Height + 85
'Option2.Value = False
Exit Sub

Err:



'On Error GoTo Err
'CommonDialog1.ShowColor
'CommonDialog1.CancelError = True
'On Error GoTo Err
   
'   Label3(Index).ForeColor = CommonDialog1.Color
   
'Exit Sub
'Err:
End Sub



Private Sub lvButtons_H6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H6(Index).hwnd, "Font", _
          vbCrLf & "Click On Me To Edit Your Font." _
          & vbCrLf & "Font, Size, Style, And Effects.", _
          0, 100
End If
End Sub

Private Sub lvButtons_H7_Click(Index As Integer)
'... Sub to choose the Text Size, Color and Style
'... for the Footer

CommonDialog1.CancelError = True
On Error GoTo Err
'... Set The Text Size, Color and Style of the Text
'... to the Font Dialog Control
CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
CommonDialog1.FontName = Label3(Index).FontName
CommonDialog1.FontBold = Label3(Index).FontBold
CommonDialog1.Color = Label3(Index).ForeColor
CommonDialog1.FontSize = Label3(Index).FontSize
CommonDialog1.FontItalic = Label3(Index).FontItalic
CommonDialog1.FontUnderline = Label3(Index).FontUnderline
'... Show the Font Dialog
CommonDialog1.ShowFont

'... Set the User Chossen Font Style, Size and Color
'... to the Fotter
Label6(Index).Caption = Trim(Label6(Index).Caption)
Label6(Index).Alignment = 2

If CommonDialog1.FontBold = True Then
  Label6(Index).FontBold = True
  Else
    Label6(Index).FontBold = False
        End If

If CommonDialog1.FontItalic = True Then
  Label6(Index).FontItalic = True
  Else
    Label6(Index).FontItalic = False
        End If
        
If CommonDialog1.FontUnderline = True Then
  Label6(Index).FontUnderline = True
  Else
    Label6(Index).FontUnderline = False
        End If
Label6(Index).FontName = CommonDialog1.FontName
Label6(Index).ForeColor = CommonDialog1.Color
Label6(Index).FontSize = CommonDialog1.FontSize
'... We want to let the Label to auto Size
'... to the New Font Size
Label6(Index).AutoSize = True
'... Now we want to set it to the Width
'... ot the Picture Boxes
'Label6(Index).AutoSize = False
Label6(Index).Height = Label6(Index).Height + 85

Label6(Index).Top = Main.pic1(0).Height - Label6(Index).Height
Label6(Index).Left = 0
Label6(Index).Width = Main.pic1(0).Width
Label6(Index).Height = Label6(Index).Height + 85
'Option2.Value = False
Exit Sub

Err:
End Sub

Private Sub lvButtons_H7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Not MMove Or c.Alive = False Then
          MMove = True
          ApplyDemoValues
          c.ShowToolTip lvButtons_H7(Index).hwnd, "Font", _
          vbCrLf & "Click On Me To Edit Your Font." _
          & vbCrLf & "Font, Size, Style, And Effects.", _
          0, 100
End If
End Sub

Private Sub lvButtons_H8_Click()

CommonDialog1.CancelError = True
On Error GoTo Err
CommonDialog1.InitDir = Path1 & "Graphics\"
CommonDialog1.Filter = "Pictures (*.jpg;*.gif;*.bmp;*.ico)|*.jpg;*.gif;*.bmp;*.ico"
CommonDialog1.ShowOpen

Image1.Height = 1000
Image1.Width = 1000
        
        Image1.Picture = LoadPicture(CommonDialog1.FileName)


Exit Sub
Err:
End Sub

Private Sub lvButtons_H9_Click()

Image1.Height = 700
Image1.Width = 700

Image1.Picture = Picture2.Picture

End Sub

Private Sub Option4_Click(Index As Integer)

If Index = 0 Then
MsgBox ("The Barcode Is Using The Text From Label3."), vbInformation, "Barcode"
'Barcode
Label5.Visible = False
Line1.Visible = False
Image3.Visible = True
cl.Code128 BARCODE, 3, Text1(2).Text, False
Image3.Left = (CardB.Width - Image3.Width) / 2    ' Center form horizontally.
Image3.Visible = True
Image3.Picture = BARCODE.Image
End If

If Index = 1 Then
Label5.Visible = True
Line1.Visible = True
Image3.Visible = False
End If

If Index = 2 Then
Label5.Visible = False
Line1.Visible = False
Image3.Visible = False
End If

End Sub

Private Sub Photo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

distX = x - Photo.Left
distY = y - Photo.Top
Clicked = True

End Sub

Private Sub Photo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Clicked Then
           Photo.Left = x - distX
           Photo.Top = y - distY
           Photo.Refresh
End If

End Sub

Private Sub Photo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Clicked = False

End Sub

Private Sub Form_Load()

On Error Resume Next
'... Set Form to center of Screen
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

Clipboard.Clear

Call CreateDll
Call LoadDatabase

PhotoPrint = ""
    
px = Screen.TwipsPerPixelX
py = Screen.TwipsPerPixelY
    
    xOffset = -1
    yOffset = -1

Set c = New ExToolTip             '//-- Creates a new instance of theclass.

Call CenLogo
Call gettemps

HScroll3.Value = 100
HScroll2.Value = 100

Call RoddersMotion.Vision_VFWstart(Picture1)

Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload PrintPic

Call RoddersMotion.Vision_VFWstop

End Sub


Private Sub HScroll2_Change()
'... Resize the Width of Image

HScroll2.Min = 300
HScroll2.Max = 4700

HScroll2.SmallChange = (HScroll2.Max / 20) + 1
HScroll2.LargeChange = (HScroll2.Max / 5) + 1

Logo.Width = HScroll2.Value
Logo.Width = Logo.Width

'Call CenLogo

 
End Sub

Private Sub HScroll3_Change()
'... Resize the Height of Image

HScroll3.Min = 300
HScroll3.Max = 3000

HScroll3.SmallChange = (HScroll3.Max / 20) + 1
HScroll3.LargeChange = (HScroll3.Max / 5) + 1

Logo.Height = HScroll3.Value
Logo.Height = Logo.Height

'Call CenLogo

End Sub

Private Sub lvButtons_H5_Click()

CommonDialog1.CancelError = True
On Error GoTo Err
CommonDialog1.InitDir = Path1 & "Graphics\"
CommonDialog1.Filter = "Pictures (*.jpg;*.gif;*.bmp;*.ico)|*.jpg;*.gif;*.bmp;*.ico"
CommonDialog1.ShowOpen

Logo.Height = 700
Logo.Width = 700
Call CenLogo
 
        Logo.Picture = LoadPicture(CommonDialog1.FileName)

'Call CenLogo

Exit Sub
Err:
End Sub

Private Sub PicSrc_DblClick()
Call RoddersMotion.Vision_VFWFormatDialog
End Sub

Private Sub PicSrc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   
   If x > xOffset And x < xOffset + 1355 Then
        If y > yOffset And y < yOffset + 1625 Then
            CurX = x
            CurY = y
            Clicked = True
        End If
    End If
 
End Sub

Private Sub PicSrc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Clicked Then
        xOffset = xOffset + (x - CurX)
        yOffset = yOffset + (y - CurY)
        CurX = x
        CurY = y
        PicSrc.Refresh
    End If

End Sub

Private Sub PicSrc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Clicked = False
     PicDes.Cls
     i = BitBlt(PicDes.hDC, -2, -2, Int(Abs(yOffset + 1625) / px), Int(Abs(xOffset + 1355) / py), PicSrc.hDC, Int(xOffset / px), Int(yOffset / py), SRCCOPY)
     SavePicture PicDes.Image, App.Path & "\Photo.bmp"
     
     Photo.Picture = PicDes.Image
    
End Sub

Private Sub PicSrc_Paint()
    PicSrc.Line (xOffset, yOffset)-(xOffset + 1355, yOffset + 1625), RGB(0, 0, 255), B
End Sub

Private Sub Text1_Change(Index As Integer)

Label3(Index).Caption = Text1(Index).Text

If Index = 2 Then
cl.Code128 BARCODE, 3, Text1(Index).Text, False
Image3.Left = (CardB.Width - Image3.Width) / 2    ' Center form horizontally.
Image3.Picture = BARCODE.Image
End If

End Sub

Private Sub Text2_Change(Index As Integer)
Label6(Index).Caption = Text2(Index).Text

End Sub

Private Sub Timer1_Timer()
  DoEvents
  Call RoddersMotion.Vision_Motion(Picture1, PicSrc, picResult)

End Sub
Sub CenLogo()

Logo.Top = CardF.Height / 2 - Logo.Height / 2
Logo.Left = CardF.Width / 2 - Logo.Width / 2

End Sub
Private Function Blend(ByVal Color1 As Long, ByVal Color2 As Long) As Long
    Dim r As Long, G As Long, B As Long
    B = ((((Color2 \ &H10000) And &HFF) * 50) + _
    (((Color1 \ &H10000) And &HFF) * 50)) \ 100
    G = ((((Color2 \ &H100) And &HFF) * 50) + _
    (((Color1 \ &H100) And &HFF) * 50)) \ 100
    r = (((Color2 And &HFF) * 50) + ((Color1 And &HFF) * 50)) \ 100
    Blend = RGB(r, G, B)
End Function

Sub gettemps()

'On Error GoTo hErr

Combo1.Clear

strSQL = "SELECT * FROM Templates"
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)
If gSn.RecordCount > 0 Then
Do While Not gSn.EOF
Combo1.AddItem gSn.Fields("Name")
gSn.MoveNext
DoEvents
Loop
gSn.Close

Combo1.ListIndex = 0

End If


End Sub
Private Sub ApplyDemoValues() '//--Default Values used in Demo
        
        c.DelayTime = 26
        c.KillTime = 513
        c.TextColor = &H80000012
'        Set c.Picture = Form1.ToolTipPic
        c.BackStyle = 0
        c.Font.Name = "Arial"
        c.Shadow = True
        c.ToolTipStyle = True

End Sub
