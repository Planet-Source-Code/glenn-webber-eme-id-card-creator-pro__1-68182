VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub ListView1_Click()

Dim pic As StdPicture

On Error Resume Next

strSQL = "SELECT * FROM Users WHERE Label1=" & Quote & ListView1.SelectedItem & Quote
Set gSn = db.OpenRecordset(strSQL) ', dbOpenSnapshot)

PhotoPrint() = gSn.Fields("Photo")
        Set pic = PictureFromByteStream(PhotoPrint)
        Set Form1.Photo.Picture = pic

PhotoPrint() = gSn.Fields("Logo")
        Set pic = PictureFromByteStream(PhotoPrint)
        Set Form1.Logo.Picture = pic

PhotoPrint() = gSn.Fields("IM1")
        Set pic = PictureFromByteStream(PhotoPrint)
        Set Form1.Image1.Picture = pic
        
 PhotoPrint() = gSn.Fields("IM2")
        Set pic = PictureFromByteStream(PhotoPrint)
        Set Form1.Image2.Picture = pic

Form1.Text1(0).Text = gSn.Fields("Label1")
Form1.Text1(1).Text = gSn.Fields("Label2")
Form1.Text1(2).Text = gSn.Fields("Label3")
Form1.Text1(3).Text = gSn.Fields("Label4")
Form1.Text1(4).Text = gSn.Fields("Label5")
Form1.Text1(5).Text = gSn.Fields("Label6")
Form1.Text1(6).Text = gSn.Fields("Label7")
Form1.Text1(7).Text = gSn.Fields("Label8")
Form1.Text2(0).Text = gSn.Fields("Label9")
Form1.Text2(1).Text = gSn.Fields("Label10")
Form1.Text2(2).Text = gSn.Fields("Label11")
Form1.Text2(3).Text = gSn.Fields("Label12")
Form1.Text2(4).Text = gSn.Fields("Label13")
Form1.Text2(5).Text = gSn.Fields("Label14")
Form1.Text2(6).Text = gSn.Fields("Label15")
Form1.Text2(7).Text = gSn.Fields("Label16")

gSn.Close

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next

If ListView1.SortOrder = lvwAscending Then
   ListView1.SortOrder = lvwDescending
Else
   ListView1.SortOrder = lvwAscending
End If

ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Sorted = True

End Sub
