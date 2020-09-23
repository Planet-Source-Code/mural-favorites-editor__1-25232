VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Favorites Editor By: MuraL"
   ClientHeight    =   3255
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   1800
      Pattern         =   "*.ant*"
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Favorites Editor v.1.0   "
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "ghgfh"
            TextSave        =   "ghgfh"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0114
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0228
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "URL"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "Add New Favorite"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remove"
            Object.ToolTipText     =   "Remove Selected Favorite"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "F&ile"
      Begin VB.Menu new 
         Caption         =   "N&ew"
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu endit 
         Caption         =   "Exi&t"
      End
   End
   Begin VB.Menu menu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu openinbrowsr 
         Caption         =   "&Visit..."
      End
      Begin VB.Menu serperpeorpeorpegjvkf 
         Caption         =   "-"
      End
      Begin VB.Menu addit 
         Caption         =   "&Add"
      End
      Begin VB.Menu remove 
         Caption         =   "R&emove"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu toolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu statusbar 
         Caption         =   "Statusba&r"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Form2.Show vbModal
End Sub

Private Sub addit_Click()
Form2.Show vbModal
End Sub

Private Sub endit_Click()
Unload Form2
Unload Form1
End

End Sub

Private Sub Form_Load()
Call Activatee
End Sub

Private Sub ListView1_DblClick()
Dim Title, Des, URL As String
Form1.File1.Path = App.Path
Form1.File1.Refresh


Open App.Path + "\" + Form1.ListView1.SelectedItem + ".ant" For Input As #1
Input #1, Title
Input #1, Des
Input #1, URL
Form2.Text1.Text = Title
Form2.Text2.Text = Des
Form2.Text3.Text = URL
Form2.Show vbModal
Close #1

End Sub


Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu Me.menu
End If
End Sub


Private Sub new_Click()
Form2.Show vbModal
End Sub

Private Sub openinbrowsr_Click()
Dim Title, Des, URL As String
Form1.File1.Path = App.Path
Form1.File1.Refresh


Open App.Path + "\" + Form1.ListView1.SelectedItem + ".ant" For Input As #1
Input #1, Title
Input #1, Des
Input #1, URL

Shell ("start " + URL)



Close #1
End Sub


Private Sub remove_Click()
On Error Resume Next
Kill App.Path + "\" + ListView1.SelectedItem + ".ant"
ListView1.ListItems.remove ListView1.SelectedItem.Index
Call Activatee
End Sub

Private Sub statusbar_Click()
If statusbar.Checked = True Then
statusbar.Checked = False
StatusBar1.Visible = False
Form1.Height = Form1.Height - 255
Else
statusbar.Checked = True
StatusBar1.Visible = True
Form1.Height = Form1.Height + 255
End If
End Sub

Private Sub toolbar_Click()
If toolbar.Checked = True Then
toolbar.Checked = False
Toolbar1.Visible = False
ListView1.Top = 0
Form1.Height = Form1.Height - 420
Else
Toolbar1.Visible = True
ListView1.Top = 480
toolbar.Checked = True
Form1.Height = Form1.Height + 420
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "new"
Form2.Show vbModal
Case "remove"
On Error Resume Next
Kill App.Path + "\" + ListView1.SelectedItem + ".ant"
ListView1.ListItems.remove ListView1.SelectedItem.Index
Call Activatee
End Select

End Sub


