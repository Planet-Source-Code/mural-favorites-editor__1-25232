VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Favorite"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   275
      Left            =   2280
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Fields"
      Height          =   275
      Left            =   840
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   275
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   1720
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Title"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Description"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "URL Address"
      TabPicture(2)   =   "Form2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   5
         Top             =   480
         Width           =   2535
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   0
            TabIndex        =   6
            Text            =   "http://"
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   3
         Top             =   480
         Width           =   2535
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2895
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   360
            TabIndex        =   2
            Top             =   120
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

If Text1.Text = "" Then
MsgBox "Not all the form fields are completly filled out!", vbExclamation, "Error"
Else
If Text2.Text = "" Then
MsgBox "Not all the form fields are completly filled out!", vbExclamation, "Error"
Else
If Text3.Text = "" Then
MsgBox "Not all the form fields are completly filled out!", vbExclamation, "Error"
Else

Open App.Path + "\" + Text1.Text + ".ant" For Output As #1
    Print #1, Text1.Text
    Print #1, Text2.Text
    Print #1, Text3.Text
Close
Call Activatee
Text1.Text = ""
Text2.Text = ""
Text3.Text = "http://"
Text3.SelStart = Len(Text3.Text)
Form2.Hide

End If
    End If
        End If
            
        
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = "http://"
Text3.SelStart = Len(Text3.Text)
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = "http://"
Text3.SelStart = Len(Text3.Text)
Form2.Hide
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
Text1.SetFocus
End If
If SSTab1.Tab = 1 Then
Text2.SetFocus
End If
If SSTab1.Tab = 2 Then
Text3.SetFocus
Text3.SelStart = Len(Text3.Text)
End If
End Sub

Private Sub Timer1_Timer()
Text1.SetFocus
Timer1.Enabled = False
End Sub


