VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Files"
   ClientHeight    =   5520
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3195
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox txtFile 
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   120
         Width           =   4815
      End
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Text            =   "*.*"
         ToolTipText     =   "Select File filter"
         Top             =   720
         Width           =   1815
      End
      Begin VB.FileListBox RasFile 
         Height          =   2040
         Left            =   600
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         ToolTipText     =   "Select Files"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "File name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "File Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.DirListBox RasDir 
         Height          =   990
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
      Begin VB.DriveListBox RASDrive 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
 Unload Me
End Sub

Private Sub cmbFilter_Change()
 RasFile.Pattern = cmbFilter.Text
 txtFile.Text = ""
 Me.Caption = RasFile.Filename
 RasFile.Refresh
End Sub

Private Sub cmbFilter_Click()
   txtFile.Text = ""
   Me.Caption = RasFile.Filename
   RasFile.Pattern = cmbFilter.Text
   RasFile.Refresh
End Sub

Private Sub Form_Load()
 txtFile.Text = ""
 Me.Caption = RasFile.Filename
 With cmbFilter
  .Clear
  .Text = "*.*"
  .AddItem "*.*"
  .AddItem "*.par"
  .AddItem "*.dic"
  .AddItem "*.dep"
  .AddItem "*.dat"
  .AddItem "*.inp"
  .AddItem "*.ca?"
  .AddItem "*.exe"
 End With
 
End Sub

Private Sub OKButton_Click()
 'MsgBox "Selected : " & RasFile.Filename
 Unload Me
End Sub

Private Sub RasDir_Change()
 RasFile.Path = RasDir.Path
 txtFile.Text = ""
 Me.Caption = RasFile.Filename
End Sub

Private Sub RASDrive_Change()
  RasDir.Path = RASDrive.Drive
  txtFile.Text = ""
  Me.Caption = RasFile.Filename
End Sub

Private Sub RasFile_DblClick()
   txtFile.Text = RasDir & RasFile.Filename
   Me.Caption = RasFile.Filename
   Unload Me
End Sub

Private Sub RasFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtFile.Text = RasDir & RasFile.Filename
  Me.Caption = RasFile.Filename
End Sub
