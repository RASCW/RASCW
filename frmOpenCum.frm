VERSION 5.00
Begin VB.Form frmOpenCum 
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "frmOpenCum.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1185
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_Cum 
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   252
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   732
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   252
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txt_Cum 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
   Begin VB.ComboBox cboCum 
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenCum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCum_Click()
File_Cum.Pattern = "*.Cum"
End Sub

Private Sub cmdBrowse_Click()
cboCum.AddItem "Cum files (*.Cum)"
cboCum.ListIndex = 0
File_Cum.Visible = True
End Sub

Private Sub cmdCancel_Click()
Unload frmOpenCum
End Sub

Private Sub cmdOpen_Click()
If txt_Cum.Text <> "" Then
   txt_file_cum = txt_Cum.Text
   Unload frmOpenCum
Else
   Beep
End If
End Sub

Private Sub File_Cum_Click()
txt_Cum.Text = File_Cum.Filename
File_Cum.Visible = False
End Sub

Private Sub Form_Load()
 'File_Cum.Path = App.Path
End Sub
