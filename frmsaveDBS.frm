VERSION 5.00
Begin VB.Form frmSaveDBS 
   Caption         =   "Create a Database"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   Icon            =   "frmsaveDBS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_DBS 
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   732
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   252
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   732
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   252
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   732
   End
   Begin VB.TextBox txtDBS 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
   Begin VB.ComboBox cboDBS 
      Height          =   288
      Left            =   0
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Give a name without ext."
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   1812
   End
End
Attribute VB_Name = "frmSaveDBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboDBS_Click()
File_DBS.Pattern = "*.mdb"
End Sub

Private Sub cmdBrowse_Click()
cboDBS.AddItem "MDB files (*.mdb)"
cboDBS.ListIndex = 0
File_DBS.Visible = True
End Sub

Private Sub cmdCancel_Click()
myDBName = ""
Unload frmSaveDBS
End Sub

 

Private Sub cmdSave_Click()
If txtDBS.Text <> "" Then
myDBName = Trim(txtDBS.Text) + ".mdb"
Unload frmSaveDBS
Else
Beep
End If
End Sub

Private Sub File_DBS_Click()
Dim ll As Integer
ll = Len(File_DBS.Filename)
txtDBS.Text = Mid(Trim(File_DBS.Filename), 1, ll - 4)
File_DBS.Visible = False
End Sub

