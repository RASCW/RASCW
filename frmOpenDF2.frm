VERSION 5.00
Begin VB.Form frmOpenDF2 
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmOpenDF2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_DF2 
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   252
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   732
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   252
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtDF2 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.ComboBox cboDF2 
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenDF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDF2_Click()
    File_DF2.Pattern = "*.DF2"
End Sub

Private Sub cmdBrowse_Click()
    cboDF2.AddItem "DF2 files (*.DF2)"
    cboDF2.ListIndex = 0
    File_DF2.Visible = True
End Sub

Private Sub cmdCancel_Click()
  Unload frmOpenDF2
End Sub

Private Sub cmdOpen_Click()
    If txtDF2.Text <> "" Then
        txt_file_df2 = txtDF2.Text
        Check = txt_file_df2
        
        Unload frmOpenDF2
    Else
        Beep
    End If
End Sub

Private Sub File_DF2_Click()
    txtDF2.Text = File_DF2.Filename
    File_DF2.Visible = False
End Sub

