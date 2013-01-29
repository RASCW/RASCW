VERSION 5.00
Begin VB.Form frmOpenDF1 
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmOpenDF1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_DF1 
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
   Begin VB.TextBox txtDF1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.ComboBox cboDF1 
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenDF1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDF1_Click()
    File_DF1.Pattern = "*.df1"
End Sub

Private Sub cmdBrowse_Click()
    cboDF1.AddItem "DF1 files (*.df1)"
    cboDF1.ListIndex = 0
    File_DF1.Visible = True
End Sub

Private Sub cmdCancel_Click()
  Unload frmOpenDF1
End Sub

Private Sub cmdOpen_Click()
    If txtDF1.Text <> "" Then
        txt_file_df1 = txtDF1.Text
        Check = txt_file_df1
        
        Unload frmOpenDF1
    Else
        Beep
    End If
End Sub

Private Sub File_DF1_Click()
    txtDF1.Text = File_DF1.Filename
    File_DF1.Visible = False
End Sub

