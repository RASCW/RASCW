VERSION 5.00
Begin VB.Form frmOpenRank 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3210
   Icon            =   "frmOpenRank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_Den 
      Height          =   870
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   852
   End
   Begin VB.CommandButton cmdOpenDen 
      Caption         =   "Open"
      Height          =   252
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   852
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   252
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin VB.TextBox Txt_Den 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2052
   End
   Begin VB.ComboBox cboDen 
      Height          =   288
      Left            =   2040
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDen_Click()
    File_Den.Pattern = "*.Opt"
End Sub

Private Sub cmdBrowse_Click()
    cboDen.AddItem "Opt files (*.Opt)"
    cboDen.ListIndex = 0
    File_Den.Visible = True

End Sub

Private Sub cmdCancel_Click()
    CancelKeyPress = 1
    Unload frmOpenRank
End Sub

Private Sub cmdOpenDen_Click()
    If Txt_Den.Text <> "" Then
        txt_file_den = Txt_Den.Text
        OpenFileKey = 1
        Unload frmOpenRank
    Else
        OpenFileKey = 0
        Beep
    End If
End Sub

Private Sub File_Den_Click()
    Txt_Den.Text = File_Den.Filename
    File_Den.Visible = False
End Sub

