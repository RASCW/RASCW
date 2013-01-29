VERSION 5.00
Begin VB.Form frmOpenRan1 
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmOpenRan1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_ran1 
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
   Begin VB.TextBox txtRan1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.ComboBox cboRan1 
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenRan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRan1_Click()
    File_ran1.Pattern = "*." + file_type
End Sub

Private Sub cmdBrowse_Click()
    cboRan1.AddItem "Data files (*." + file_type + ")"
    cboRan1.ListIndex = 0
    File_ran1.Visible = True
End Sub

Private Sub cmdCancel_Click()
    CancelKeyPress = 1
    Unload frmOpenRan1
End Sub

Private Sub cmdOpen_Click()
    If txtRan1.Text <> "" Then
        txt_file_ran1 = txtRan1.Text
        Check = txt_file_ran1
        
        Unload frmOpenRan1
    Else
        Beep
    End If
End Sub

Private Sub File_ran1_Click()
    txtRan1.Text = File_ran1.Filename
    File_ran1.Visible = False
End Sub

Private Sub Form_Load()
    txt_file_ran1 = ""
End Sub
