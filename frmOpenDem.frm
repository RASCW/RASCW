VERSION 5.00
Begin VB.Form frmOpenDem 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1815
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmOpenDem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_Den 
      Height          =   870
      Left            =   120
      TabIndex        =   5
      Top             =   390
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Top             =   1380
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpenDen 
      Caption         =   "Open"
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   750
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Txt_Den 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox cboDem 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1380
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "File type :"
      Height          =   345
      Left            =   150
      TabIndex        =   6
      Top             =   1110
      Width           =   2985
   End
End
Attribute VB_Name = "frmOpenDem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDem_Click()
    Select Case cboDem.ListIndex
    Case 0
        File_Den.Pattern = "*.dem"
    Case 1
        File_Den.Pattern = "*.dn2"
    End Select
End Sub

Private Sub cmdBrowse_Click()
    If File_Den.Visible Then
       File_Den.Visible = False
    Else
       File_Den.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    CancelKeyPress = 1
    Unload frmOpenDem
End Sub

Private Sub cmdOpenDen_Click()
    If Txt_Den.Text <> "" Then
        txt_file_dem = Txt_Den.Text
        OpenFileKey = 1
        Unload frmOpenDem
    Else
        OpenFileKey = 0
        Beep
    End If
End Sub

Private Sub File_Den_Click()
    Txt_Den.Text = File_Den.Filename
    File_Den.Visible = False
End Sub

Private Sub Form_Load()
    cboDem.AddItem "DEM files (*.dem)"
    cboDem.ListIndex = 0
    cboDem.AddItem "DN2 files (*.dn2)"

End Sub
