VERSION 5.00
Begin VB.Form frmOpenDI1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4500
   Icon            =   "frmOpenDI1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_DI1 
      Height          =   870
      Left            =   240
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2910
      TabIndex        =   3
      Top             =   1350
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   345
      Left            =   2910
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   150
      Width           =   1365
   End
   Begin VB.TextBox txtDI1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox cboDI1 
      Height          =   315
      ItemData        =   "frmOpenDI1.frx":0442
      Left            =   240
      List            =   "frmOpenDI1.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1350
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "File type :"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1110
      Width           =   2355
   End
End
Attribute VB_Name = "frmOpenDI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDI1_Click()
    Select Case cboDI1.ListIndex
    Case 0
        File_DI1.Pattern = "*.di1"
    Case 1
        File_DI1.Pattern = "*.di2"
    End Select
End Sub

Private Sub cmdBrowse_Click()
'    cboDI1.AddItem "DI1 files (*.DI1)"
'    cboDI1.ListIndex = 0
    If File_DI1.Visible Then
       File_DI1.Visible = False
    Else
       File_DI1.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmOpenDI1
End Sub

Private Sub cmdOpen_Click()
    If txtDI1.Text <> "" Then
        txt_file_di1 = txtDI1.Text
        Check = txt_file_di1
        Unload frmOpenDI1
    Else
        Beep
    End If
End Sub

Private Sub File_DI1_Click()
    txtDI1.Text = File_DI1.Filename
    File_DI1.Visible = False
End Sub

Private Sub Form_Load()
    cboDI1.AddItem "DI1 files (*.di1)"
    cboDI1.ListIndex = 0
    cboDI1.AddItem "DI2 files (*.di2)"
    'cboDI1.ListIndex = 1
End Sub
