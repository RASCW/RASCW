VERSION 5.00
Begin VB.Form frmOpenOutputFiles 
   Caption         =   "Open Output Files"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Txt_Output 
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2052
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   252
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   852
   End
   Begin VB.CommandButton cmdOpenOutput 
      Caption         =   "Open"
      Height          =   252
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   852
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   852
   End
   Begin VB.FileListBox File_Output 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2052
   End
End
Attribute VB_Name = "frmOpenOutputFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboOutput_Click()
  File_Output.Pattern = "*.*"
End Sub

Private Sub cmdBrowse_Click()
  cboOutput.AddItem "Output files (*.*)"
  cboOutput.ListIndex = 0
  File_Output.Visible = True

End Sub

Private Sub cmdCancel_Click()
   Unload frmOpenOutputFiles
End Sub

Private Sub cmdOpenOutput_Click()
Dim X
Dim StrOutput As String

If Txt_Output.Text <> "" Then
  txt_file_Output = Txt_Output.Text
  StrOutput = "notepad.exe " + txt_file_Output
  X = Shell(StrOutput, vbNormalFocus)
  Unload frmOpenOutputFiles
Else
  Beep
End If
End Sub

Private Sub File_Output_Click()
Txt_Output.Text = File_Output.Filename
File_Output.Visible = False
End Sub

