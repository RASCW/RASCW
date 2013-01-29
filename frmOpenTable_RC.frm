VERSION 5.00
Begin VB.Form frmOpenTable_RC 
   Caption         =   "Enter File Name For Rasc Input"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "frmOpenTable_RC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_ran1 
      Height          =   675
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1212
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Apply"
      Height          =   252
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1212
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Select Name"
      Height          =   252
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1212
   End
   Begin VB.TextBox txtRan1 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.ComboBox cboRan1 
      Height          =   288
      Left            =   720
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenTable_RC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRan1_Click()
File_ran1.Pattern = "*.dat"
End Sub

Private Sub cmdBrowse_Click()
cboRan1.AddItem "Ran files (*.dat)"
cboRan1.ListIndex = 0
File_ran1.Visible = True
End Sub

Private Sub cmdCancel_Click()
RascInputFile = ""
Unload frmOpenTable_RC
End Sub

Private Sub cmdOpen_Click()
If txtRan1.Text <> "" Then
RascInputFile = Trim(txtRan1.Text) + ".dat"

Unload frmOpenTable_RC
Else
Beep
End If
End Sub

Private Sub File_ran1_Click()

Dim filelen As Integer
Dim tempName As String
tempName = Trim(File_ran1.Filename)
filelen = Len(tempName)
If filelen > 4 Then
txtRan1.Text = Mid(tempName, 1, filelen - 4)
'txtRan1.Text = File_ran1.Filename
End If
File_ran1.Visible = False
End Sub

Private Sub Form_Load()
RascInputFile = ""
End Sub
