VERSION 5.00
Begin VB.Form frmOpenMake 
   Caption         =   "Import from old Makedat"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmOpenMake.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboMake 
      Height          =   288
      Left            =   2040
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.FileListBox File_New_Make 
      Height          =   480
      Left            =   2880
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   252
      Left            =   4560
      TabIndex        =   8
      Top             =   240
      Width           =   732
   End
   Begin VB.TextBox txt_New_Make 
      Height          =   288
      Left            =   2880
      TabIndex        =   7
      Text            =   " "
      Top             =   240
      Width           =   1692
   End
   Begin VB.FileListBox File_ran1 
      Height          =   480
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   5400
      TabIndex        =   3
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "OK"
      Height          =   252
      Left            =   5400
      TabIndex        =   2
      Top             =   360
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
   Begin VB.TextBox txtRan1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
   Begin VB.ComboBox cboRan1 
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Note: Old Makedat *.lst and *.dic files Required"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "New Project Name"
      Height          =   252
      Left            =   2880
      TabIndex        =   11
      Top             =   0
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "Old MakeDat File Name"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   1812
   End
End
Attribute VB_Name = "frmOpenMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub cboMake_Click()
File_New_Make.Pattern = "*.mdb"
End Sub

Private Sub cboRan1_Click()
File_ran1.Pattern = "*.lst"
End Sub

Private Sub cmdBrowse_Click()
cboRan1.AddItem "LST files (*.lst)"
cboRan1.ListIndex = 0
File_ran1.Visible = True
End Sub

Private Sub cmdCancel_Click()
Old_Makedat_File = ""
New_Makedat_File = ""
Unload frmOpenMake
End Sub

Private Sub cmdOpen_Click()
If txtRan1.Text <> "" And txt_New_Make.Text <> "" Then
Old_Makedat_File = Trim(txtRan1.Text) + ".lst"
New_Makedat_File = Trim(txt_New_Make.Text) + ".mdb"
Unload frmOpenMake
Else
Beep
Exit Sub
End If
End Sub

Private Sub Command1_Click()
cboMake.AddItem "MDB files (*.mdb)"
cboMake.ListIndex = 0
File_New_Make.Visible = True
End Sub
 

Private Sub File_New_Make_Click()

Dim filelen As Integer
Dim tempName As String
tempName = Trim(File_New_Make.Filename)
filelen = Len(tempName)
If filelen > 4 Then
txt_New_Make.Text = Mid(tempName, 1, filelen - 4)
'txt_New_Make.Text = File_New_Make.Filename
End If
File_New_Make.Visible = False
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
Old_Makedat_File = ""
New_Makedat_File = ""
End Sub

