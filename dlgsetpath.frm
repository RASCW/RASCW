VERSION 5.00
Begin VB.Form dlgsetpath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Project Space"
   ClientHeight    =   3750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3855
   Icon            =   "dlgsetpath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   288
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   2772
   End
   Begin VB.DirListBox Dir1 
      Height          =   2016
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3372
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   960
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Driver:"
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   2820
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Folder"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "dlgsetpath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
'        MDIFrmCascRasc.ActiveTimer.Enabled = False
'        MDIFrmCascRasc.DeActiveTimer.Enabled = False
End Sub

Private Sub Form_Load()
    Dir1.Path = CurDir
    Text1.Text = Dir1.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
'        MDIFrmCascRasc.ActiveTimer.Enabled = True
'        MDIFrmCascRasc.DeActiveTimer.Enabled = True
End Sub

Private Sub OKButton_Click()
    ChDir Dir1.Path
    ChDrive (Drive1.Drive)
    DialogInitPath = CurDir  'For load text files dialog
    ChartSaveInitPath = CurDir 'For Chart Save As dialog
    ChartOpenInitPath = CurDir 'For ChartShow Open dialog
    CurrentDir = Dir1.Path
    CurrentDrive = Drive1.Drive
    Unload Me
End Sub
