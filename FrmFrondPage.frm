VERSION 5.00
Begin VB.Form FrmFrontPage 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   8040
   ClientLeft      =   -8160
   ClientTop       =   1.96800e5
   ClientWidth     =   11145
   Icon            =   "FrmFrondPage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H8000000D&
      Height          =   8175
      Left            =   -90
      Picture         =   "FrmFrondPage.frx":0442
      ScaleHeight     =   8115
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   -30
      Width           =   11295
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Continue"
         Default         =   -1  'True
         Height          =   375
         Left            =   9390
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " RASC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   37.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Index           =   0
         Left            =   2220
         TabIndex        =   8
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Biostratigraphic Zonation and Correlation of Fossil Events"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "by"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   5550
         TabIndex        =   6
         Top             =   3270
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "F. P. Agterberg,  F. M. Gradstein, Q. Cheng  and  G. Liu "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   2400
         TabIndex        =   5
         Top             =   3750
         Width           =   6975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright  1998, 2007   F. P. Agterberg  and  F. M. Gradstein "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   6600
         Width           =   6495
      End
      Begin VB.Image Image1 
         Height          =   1650
         Left            =   4320
         Picture         =   "FrmFrondPage.frx":368B2
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " CASC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   37.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Index           =   1
         Left            =   6870
         TabIndex        =   3
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " and"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   38.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Index           =   2
         Left            =   4950
         TabIndex        =   2
         Top             =   1500
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmFrontPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()
End Sub

Private Sub Command2_Click()
        Unload FrmFrontPage
End Sub

Private Sub Command3_Click()
        FrmFrontPage.Hide
        FrmFrontPage1.Show 1
End Sub

Private Sub Form_Activate()
'        MDIFrmCascRasc.ActiveTimer.Enabled = False
'        MDIFrmCascRasc.DeActiveTimer.Enabled = False
End Sub

Private Sub Form_Resize()
        'Dim Heightfactor As Double
        'Dim Widthfactor As Double
        'Heightfactor = FrmFrontPage.Height - 100
        '   If Heightfactor < 0 Then
        '   Heightfactor = 0
        '   End If
        'Widthfactor = FrmFrontPage.Width - 100
        '   If Widthfactor < 0 Then
        '   Widthfactor = 0
        '   End If
        'Frame1.Width = Widthfactor
        'Frame1.Height = Heightfactor
End Sub



Private Sub Form_Unload(Cancel As Integer)
    MdiFrmFirstTimeLoad = 0
    If MDIFrmCascRasc.mnuToolbar.Checked Then
        'ToolBar.Show
        ShowWindow ToolBar.hwnd, SW_SHOW
        MDIFrmCascRasc.ActiveTimer.Enabled = True
        MDIFrmCascRasc.DeActiveTimer.Enabled = True
    End If
    ShowWindow MDIFrmCascRasc.hwnd, SW_SHOW

End Sub
