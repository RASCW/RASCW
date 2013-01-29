VERSION 5.00
Begin VB.Form FrmFrontPage1 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   8025
   ClientLeft      =   -8160
   ClientTop       =   1.96800e5
   ClientWidth     =   11100
   Icon            =   "FrmFrondPage1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmFrondPage1.frx":0442
   ScaleHeight     =   8100.141
   ScaleMode       =   0  'User
   ScaleWidth      =   10452.85
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   510
      Top             =   3720
   End
   Begin VB.PictureBox Picture1 
      Height          =   8115
      Left            =   -60
      Picture         =   "FrmFrondPage1.frx":2E86F
      ScaleHeight     =   537
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   743
      TabIndex        =   0
      Top             =   -30
      Width           =   11205
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Cancel          =   -1  'True
         Caption         =   "Continue"
         Default         =   -1  'True
         Height          =   375
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7290
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " RASC "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   37.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   975
         Index           =   0
         Left            =   1710
         TabIndex        =   7
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Version 20"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   4620
         TabIndex        =   6
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " RASC     for   Ranking, Scaling, and Variance Analysis    "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   1890
         TabIndex        =   5
         Top             =   4080
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " CASC     for    Correlation and Standard-error Calculation "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   5220
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " CASC "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   37.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   975
         Index           =   1
         Left            =   6630
         TabIndex        =   3
         Top             =   1770
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H0080FF80&
         Height          =   975
         Index           =   2
         Left            =   4590
         TabIndex        =   2
         Top             =   1800
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmFrontPage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptext1, temptext2, temptext3, temptext4 As String
Dim length_txt1, length_txt2, length_txt3, length_txt4 As Integer
Dim time_loop As Integer
Dim tempstr1, tempstr2, tempStr3, tempStr4 As Integer

Private Sub Text1_Change()
End Sub


Private Sub Command3_Click()
    Unload FrmFrontPage1
    MdiFrmFirstTimeLoad = 0
    If MDIFrmCascRasc.mnuToolbar.Checked Then
        ToolBar.Show
        MDIFrmCascRasc.ActiveTimer.Enabled = True
        MDIFrmCascRasc.DeActiveTimer.Enabled = True
    End If
    MDIFrmCascRasc.SetFocus
End Sub

Private Sub Form_Load()
tempstr1 = 0
tempstr2 = 0
tempStr3 = 0
tempStr4 = 0

'temptext1 = Label2.Caption
temptext2 = Label4.Caption
temptext3 = Label11.Caption
'temptext4 = Label12.Caption


'Label2.Caption = ""
Label4.Caption = ""
Label11.Caption = ""
'Label12.Caption = ""
'Label2.Visible = True
Label4.Visible = True
Label11.Visible = True
'Label12.Visible = True
time_loop = 0
End Sub



Private Sub Label12_Click()

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

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

Private Sub Timer1_Timer()
 If tempstr1 * tempstr2 * tempStr3 * tempStr4 <> 1 Then
   
   
   time_loop = time_loop + 1
 
 'If time_loop < Len(temptext1) + 1 Then
'    Label2.Caption = Mid(temptext1, 1, time_loop)
 'Else
 '   tempstr1 = 1
 'End If
 If time_loop < Len(temptext2) + 1 Then
    Label4.Caption = Mid(temptext2, 1, time_loop)
 Else
    tempstr2 = 1
 End If
 If time_loop < Len(temptext3) + 1 Then
    Label11.Caption = Mid(temptext3, 1, time_loop)
 Else
    tempStr3 = 1
 End If
 'If time_loop < Len(temptext4) + 1 Then
 'Label12.Caption = Mid(temptext4, 1, time_loop)
 'Else
 'tempStr4 = 1
 'End If
 
 Else
    Exit Sub
 End If
End Sub
