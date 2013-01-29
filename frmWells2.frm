VERSION 5.00
Begin VB.Form frmWells2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display of Casc Results"
   ClientHeight    =   6075
   ClientLeft      =   150
   ClientTop       =   585
   ClientWidth     =   10095
   Icon            =   "frmWells2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   612
      Left            =   5400
      TabIndex        =   98
      Top             =   5250
      Width           =   1815
      Begin VB.CheckBox ChkLabels 
         Caption         =   "Edit Well Labels"
         Height          =   252
         Left            =   240
         TabIndex        =   99
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ListBox List_add_events 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1590
      Left            =   3960
      MultiSelect     =   2  'Extended
      TabIndex        =   97
      Top             =   3480
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.Frame Frame3 
      Caption         =   " "
      Height          =   612
      Left            =   120
      TabIndex        =   93
      Top             =   5250
      Width           =   5055
      Begin VB.OptionButton OptNormal 
         Caption         =   "O.D. + Error Bars"
         Height          =   252
         Left            =   1800
         TabIndex        =   100
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptObs 
         Caption         =   "Observed Depths"
         Height          =   252
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   1572
      End
      Begin VB.OptionButton OptProb 
         Caption         =   "Probable Depths"
         Height          =   252
         Left            =   3360
         TabIndex        =   94
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8880
      TabIndex        =   49
      Top             =   5460
      Width           =   852
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   7650
      TabIndex        =   48
      Top             =   5460
      Width           =   852
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   19
      Left            =   2640
      TabIndex        =   41
      Top             =   4590
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   18
      Left            =   1920
      TabIndex        =   40
      Top             =   4590
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   17
      Left            =   1200
      TabIndex        =   39
      Top             =   4590
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   16
      Left            =   480
      TabIndex        =   38
      Top             =   4590
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   15
      Left            =   3360
      TabIndex        =   37
      Top             =   4230
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   14
      Left            =   2640
      TabIndex        =   36
      Top             =   4230
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   13
      Left            =   1920
      TabIndex        =   35
      Top             =   4230
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   12
      Left            =   1200
      TabIndex        =   34
      Top             =   4230
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   11
      Left            =   480
      TabIndex        =   33
      Top             =   4200
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   10
      Left            =   3360
      TabIndex        =   32
      Top             =   3870
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   9
      Left            =   2640
      TabIndex        =   31
      Top             =   3870
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   8
      Left            =   1920
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   7
      Left            =   1200
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   6
      Left            =   480
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   5
      Left            =   3360
      TabIndex        =   27
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   4
      Left            =   2640
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   3
      Left            =   1920
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   2
      Left            =   1200
      TabIndex        =   24
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   1
      Left            =   480
      TabIndex        =   23
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtevent 
      Height          =   288
      Index           =   20
      Left            =   3360
      TabIndex        =   22
      Top             =   4590
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   20
      Left            =   3360
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   16
      Left            =   480
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   17
      Left            =   1200
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   18
      Left            =   1920
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   19
      Left            =   2640
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   14
      Left            =   2640
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   13
      Left            =   1920
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   12
      Left            =   1200
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   11
      Left            =   480
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   10
      Left            =   3360
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   9
      Left            =   2640
      TabIndex        =   11
      Top             =   1404
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   8
      Left            =   1920
      TabIndex        =   10
      Top             =   1404
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   7
      Left            =   1200
      TabIndex        =   9
      Top             =   1404
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   1404
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   5
      Left            =   3360
      TabIndex        =   7
      Top             =   1044
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   4
      Left            =   2640
      TabIndex        =   6
      Top             =   1044
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   15
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   3
      Left            =   1920
      TabIndex        =   4
      Top             =   1044
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   1044
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox well 
      Height          =   288
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1044
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Total_event 
      Height          =   288
      Left            =   2640
      TabIndex        =   1
      Top             =   3150
      Width           =   492
   End
   Begin VB.TextBox total_well 
      Height          =   288
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   492
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Wells"
      Height          =   2535
      Left            =   120
      TabIndex        =   42
      Top             =   240
      Width           =   9852
      Begin VB.CommandButton cmdLabeladd 
         Caption         =   "Accept"
         Height          =   255
         Left            =   7830
         TabIndex        =   102
         Top             =   510
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TxtWellLabels 
         Height          =   1488
         Left            =   7770
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   101
         Text            =   "frmWells2.frx":0442
         Top             =   840
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.ListBox List_add_well 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1590
         Left            =   3840
         MultiSelect     =   2  'Extended
         TabIndex        =   96
         Top             =   840
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.CommandButton cmdWellDic 
         Caption         =   "List of Wells"
         Height          =   252
         Left            =   4920
         TabIndex        =   46
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label well_label 
         Caption         =   " 20"
         Height          =   252
         Index           =   19
         Left            =   3000
         TabIndex        =   71
         Top             =   1920
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 19"
         Height          =   252
         Index           =   18
         Left            =   2280
         TabIndex        =   70
         Top             =   1920
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 18"
         Height          =   252
         Index           =   17
         Left            =   1560
         TabIndex        =   69
         Top             =   1920
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 17"
         Height          =   252
         Index           =   16
         Left            =   840
         TabIndex        =   68
         Top             =   1920
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label well_label 
         Caption         =   " 16"
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   67
         Top             =   1920
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 15"
         Height          =   252
         Index           =   14
         Left            =   3000
         TabIndex        =   66
         Top             =   1560
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 14"
         Height          =   252
         Index           =   13
         Left            =   2280
         TabIndex        =   65
         Top             =   1560
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 13"
         Height          =   252
         Index           =   12
         Left            =   1560
         TabIndex        =   64
         Top             =   1560
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 12"
         Height          =   252
         Index           =   11
         Left            =   840
         TabIndex        =   63
         Top             =   1560
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 11"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   62
         Top             =   1560
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   " 10"
         Height          =   252
         Index           =   9
         Left            =   3000
         TabIndex        =   61
         Top             =   1200
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "   9"
         Height          =   252
         Index           =   8
         Left            =   2280
         TabIndex        =   60
         Top             =   1200
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "  8"
         Height          =   252
         Index           =   7
         Left            =   1560
         TabIndex        =   57
         Top             =   1200
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "  7"
         Height          =   252
         Index           =   6
         Left            =   840
         TabIndex        =   56
         Top             =   1200
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "  6"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "  5 "
         Height          =   252
         Index           =   4
         Left            =   3000
         TabIndex        =   54
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "   4"
         Height          =   252
         Index           =   3
         Left            =   2280
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "  3"
         Height          =   252
         Index           =   2
         Left            =   1560
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label well_label 
         Caption         =   "  2"
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   51
         Top             =   840
         Visible         =   0   'False
         Width           =   372
         WordWrap        =   -1  'True
      End
      Begin VB.Label well_label 
         Caption         =   "  1"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label Label1 
         Caption         =   "Number of  Wells Displayed "
         Height          =   252
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   2172
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Events"
      Height          =   2445
      Left            =   120
      TabIndex        =   43
      Top             =   2790
      Width           =   9852
      Begin VB.CommandButton cmdEventDic 
         Caption         =   "List of Events"
         Height          =   252
         Left            =   4950
         TabIndex        =   47
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label event_label 
         Caption         =   " 15"
         Height          =   252
         Index           =   14
         Left            =   3000
         TabIndex        =   92
         Top             =   1440
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 20"
         Height          =   252
         Index           =   19
         Left            =   3000
         TabIndex        =   91
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 16"
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   90
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 17"
         Height          =   252
         Index           =   16
         Left            =   840
         TabIndex        =   89
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 18"
         Height          =   252
         Index           =   17
         Left            =   1560
         TabIndex        =   88
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 19"
         Height          =   252
         Index           =   18
         Left            =   2280
         TabIndex        =   87
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 14"
         Height          =   252
         Index           =   13
         Left            =   2280
         TabIndex        =   86
         Top             =   1440
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   " 13"
         Height          =   252
         Index           =   12
         Left            =   1560
         TabIndex        =   85
         Top             =   1440
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 12"
         Height          =   252
         Index           =   11
         Left            =   840
         TabIndex        =   84
         Top             =   1440
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   " 11"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   83
         Top             =   1440
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.Label event_label 
         Caption         =   " 10"
         Height          =   252
         Index           =   9
         Left            =   3000
         TabIndex        =   82
         Top             =   1080
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  9"
         Height          =   252
         Index           =   8
         Left            =   2280
         TabIndex        =   81
         Top             =   1080
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  8"
         Height          =   252
         Index           =   7
         Left            =   1560
         TabIndex        =   80
         Top             =   1080
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  7"
         Height          =   252
         Index           =   6
         Left            =   840
         TabIndex        =   79
         Top             =   1080
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  6"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   78
         Top             =   1080
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  5"
         Height          =   252
         Index           =   4
         Left            =   3000
         TabIndex        =   77
         Top             =   720
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  4"
         Height          =   252
         Index           =   3
         Left            =   2280
         TabIndex        =   76
         Top             =   720
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  3"
         Height          =   252
         Index           =   2
         Left            =   1560
         TabIndex        =   74
         Top             =   720
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "  2"
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   73
         Top             =   720
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label event_label 
         Caption         =   "   1"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   720
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Events Correlated"
         Height          =   252
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   2292
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   252
      Left            =   4440
      TabIndex        =   75
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   252
      Left            =   5160
      TabIndex        =   59
      Top             =   1200
      Width           =   252
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   252
      Left            =   4440
      TabIndex        =   58
      Top             =   1200
      Width           =   252
   End
End
Attribute VB_Name = "frmWells2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim act_Num_well As Integer
Dim total_Num_event As Integer
Dim act_Num_event As Integer
Dim add_events_num() As Integer
Dim add_events_name() As String * 50
Dim add_well_num() As Integer
Dim add_well_name() As String * 50
Dim well_label_selected() As Integer
Dim editlabel As Boolean



Private Sub event_Change(Index As Integer)

End Sub

Private Sub ChkLabels_Click()
    If ChkLabels.Value = 1 Then
      TxtWellLabels.Visible = True
      'Lblwelllabel.Visible = True
      cmdLabeladd.Visible = True
      
    'editlabel = True
    
     ' Cmdwelllabel.Visible = True
     'cmdEventLabel.Visible = True
     Else
      TxtWellLabels.Visible = False
      cmdLabeladd.Visible = False
    'editlabel = False
      'List_well_labels.Visible = False
      'List_event_labels.Visible = False
      'Cmdwelllabel.Visible = False
      'cmdEventLabel.Visible = False
    End If
End Sub




Private Sub chkTable_Click()

End Sub

Private Sub cmdApply_Click()
Dim I As Integer

If total_well.Text <> "" And Total_event.Text <> "" Then
    
        max_well_num2 = Int(Trim(total_well.Text))
        max_event_num2 = Int(Trim(Total_event.Text))
        
        For I = 1 To max_well_num2
            If well(I).Text = "" Then
               Beep
               Exit Sub
            Else
            
                 well_numbers2(I) = well(I).Text
            End If
        Next I
            
            
        For I = 1 To max_event_num2
            If txtevent(I).Text = "" Then
                    Beep
                    Exit Sub
             Else
                 event_numbers2(I) = txtevent(I).Text
            End If
        Next I
        If OptProb.Value = False And OptObs.Value = False And OptNormal.Value = False Then
              Beep
              MsgBox "Select an option from Observed depths, Prob. depths or Normality test"
              Exit Sub
        End If
       prob_opt2 = OptProb.Value
       obs_opt2 = OptObs.Value
       norm_opt2 = OptNormal.Value
       testing_flag2 = 1
    
       Unload frmWells2
Else
        Beep
End If
'testing for the option prob and obs
'add well labels
   
End Sub

Private Sub cmdCancel_Click()
    testing_flag2 = 0
    Unload frmWells2
End Sub

Private Sub cmdEventDic_Click()
Dim I As Integer
Dim k As Integer
k = 0
If cmdEventDic.Caption = "List of Events" Then
    List_add_events.Visible = True
    cmdEventDic.Caption = "Add Events"
 Else
    For I = 0 To List_add_events.ListCount - 1
        If List_add_events.Selected(I) Then
           k = k + 1
        End If
     Next I
     If k > 20 Then
            Beep
            MsgBox "Maximum number of selected events is 20"
            List_add_events.Visible = False
           cmdEventDic.Caption = "List of Events"
           Exit Sub
      Else
            k = 0
            For I = 0 To List_add_events.ListCount - 1
               If List_add_events.Selected(I) Then
               k = k + 1
               txtevent(k).Text = Trim(Mid(List_add_events.List(I), 1, 5))     'add_events_num(I + 1)
               'add event names into flattening list box
               event_names2(k) = Trim(Mid(List_add_events.List(I), 11, Len(List_add_events.List(I)) - 10))   'add_events_name(I + 1)
                 
               txtevent(k).Visible = True
               event_label(k - 1).Visible = True
               End If
            Next I
                Total_event.Text = Str(k)
                For I = k To 19
                    txtevent(I + 1).Visible = False
                    event_label(I).Visible = False
                Next I
               
               List_add_events.Visible = False
               cmdEventDic.Caption = "List of Events"
            End If
  
 End If
 
 
End Sub



Private Sub cmdWellDic_Click()
Dim I, k As Integer
Dim well_labels As String
k = 0
If cmdWellDic.Caption = "List of Wells" Then
    List_add_well.Visible = True
    cmdWellDic.Caption = "Add Wells"
Else
    For I = 0 To List_add_well.ListCount - 1
    
        well_label_selected(I + 1) = 0  'test whether I + 1 is selected or not
        If List_add_well.Selected(I) Then
        k = k + 1
        If k > 20 Then
           MsgBox "You have selected more than 20 wells."
           k = 20
           GoTo proc1:
        End If
        well(k).Text = add_well_num(I + 1)
        well(k).Visible = True
        well_label(k - 1).Visible = True
        well_labels = well_labels + Space(2 * (4 - Len(Str(add_well_num(I + 1))))) _
                    + Str(add_well_num(I + 1)) + Space(5) + Trim(add_well_name(I + 1)) + _
                    Chr(13) + Chr(10)
        well_label_selected(I + 1) = 1  'test I + 1 label is selected
         End If
     Next I
proc1:
        TxtWellLabels.Text = well_labels    ' add labels into edit label text box
      '  TxtWellLabels.Visible = True
       
        List_add_well.Visible = False
         total_well.Text = Str(k)
        For I = k To 19
        well(I + 1).Visible = False
        well_label(I).Visible = False
        Next I
     cmdWellDic.Caption = "List of Wells"
       
     ' update the events list so that only event occur on wells are displayed
     List_add_events.Clear
     List_events
     For I = 0 To 19
             txtevent(I + 1).Visible = False
             event_label(I).Visible = False
     Next I

 
 End If


End Sub



Private Sub cmdLabeladd_Click()
Dim k As Integer
Dim label As String
Dim I As Integer
Dim J As Integer
Dim Temp As String
Dim txtlength As Integer
  txtlength = Len(Trim(TxtWellLabels.Text))
 I = 1
 If txtlength > 0 Then
  Temp = Mid(Trim(TxtWellLabels.Text), I, 1)
  txtlength = txtlength - 1
  
  For k = 1 To Int(Trim(total_well.Text))
     If txtlength > 0 Then
     Temp = Mid(Trim(TxtWellLabels.Text), I, 1)
     txtlength = txtlength - 1
     label = ""
       Do While (Temp <> Chr(13)) And (txtlength > 0)
       label = label + Temp
       I = I + 1
       Temp = Mid(Trim(TxtWellLabels.Text), I, 1)
       txtlength = txtlength - 1
       
       Loop
     ' MsgBox label
       well_label_used2(k) = label
         I = I + 1
    End If
  Next k
  
  '
 '
 '
 '
 ' For k = 1 To List_add_well.ListCount
 '   If (well_label_selected(k) = 1) And (txtlength > 0) Then
 '  temp = Mid(Trim(TxtWellLabels.Text), I, 1)
 '  txtlength = txtlength - 1
 '    label = ""
 '      Do While (temp <> Chr(13)) And (txtlength > 0)
 '      label = label + temp
 '      I = I + 1
 '      temp = Mid(Trim(TxtWellLabels.Text), I, 1)
 ''      txtlength = txtlength - 1
 '
 '      Loop
 '    ' MsgBox label
 '   well_label_used2(k) = label
 '   I = I + 1
 '   End If
 ' Next k
cmdLabeladd.Visible = False
TxtWellLabels.Visible = False
 End If
End Sub


Private Sub form_Click()
Dim I As Integer
 
 If total_well.Text <> "" Then
    If total_well.Text > 20 Then
    Beep
    MsgBox "Maximum number of selected wells is 20"
    Else
    
    For I = 1 To 20
        If I <= total_well.Text Then
        well(I).Visible = True
        well_label(I - 1).Visible = True
        Else
        well(I).Visible = False
        well_label(I - 1).Visible = False
        End If
        Next I
    End If
 End If
 If Total_event.Text <> "" Then
    If Total_event.Text > 20 Then
    Beep
    MsgBox "Maximum number of selected events is 20"
    Else
    For I = 1 To 20
        If I <= Total_event.Text Then
        txtevent(I).Visible = True
        event_label(I - 1).Visible = True
        Else
          txtevent(I).Visible = False
        event_label(I - 1).Visible = False
        End If
    Next I
   End If
  End If


End Sub

Private Sub Form_Load()
        Dim I As Integer
        If modifyflag2 = 1 Then
         total_well.Text = Str(max_well_num2)
         Total_event.Text = Str(max_event_num2)
           For I = 1 To max_event_num2
              event_label(I - 1).Visible = True
              txtevent(I).Text = Str(event_numbers2(I))
              txtevent(I).Visible = True
           Next I
            For I = 1 To max_well_num2
            well_label(I - 1).Visible = True
            well(I).Text = Str(well_numbers2(orders2(I)))
            well(I).Visible = True
            Next I
               If prob_opt2 = 1 Then
                OptProb.Value = True
                Else
                OptObs.Value = False
                End If
        Else
          For I = 1 To 20
            well(I).Visible = False
            well_label(I - 1).Visible = False
            txtevent(I).Visible = False
            event_label(I - 1).Visible = False
        Next I
        End If
        read_well_event
End Sub



Private Sub Frame1_Click()
Dim I As Integer
 If total_well.Text <> "" Then
    If total_well.Text > 20 Then
       Beep
       MsgBox "Maximum number of wells is 20"
       Else
       
        For I = 1 To 20
            If I <= total_well.Text Then
            well(I).Visible = True
            well_label(I - 1).Visible = True
            Else
            well(I).Visible = False
            well_label(I - 1).Visible = False
            End If
            
        Next I
    End If
    
 End If
 If Total_event.Text <> "" Then
        If Total_event.Text > 20 Then
          Beep
          MsgBox "Maximum number of events is 20"
          Else
          
    For I = 1 To 20
        If I <= Total_event.Text Then
        txtevent(I).Visible = True
        event_label(I - 1).Visible = True
        Else
          txtevent(I).Visible = False
        event_label(I - 1).Visible = False
        End If
        
    Next I
    End If
 End If
End Sub

Private Sub Frame2_Click()
Dim I As Integer
 If total_well.Text <> "" Then
    If total_well.Text > 20 Then
       Beep
       MsgBox "Maximum number of wells is 20"
    Else
    
    For I = 1 To 20
        If I <= total_well.Text Then
        well(I).Visible = True
        well_label(I - 1).Visible = True
        Else
        well(I).Visible = False
        well_label(I - 1).Visible = False
        End If
 
    Next I
   End If
 End If
 If Total_event.Text <> "" Then
    If Total_event.Text > 20 Then
        Beep
        MsgBox "Maximum number of events is 20"
        Else
    For I = 1 To 20
        If I <= Total_event.Text Then
        txtevent(I).Visible = True
        event_label(I - 1).Visible = True
        Else
          txtevent(I).Visible = False
        event_label(I - 1).Visible = False
        End If
   Next I
    End If
 End If
End Sub


Public Sub read_well_event()    'read events and event numbers
Dim I, J, k As Integer          ' in the list_add_event
Dim Temp As String
Dim test_num As Integer
Dim FileNum As Integer

FileNum = FreeFile
If txt_file_ran1 <> "" Then
    Open CurDir + "\" + txt_file_ran1 For Input As FileNum
Else
    MsgBox "Input file does not exist, please try again"
    Exit Sub
End If

Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
'    MsgBox temp
act_Num_well = Val(Mid(Temp, 24, 4))           'read the number of well
'    MsgBox Str(act_Num_well)
Input #FileNum, Temp

Input #FileNum, Temp    ' number of events
act_Num_event = Val(Mid(Temp, 25, 4))
total_Num_event = act_Num_event  'here save the total event number

'MsgBox Str(act_Num_event)

ReDim add_well_num(1 To act_Num_well)
ReDim add_well_name(1 To act_Num_well)
ReDim well_label_used2(1 To act_Num_well)
ReDim well_label_selected(1 To act_Num_well)
ReDim add_events_num(1 To act_Num_event)
ReDim add_events_name(1 To act_Num_event)


For I = 1 To act_Num_well
    add_well_num(I) = -1
    add_well_name(I) = "**"
    well_label_used2(I) = "Well No." + Str(I)
Next I

For J = 1 To act_Num_event
 add_events_num(J) = -1
 add_events_name(J) = "**"
Next J

'read entire wells and events
  
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp



For I = 1 To act_Num_well       'every wells
    Line Input #FileNum, Temp        'read in the number and name of well
           add_well_num(I) = Mid(Temp, 1, 3)
           add_well_name(I) = Mid(Temp, 5, 50)
    ' well_label_used2(I) = add_well_name(I)
    ' skip 4 lines
           Input #FileNum, Temp
           Input #FileNum, Temp
           Input #FileNum, Temp
           Input #FileNum, Temp
    
     For J = 1 To act_Num_event
          Line Input #FileNum, Temp
          If Temp <> " * * *" Then
                ' Old statement:
                        '  Input #FileNum, Temp
                        '  If Mid(Temp,1, 1) <> "*" Then
              add_events_num(J) = Mid(Temp, 42, 4)   ' add_events_number -- event number
              add_events_name(J) = Mid(Temp, 46, 50)   'event name
          End If
      Next J
      For J = 1 To 3
       If Not EOF(FileNum) Then
       Input #FileNum, Temp
       End If
      Next J
    Next I
   
Close FileNum

' put the well number and name into list_add_well
   For I = 1 To act_Num_well
      List_add_well.AddItem Space(2 * (4 - Len(Str(add_well_num(I))))) + Str(add_well_num(I)) + Space(5) + add_well_name(I)
   Next I
      
 'put events into list_add_event
    For J = 1 To act_Num_event
        If add_events_num(J) <> -1 Then
             List_add_events.AddItem Space(5 - Len(Str(add_events_num(J)))) + Str(add_events_num(J)) + Space(5) + add_events_name(J)
        End If
    Next J
End Sub





Private Sub Option1_Click()

End Sub

Public Sub List_events()    ' List the events occur on the wells selected
Dim I, J, k As Integer          ' in the list_add_event
Dim Temp As String
Dim test_num As Integer
Dim FileNum As Integer
Dim found As Integer

FileNum = FreeFile
If txt_file_ran1 <> "" Then
    Open CurDir + "\" + txt_file_ran1 For Input As FileNum
Else
    MsgBox "Input file does not exist, please try again"
    Exit Sub
End If

    
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
'    MsgBox temp
act_Num_well = Val(Mid(Temp, 24, 4)) 'read the number of well
' MsgBox Str(act_Num_well)
Input #FileNum, Temp

Input #FileNum, Temp    ' number of events
act_Num_event = Val(Mid(Temp, 25, 4))
total_Num_event = act_Num_event  'here save the total event number

'MsgBox Str(act_Num_event)

ReDim add_well_num(1 To act_Num_well)
ReDim add_well_name(1 To act_Num_well)
ReDim add_events_num(1 To act_Num_event)
ReDim add_events_name(1 To act_Num_event)

For J = 1 To act_Num_event
 add_events_num(J) = -1
 add_events_name(J) = "**"
Next J

'read entire wells and events
  
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp



For I = 1 To act_Num_well       'read through every wells
    Input #FileNum, Temp        'read in the number and name of well
           add_well_num(I) = Mid(Temp, 1, 2)
           add_well_name(I) = Mid(Temp, 3, 50)
    'skip 4 lines
           Input #FileNum, Temp
           Input #FileNum, Temp
           Input #FileNum, Temp
           Input #FileNum, Temp
    found = 0   'flag variable
        For k = 1 To total_well.Text
          If add_well_num(I) = well(k).Text Then
            found = 1
          End If
        Next k
        
       For J = 1 To act_Num_event
          Line Input #FileNum, Temp
          
           If found = 1 Then   'well number is selected
                      
     '       If Mid(temp, 1, 1) <> "***" Then
                If Temp <> " * * *" Then
              add_events_num(J) = Mid(Temp, 42, 4) ' add_events_number -- event number
              add_events_name(J) = Mid(Temp, 46, 50) 'event name
             End If
           End If
        Next J
      For J = 1 To 3
       If Not EOF(FileNum) Then
           Input #FileNum, Temp
          End If
      Next J
 Next I
   
Close FileNum

 'put events into list_add_event
    For J = 1 To act_Num_event
        If add_events_num(J) <> -1 Then
          List_add_events.AddItem Space(5 - Len(Str(add_events_num(J)))) + Str(add_events_num(J)) + Space(5) + add_events_name(J)
       End If
    Next J
End Sub


