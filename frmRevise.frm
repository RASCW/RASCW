VERSION 5.00
Begin VB.Form frmRevise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Parameters to run RASC"
   ClientHeight    =   5400
   ClientLeft      =   1200
   ClientTop       =   2595
   ClientWidth     =   10125
   Icon            =   "frmRevise.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox FileOut 
      Height          =   870
      Left            =   5880
      TabIndex        =   94
      Top             =   900
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame FrameUnit 
      Caption         =   "Unit to be used in CASC"
      Height          =   525
      Left            =   4920
      TabIndex        =   104
      Top             =   1230
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "feet"
         Height          =   285
         Left            =   2280
         TabIndex        =   105
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "meters"
         Height          =   255
         Left            =   1140
         TabIndex        =   106
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear UE/MH Selections"
      Height          =   315
      Left            =   5460
      TabIndex        =   107
      Top             =   1950
      Width           =   2025
   End
   Begin VB.CommandButton cmdMarkerEvents 
      BackColor       =   &H00FF80FF&
      Caption         =   "Select Marker Events"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   2370
      Width           =   1965
   End
   Begin VB.ListBox ListUniEvent 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2580
      ItemData        =   "frmRevise.frx":0442
      Left            =   5250
      List            =   "frmRevise.frx":0444
      MultiSelect     =   2  'Extended
      TabIndex        =   102
      Top             =   2730
      Width           =   4815
   End
   Begin VB.CommandButton cmdUniEvent 
      BackColor       =   &H00FF8080&
      Caption         =   "Select Unique Events"
      Height          =   315
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   2370
      Width           =   2025
   End
   Begin VB.TextBox txtM 
      Height          =   285
      Left            =   960
      TabIndex        =   97
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtUE 
      Height          =   285
      Left            =   960
      TabIndex        =   96
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboOut 
      Height          =   315
      Left            =   7410
      TabIndex        =   95
      Text            =   "Combo1"
      Top             =   210
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Browse"
      Height          =   285
      Left            =   7380
      TabIndex        =   93
      Top             =   630
      Width           =   720
   End
   Begin VB.TextBox txtOut 
      Height          =   288
      Left            =   5880
      TabIndex        =   92
      Top             =   630
      Width           =   1485
   End
   Begin VB.CommandButton cmdDictionary 
      Caption         =   "&Dictionary"
      Height          =   315
      Left            =   8400
      TabIndex        =   51
      Top             =   1950
      Width           =   1425
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   40
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   50
      Text            =   "0"
      Top             =   5016
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   39
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   49
      Text            =   "0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   38
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   48
      Text            =   "0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   37
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   47
      Text            =   "0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   36
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   46
      Text            =   "0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   35
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   45
      Text            =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   34
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   44
      Text            =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   33
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   43
      Text            =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   32
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   42
      Text            =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   31
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   41
      Text            =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   30
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   40
      Text            =   "0"
      Top             =   4320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   29
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   39
      Text            =   "0"
      Top             =   4320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   28
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   38
      Text            =   "0"
      Top             =   4320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   27
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   37
      Text            =   "0"
      Top             =   4320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   26
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   36
      Text            =   "0"
      Top             =   4320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   25
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   35
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   24
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   34
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   23
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   33
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   22
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   32
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   21
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   31
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   11
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   30
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   12
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   29
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   13
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   28
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   14
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   27
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   15
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   26
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   16
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   25
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   17
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   24
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   18
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   23
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   19
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   22
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   20
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   21
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   10
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   9
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   19
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   8
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   18
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   7
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   17
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   6
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   16
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   5
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   15
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   4
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   3
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   2
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   12
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Txt 
      Height          =   288
      Index           =   1
      Left            =   1800
      MaxLength       =   4
      ScrollBars      =   1  'Horizontal
      TabIndex        =   11
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CheckBox ChkCasc 
      Caption         =   "Use CASC"
      Height          =   492
      Left            =   3720
      TabIndex        =   10
      Top             =   1230
      Width           =   1212
   End
   Begin VB.CheckBox ChkScale 
      Caption         =   "Scaling"
      Height          =   492
      Left            =   3720
      TabIndex        =   9
      Top             =   750
      Width           =   972
   End
   Begin VB.CheckBox ChkUnique 
      Caption         =   "Unique Events (UE) or Marker Horizons (MH)"
      Height          =   492
      Left            =   3720
      TabIndex        =   8
      Top             =   240
      Width           =   1932
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8550
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdApplyRevise 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8550
      TabIndex        =   6
      Top             =   270
      Width           =   1335
   End
   Begin VB.TextBox txtMinWell 
      Height          =   288
      Left            =   2880
      TabIndex        =   4
      Top             =   1290
      Width           =   612
   End
   Begin VB.TextBox txtMinEvent 
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Top             =   810
      Width           =   612
   End
   Begin VB.TextBox txtTotalWell 
      Height          =   288
      Left            =   2880
      TabIndex        =   0
      Top             =   330
      Width           =   612
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   60
      X2              =   10050
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Save INP file as :"
      Height          =   255
      Left            =   5910
      TabIndex        =   100
      Top             =   360
      Width           =   1365
   End
   Begin VB.Label lblM 
      BackColor       =   &H00FF80FF&
      Caption         =   "# of MH"
      Height          =   252
      Left            =   360
      TabIndex        =   99
      Top             =   3960
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lblUE 
      BackColor       =   &H00FF8080&
      Caption         =   "# of UE"
      Height          =   255
      Left            =   360
      TabIndex        =   98
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lab 
      Caption         =   "20"
      Height          =   372
      Index           =   40
      Left            =   4440
      TabIndex        =   91
      Top             =   5040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "19"
      Height          =   372
      Index           =   39
      Left            =   3720
      TabIndex        =   90
      Top             =   5040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "18"
      Height          =   372
      Index           =   38
      Left            =   3000
      TabIndex        =   89
      Top             =   5040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "17"
      Height          =   372
      Index           =   37
      Left            =   2280
      TabIndex        =   88
      Top             =   5040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "16"
      Height          =   372
      Index           =   36
      Left            =   1560
      TabIndex        =   87
      Top             =   5040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "15"
      Height          =   372
      Index           =   35
      Left            =   4440
      TabIndex        =   86
      Top             =   4680
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "14"
      Height          =   372
      Index           =   34
      Left            =   3720
      TabIndex        =   85
      Top             =   4680
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "13"
      Height          =   372
      Index           =   33
      Left            =   3000
      TabIndex        =   84
      Top             =   4680
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "12"
      Height          =   372
      Index           =   32
      Left            =   2280
      TabIndex        =   83
      Top             =   4680
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "11"
      Height          =   372
      Index           =   31
      Left            =   1560
      TabIndex        =   82
      Top             =   4680
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "10"
      Height          =   372
      Index           =   30
      Left            =   4440
      TabIndex        =   81
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 9"
      Height          =   372
      Index           =   29
      Left            =   3720
      TabIndex        =   80
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 8"
      Height          =   372
      Index           =   28
      Left            =   3000
      TabIndex        =   79
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 7"
      Height          =   372
      Index           =   27
      Left            =   2280
      TabIndex        =   78
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 6"
      Height          =   372
      Index           =   26
      Left            =   1560
      TabIndex        =   77
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 5"
      Height          =   372
      Index           =   25
      Left            =   4440
      TabIndex        =   76
      Top             =   3960
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 4"
      Height          =   372
      Index           =   24
      Left            =   3720
      TabIndex        =   75
      Top             =   3960
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 3"
      Height          =   372
      Index           =   23
      Left            =   3000
      TabIndex        =   74
      Top             =   3960
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 2"
      Height          =   372
      Index           =   22
      Left            =   2280
      TabIndex        =   73
      Top             =   3960
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "1"
      Height          =   252
      Index           =   21
      Left            =   1590
      TabIndex        =   72
      Top             =   3960
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Lab 
      Caption         =   "20"
      Height          =   372
      Index           =   20
      Left            =   4440
      TabIndex        =   71
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "19"
      Height          =   372
      Index           =   19
      Left            =   3720
      TabIndex        =   70
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "18"
      Height          =   372
      Index           =   18
      Left            =   3000
      TabIndex        =   69
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "17"
      Height          =   372
      Index           =   17
      Left            =   2280
      TabIndex        =   68
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "16"
      Height          =   372
      Index           =   16
      Left            =   1560
      TabIndex        =   67
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "15"
      Height          =   372
      Index           =   15
      Left            =   4440
      TabIndex        =   66
      Top             =   3120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "14"
      Height          =   372
      Index           =   14
      Left            =   3720
      TabIndex        =   65
      Top             =   3120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "13"
      Height          =   372
      Index           =   13
      Left            =   3000
      TabIndex        =   64
      Top             =   3120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "12"
      Height          =   372
      Index           =   12
      Left            =   2280
      TabIndex        =   63
      Top             =   3120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "11"
      Height          =   372
      Index           =   11
      Left            =   1560
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "10"
      Height          =   372
      Index           =   10
      Left            =   4440
      TabIndex        =   61
      Top             =   2760
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 9"
      Height          =   372
      Index           =   9
      Left            =   3720
      TabIndex        =   60
      Top             =   2760
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 8"
      Height          =   372
      Index           =   8
      Left            =   3000
      TabIndex        =   59
      Top             =   2760
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 7"
      Height          =   372
      Index           =   7
      Left            =   2280
      TabIndex        =   58
      Top             =   2760
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 6"
      Height          =   372
      Index           =   6
      Left            =   1580
      TabIndex        =   57
      Top             =   2760
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 5"
      Height          =   372
      Index           =   5
      Left            =   4440
      TabIndex        =   56
      Top             =   2400
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Lab 
      Caption         =   " 4"
      Height          =   372
      Index           =   4
      Left            =   3720
      TabIndex        =   55
      Top             =   2400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   " 3"
      Height          =   372
      Index           =   3
      Left            =   3000
      TabIndex        =   54
      Top             =   2400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Lab 
      Caption         =   "2"
      Height          =   252
      Index           =   2
      Left            =   2330
      TabIndex        =   53
      Top             =   2400
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Lab 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   52
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblMinWell 
      Caption         =   "Minimum Number of Wells in which each Pair of Events should occur"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1290
      Width           =   2655
   End
   Begin VB.Label lblNumEvent 
      Caption         =   "Minimum Number of Wells in which an Event should occur"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   810
      Width           =   2655
   End
   Begin VB.Label lblTotalWell 
      Caption         =   "Total Number of Wells"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   330
      Width           =   1815
   End
End
Attribute VB_Name = "frmRevise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gFileNum As Integer
Dim SaveINPFileKey As Integer, FileBrowseKey As Integer, FirstTimeCallKey As Integer, InpFileNameKey As Integer
Dim gRecordLen As Long
Dim Counter
Dim Totalwell As Integer
Dim MinEvents As Integer
Dim MinWells As Integer
Dim UniqueEvent As Integer
Dim Scaling As Integer
Dim UseCasc As Integer
Dim CurUESelection(1 To 20) As Integer
Dim CurMHSelection(1 To 20) As Integer
Dim outstr(1 To 40) As String
Dim TotalNumofUEEvents As Integer
Dim TotalNumofMHEvents As Integer
Dim TempFtM As Integer
Dim TempCount As Integer
Dim UniEvent() As String
'Dim CurUESelection(1 To 20) As Integer, CurMHSelection(1 To 20) As Integer
 

Private Sub cboOut_Click()
   FileOut.Pattern = "*.Inp"
   
End Sub

Private Sub ChkCasc_Click()
    'turn off/on the options for ft or m
    
    If ChkCasc.Value = 1 Then
       Option1.Visible = True
       Option2.Visible = True
       FrameUnit.Visible = True
    ElseIf ChkCasc.Value = 0 Then
       Option1.Visible = False
       Option2.Visible = False
       FrameUnit.Visible = False
    End If
    
    If ChkCasc.Value = 0 Then
       MsgBox " Note: If this box is not checked, you cannot continue to run CASC on this dataset. "
    End If
End Sub

Private Sub ChkUnique_Click()
    Dim I As Integer
    MinWells = Trim(txtMinWell.Text)
    MinEvents = Trim(txtMinEvent.Text)
    If ChkUnique.Value = 1 Then
       lblUE.Visible = True
       lblM.Visible = True
       txtUE.Visible = True
       txtM.Visible = True
    Else
       lblUE.Visible = False
       lblM.Visible = False
       txtUE.Visible = False
       txtM.Visible = False
       For I = 1 To 40
          Txt(I).Visible = False
          Lab(I).Visible = False
       Next I
    End If
    If ChkUnique.Value = 1 Then
       If Trim(txtUE.Text) = "" Then txtUE.Text = 0
       If Trim(txtM.Text) = "" Then txtM.Text = 0
       For I = 1 To 20
          If I <= txtUE.Text Then
           Txt(I).Visible = True
           Lab(I).Visible = True
          Else
           Txt(I).Visible = False
           Lab(I).Visible = False
          End If
          If I <= txtM.Text Then
            Txt(I + 20).Visible = True
            Lab(I + 20).Visible = True
          Else
            Txt(I + 20).Visible = False
            Lab(I + 20).Visible = False
          End If
        Next I
        cmdUniEvent.Visible = True
        cmdMarkerEvents.Visible = True
        cmdClear.Visible = True
        ListUniEvent.Visible = True
      Else
         For I = 1 To 40 Step 1
             Txt(I).Visible = False
             Lab(I).Visible = False
         Next I
        cmdUniEvent.Visible = False
        cmdMarkerEvents.Visible = False
        cmdClear.Visible = False
        ListUniEvent.Visible = False
      End If
      
End Sub

Private Sub cmdApplyRevise_Click()
Dim I As Integer
    'Save the current record.
    If Trim(txtOut.Text) = "" Then
        MsgBox "New INP File name is empty"
        Exit Sub
    Else
       Dim pos As Integer
       pos = InStr(txtOut.Text, ".")
       If pos > 0 Then
          txtOut.Text = left(txtOut.Text, pos - 1)
          If Trim(txtOut.Text) = "" Then
            Beep
            MsgBox "Wrong output file name"
            Exit Sub
          End If
       End If
    End If
    If Val(txtTotalWell.Text) <= 0 Or Val(txtTotalWell.Text) >= 100 Then
        MsgBox "Wrong number for Total Number of Wells/Sections : " + txtTotalWell.Text + Chr$(13) + "Accepted Value: 0< X <100"
        Exit Sub
    End If
    If Val(txtMinEvent.Text) <= 0 Or Val(txtMinEvent.Text) >= 100 Then
        MsgBox "Wrong number for Minimum Number of Wells in which an Event should occur: " + txtMinEvent.Text + Chr$(13) + "Accepted Value: 0< X <100"
        Exit Sub
    End If
    If Val(txtMinWell.Text) <= 0 Or Val(txtMinWell.Text) >= 100 Then
        MsgBox "Wrong number for Minimum Number of Wells in which each Pair of Events should occur: " + txtMinWell.Text + Chr$(13) + "Accepted Value: 0< X <100"
        Exit Sub
    End If
    
'Find out whether the INP filename changed
 If Trim(txtOut.Text) = Trim(frmRascW.txt_inp.Text) Then
      InpFileNameKey = 0
Else
     InpFileNameKey = 1
End If

 ' Check whether the output file is already exist.
 If Dir(CurDir + "\" + Trim(txtOut.Text) + ".Inp") <> "" And InpFileNameKey = 1 Then
     I = MsgBox("Inp Outfile " + CurDir + "\" + Trim(txtOut.Text) + ".Inp" + " already exist.  Overwrite it ?", vbYesNo)
    If I = vbNo Then
       SaveINPFileKey = 0
       Exit Sub    ' (vbNo=7;vbYes=6)
    End If
 End If

    
    SaveCurrentRecord
    
    If SaveINPFileKey = 1 Then
       txt_file_inp = Trim(txtOut.Text)
       frmRascW.txt_inp.Text = txt_file_inp
       Unload frmRevise
    End If
End Sub

Private Sub cmdCancel_Click()
        Dim gFileNum As Integer
        Dim I As Integer
        Dim Temp As String
   
   If txtOut.Text <> CurRASCParaFile(1) Then
       If Len(CurRASCParaFile(1)) > 4 And InStr(1, CurRASCParaFile(1), ".") > 0 Then
           Temp = Mid(CurRASCParaFile(1), 1, Len(CurRASCParaFile(1)) - 4)
       End If
       frmRascW.txt_inp.Text = Temp
      'Recover previous record
        gFileNum = FreeFile
        Open CurDir & "\Rasctemp" For Output As gFileNum
        Print #gFileNum, CurRASCParaFile(1)        'fileINp
        Print #gFileNum, CurRASCParaFile(2)        'fileDat
        Print #gFileNum, CurRASCParaFile(3)         'fileDic
        Print #gFileNum, CurRASCParaFile(4)         'FileOut
        Close gFileNum
   End If
   Unload frmRevise
   
   'frmRevise.Hide
End Sub


Private Sub cmdDictionary_Click()
   frmDic.Show 1
End Sub

Private Sub cmdMarkerEvents_Click()
Dim I As Integer
Dim k As Integer

Dim fileDic As String
Dim fileAll As String
Dim gFileNum As Integer
Dim filelen As Integer, TotalEventNumber As Integer, MarkerEventNumber As Integer, OccurMinEvent As Integer
Dim tempName As String
Dim Temp As String

'Set corresponding Unique Event CommandButton status
cmdUniEvent.Caption = "Select Unique Events"

    
If Trim(txtMinEvent.Text) = "" Or Val(txtMinEvent.Text) <= 0 Then
    txtMinEvent.Text = ""
    MsgBox "The criteria value for Marker event file is not valid. " + Chr$(13) + "Please set an appropriate value for Minimum Number of Wells/Sections in which an Event should occur."
   Exit Sub
End If
OccurMinEvent = Val(txtMinEvent.Text)
    
If cmdMarkerEvents.Caption = "Select Marker Events" Then
    
    'keep the event number in an array called UniEvent()
    fileAll = Trim(CurDir + "\" + txt_file_RASCout + ".all")

    ListUniEvent.Clear
    ListUniEvent.ForeColor = &HC0&
    
    gFileNum = FreeFile
    If Dir(fileAll) = "" Then
         MsgBox "Marker event file" + txt_file_RASCout + ".all is not available." + Chr$(13) + "Please check it and try again."
    Exit Sub
    End If
    
    Open fileAll For Input As gFileNum
    TotalEventNumber = 0
'    I = 0
'            While Not (EOF(gFileNum))
    If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, Temp
    End If
    If Trim(Temp) <> "" Then
        TotalEventNumber = Val(Mid(Temp, Len(Temp) - 3, 4))
    End If
            'Display gCascInput.
            'ListUniEvent.AddItem Temp
            'I = I + 1
            'Wend
    'Close gFileNum
     If TotalEventNumber = 0 Then Exit Sub
   'Display gCascInput.
    TempCount = 0
    For I = 1 To TotalEventNumber
            Line Input #gFileNum, Temp
            If Val(Mid(Temp, 5, 4)) >= OccurMinEvent Then
               TempCount = TempCount + 1
            End If
    Next I
    Close gFileNum
    ListUniEvent.AddItem "Number of Marker Events = " + Trim(TempCount)
   If TempCount = 0 Then
       Exit Sub
   End If
   
   ReDim UniEvent(1 To TempCount)
   k = 0
   Open fileAll For Input As gFileNum
    Line Input #gFileNum, Temp
        For I = 1 To TotalEventNumber
            Line Input #gFileNum, Temp
            If Val(Mid(Temp, 5, 4)) >= OccurMinEvent Then
                k = k + 1
                UniEvent(k) = Val(Mid(Temp, 1, 4))
                ListUniEvent.AddItem Temp
            End If
    Next I

    Close gFileNum
          
    ListUniEvent.Visible = True
    cmdMarkerEvents.Caption = "Add Marker Events"
    'TempCount = ListUniEvent.ListCount
    'Mark the current selections
    For I = 1 To ListUniEvent.ListCount - 1
        For k = 1 To TotalNumofMHEvents
            If Val(Mid(ListUniEvent.List(I), 1, 4)) = CurMHSelection(k) Then
                ListUniEvent.Selected(I) = True
            End If
        Next k
    Next I
 
 Else 'select records and add
    k = 0
    For I = 1 To ListUniEvent.ListCount - 1
       If ListUniEvent.Selected(I) Then
            k = k + 1
            If k < 21 Then
                Txt(k + 20).Text = UniEvent(I)
                'add event names into flattening list box
                Txt(k + 20).Visible = True
                Lab(k + 20).Visible = True
                TotalNumofMHEvents = k
                CurMHSelection(k) = Val(UniEvent(I))
          End If
       End If
    Next I
    If k > 20 Then
        MsgBox "You have selected more than 20 Marker events."
    End If
 
 
     ListUniEvent.Clear
     txtM.Text = Str(k)
     For I = k To 19
        Txt(I + 1 + 20).Visible = False
        Lab(I + 1 + 20).Visible = False
     Next I
     
     cmdMarkerEvents.Caption = "Select Marker Events"
  
 End If

End Sub

Private Sub cmdOut_Click()
    cboOut.AddItem "Inp files (*.Inp)"
    cboOut.ListIndex = 0
    If FileBrowseKey = 0 Then
       FileOut.Visible = True
       FileBrowseKey = 1
    Else
       FileOut.Visible = False
       FileBrowseKey = 0
    End If
    
End Sub
 

Private Sub cmdUniEvent_Click()
Dim I As Integer
Dim k As Integer

Dim fileDic As String
Dim FileOut As String
Dim gFileNum As Integer
Dim filelen As Integer, TotalEventNumber As Integer, OccurMinEvent As Integer
Dim tempName As String, fileAll As String
Dim Temp As String
    
'Set corresponding Marker Event CommandButton status
cmdMarkerEvents.Caption = "Select Marker Events"

If Trim(txtMinEvent.Text) = "" Or Val(txtMinEvent.Text) <= 0 Then
    txtMinEvent.Text = ""
    MsgBox "The criteria value for Unique event file is not valid. " + Chr$(13) + "Please set an appropriate value for Minimum Number of Wells/Sections in which an Event should occur."
   Exit Sub
End If
OccurMinEvent = Val(txtMinEvent.Text)
    
If cmdUniEvent.Caption = "Select Unique Events" Then
    
    'keep the event number in an array called UniEvent()
    fileAll = Trim(CurDir + "\" + txt_file_RASCout + ".all")

    ListUniEvent.Clear
    ListUniEvent.ForeColor = &HFF0000
    
    gFileNum = FreeFile
    If Dir(fileAll) = "" Then
         MsgBox "Unique event file" + txt_file_RASCout + ".all is not available." + Chr$(13) + "Please check it and try again."
         Exit Sub
    End If
    
    Open fileAll For Input As gFileNum
    TotalEventNumber = 0
'    I = 0
'            While Not (EOF(gFileNum))
    If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, Temp
    End If
    If Trim(Temp) <> "" Then
        TotalEventNumber = Val(Mid(Temp, Len(Temp) - 3, 4))
    End If
            'Display gCascInput.
            'ListUniEvent.AddItem Temp
            'I = I + 1
            'Wend
    'Close gFileNum
     If TotalEventNumber = 0 Then Exit Sub
   'Display gCascInput.
    TempCount = 0
    For I = 1 To TotalEventNumber
            Line Input #gFileNum, Temp
            If Val(Mid(Temp, 5, 4)) < OccurMinEvent Then
               TempCount = TempCount + 1
            End If
    Next I
    Close gFileNum
    ListUniEvent.AddItem "Number of Unique Events = " + Trim(TempCount)
   If TempCount = 0 Then
       Exit Sub
   End If
   
   ReDim UniEvent(1 To TempCount)
   k = 0
   Open fileAll For Input As gFileNum
    Line Input #gFileNum, Temp
        For I = 1 To TotalEventNumber
            Line Input #gFileNum, Temp
            If Val(Mid(Temp, 5, 4)) < OccurMinEvent Then
                k = k + 1
                UniEvent(k) = Val(Mid(Temp, 1, 4))
                ListUniEvent.AddItem Temp
            End If
    Next I
    Close gFileNum
          
    ListUniEvent.Visible = True
    cmdUniEvent.Caption = "Add Unique Events"
    'Mark the current selections
    For I = 1 To ListUniEvent.ListCount - 1
        For k = 1 To TotalNumofUEEvents
            If Val(Mid(ListUniEvent.List(I), 1, 4)) = CurUESelection(k) Then
                ListUniEvent.Selected(I) = True
            End If
        Next k
    Next I
    
 Else 'select records and add
    k = 0
    For I = 1 To ListUniEvent.ListCount - 1
       If ListUniEvent.Selected(I) Then
       k = k + 1
            If k < 21 Then
                Txt(k).Text = UniEvent(I)
                'add event names into flattening list box
                Txt(k).Visible = True
                Lab(k).Visible = True
                TotalNumofUEEvents = k
                CurUESelection(k) = Val(UniEvent(I))
           End If
       End If
    Next I
    If k > 20 Then
        MsgBox "You have selected more than 20 unique events."
    End If
 
 
     ListUniEvent.Clear
     txtUE.Text = Str(k)
     For I = k To 19
     Txt(I + 1).Visible = False
     Lab(I + 1).Visible = False
     Next I
     
     cmdUniEvent.Caption = "Select Unique Events"
  
 End If
End Sub

Private Sub FileOut_Click()
    Dim filelen As Integer
    Dim tempName As String
    
    tempName = Trim(FileOut.Filename)
    filelen = Len(tempName)
    If filelen > 4 Then
    txtOut.Text = Mid(tempName, 1, filelen - 4)
    'txtOut.Text = FileOut.Filename
    End If
    FileOut.Visible = False
End Sub

Private Sub form_Click()
Dim I As Integer
  If ChkUnique.Value = 1 Then
   For I = 1 To 20 Step 1
      If I <= Trim(txtUE.Text) Then
      Txt(I).Visible = True
      Lab(I).Visible = True
     Else
      Txt(I).Visible = False
       Lab(I).Visible = False
     End If
      If I <= Trim(txtM.Text) Then
        Txt(I + 20).Visible = True
        Lab(I + 20).Visible = True
      Else
        Txt(I + 20).Visible = False
        Lab(I + 20).Visible = False
      End If
    Next
 '   txtUE.Visible = True
 '   txtM.Visible = True
   Else
      For I = 1 To 40 Step 1
     Txt(I).Visible = False
      Lab(I).Visible = False
      Next
 '   txtUE.Visible = False
 '   txtM.Visible = False
    
If ChkCasc.Value = 1 Then
   Option1.Visible = True
   Option2.Visible = True
   FrameUnit.Visible = True
ElseIf ChkCasc.Value = 0 Then
   Option1.Visible = False
   Option2.Visible = False
   FrameUnit.Visible = False
End If
   
   
    
      
End If
End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    txtTotalWell.Text = ""
    txtMinEvent.Text = ""
    txtMinWell.Text = ""
       For I = 1 To 40 Step 1
         Txt(I).Visible = False
          Lab(I).Visible = False
       Next I
    SaveINPFileKey = 0
    FileBrowseKey = 0
    FirstTimeCallKey = 1
                     
    'Display the current record.
    ShowCurrentRecord
    txtOut.Text = Mid(txt_file_inp, 1, Len(txt_file_inp) - 4)

End Sub

Public Sub ShowCurrentRecord()

Dim I, J As Integer
Dim gFileNum As Integer, Istep As Integer
Dim Temp As String

TotalNumofUEEvents = 0
TotalNumofMHEvents = 0

gFileNum = FreeFile
 If Dir(txt_file_inp) = "" Then
    'Create this new empty INP file in current directory for edit
    Open CurDir & "\" + txt_file_inp For Output As gFileNum
    Close gFileNum
    
    'MsgBox "Inp File: " + CurDir + "\" + txt_file_inp + " does not exist"
    'Exit Sub
 End If
  
Open CurDir & "\" + txt_file_inp For Input As gFileNum

'Give initial values of parameters for a new file (empty INP file generated by rascw.exe)
      Totalwell = 0
      MinEvents = 0
      UniqueEvent = 1
       Scaling = 1
       UseCasc = 1
       MinWells = 0
       TempFtM = 0

Temp = ""
'Get the three records.
If EOF(gFileNum) = False Then
      Line Input #gFileNum, Temp
      If Len(Temp) = 14 Then
            Totalwell = Val(Mid(Temp, 1, 2))
            MinEvents = Val(Mid(Temp, 3, 2))
            UniqueEvent = Val(Mid(Temp, 5, 2))
             Scaling = Val(Mid(Temp, 7, 2))
             UseCasc = Val(Mid(Temp, 9, 2))
             MinWells = Val(Mid(Temp, 11, 2))
             TempFtM = Val(Mid(Temp, 13, 2))
       Else
             'judge whether the INP file is a new file
             If Len(Trim(Temp)) = 0 Or Len(Trim(Temp)) > 14 Then
                  MsgBox "The first line of current INP file is empty or incorrect, system default value will be given."
             Else
                  For I = 1 To Len(Mid(Temp, 1, 14)) Step 2
                      Select Case I
                      Case 1
                            Totalwell = Val(Mid(Temp, 1, 2))
                      Case 3
                            MinEvents = Val(Mid(Temp, 3, 2))
                       Case 5
                           UniqueEvent = Val(Mid(Temp, 5, 2))
                      Case 7
                             Scaling = Val(Mid(Temp, 7, 2))
                      Case 9
                             UseCasc = Val(Mid(Temp, 9, 2))
                      Case 11
                             MinWells = Val(Mid(Temp, 11, 2))
                      Case 13
                             TempFtM = Val(Mid(Temp, 13, 2))
                      End Select
                  Next I
             End If
       End If
End If
If EOF(gFileNum) = False Then
     Line Input #gFileNum, Temp
     If Len(Trim(Temp)) = 0 Then
         TotalNumofUEEvents = 0
     Else
         For I = 1 To Len(Temp) Step 4
           If Trim(Mid(Temp, I, 4)) <> "" Then
              TotalNumofUEEvents = TotalNumofUEEvents + 1
           Else
               Exit For
           End If
        Next I
'         If Len(Trim(Temp)) Mod 4 = 0 Then
'              TotalNumofUEEvents = Len(Trim(Temp)) / 4
'    '          MsgBox "If Mod = 0 " + Str(TotalNumofUEEvents)
'           Else
'              TotalNumofUEEvents = Int(Len(Trim(Temp)) / 4) + 1
'     '        MsgBox "MOD <> 0" + Str(TotalNumofUEEvents)
'           End If
     End If
End If
Istep = 0
If TotalNumofUEEvents > 0 Then
     For I = 1 To TotalNumofUEEvents * 4 Step 4
         Istep = Istep + 1
         CurUESelection(Istep) = Val(Mid(Temp, I, 4))
        ' Txt(I).Text = Trim(CurUESelection(Istep))
    Next
End If

If EOF(gFileNum) = False Then
   Line Input #gFileNum, Temp
   If Len(Trim(Temp)) = 0 Then
       TotalNumofMHEvents = 0
   Else
         For I = 1 To Len(Temp) Step 4
           If Trim(Mid(Temp, I, 4)) <> "" Then
              TotalNumofMHEvents = TotalNumofMHEvents + 1
           Else
               Exit For
           End If
        Next I
'       If Len(Trim(Temp)) Mod 4 = 0 Then
'          TotalNumofMHEvents = Len(Temp) / 4
'  '        MsgBox "If Mod = 0 " + Str(TotalNumofMHEvents)
'       Else
'          TotalNumofMHEvents = Int(Len(Temp) / 4) + 1
'   '      MsgBox "MOD <> 0" + Str(TotalNumofMHEvents)
'       End If
   End If
End If
Istep = 0
If TotalNumofMHEvents > 0 Then
     For I = 1 To TotalNumofMHEvents * 4 Step 4
         Istep = Istep + 1
         CurMHSelection(Istep) = Val(Mid(Temp, I, 4))
         'Txt(I + 20).Text = Trim(CurMHSelection(Istep))
    Next
End If
  
Close gFileNum
  
txtUE.Text = Str(TotalNumofUEEvents)
txtM.Text = Str(TotalNumofMHEvents)
  
    
txtTotalWell.Text = Trim(Totalwell)
txtMinEvent.Text = Trim(MinEvents)
txtMinWell.Text = Trim(MinWells)
ChkUnique.Value = Trim(UniqueEvent)
ChkScale.Value = Trim(Scaling)
ChkCasc.Value = Trim(UseCasc)

If TempFtM = 1 Then
   Option1.Value = False
   Option2.Value = True
ElseIf TempFtM = 0 Then
   Option1.Value = True
   Option2.Value = False
End If
If ChkCasc.Value = 1 Then
   Option1.Visible = True
   Option2.Visible = True
   FrameUnit.Visible = True
ElseIf ChkCasc.Value = 0 Then
   Option1.Visible = False
   Option2.Visible = False
   FrameUnit.Visible = False
End If


If ChkUnique.Value = 1 Then
    lblUE.Visible = True
    lblM.Visible = True
    txtUE.Visible = True
    txtM.Visible = True
 
    For I = 1 To txtUE.Text Step 1
         Txt(I).Visible = True
         Txt(I).Text = Trim(CurUESelection(I))
         Lab(I).Visible = True
    Next I
    
    For I = 1 To txtM.Text Step 1
        Txt(I + 20).Visible = True
        Txt(I + 20).Text = Trim(CurMHSelection(I))
        Lab(I + 20).Visible = True
     Next I
        cmdUniEvent.Visible = True
        cmdMarkerEvents.Visible = True
        ListUniEvent.Visible = True
Else
         For I = 1 To 40 Step 1
             Txt(I).Visible = False
             Lab(I).Visible = False
         Next I
        cmdUniEvent.Visible = False
        cmdMarkerEvents.Visible = False
        ListUniEvent.Visible = False
End If
   
Close gFileNum

End Sub

Public Sub SaveCurrentRecord()
'Fill gCascInput with the currently displayed data
Dim I As Integer
Dim gFileNum As Integer
Dim addspace As String
Dim wordlen As Integer
Dim strapp1 As String
Dim strapp2 As String
Dim tempSpace As String
Dim Opt_Ft_M As String

' ' Check whether the output file is already exist.
' If Dir(CurDir + "\" + Trim(txtOut.Text) + ".Inp") <> "" Then
'     I = MsgBox("Inp Outfile " + CurDir + "\" + Trim(txtOut.Text) + ".Inp" + " already exist.  Overwrite it ?", vbYesNo)
'    If I = vbNo Then
'       SaveINPFileKey = 0
'       Exit Sub    ' (vbNo=7;vbYes=6)
'    End If
' End If

gFileNum = FreeFile

Open CurDir + "\" + Trim(txtOut.Text) + ".Inp" For Output As gFileNum

Totalwell = txtTotalWell.Text
MinEvents = txtMinEvent.Text
MinWells = txtMinWell.Text
UniqueEvent = ChkUnique.Value
Scaling = ChkScale.Value
UseCasc = ChkCasc.Value
If Option1.Value = True Then
   Opt_Ft_M = "0"
Else
   Opt_Ft_M = "1"
End If
   

'Save gCascInput to the current record
'check the length of totalwell
tempSpace = ""
tempSpace = Space(2 - Len(Trim(Totalwell)))
'According to Fortran I2 format
Print #gFileNum, tempSpace + Trim(Totalwell); Space(2 - Len(Trim(MinEvents))); Trim(MinEvents); _
      Space(2 - Len(Trim(UniqueEvent))); Trim(UniqueEvent) + Space(2 - Len(Trim(Scaling))) + Trim(Scaling); _
      Space(2 - Len(Trim(UseCasc))) + Trim(UseCasc); Space(2 - Len(Trim(MinWells))); Trim(MinWells); Space(2 - Len(Trim(Opt_Ft_M))); Opt_Ft_M

strapp1 = ""

For I = 1 To 20
     If I <= Val(txtUE.Text) Then
         outstr(I) = Txt(I).Text
         wordlen = Len(Trim(outstr(I)))
         addspace = Space(4 - wordlen)
         strapp1 = strapp1 + addspace + Trim(outstr(I))
      Else
         outstr(I) = ""
     End If
Next I

strapp2 = ""
For I = 1 To 20
     If I <= Val(txtM.Text) Then
         outstr(I + 20) = Trim(Txt(I + 20).Text)
         wordlen = Len(Trim(outstr(I + 20)))
         addspace = Space(4 - wordlen)
         strapp2 = strapp2 + addspace + Trim(outstr(I + 20))
      Else
         outstr(I + 20) = ""
     End If
Next I
 
Print #gFileNum, strapp1
Print #gFileNum, strapp2

Close gFileNum
SaveINPFileKey = 1
End Sub


Private Sub txtMinEvent_Change()
 Dim I As Integer
   If FirstTimeCallKey <> 1 Then
        If txtMinEvent.Text <> "" Then
            If Val(txtMinEvent.Text) <= 0 Or Val(txtMinEvent.Text) >= 100 Then
                MsgBox "Wrong number for Minimum Number of Wells in which an Event should occur: " + txtMinEvent.Text + Chr$(13) + "Accepted Value: 0< X <100"
                Exit Sub
            End If
        End If
        RecordModify
        RecordReflesh
        ListReflesh
  End If
  FirstTimeCallKey = 0

End Sub

Private Sub txtMinWell_Change()
    If txtMinWell.Text <> "" Then
       If Val(txtMinWell.Text) <= 0 Or Val(txtMinWell.Text) >= 100 Then
           MsgBox "Wrong number for Minimum Number of Wells in which each Pair of Events should occur: " + txtMinWell.Text + Chr$(13) + "Accepted Value: 0< X <100"
           Exit Sub
       End If
    End If
End Sub

Private Sub txtTotalWell_Change()
    If txtTotalWell.Text <> "" Then
        If Val(txtTotalWell.Text) <= 0 Or Val(txtTotalWell.Text) >= 100 Then
            MsgBox "Wrong number for Total Number of Wells/Sections : " + txtTotalWell.Text + Chr$(13) + "Accepted Value: 0< X <100"
            Exit Sub
        End If
    End If
End Sub

Private Sub txtUE_Change()
 Dim I As Integer
 If Trim(txtUE.Text) = "" Then txtUE.Text = 0
   For I = 1 To 20
      If I <= txtUE.Text Then
       Txt(I).Visible = True
       Lab(I).Visible = True
      Else
       Txt(I).Visible = False
       Lab(I).Visible = False
      End If
   Next
End Sub

Private Sub RecordReflesh()
 Dim I As Integer
        txtUE.Text = Str(TotalNumofUEEvents)
        For I = 1 To 20
          If I <= TotalNumofUEEvents Then
           Txt(I).Visible = True
           Txt(I).Text = Trim(CurUESelection(I))
           Lab(I).Visible = True
          Else
           Txt(I).Visible = False
           Lab(I).Visible = False
          End If
          txtM.Text = Str(TotalNumofMHEvents)
          If I <= TotalNumofMHEvents Then
            Txt(I + 20).Visible = True
            Txt(I + 20).Text = Trim(CurMHSelection(I))
            Lab(I + 20).Visible = True
          Else
            Txt(I + 20).Visible = False
            Lab(I + 20).Visible = False
          End If
        Next I
End Sub

Private Sub RecordModify()
  Dim I, J
  Dim fileAll As String
  Dim gFileNum As Integer
  Dim filelen As Integer, TotalEventNumber As Integer, MarkerEventNumber As Integer, OccurMinEvent As Integer
  Dim tempName As String
  Dim Temp As String, TempUECount As Integer, TempMHCount As Integer
  Dim TempUESelection(1 To 20) As Integer
  Dim TempMHSelection(1 To 20) As Integer

   If Trim(txtMinEvent.Text) = "" Or Val(txtMinEvent.Text) <= 0 Then
        txtMinEvent.Text = ""
        MsgBox "The criteria value for Marker event file is not valid. " + Chr$(13) + "Please set an appropriate value for Minimum Number of Wells/Sections in which an Event should occur."
        Exit Sub
   End If
   OccurMinEvent = Val(txtMinEvent.Text)

    'keep the event number in an array called UniEvent()
    fileAll = Trim(CurDir + "\" + txt_file_RASCout + ".all")

    gFileNum = FreeFile
    If Dir(fileAll) = "" Then
         MsgBox "Marker event file" + txt_file_RASCout + ".all is not available." + Chr$(13) + "Please check it and try again."
        Exit Sub
    End If
    
    Open fileAll For Input As gFileNum
    TotalEventNumber = 0
    If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, Temp
    End If
    If Trim(Temp) <> "" Then
        TotalEventNumber = Val(Mid(Temp, Len(Temp) - 3, 4))
    End If
    If TotalEventNumber = 0 Then Exit Sub
   'Display gCascInput.
    TempUECount = 0
    TempMHCount = 0
    For I = 1 To TotalEventNumber
            Line Input #gFileNum, Temp
            For J = 1 To TotalNumofUEEvents
               If Val(Mid(Temp, 1, 4)) = CurUESelection(J) And Val(Mid(Temp, 5, 4)) < OccurMinEvent Then
                   TempUECount = TempUECount + 1
                   TempUESelection(TempUECount) = CurUESelection(J)
               End If
            Next J
            For J = 1 To TotalNumofMHEvents
               If Val(Mid(Temp, 1, 4)) = CurMHSelection(J) And Val(Mid(Temp, 5, 4)) >= OccurMinEvent Then
                   TempMHCount = TempMHCount + 1
                   TempMHSelection(TempMHCount) = CurMHSelection(J)
               End If
            Next J
    Next I
    Close gFileNum
'Modify and Save result
    TotalNumofUEEvents = TempUECount
    For I = 1 To TempUECount
        CurUESelection(I) = TempUESelection(I)
    Next I
    TotalNumofMHEvents = TempMHCount
    For I = 1 To TempMHCount
        CurMHSelection(I) = TempMHSelection(I)
    Next I
    
End Sub

Private Sub ListReflesh()
        ListUniEvent.Clear
        cmdMarkerEvents.Caption = "Select Marker Events"
        cmdUniEvent.Caption = "Select Unique Events"
End Sub

Private Sub cmdClear_Click()
        ListUniEvent.Clear
        cmdMarkerEvents.Caption = "Select Marker Events"
        cmdUniEvent.Caption = "Select Unique Events"
        TotalNumofUEEvents = 0
        TotalNumofMHEvents = 0
        RecordReflesh
End Sub

