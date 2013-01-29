VERSION 5.00
Begin VB.Form frmChartTable1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   435
      Left            =   3930
      TabIndex        =   5
      Top             =   3060
      Width           =   2805
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy"
      Height          =   465
      Left            =   3930
      TabIndex        =   4
      Top             =   2400
      Width           =   2805
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   525
      Left            =   3930
      TabIndex        =   3
      Top             =   1680
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   3930
      TabIndex        =   2
      Top             =   1020
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3960
      TabIndex        =   1
      Top             =   300
      Width           =   2655
   End
   Begin VB.TextBox DataText 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmChartTable1.frx":0000
      Top             =   30
      Width           =   6945
   End
End
Attribute VB_Name = "frmChartTable1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DataText.Text = "line 1" + Chr$(vbKeyControl) + Chr$(13)
    
End Sub

Private Sub Command2_Click()
    DataText.Text = "line 2" + Chr$(85)

End Sub

Private Sub Command3_Click()
    DataText.Text = "line 3" + Chr$(85)

End Sub

Private Sub Command5_Click()
    DataText.Text = ""

End Sub
