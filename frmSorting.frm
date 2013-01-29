VERSION 5.00
Begin VB.Form frmSorting 
   Caption         =   "Sorting Data"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frmSorting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   612
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Ascending"
      Height          =   252
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   1212
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Descending"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1452
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   252
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   612
   End
   Begin VB.ListBox List1 
      DataField       =   " "
      DataSource      =   "Data1"
      Height          =   645
      ItemData        =   "frmSorting.frx":0442
      Left            =   240
      List            =   "frmSorting.frx":046A
      TabIndex        =   3
      Top             =   480
      Width           =   2052
   End
   Begin VB.TextBox txtSortField 
      Height          =   288
      Left            =   2640
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "Select item to sort"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   ">>"
      Height          =   252
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   252
   End
End
Attribute VB_Name = "frmSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
sort_method = ""
frmSorting.Hide
End Sub

Private Sub cmdOK_Click()
sort_field = txtSortField.Text
    If Option1.Value = True Then
       sort_method = "DESC"
    ElseIf Option2.Value = True Then
       sort_method = "ASC"
    Else
       MsgBox "You have to select a sorting method"
    Exit Sub
    End If
   frmSorting.Hide
End Sub


Private Sub List1_Click()
txtSortField.Text = List1.Text
End Sub
