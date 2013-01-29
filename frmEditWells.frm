VERSION 5.00
Begin VB.Form frmEditWells 
   Caption         =   "Edit Wells"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "frmEditWells.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   372
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   732
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   372
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   372
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.ListBox lstWellNames 
      DataField       =   "well_name"
      DataSource      =   " "
      Height          =   1230
      ItemData        =   "frmEditWells.frx":0442
      Left            =   120
      List            =   "frmEditWells.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   3012
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
End
Attribute VB_Name = "frmEditWells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempName As String

Private Sub Form_Load()
Dim TabIndex As Integer
Dim TabCount As Integer

'Set dbsSAGA = OpenDatabase(myDBName)

'Data1.DatabaseName = myDBName
'    With Data1.Database
'    TabCount = 0
'       If .TableDefs.count = 0 Then
'          MsgBox "The database:  " + myDBName + " is empty", vbExclamation
'          Exit Sub
'       End If
'        For TabIndex = 0 To .TableDefs.count - 1
'               If Left$(.TableDefs(TabIndex).Name, 4) = "head" Then
'              Label1.DataField = .TableDefs(TabIndex).Fields(0).Name
'               Data1.RecordSource = .TableDefs(TabIndex).Name
'                tempName = Label1.Caption
'                lstWellNames.AddItem tempName
'
'                TabCount = TabCount + 1
'            End If
'        Next TabIndex
'     End With
End Sub


