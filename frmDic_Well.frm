VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDic_Well 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dictionary"
   ClientHeight    =   5244
   ClientLeft      =   5316
   ClientTop       =   1152
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5244
   ScaleWidth      =   5880
   Begin VB.CommandButton Command1 
      Caption         =   "Sort"
      Height          =   252
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   612
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDic_Well.frx":0000
      Height          =   4572
      Left            =   240
      OleObjectBlob   =   "frmDic_Well.frx":0010
      TabIndex        =   3
      Top             =   600
      Width           =   5412
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\rascw991\testips.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   4440
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Event_name"
      Top             =   120
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox txtSearch 
      Height          =   288
      Left            =   240
      TabIndex        =   2
      Top             =   84
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   252
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   252
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "frmDic_Well"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim textToSearch As String * 40
Dim MyPos As String * 40

Private Sub cmdCancel_Click()
'RichTextDic.Text = ""
RichTextDic.Refresh
Unload frmDic
End Sub

Private Sub CmdHelp_Click()
MsgBox "Select a Dic file"
End Sub

Private Sub cmdSearch_Click()
 Dim FoundPos As Integer
 Dim FoundLine As Integer
    ' Find the text specified in the TextBox control.
   FoundPos = RichTextDic.Find(txtSearch.Text, 1, , (rtfWholeWord = 1 Or rtfWholeWord = 0) _
           And (rtfNoHighlight = 6))

    ' Show message based on whether the text was found or not.

    'If FoundPos <> -1 Then
        ' Returns number of line containing found text.
     '   FoundLine = RichTextDic.GetLineFromChar(FoundPos)
     '   MsgBox "Word found on line " & CStr(FoundLine)
    'Else
     '   MsgBox "Word not found."

    'End If
txtSearch.Text = ""
End Sub





Private Sub Command1_Click()

Data1.RecordSource = "SELECT Event_Name, Event_Number from " + "Event_Name" '+ " order by " + sort_field + " " + sort_method + ";"
     Data1.Refresh
End Sub

Private Sub Form_Load()
Dim dbsMyTemp As Database
Dim rstMytemp As Recordset

Data1.DatabaseName = myDBName

Data1.RecordSource = "Event_Name"

'RichTextDic.Text = Data1.Recordset(1).Value




End Sub


 
