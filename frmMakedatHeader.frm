VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMakedatheader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Well Headers"
   ClientHeight    =   7440
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frmMakedatHeader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9480
   Visible         =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmMakedatHeader.frx":0442
      Height          =   7215
      Left            =   120
      OleObjectBlob   =   "frmMakedatHeader.frx":0456
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\users\frits\qiuming\test35.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "header"
      Top             =   240
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmMakedatheader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myDB As Database
Dim myST As Recordset
Dim AddUpdate As String
Dim table_Changeable As Boolean
Dim tmp_Well_Name As String
Dim tmp_Well_Header As String





Private Sub Form_Load()
Dim NumOfWell, RowOfTab As Integer
Dim I As Integer
Dim rstHeader As Recordset

 
Set DBsSAGA = OpenDatabase(myDBName)
   If DBsSAGA.TableDefs.count > 2 Then
     NumOfWell = DBsSAGA.TableDefs.count - 3
   Else
     MsgBox "It is not right database"
     Exit Sub
   End If
   
   With DBsSAGA
        ' Open table-type Recordset and show RecordCount
        ' property.
     Set rstHeader = .OpenRecordset("header")
   With rstHeader
     If .EOF = True And NumOfWell > 3 Then
       For I = 1 To NumOfWell - 3
        .AddNew
        .Fields(0) = Str(I)
        .Fields(1) = ""
        .Fields(2) = ""
        .Fields(3) = 0
        .Fields(4) = ""
        .Fields(5) = 0
        .Fields(6) = ""
        .Update
        .Bookmark = .LastModified
       Next I
     End If
   End With
   End With
 DBsSAGA.Close
 
DBGrid1.Visible = False

Data1.DatabaseName = myDBName
Data1.RecordSource = "header"

DBGrid1.Visible = True


End Sub


