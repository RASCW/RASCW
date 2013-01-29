VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DbList32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMakedatWell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Well Editor"
   ClientHeight    =   7410
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9795
   Icon            =   "frmMakedatWell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9795
   Visible         =   0   'False
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmMakedatWell.frx":0442
      Height          =   615
      Left            =   3360
      OleObjectBlob   =   "frmMakedatWell.frx":0456
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Header"
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtEventNumber 
      DataField       =   "Event_Number"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text14"
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   252
      Left            =   8280
      TabIndex        =   23
      Top             =   2280
      Width           =   852
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmMakedatWell.frx":0E29
      Height          =   4455
      Left            =   240
      OleObjectBlob   =   "frmMakedatWell.frx":0E3D
      TabIndex        =   0
      Top             =   2880
      Width           =   9255
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\rascw991\ips2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Event_name"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\rasccasc\june28.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "header"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fossil Events in Samples"
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   9252
      Begin VB.TextBox txtTemp 
         DataField       =   "Event_Number"
         DataSource      =   "Data2"
         Height          =   288
         Left            =   2280
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1800
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "c:\rascw991\ips2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   324
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Event_name"
         Top             =   2040
         Visible         =   0   'False
         Width           =   912
      End
      Begin VB.TextBox txtFossil_Name 
         DataField       =   "Fossil_Names"
         DataSource      =   "Data1"
         Height          =   288
         HideSelection   =   0   'False
         Left            =   120
         TabIndex        =   26
         Text            =   " "
         Top             =   1440
         Width           =   2892
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frmMakedatWell.frx":1810
         DataField       =   "Fossil_Name"
         DataSource      =   "Data5"
         Height          =   288
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3132
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "Fossil_Name"
         Text            =   ""
      End
      Begin VB.TextBox txtE_Qualifier 
         DataField       =   "Event_Qualifier"
         DataSource      =   "Data1"
         Height          =   288
         Left            =   5760
         TabIndex        =   5
         Text            =   " "
         Top             =   600
         Width           =   1092
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   5760
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   600
         Width           =   1332
      End
      Begin VB.TextBox txtQuality 
         DataField       =   "Sample_Quality"
         DataSource      =   "Data1"
         Height          =   288
         Left            =   3720
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox txtCount 
         DataField       =   "Count"
         DataSource      =   "Data1"
         Height          =   288
         Left            =   7560
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1332
      End
      Begin VB.TextBox txtRel_Abundance 
         DataField       =   "Rel_Abundance"
         DataSource      =   "Data1"
         Height          =   288
         Left            =   5760
         TabIndex        =   3
         Text            =   " "
         Top             =   1440
         Width           =   2892
      End
      Begin VB.TextBox txtType 
         DataField       =   "Sample_type"
         DataSource      =   "Data1"
         Height          =   288
         Left            =   1800
         TabIndex        =   4
         Text            =   " "
         Top             =   600
         Width           =   1212
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add "
         Height          =   252
         Left            =   3720
         TabIndex        =   19
         Top             =   1920
         Width           =   852
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   252
         Left            =   6600
         TabIndex        =   18
         Top             =   1920
         Width           =   852
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   252
         Left            =   5160
         TabIndex        =   17
         Top             =   1920
         Width           =   852
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   252
         Left            =   8160
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.TextBox TxtDepth 
         DataField       =   "Sample_depth"
         DataSource      =   "Data1"
         Height          =   288
         Left            =   120
         TabIndex        =   9
         Text            =   "1"
         Top             =   600
         Width           =   1332
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\rasccasc\june28.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "well1_inf"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.ComboBox Combo3 
         Height          =   288
         ItemData        =   "frmMakedatWell.frx":1824
         Left            =   1800
         List            =   "frmMakedatWell.frx":1826
         TabIndex        =   7
         Text            =   " "
         Top             =   600
         Width           =   1452
      End
      Begin VB.ComboBox Combo7 
         Height          =   288
         ItemData        =   "frmMakedatWell.frx":1828
         Left            =   5760
         List            =   "frmMakedatWell.frx":182A
         TabIndex        =   6
         Text            =   " "
         Top             =   1440
         Width           =   3132
      End
      Begin VB.ComboBox Combo4 
         Height          =   288
         Left            =   3720
         TabIndex        =   2
         Text            =   " "
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label13 
         Caption         =   "Event Qualifier"
         Height          =   252
         Index           =   0
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label14 
         Caption         =   " Event Number"
         Height          =   252
         Left            =   3720
         TabIndex        =   22
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label Label13 
         Caption         =   " Count"
         Height          =   252
         Index           =   1
         Left            =   7560
         TabIndex        =   21
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label7 
         Caption         =   "Sample Depth"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Sample Type"
         Height          =   252
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "Sample Quality"
         Height          =   252
         Left            =   3720
         TabIndex        =   13
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "Event Name"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label Label12 
         Caption         =   "Relative Abundance"
         Height          =   252
         Left            =   5760
         TabIndex        =   11
         Top             =   1200
         Width           =   1932
      End
   End
End
Attribute VB_Name = "frmMakedatWell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myDB As Database
Dim myST As Recordset
Dim AddUpdate As String
Dim tmp_Well_Name As String
Dim Cap_Well As String


Private Sub Check1_Click()
If Check1.Value = 1 Then
DBGrid2.Visible = True
Else
DBGrid2.Visible = False
End If
End Sub

Private Sub cmdAddNew_Click()
Dim result
On Error GoTo Errorhandler

    If AddUpdate = "Add" Then
       result = MsgBox("Do you want to add new data? ", vbYesNo, "Confirm Adding")
       If result = vbNo Then Exit Sub
       Data1.Recordset.AddNew
        
        TxtDepth.Text = "0"
        txtEventNumber.Text = "0"
        txtCount.Text = "0"
'       Text10.Text = ""
'       Text11.Text = ""
'       Text12.Text = ""
        AddUpdate = "Update"
        cmdAddNew.Caption = AddUpdate
    ElseIf AddUpdate = "Update" Then
        If Check() = True Then
           Data1.Recordset.Update
           Data1.Recordset.Bookmark = Data1.Recordset.LastModified
           
           Data1.Refresh
           AddUpdate = "Add"
           cmdAddNew.Caption = AddUpdate
           cmdEdit.Visible = True
           cmdDelete.Visible = True
        
         End If
     End If
  
  
     
     Exit Sub
Errorhandler:
    MsgBox "Error in add records", vbCritical, "Error"
End Sub


'Private Sub cmdCancel_Click()
'myDBName = ""
'Unload frmMakedatWell'

'End Sub

Private Sub cmdDelete_Click()
Dim result
On Error GoTo Errorhandler

    result = MsgBox("Do you want to delete the current record? ", vbYesNo, "Confirm Deleting")
    If result = vbNo Then Exit Sub
    Data1.Recordset.Delete
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    Data1.Refresh
Errorhandler:

End Sub

Private Sub cmdEdit_Click()
Dim response
On Error GoTo Errorhandler
   response = MsgBox("Do you want to modify the current record?", vbOKCancel)
   If response = vbCancel Then
   Exit Sub
   Else
        If Check() = True Then
        Data1.Recordset.Edit
        Data1.Recordset.Update
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        Data1.Refresh
        Else
         Exit Sub
        End If
  
  End If
  
Errorhandler:

End Sub



Private Sub cmdHeader_Click()
DBGrid2.Visible = True
End Sub

 

Private Sub cmdSort_Click()
Dim response
   response = MsgBox("Do you want to sort the data?", vbOKCancel, "")
   If response = vbCancel Then
   Exit Sub
   Else
        frmSortingW.Show 1
         If sort_method = "" Then Exit Sub
      '    Data1.RecordSource = "select * from well_info order by " + sort_field + " " + sort_method + ";"
          
  '     Data1.RecordSource = "SELECT " + tmp_Well_Name + ".[Sample_depth], " + tmp_Well_Name + ".[Sample_type], " + _
  '     " Event_Name.[Fossil_name], event_Name.[Event_type]," + _
  '      tmp_Well_Name + ".[Event_number], " + tmp_Well_Name + ".[Sample_quality], " + _
  '      tmp_Well_Name + ".[Rel_abundance], " + tmp_Well_Name + ".[Count], " + _
  '      tmp_Well_Name + ".[Event_qualifier] " + _
  '     " from " + tmp_Well_Name + "INNER JOIN Event_Name ON Event_Name.[Event_number] " + _
  '     "= " + tmp_Well_Name + ".[Event_number] order by " + tmp_Well_Name + "." + sort_field + _
  '     " " + sort_method + ";"
    
     Data1.RecordSource = "SELECT * from " + tmp_Well_Name + " order by " + sort_field + " " + sort_method + ";"
     Data1.Refresh
    End If

End Sub




Private Sub Combo2_Click()
Dim gEventQ As String
  gEventQ = Combo2.Text
    Select Case gEventQ
       Case "uncertain"
        txtE_Qualifier.Text = "?"
       Case "affinis"
        txtE_Qualifier.Text = "aff"
       Case "confer"
        txtE_Qualifier.Text = "cf"
       Case Else
        MsgBox "Select event qualifier from browser!", vbExclamation
     End Select
End Sub

Private Sub Combo3_Click()
txtType.Text = Left$(Combo3.Text, 1)

End Sub

Private Sub Combo4_Click()
txtQuality.Text = Left$(Combo4.Text, 1)
End Sub

Private Sub Combo7_Click()
txtRel_Abundance.Text = Left$(Combo7.Text, 1)

End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub DBCombo1_Change()
Data2.RecordSource = "SELECT Event_Number from Event_Name Where Fossil_Name = '" & DBCombo1.Text & "';"
 
Data2.Refresh
End Sub

Private Sub DBCombo1_Click(Area As Integer)
txtFossil_Name.Visible = False

txtFossil_Name.Text = DBCombo1.Text
Data2.RecordSource = "SELECT Event_Number from Event_Name Where Fossil_Name = '" & DBCombo1.Text & "';"
 
Data2.Refresh
End Sub



 

Private Sub DBCombo2_Click(Area As Integer)
'Text14.Text = DBCombo2.Text
End Sub




Private Sub Form_Load()
Dim tempstr As String
Dim dbsMyTemp As Database
Dim rstMytemp As Recordset
Dim I, NumOfrecord As Integer
Dim Cap_Well_Local As String
Dim rst1 As Recordset
'Well name
Cap_Well_Local = Well_Caption
'Well_Caption
'Name of well table
tmp_Well_Name = Well_Name_Temp

Data1.DatabaseName = myDBName
Data3.DatabaseName = myDBName
Data5.DatabaseName = myDBName

Data2.DatabaseName = myDBName
'Data2.RecordSource = "SELECT Event_Number from Event_Name Where Fossil_Name = '" & Trim(DBCombo1.Text) & "';"
 'Data2.Refresh

'Data1.RecordSource = tmp_Well_Name
Data1.RecordSource = "SELECT * from " + tmp_Well_Name + " order by " + "Sample_depth" + " " + "ASC" + ";"
    Data1.Refresh
'Data3.RecordSource = "header"
Data3.RecordSource = "SELECT * from header Where Well_Name = '" & Cap_Well_Local & "';"
 Data3.Refresh

Data5.RecordSource = "select fossil_name from Event_name order by fossil_name"

'Set dbsMyTemp = OpenDatabase(myDBName)
'    With dbsMyTemp
'        ' Open table-type Recordset and show RecordCount
'        ' property.
'     Set rstMytemp = .OpenRecordset("Event_name")
'
'       If rstMytemp.EOF = False Then
'          rstMytemp.MoveLast
'          NumOfrecord = rstMytemp.RecordCount
'       End If
'   End With
 '   If NumOfrecord > 0 Then
 '      rstMytemp.MoveFirst
 '      For I = 0 To NumOfrecord - 1
 '       If IsNull(rstMytemp.Fields(0).Value) = True And IsNull(rstMytemp.Fields(2).Value) = True Then
 '          tempstr = ""
 '        ElseIf IsNull(rstMytemp.Fields(0).Value) = True Then
 '          tempstr = rstMytemp.Fields(2).Value
 '        ElseIf IsNull(rstMytemp.Fields(2).Value) = True Then
 '           tempstr = rstMytemp.Fields(0).Value
 '        Else
 '           tempstr = rstMytemp.Fields(0).Value + " " + rstMytemp.Fields(2).Value
 '        End If
 '
 '      Combo1.AddItem tempstr, I
 '         If rstMytemp.EOF = False Then
 '         rstMytemp.MoveNext
 '         End If
 '      Next I
 '    End If
 ' dbsMyTemp.Close
'add item to combo1

 
AddUpdate = "Add"

Combo2.AddItem "uncertain", 0
Combo2.AddItem "affinis", 1
Combo2.AddItem "confer", 2
Combo3.AddItem "ditch cuttings", 0
Combo3.AddItem "side-wall core", 1
Combo3.AddItem "core", 2
Combo3.AddItem "unknown/not assigned", 3
Combo3.AddItem "outcrop", 4
Combo3.AddItem "e-log", 5
Combo3.AddItem "acoustic (seismic)", 6
Combo3.AddItem "magnetochron", 7

Combo4.AddItem "good", 0
Combo4.AddItem "normal", 1
Combo4.AddItem "bad", 2

Combo7.AddItem "incidental", 0
'Combo7.AddItem "few specimens", 1
Combo7.AddItem "rare specimens", 1
Combo7.AddItem "common specimens", 2
Combo7.AddItem "frequent specimens", 3
Combo7.AddItem "abundant specimens", 4
Combo7.AddItem "dominant specimens", 5


End Sub



Public Function Check() As Boolean
If Trim(TxtDepth.Text) = "" Or IsNumeric(Trim(TxtDepth.Text)) = False Then
     MsgBox "Sample depth must be numeric"
     Check = False
ElseIf txtEventNumber.Text = "" Or IsNumeric(Trim(txtEventNumber.Text)) = False Then
      MsgBox "Event Number must be numeric", vbExclamation
     Check = False
ElseIf Trim(txtCount.Text) = "" Or IsNumeric(Trim(txtCount.Text)) = False Then
      MsgBox "Count must be numeric", vbExclamation
   Check = False
Else
   Check = True
End If
End Function
 
Private Sub txtFossil_Name_Change()
txtFossil_Name.Visible = True

End Sub

Private Sub txtTemp_Change()
txtEventNumber.Text = txtTemp.Text
End Sub
