VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DbList32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMakedatEvent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fossil Dictionary Editor"
   ClientHeight    =   6840
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmMakedatEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9960
   Visible         =   0   'False
   Begin VB.CommandButton CmdSort 
      Caption         =   "Sort"
      Height          =   252
      Left            =   8520
      TabIndex        =   34
      Top             =   1920
      Width           =   972
   End
   Begin VB.TextBox txtEventType 
      DataField       =   "Event_Type"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   3240
      TabIndex        =   32
      Text            =   " "
      Top             =   480
      Width           =   2052
   End
   Begin VB.ComboBox cboEventType 
      Height          =   288
      Left            =   3240
      TabIndex        =   31
      Text            =   "Combo1"
      Top             =   480
      Width           =   2292
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmMakedatEvent.frx":0442
      Height          =   3972
      Left            =   480
      OleObjectBlob   =   "frmMakedatEvent.frx":0456
      TabIndex        =   30
      Top             =   2280
      Width           =   9012
   End
   Begin VB.TextBox txtFossilNumber 
      DataField       =   "Event_Number"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   " "
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox txtImage 
      DataField       =   "Image"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   7080
      TabIndex        =   24
      Text            =   " "
      Top             =   1920
      Width           =   732
   End
   Begin VB.ComboBox cboImage 
      Height          =   288
      Left            =   7080
      TabIndex        =   27
      Text            =   "Combo2"
      Top             =   1920
      Width           =   972
   End
   Begin VB.TextBox txtSynonyms 
      DataField       =   "Synonyms"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   5760
      TabIndex        =   13
      Text            =   " "
      Top             =   1920
      Width           =   732
   End
   Begin VB.ComboBox cboSynonyms 
      Height          =   288
      Left            =   5760
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   1920
      Width           =   972
   End
   Begin VB.TextBox txtReference 
      DataField       =   "Reference"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   6720
      TabIndex        =   22
      Text            =   " "
      Top             =   1200
      Width           =   1332
   End
   Begin VB.TextBox txtClassOfFossil 
      DataField       =   "Class_of_Fossils"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   480
      TabIndex        =   21
      Text            =   " "
      Top             =   1200
      Width           =   2292
   End
   Begin VB.TextBox txtFossilNames 
      DataField       =   "Fossil_Name"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   480
      TabIndex        =   20
      Text            =   " "
      Top             =   480
      Width           =   2292
   End
   Begin MSDBCtls.DBCombo DBCboFossilNames 
      DataField       =   "EventName"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   480
      TabIndex        =   19
      Top             =   480
      Width           =   2532
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DBCombo1"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "  "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Event_name"
      Top             =   120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   252
      Left            =   8520
      TabIndex        =   18
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   8520
      TabIndex        =   17
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   252
      Left            =   8520
      TabIndex        =   16
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   252
      Left            =   8520
      TabIndex        =   15
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox txtAge 
      DataField       =   "Age"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   3240
      TabIndex        =   11
      Text            =   " "
      Top             =   1920
      Width           =   2292
   End
   Begin VB.TextBox txtBiozone 
      DataField       =   "Biozone"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   480
      TabIndex        =   9
      Text            =   " "
      Top             =   1920
      Width           =   2532
   End
   Begin VB.TextBox txtLocality 
      DataField       =   "Type_Locality"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   3240
      TabIndex        =   7
      Text            =   " "
      Top             =   1200
      Width           =   2292
   End
   Begin VB.ComboBox cboClassOfFossil 
      Height          =   288
      Left            =   480
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2532
   End
   Begin VB.TextBox txtYear 
      DataField       =   "Year"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   5760
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox txtAuthor 
      DataField       =   "Authors"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   6720
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label Label11 
      Caption         =   "Event Type"
      Height          =   252
      Left            =   3240
      TabIndex        =   33
      Top             =   240
      Width           =   972
   End
   Begin VB.Label lblNumber 
      Caption         =   "Event  #"
      Height          =   252
      Left            =   5760
      TabIndex        =   29
      Top             =   240
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "Image"
      Height          =   252
      Left            =   7080
      TabIndex        =   25
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Reference"
      Height          =   252
      Left            =   6720
      TabIndex        =   23
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label10 
      Caption         =   "Synonyms"
      Height          =   252
      Left            =   5760
      TabIndex        =   14
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Label Label8 
      Caption         =   "Age"
      Height          =   252
      Left            =   3240
      TabIndex        =   12
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "Biozone"
      Height          =   252
      Left            =   480
      TabIndex        =   10
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label6 
      Caption         =   "Type Locality"
      Height          =   252
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "Class of Fossils"
      Height          =   252
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   2292
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      Height          =   252
      Left            =   5760
      TabIndex        =   4
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Author(s)"
      Height          =   252
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label lblFossilNames 
      Caption         =   " Fossil Name"
      Height          =   252
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1932
   End
End
Attribute VB_Name = "frmMakedatEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddUpdate As String
Dim sPreviousRecord As String



Private Sub cboClassOfFossil_Click()
Dim Fossil As String
'MsgBox "Frits, this is test#10 "
Fossil = cboClassOfFossil.Text
Select Case Fossil
  Case "Planktonic Foraminifers"
      txtClassOfFossil.Text = "PF"
  Case "Calcareous Benthic Foraminifers"
       txtClassOfFossil.Text = "CF"
Case "Agglutinated Benthic Foraminifers"
      txtClassOfFossil.Text = "AF"
Case "Nannofossils (Coccoliths)"
     txtClassOfFossil.Text = "NA"
Case "Dinoflagellate Cysts"
    txtClassOfFossil.Text = "DC"
Case "Spores/Pollen"
    txtClassOfFossil.Text = "SP"
Case "Radiolarians"
    txtClassOfFossil.Text = "RA"
Case "Diatoms"
    txtClassOfFossil.Text = "DI"
Case "Silicoflagellates"
    txtClassOfFossil.Text = "SI"
Case "Bolboforma"
    txtClassOfFossil.Text = "BO"
Case "Conodonts"
    txtClassOfFossil.Text = "CO"
Case "Graptolites"
    txtClassOfFossil.Text = "GR"
Case "Ostracods"
    txtClassOfFossil.Text = "OS"
Case "Ammonites"
    txtClassOfFossil.Text = "AM"
Case "Belemnites"
    txtClassOfFossil.Text = "BE"
Case "Pelecypods"
    txtClassOfFossil.Text = "PE"
Case "Brachiopods"
    txtClassOfFossil.Text = "BA"
Case "Crinoids"
    txtClassOfFossil.Text = "CR"
Case "Echinoids"
    txtClassOfFossil.Text = "EC"
Case "Sponges"
    txtClassOfFossil.Text = "SP"
Case "Magnetochrons"
    txtClassOfFossil.Text = "MA"
Case "E-log Events", 21
    txtClassOfFossil.Text = "EL"
Case "Seismic Events", 22
    txtClassOfFossil.Text = "SE"
Case "Lithologic Events", 23
    txtClassOfFossil.Text = "LI"
Case "Miscellaneous", 24
    txtClassOfFossil.Text = "MI"
Case Else
    MsgBox "Select one class from browser"
    Exit Sub
End Select
End Sub

Private Sub cboEventType_Click()
Dim gEventType As String
    gEventType = cboEventType.Text
    Select Case gEventType
        Case "Last occurrence"
            txtEventType.Text = "LO"
        Case "Last consistent and/or last common occurrence"
            txtEventType.Text = "LCO"
        Case "Last abundant occurrence"
            txtEventType.Text = "LAO"
        Case "Acme (peak occurrence)"
            txtEventType.Text = "AC"
        Case "First abundant occurrence"
            txtEventType.Text = "FAO"
        Case "First consistent and/or first common occurrence"
            txtEventType.Text = "FCO"
        Case "First occurrence"
            txtEventType.Text = "FO"
        Case Else
             MsgBox "Select a type from broswer!", vbExclamation
     End Select
End Sub

Private Sub cboImage_Click()
txtImage.Text = Left$(cboImage.Text, 1)
End Sub

Private Sub cboSynonyms_Click()
txtSynonyms.Text = Left$(cboSynonyms.Text, 1)

End Sub

Private Sub Command1_Click()
'CommonDialog1.ShowColor
End Sub

Private Sub cmdAdd_Click()
Dim result
Dim test
Dim Currentpos As Integer
Dim rstEvent_Name As Recordset
    If AddUpdate = "Add" Then
       result = MsgBox("Do you want to add a new fossil? ", vbYesNo, "Confirm Adding")
       If result = vbNo Then Exit Sub
       Currentpos = Data1.Recordset.RecordCount
       
       Data1.Recordset.AddNew
             
       txtFossilNames.Text = ""
       txtAuthor.Text = ""
       txtYear.Text = "0"
       txtLocality.Text = ""
       txtBiozone.Text = ""
       txtClassOfFossil.Text = ""
       txtReference.Text = ""
       txtAge.Text = ""
       txtSynonyms.Text = ""
       txtImage.Text = ""
       txtEventType.Text = ""
       txtFossilNumber.Text = Str(Currentpos + 1)
       cmdEdit.Visible = False
       cmdDelete.Visible = False
       cmdSort.Visible = False
       AddUpdate = "Update"
       cmdAdd.Caption = AddUpdate
            
    ElseIf AddUpdate = "Update" Then
      If Check() = True Then
       
       Data1.Recordset.Update
      Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    
           
    '  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
         Data1.Refresh
      AddUpdate = "Add"
      cmdAdd.Caption = AddUpdate
      
      cmdEdit.Visible = True
      cmdDelete.Visible = True
      cmdSort.Visible = True
       
      Else
        Exit Sub
      End If
      
    End If
    
End Sub

'Private Sub Command4_Click()
'Unload frmMakedatEvent
'End Sub

Private Sub cmdCancel_Click()
myDBName = ""
Unload frmMakedatEvent

End Sub

Private Sub cmdDelete_Click()
Dim result
Dim response
Dim Currentpos As Integer
'On Error GoTo ErrorHandler
   response = MsgBox("Do you want to change the current name to dummy?", vbOKCancel, "")
   If response = vbCancel Then
   Exit Sub
   Else
        Currentpos = Data1.Recordset.AbsolutePosition
        
        Data1.Refresh
        Data1.Recordset.Move (Currentpos)
        txtFossilNames.Text = "dummy"
        Data1.Recordset.Edit
        Data1.Recordset.Update
          Data1.Refresh
   End If
  

  '  result = MsgBox("Do you want to change the current name to dummy? ", vbYesNo, "Confirm Deleting")
  '  If result = vbNo Then Exit Sub
 '  Data1.Recordset.Edit
  '  Data1.Recordset.Fields(0).Value = "Dummy"
 '  Data1.Recordset.Delete
  '  Data1.Refresh
End Sub

Private Sub cmdEdit_Click()
Dim response
'On Error GoTo ErrorHandler
   response = MsgBox("Do you want to modify the current record?", vbOKCancel, "")
   If response = vbCancel Then
   Exit Sub
   Else
        If Check() = True Then
        Data1.Recordset.Edit
        Data1.Recordset.Update
          Data1.Refresh
        Else
         Exit Sub
        End If
   End If
'ErrorHandler:
'  MsgBox "Use add button!", vbExclamation
 
End Sub

Private Sub cmdPrint_Click()

End Sub



Private Sub cmdSearch_Click()

End Sub

Private Sub cmdSort_Click()
Dim response
   response = MsgBox("Do you want to sort the data?", vbOKCancel, "")
   If response = vbCancel Then
   Exit Sub
   Else
        frmSorting.Show 1
         If sort_method = "" Then Exit Sub
          Data1.RecordSource = "select * from Event_name order by " + sort_field + " " + sort_method + ";"
          
          Data1.Refresh
    End If
End Sub

Private Sub Form_Load()
 

Set DBsSAGA = OpenDatabase(myDBName)

 DBsSAGA.Close
 

Data1.DatabaseName = myDBName
 
 
'Data1.DatabaseName = myDBName
    
AddUpdate = "Add"
cboClassOfFossil.AddItem "Planktonic Foraminifers", 0
cboClassOfFossil.AddItem "Calcareous Benthic Foraminifers", 1
cboClassOfFossil.AddItem "Agglutinated Benthic Foraminifers", 2
cboClassOfFossil.AddItem "Nannofossils (Coccoliths)", 3
cboClassOfFossil.AddItem "Dinoflagellate Cysts", 4
cboClassOfFossil.AddItem "Spores/Pollen", 5
cboClassOfFossil.AddItem "Radiolarians", 6
cboClassOfFossil.AddItem "Diatoms", 7
cboClassOfFossil.AddItem "Silicoflagellates", 8
cboClassOfFossil.AddItem "Bolboforma", 9
cboClassOfFossil.AddItem "Conodonts", 10
cboClassOfFossil.AddItem "Graptolites", 11
cboClassOfFossil.AddItem "Ostracods", 12
cboClassOfFossil.AddItem "Ammonites", 13
cboClassOfFossil.AddItem "Belemnites", 14
cboClassOfFossil.AddItem "Pelecypods", 15
cboClassOfFossil.AddItem "Brachiopods", 16
cboClassOfFossil.AddItem "Crinoids", 17
cboClassOfFossil.AddItem "Echinoids", 18
cboClassOfFossil.AddItem "Sponges", 19
cboClassOfFossil.AddItem "Magnetochrons", 20
cboClassOfFossil.AddItem "E-log Events", 21
cboClassOfFossil.AddItem "Seismic Events", 22
cboClassOfFossil.AddItem "Lithologic Events", 23
cboClassOfFossil.AddItem "Miscellaneous", 24
 
 cboSynonyms.AddItem "yes", 0
 cboSynonyms.AddItem "no", 1
 cboImage.AddItem "yes", 0
 cboImage.AddItem "no", 1
 
cboEventType.AddItem "Last occurrence", 0
cboEventType.AddItem "Last consistent and/or last common occurrence", 1
cboEventType.AddItem "Last abundant occurrence", 2
cboEventType.AddItem "Acme (peak occurrence)", 3
cboEventType.AddItem "First abundant occurrence", 4
cboEventType.AddItem "First consistent and/or first common occurrence", 5
cboEventType.AddItem "First occurrence", 6
 
 
 
End Sub

Public Function Check() As Boolean
 
 ' On Error GoTo ErrorHandler
   'If Trim(txtFossilNames.Text) = "" Or _
       Trim(txtAuthor.Text) = "" Or
   
   '    Trim(txtLocality.Text) = "" Or _
    '   Trim(txtBiozone.Text) = "" Or _
    '   Trim(txtClassOfFossil.Text) = "" Or _
    '   Trim(txtReference.Text) = "" Or _
    '   Trim(txtAge.Text) = "" Or _
    '   Trim(txtSynonyms.Text) = "" Or _
    '   Trim(txtImage.Text) = "" Or
     If Trim(txtYear.Text) = "" Or _
        IsNumeric(txtYear.Text) = False Or _
        Trim(txtFossilNumber.Text) = "" Or _
        IsNumeric(txtFossilNumber.Text) = False Then
        MsgBox "Year and Fossil number can not be left empty or non numeric!", vbExclamation
       Check = False
    Else
       Check = True
   End If
'ErrorHandler:
 '    check = False
   
End Function

