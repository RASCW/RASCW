VERSION 5.00
Begin VB.Form frmCor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Parameters and Data Files to Run Cor"
   ClientHeight    =   4005
   ClientLeft      =   375
   ClientTop       =   1035
   ClientWidth     =   5970
   Icon            =   "frmCor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5970
   Begin VB.FileListBox File_list3 
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   2452
   End
   Begin VB.FileListBox File_list2 
      Height          =   285
      Left            =   480
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   2452
   End
   Begin VB.FileListBox File_List1 
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   2452
   End
   Begin VB.CommandButton cmdRevise 
      Caption         =   "Yes"
      Height          =   372
      Left            =   5040
      TabIndex        =   19
      Top             =   360
      Width           =   732
   End
   Begin VB.TextBox txt_Out 
      Height          =   372
      Left            =   480
      TabIndex        =   18
      Top             =   2880
      Width           =   2452
   End
   Begin VB.TextBox txt_dic 
      Height          =   372
      Left            =   480
      TabIndex        =   17
      Top             =   2040
      Width           =   2452
   End
   Begin VB.TextBox txt_Dat 
      Height          =   372
      Left            =   480
      TabIndex        =   16
      Top             =   1200
      Width           =   2452
   End
   Begin VB.TextBox txt_inp 
      Height          =   372
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   2452
   End
   Begin VB.FileListBox File_list4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   2452
   End
   Begin VB.CommandButton cmd_Dic 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   732
   End
   Begin VB.CommandButton cmdTemp 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2880
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton cmdApplyRasc 
      Caption         =   "&Apply"
      Height          =   372
      Left            =   5040
      TabIndex        =   7
      Top             =   1560
      Width           =   732
   End
   Begin VB.CommandButton CmdCancelRasc 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   5040
      TabIndex        =   6
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton cmdRascDat 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   732
   End
   Begin VB.ComboBox cboFileType 
      Height          =   288
      Left            =   2160
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdRascInp 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   732
   End
   Begin VB.Label lblRevise 
      Alignment       =   2  'Center
      Caption         =   "Revise Parameter File"
      Height          =   456
      Left            =   3840
      TabIndex        =   20
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label lblOut 
      Caption         =   "Output File"
      Height          =   372
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   972
   End
   Begin VB.Label lbldat 
      Caption         =   "Rasc Data File"
      Height          =   372
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1272
   End
   Begin VB.Label lblinp 
      Caption         =   "Parameter File"
      Height          =   372
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1272
   End
   Begin VB.Label lblhout 
      Caption         =   "Rasc Output File"
      Height          =   372
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1272
   End
End
Attribute VB_Name = "frmCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim gRascInput As RascInput
Dim gFileNum As Integer
Dim gRecordLen As Long

Private Sub cboFileType_Click()

Select Case cboFileType.ListIndex
 Case 0
    File_List1.Pattern = "*.Par"
 Case 1
    File_list2.Pattern = "*.Dat"
 Case 2
    File_list3.Pattern = "*h.Out"
 Case 3
    File_list4.Pattern = "*.out"
    
 End Select
  
End Sub



Private Sub CmdCancelRasc_Click()
Unload frmCor
End Sub

Private Sub cmdRevise_Click()
If txt_inp.Text <> "" Then
txt_file_inp = txt_inp.Text
Else
   MsgBox "You did not choose any files yet", vbOKOnly
 Exit Sub

End If
txt_file_hout = txt_dic.Text    'h.out file
  If txt_file_hout = "" Then
    MsgBox "h.Out input is empty", vbOKOnly
  Exit Sub
  
  End If
  
frmReviseCor.Show 1
End Sub

Private Sub cmdTemp_Click()
cboFileType.AddItem "Par files (*.Par)"
cboFileType.AddItem "Dat files (*.Dat)"
cboFileType.AddItem "h.Out files (*h.Out)"
cboFileType.AddItem "Out files (*.out)"
cboFileType.ListIndex = 3
File_list4.Visible = True
End Sub

Private Sub cmdApplyRasc_Click()
'Save the current record.
If SaveCurrentRecord = False Then Exit Sub
'close the file.
RunCorw
Unload frmCor

End Sub

Private Sub cmdCancel_Click()
End

End Sub

Private Sub cmdRascDat_Click()
cboFileType.AddItem "Par files (*.Par)"
cboFileType.AddItem "Dat files (*.Dat)"
cboFileType.AddItem "h.Out files (*h.Out)"
cboFileType.AddItem "Out files (*.out)"
cboFileType.ListIndex = 1
File_list2.Visible = True
End Sub

Private Sub cmd_Dic_Click()
cboFileType.AddItem "Par files (*.Par)"
cboFileType.AddItem "Dat files (*.Dat)"
cboFileType.AddItem "h.Out files (*h.Out)"
cboFileType.AddItem "Out files (*.   )"
cboFileType.ListIndex = 2
File_list3.Visible = True
End Sub

Private Sub cmdHelpRasc_Click()
MsgBox "To be added"

End Sub

Private Sub cmdRascInp_Click()
cboFileType.AddItem "Par files (*.Par)"
cboFileType.AddItem "Dat files (*.Dat)"
cboFileType.AddItem "h.Out files (*h.Out)"
cboFileType.AddItem "Out files (*.   )"
cboFileType.ListIndex = 0
File_List1.Refresh
File_List1.Visible = True
End Sub


Private Sub File_List1_Click()
txt_inp.Text = File_List1.Filename
File_List1.Visible = False
End Sub



Private Sub File_List2_Click()
txt_Dat.Text = File_list2.Filename
File_list2.Visible = False
End Sub




Private Sub File_List3_Click()
txt_dic.Text = File_list3.Filename
File_list3.Visible = False
End Sub



Private Sub File_List4_Click()
txt_Out.Text = File_list4.Filename
File_list4.Visible = False
End Sub


Private Sub File1_Click()

End Sub

Private Sub Form_Load()
txt_inp.Text = ""   '.par file
txt_Dat.Text = ""   '.dat file
txt_dic.Text = ""   'h.out file
txt_Out.Text = ""   'out file for cor

'File_List1.Path = App.Path
'File_list2.Path = App.Path
'File_list3.Path = App.Path
'File_list4.Path = App.Path

ShowCurrentRecord

End Sub

Private Function SaveCurrentRecord() As Boolean
Dim fileDat As String
Dim fileINp As String
Dim fileDic As String
Dim FileOut As String
Dim gFileNum As Integer

If Trim(txt_inp.Text) = "" Or _
Trim(txt_Dat.Text) = "" Or _
Trim(txt_dic.Text) = "" Or _
Trim(txt_Out.Text) = "" Then
MsgBox "Input and output boxes can not be empty"
SaveCurrentRecord = False
Exit Function
End If

'Fill gCascInput with the currently displayed data
gFileNum = FreeFile
Open CurDir & "\cortemp" For Output As gFileNum
fileINp = txt_inp.Text
fileDat = txt_Dat.Text
fileDic = txt_dic.Text
FileOut = txt_Out.Text

'Save gCascInput to the current record
Print #gFileNum, fileINp
Print #gFileNum, fileDic
Print #gFileNum, fileDat
Print #gFileNum, FileOut
Close gFileNum
SaveCurrentRecord = True
End Function

Private Sub ShowCurrentRecord()
Dim fileDat As String
Dim fileINp As String
Dim fileDic As String
Dim FileOut As String
Dim gFileNum As Integer
'Fill gCascInput with the data of the current record

gFileNum = FreeFile
If Dir("cortemp") = "" Then
    Open CurDir & "\Cortemp" For Output As gFileNum
    Close gFileNum
End If

Open CurDir & "\Cortemp" For Input As gFileNum
If Not EOF(gFileNum) Then
Input #gFileNum, fileINp, fileDic, fileDat, FileOut
'Display gCascInput.
txt_inp.Text = Trim(fileINp)
txt_Dat.Text = Trim(fileDat)
txt_dic.Text = Trim(fileDic)
txt_Out.Text = Trim(FileOut)
Else
txt_inp.Text = ""
txt_Dat.Text = ""
txt_dic.Text = ""
txt_Out.Text = ""
End If
Close gFileNum

End Sub

Public Sub RunCorw()
Dim X, Y
X = Shell(App.Path & "\corw.exe", vbHide)
Y = MsgBox(" Done!", 0, "")
End Sub

Public Sub txt_inp_Change()

End Sub
