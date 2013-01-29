VERSION 5.00
Begin VB.Form frmCascinput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter File Names to Run CASC  "
   ClientHeight    =   4515
   ClientLeft      =   3225
   ClientTop       =   2040
   ClientWidth     =   6900
   HelpContextID   =   300000001
   Icon            =   "frmCascinput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6900
   Begin VB.CommandButton cmdParaEdit 
      Caption         =   "&Edit CASC Parameter File"
      Height          =   375
      Left            =   3660
      TabIndex        =   15
      Top             =   1740
      Width           =   2325
   End
   Begin VB.FileListBox FileList3 
      Height          =   870
      Left            =   240
      TabIndex        =   14
      Top             =   2100
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtPara 
      Height          =   345
      Left            =   240
      TabIndex        =   12
      Top             =   1740
      Width           =   2415
   End
   Begin VB.CommandButton cmdCASCParameter 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1740
      Width           =   765
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2640
      TabIndex        =   10
      Top             =   3120
      Width           =   765
   End
   Begin VB.FileListBox FileList2 
      Height          =   870
      Left            =   210
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.ComboBox cboFileTypeDat 
      Height          =   315
      ItemData        =   "frmCascinput.frx":0442
      Left            =   2790
      List            =   "frmCascinput.frx":0444
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   3810
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdDat 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2670
      TabIndex        =   7
      Top             =   390
      Width           =   732
   End
   Begin VB.FileListBox FileList1 
      Height          =   870
      Left            =   270
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.TextBox txtOut 
      Height          =   372
      Left            =   210
      TabIndex        =   5
      Top             =   3120
      Width           =   2430
   End
   Begin VB.TextBox txtDat 
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   390
      Width           =   2430
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   5490
      TabIndex        =   1
      Top             =   3810
      Width           =   1125
   End
   Begin VB.CommandButton cmdApplyCasc 
      Caption         =   "&Apply"
      Height          =   372
      Left            =   4200
      TabIndex        =   0
      Top             =   3810
      Width           =   1155
   End
   Begin VB.Label LabelParameter 
      Caption         =   "CASC Parameter File"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1530
      Width           =   2385
   End
   Begin VB.Label lblOut 
      Caption         =   "RASC Output File"
      Height          =   375
      Left            =   210
      TabIndex        =   3
      Top             =   2880
      Width           =   2430
   End
   Begin VB.Label lbldat 
      Caption         =   "RASC Data File"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   2415
   End
End
Attribute VB_Name = "frmCascinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables must be declared
Option Explicit
'Declare the variables that should be visible in all
'the procedures of the form
'Dim gCascInput As CascInput
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim FileBrowseKey1 As Integer, FileBrowseKey2 As Integer, FileBrowseKey3 As Integer, CASCWRunSuccessKey As Integer


Private Sub cboFileTypeDat_Click()
        Select Case cboFileTypeDat.ListIndex
         Case 0
            FileList1.Pattern = "*.Dat"
         Case 1
            FileList2.Pattern = "*i.Out"
         Case 2
            FileList3.Pattern = "*.Inc"
         End Select
End Sub

Private Sub cmdApplyCasc_Click()
        'Save the current record.
        If Trim(txtDat.Text) = "" Or _
           Trim(txtOut.Text) = "" Or Trim(TxtPara.Text = "") Then
           MsgBox "Some of the input and output file boxes are empty. Please give or select these filenames and try again."
           Exit Sub
        End If
        SaveCurrentRecord
        
        'Close the file.
        
        RunCascw
        If CASCWRunSuccessKey = 1 Then
             Unload frmCascinput
             CASCWRunSuccessKey = 0
        End If

End Sub


Private Sub cmdCancel_Click()
    Unload frmCascinput
End Sub

Private Sub cmdCASCParameter_Click()
    RefleshDir
    cboFileTypeDat.ListIndex = 2
    FileList3.Path = CurrentDir
    
    FileList3.Pattern = "*.*"
    FileList3.Pattern = "*.Inc"

   If FileList3.Visible = False Then
        FileBrowseKey3 = 0
    End If
   If FileBrowseKey3 = 0 Then
        FileList1.Visible = False
        FileList2.Visible = False
        FileList3.Visible = True
        FileBrowseKey3 = 1
    Else
       FileBrowseKey3 = 0
       FileList3.Visible = False
    End If

End Sub

Private Sub cmdDat_Click()
    RefleshDir
    cboFileTypeDat.ListIndex = 0
    FileList1.Path = CurrentDir
    'Active filelist flesh
    FileList1.Pattern = "*.*"
    FileList1.Pattern = "*.Dat"

    If FileList1.Visible = False Then
        FileBrowseKey1 = 0
    End If
    If FileBrowseKey1 = 0 Then
        FileList1.Visible = True
        FileList2.Visible = False
        FileList3.Visible = False
        FileBrowseKey1 = 1
    Else
       FileBrowseKey1 = 0
       FileList1.Visible = False
    End If
End Sub

Private Sub cmdDep_Click()
    RefleshDir
    FileList1.Visible = False
    FileList3.Visible = False
'        cboFileTypeDat.AddItem "Dat files (*.Dat)"
'        cboFileTypeDat.AddItem "i.out files (*i.out)"
'        cboFileTypeDat.AddItem "Inc files (*.Inc)"
    cboFileTypeDat.ListIndex = 1
    FileList2.Path = CurrentDir
    FileList2.Visible = True
End Sub

Private Sub cmdHelpCasc_Click()
  MsgBox "To be added"

End Sub



Private Sub Command1_Click()

End Sub

Private Sub cmdOut_Click()
    RefleshDir
    cboFileTypeDat.ListIndex = 1
    FileList2.Path = CurrentDir
        
    FileList2.Pattern = "*.*"
    FileList2.Pattern = "*i.Out"
        
    If FileList2.Visible = False Then
        FileBrowseKey2 = 0
    End If
    If FileBrowseKey2 = 0 Then
        FileList1.Visible = False
        FileList2.Visible = True
        FileList3.Visible = False
        FileBrowseKey2 = 1
    Else
       FileBrowseKey2 = 0
       FileList2.Visible = False
    End If
End Sub

Private Sub cmdParaEdit_Click()
Dim Y
   If Trim(TxtPara.Text) = "" Then
       Y = MsgBox("The CASC Parameter file name is empty, please Select or Give a file name and try again.", 0, "")
       Exit Sub
   End If
   txt_file_inc = Trim(TxtPara.Text)
   FileList1.Visible = False
   FileList2.Visible = False
   FileList3.Visible = False
    frmCASCParameter.Show 1
End Sub

Private Sub FileList1_Click()
Dim filelen As Integer
Dim tempName As String

tempName = Trim(FileList1.Filename)
filelen = Len(tempName)
If filelen > 4 Then
    txtDat.Text = Mid(tempName, 1, filelen - 4)
End If
'txtDat.Text = FileList1.Filename
FileList1.Visible = False
End Sub

Private Sub FileList2_Click()
Dim filelen As Integer
Dim tempName As String
tempName = Trim(FileList2.Filename)
filelen = Len(tempName)
If filelen > 4 Then
    txtOut.Text = Mid(tempName, 1, filelen - 4)
End If
'txtDep.Text = FileList2.Filename
FileList2.Visible = False
End Sub

 

Private Sub FileList3_Click()
    Dim filelen As Integer
    Dim tempName As String
    tempName = Trim(FileList3.Filename)
    filelen = Len(tempName)
    If filelen > 4 Then
        TxtPara.Text = Mid(tempName, 1, filelen - 4)
        CurCASCParaFile = TxtPara.Text
    End If
    'txtDep.Text = FileList2.Filename
    FileList3.Visible = False
    
End Sub

Private Sub Form_Activate()
        CurWindowNum = 28
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Load()
  'FileList1.Path = App.Path
  'FileList2.Path = App.Path
    txtDat.Text = ""
    txtOut.Text = ""
    TxtPara.Text = ""
    CurCASCParaFile = ""
    
    cboFileTypeDat.AddItem "Dat files (*.Dat)"
    cboFileTypeDat.AddItem "i.out files (*i.out)"
    cboFileTypeDat.AddItem "Inc files (*.Inc)"

'Calculate the length of a record.
'gRecordLen = Len(gCascInput)

'Get the next available file number
'gFileNum = FreeFile

'Open the file for random-access. If the file
'does not exist then it is created.

'Open "casctemp" For Random As gFileNum Len = gRecordLen

'Display the current record.
    ShowCurrentRecord
    FileBrowseKey1 = 0
    FileBrowseKey2 = 0
    FileBrowseKey3 = 0
    CASCWRunSuccessKey = 0
    
    CurWindowNum = 28
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd


End Sub

Public Sub SaveCurrentRecord()
        Dim InputDat As String
        Dim InputDep As String
        Dim InputOut As String
        Dim InputInc As String
        Dim gFileNum As Integer
        Dim Temp As String
        Dim FileNameLen As Integer
        'Fill gCascInput with the currently displayed data
        gFileNum = FreeFile
        Open CurDir & "\casctemp" For Output As gFileNum
        
        InputDat = Trim(txtDat.Text) + ".Dat"
        InputDep = Trim(txtDat.Text) + ".Dep"
        InputOut = Trim(txtOut.Text) + ".Out"
        InputInc = Trim(TxtPara.Text) + ".Inc"
        
        'Save gCascInput to the current record
        Print #gFileNum, InputDat
        Print #gFileNum, InputDep
        Print #gFileNum, InputOut
        FileNameLen = Len(Trim(InputOut))
        If FileNameLen <= 5 Then
             Close gFileNum
             Exit Sub
        End If
        Temp = Mid(InputOut, 1, FileNameLen - 5)
        Print #gFileNum, Temp + ".sc1"
        Print #gFileNum, Temp + ".sc2"
        Print #gFileNum, InputInc
        Close gFileNum

End Sub

Public Sub ShowCurrentRecord()
        Dim InputDat As String
        Dim InputDep As String
        Dim InputOut As String
        Dim InputOut1 As String
        Dim InputOut2 As String
        Dim InputInc As String
        Dim gFileNum As Integer
        Dim filelen As Integer
        Dim tempName As String
        
        'Fill gCascInput with the data of the current record
        gFileNum = FreeFile
        
        If Dir("casctemp") = "" Then
            Open CurDir & "\casctemp" For Output As gFileNum
            Close gFileNum
        End If
        Open CurDir & "\casctemp" For Input As gFileNum
        If Not (EOF(gFileNum)) Then
            Input #gFileNum, InputDat
        End If
        If Not (EOF(gFileNum)) Then
            Input #gFileNum, InputDep
        End If
        If Not (EOF(gFileNum)) Then
            Input #gFileNum, InputOut
        End If
       'Read *.Inc parameterfile and be compliant with old version
        If Not (EOF(gFileNum)) Then
            Input #gFileNum, InputOut1
        End If
        If Not (EOF(gFileNum)) Then
            Input #gFileNum, InputOut2
        End If
        If Not (EOF(gFileNum)) Then
            Input #gFileNum, InputInc
        End If
        
        Close gFileNum
        
        tempName = Trim(InputDat)
        filelen = Len(tempName)
        If filelen > 4 Then
            txtDat.Text = Mid(tempName, 1, filelen - 4)
        End If
        'Display gCascInput.
        'txtDat.Text = Trim(InputDat)
        'txtDep.Text = Trim(InputDep)
        tempName = Trim(InputOut)
        filelen = Len(tempName)
        If filelen > 4 Then
            txtOut.Text = Mid(tempName, 1, filelen - 4)
        End If
        tempName = Trim(InputInc)
        filelen = Len(tempName)
        If filelen > 4 Then
           TxtPara.Text = Mid(tempName, 1, filelen - 4)
            CurCASCParaFile = TxtPara.Text
        End If
        
        
        'txtOut.Text = Trim(InputOut)
        
        
        
End Sub

Public Sub RunCascw()
    Dim X, Y
    
    If Dir(App.Path & "\cascw.exe") = "" Then
        MsgBox App.Path & "\cascw.exe" + " was not found. Please check it and try again. "
        Exit Sub
    End If
     X = Shell(App.Path & "\cascw.exe", vbHide)
     Y = MsgBox("Current CASC job is done!" + Chr$(13) + "The results are saved in directory: " + CurDir, 0, "")
     CASCWRunSuccessKey = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CurWindowNum = 28
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub

Private Sub TxtPara_Change()
   ' CurCASCParaFile = TxtPara.Text
End Sub
