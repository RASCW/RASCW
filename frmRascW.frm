VERSION 5.00
Begin VB.Form frmRascW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Parameter, Data and Output File Names to Run RASC"
   ClientHeight    =   4530
   ClientLeft      =   375
   ClientTop       =   1035
   ClientWidth     =   7650
   HelpContextID   =   200000001
   Icon            =   "frmRascW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7650
   Begin VB.CheckBox RascwCheck 
      Caption         =   "Use RASC Version 18"
      Height          =   405
      Left            =   4440
      TabIndex        =   16
      Top             =   2610
      Width           =   2385
   End
   Begin VB.FileListBox File_list2 
      Height          =   870
      Left            =   480
      TabIndex        =   10
      Top             =   2130
      Visible         =   0   'False
      Width           =   2452
   End
   Begin VB.FileListBox File_List1 
      Height          =   870
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmdRevise 
      Caption         =   "&Edit RASC Parameter File"
      Height          =   372
      Left            =   4260
      TabIndex        =   15
      Top             =   360
      Width           =   2355
   End
   Begin VB.TextBox txt_Out 
      Height          =   372
      Left            =   480
      TabIndex        =   14
      Top             =   3150
      Width           =   2452
   End
   Begin VB.TextBox txt_Dat 
      Height          =   372
      Left            =   480
      TabIndex        =   13
      Top             =   1770
      Width           =   2452
   End
   Begin VB.TextBox txt_inp 
      Height          =   372
      Left            =   480
      TabIndex        =   12
      Top             =   360
      Width           =   2452
   End
   Begin VB.FileListBox File_list3 
      Height          =   870
      Left            =   480
      TabIndex        =   11
      Top             =   3540
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2940
      TabIndex        =   9
      Top             =   3150
      Width           =   732
   End
   Begin VB.CommandButton cmdApplyRasc 
      Caption         =   "&Apply"
      Height          =   372
      Left            =   4440
      TabIndex        =   6
      Top             =   3690
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancelRasc 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   5910
      TabIndex        =   5
      Top             =   3690
      Width           =   1275
   End
   Begin VB.CommandButton cmdRascDat 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2940
      TabIndex        =   2
      Top             =   1770
      Width           =   732
   End
   Begin VB.ComboBox cboFileType 
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdRascInp 
      Caption         =   "Browse"
      Height          =   372
      Left            =   2940
      TabIndex        =   0
      Top             =   360
      Width           =   732
   End
   Begin VB.Label lblOut 
      Caption         =   "RASC Output File"
      Height          =   375
      Index           =   1
      Left            =   510
      TabIndex        =   7
      Top             =   2880
      Width           =   1485
   End
   Begin VB.Label lbldat 
      Caption         =   "RASC Data File"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1500
      Width           =   1605
   End
   Begin VB.Label lblInp 
      Caption         =   "RASC Parameter File"
      Height          =   375
      Left            =   510
      TabIndex        =   3
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "frmRascW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim gRascInput As RascInput
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim FileBrowseKey1 As Integer, FileBrowseKey2 As Integer, FileBrowseKey3 As Integer, RASCWRunSuccessKey As Integer

Private Sub cboFileType_Click()

Select Case cboFileType.ListIndex
 Case 0
    File_List1.Pattern = "*.Inp"
 Case 1
    File_list2.Pattern = "*.Dat"
 Case 2
    File_list3.Pattern = "*.Dic"
 Case 3
    File_list3.Pattern = "*.Out"
 
 End Select
  
End Sub

Private Sub CmdCancelRasc_Click()
    Unload frmRascW
End Sub


Private Sub cmdRevise_Click()
        
    File_List1.Visible = False
    File_list2.Visible = False
    File_list3.Visible = False
        
        If Trim(txt_inp.Text) <> "" Then
            txt_file_inp = Trim(txt_inp.Text) + ".inp"
        Else
           MsgBox "The RASC Parameter file name is empty, please Select or Give a file name and try again.", vbOKOnly
           Exit Sub
        End If
        '**************************
        ' To run RASC to generate a dic file to be used for revision
        If Trim(txt_inp.Text) = "" Or _
           Trim(txt_Dat.Text) = "" Or _
           Trim(txt_Out.Text) = "" Then
           MsgBox "Some of the input and output boxes are empty. Please give or select these filenames and try again."
           Exit Sub
        End If
        'The data file is necessary
        If Dir(CurDir + "\" + Trim(txt_Dat.Text) + ".Dat") = "" Then
           MsgBox "The Data file: " + CurDir + "\" + Trim(txt_Dat.Text) + ".Dat" + " does not exist. Please check it and try again."
           Exit Sub
        End If
        
        'Save previous record temporarily
        Dim gFileNum As Integer
        Dim I As Integer
           
        'Fill gCascInput with the data of the current record
        
        gFileNum = FreeFile
        If Dir("Rasctemp") = "" Then
            Open CurDir & "\Rasctemp" For Output As gFileNum
            Close gFileNum
        End If
        For I = 1 To 4
                 CurRASCParaFile(I) = ""
        Next I
        Open CurDir & "\Rasctemp" For Input As gFileNum
        If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, CurRASCParaFile(1)        'fileINp
        End If
        If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, CurRASCParaFile(2)        'fileDat
        End If
        If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, CurRASCParaFile(3)         'fileDic
        End If
        If Not (EOF(gFileNum)) Then
            Line Input #gFileNum, CurRASCParaFile(4)         'FileOut
        End If
        Close gFileNum
        
        'Save user selected existing INP file
        If Dir(CurDir + "\" + txt_file_inp + ".inp") <> "" Then
             CurRASCParaFile(1) = txt_file_inp
        End If


        SaveCurrentRecord
        RunRascw
        'MsgBox " Noisy data files may have Warnings under \Tables\Warnings!", 0, ""
        '******************************
        
        'It was required that DAT file and DIC file have the same main name.
        txt_file_dic = Trim(txt_Dat.Text) + ".dic"
        If txt_file_dic = "" Then
             MsgBox "Dic input is empty", vbOKOnly
             Exit Sub
        End If
        
        txt_file_RASCout = Trim(txt_Out.Text)
        
        If txt_file_inp <> "" Then
           frmRevise.Show 1
        End If
      ' txt_inp.Text = txt_file_inp
      ' txt_file_inp = ""
End Sub

Private Sub cmdOut_Click()
  cboFileType.ListIndex = 3
   File_list3.Path = CurrentDir
   File_list3.Pattern = "*.*"
    File_list3.Pattern = "*.Out"

   If File_list3.Visible = False Then
        FileBrowseKey3 = 0
    End If
   If FileBrowseKey3 = 0 Then
        File_List1.Visible = False
        File_list2.Visible = False
        File_list3.Visible = True
        FileBrowseKey3 = 1
    Else
       FileBrowseKey3 = 0
       File_list3.Visible = False
    End If
End Sub

Private Sub cmdApplyRasc_Click()
 
    File_List1.Visible = False
    File_list2.Visible = False
    File_list3.Visible = False
    
    'Save the current record.
    If Trim(txt_inp.Text) = "" Or _
       Trim(txt_Dat.Text) = "" Or _
       Trim(txt_Out.Text) = "" Then
       MsgBox "Some of the input and output boxes are empty. Please give or select these filenames and try again."
       frmRascW.SetFocus
       Exit Sub
    End If
    
    If Dir(CurDir + "\" + Trim(txt_Dat.Text) + ".Dat") = "" Then
       MsgBox "The Data file: " + CurDir + "\" + Trim(txt_Dat.Text) + ".Dat" + " does not exist. Please check it and try again."
       Exit Sub
    End If
    
    SaveCurrentRecord
    RunRascw
    If RascwCheck.Value = Checked Then
       RASCWVersionControlKey = 1
    Else
       RASCWVersionControlKey = 0
    End If
    
    If RASCWRunSuccessKey = 1 Then
        MsgBox "Current RASC job is done!" + Chr$(13) + "The results are saved in directory: " + CurDir, 0, ""
        Unload frmRascW
        RASCWRunSuccessKey = 0
    End If

End Sub

Private Sub cmdCancel_Click()
End

End Sub

Private Sub cmdRascDat_Click()
'    cboFileType.AddItem "Inp files (*.Inp)"
'    cboFileType.AddItem "Dat files (*.Dat)"
'    cboFileType.AddItem "Dic files (*.Dic)"
'    cboFileType.AddItem "Out files (*.out)"
    cboFileType.ListIndex = 1
    File_list2.Path = CurrentDir
    File_list2.Pattern = "*.*"
    File_list2.Pattern = "*.Dat"

    If File_list2.Visible = False Then
        FileBrowseKey2 = 0
    End If
    If FileBrowseKey2 = 0 Then
        File_List1.Visible = False
        File_list2.Visible = True
        File_list3.Visible = False
        FileBrowseKey2 = 1
    Else
       FileBrowseKey2 = 0
       File_list2.Visible = False
    End If
End Sub

Private Sub cmd_Dic_Click()
'    cboFileType.AddItem "Inp files (*.Inp)"
'    cboFileType.AddItem "Dat files (*.Dat)"
'    cboFileType.AddItem "Dic files (*.Dic)"
'    cboFileType.AddItem "Out files (*.out )"
'    cboFileType.ListIndex = 2
'    File_list3.Path = CurrentDir
'    File_list3.Visible = True
'    File_List1.Visible = False
'    File_list2.Visible = False
End Sub

Private Sub cmdHelpRasc_Click()
   MsgBox "To be added"

End Sub

Private Sub cmdRascInp_Click()
'    cboFileType.AddItem "Inp files (*.Inp)"
'    cboFileType.AddItem "Dat files (*.Dat)"
'    cboFileType.AddItem "Dic files (*.Dic)"
'    cboFileType.AddItem "Out files (*.out )"
    cboFileType.ListIndex = 0
    File_List1.Path = CurrentDir
    File_List1.Pattern = "*.*"
    File_List1.Pattern = "*.Inp"
    If File_List1.Visible = False Then
        FileBrowseKey1 = 0
    End If
    If FileBrowseKey1 = 0 Then
        File_List1.Visible = True
        File_list2.Visible = False
        File_list3.Visible = False
        FileBrowseKey1 = 1
    Else
       FileBrowseKey1 = 0
       File_List1.Visible = False
    End If

End Sub


Private Sub File_List1_Click()
    Dim filelen As Integer
    Dim tempName As String
    tempName = Trim(File_List1.Filename)
    filelen = Len(tempName)
    If filelen > 4 Then
        txt_inp.Text = Mid(tempName, 1, filelen - 4)
        File_List1.Visible = False
    Else
        Exit Sub
    End If

End Sub

Private Sub File_List1_LostFocus()
   File_List1.Visible = False
End Sub
Private Sub File_List2_LostFocus()
   File_list2.Visible = False
End Sub
Private Sub File_List3_LostFocus()
   File_list3.Visible = False
End Sub

Private Sub File_List2_Click()
    Dim filelen As Integer
    Dim tempName As String
    tempName = Trim(File_list2.Filename)
    filelen = Len(tempName)
    If filelen > 4 Then
        txt_Dat.Text = Mid(tempName, 1, filelen - 4)
        File_list2.Visible = False
    Else
        Exit Sub
    End If
End Sub

Private Sub File_List3_Click()
    Dim tempName As String
    Dim filelen As Integer
    
    tempName = Trim(File_list3.Filename)
    filelen = Len(tempName)
    If filelen > 4 Then
        txt_Out.Text = Mid(tempName, 1, filelen - 4)
    End If
    File_list3.Visible = False
End Sub


Private Sub Form_Activate()
        CurWindowNum = 27
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Load()

    
    txt_inp.Text = ""
    txt_Dat.Text = ""
    txt_Out.Text = ""
    
    cboFileType.AddItem "Inp files (*.Inp)"
    cboFileType.AddItem "Dat files (*.Dat)"
    cboFileType.AddItem "Dic files (*.Dic)"
    cboFileType.AddItem "Out files (*.out )"

    
    'Display the current record.
    ShowCurrentRecord
    FileBrowseKey1 = 0
    FileBrowseKey2 = 0
    FileBrowseKey3 = 0
    RASCWRunSuccessKey = 0
    
    If RASCWVersionControlKey = 1 Then
        RascwCheck.Value = Checked
    Else
        RascwCheck.Value = Unchecked
    End If
    
    CurWindowNum = 27
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

    
End Sub

Private Sub SaveCurrentRecord()
    Dim fileDat As String
    Dim fileINp As String
    Dim fileDic As String
    Dim FileOut As String
    Dim gFileNum As Integer
    
    'Fill gCascInput with the currently displayed data
    gFileNum = FreeFile
    Open CurDir & "\Rasctemp" For Output As #gFileNum
    fileINp = Trim(txt_inp.Text) + ".Inp"
    fileDat = Trim(txt_Dat.Text) + ".Dat"
    fileDic = Trim(txt_Dat.Text) + ".Dic"
    FileOut = Trim(txt_Out.Text)
    txt_file_RASCout = FileOut
    'Save gCascInput to the current record
    Print #gFileNum, fileINp
    Print #gFileNum, fileDat
    Print #gFileNum, fileDic
    Print #gFileNum, FileOut
    
    Close #gFileNum
End Sub

Private Sub ShowCurrentRecord()
    Dim fileDat As String
    Dim fileINp As String
    Dim fileDic As String
    Dim FileOut As String
    Dim gFileNum As Integer
    Dim filelen As Integer
    Dim tempName As String
    
    
    'Fill gCascInput with the data of the current record
    
    gFileNum = FreeFile
    If Dir("Rasctemp") = "" Then
        Open CurDir & "\Rasctemp" For Output As gFileNum
        Close gFileNum
    End If
    Open CurDir & "\Rasctemp" For Input As gFileNum
    If Not (EOF(gFileNum)) Then
        Line Input #gFileNum, fileINp
    End If
    If Not (EOF(gFileNum)) Then
        Line Input #gFileNum, fileDat
    End If
    If Not (EOF(gFileNum)) Then
        Line Input #gFileNum, fileDic
    End If
    If Not (EOF(gFileNum)) Then
        Line Input #gFileNum, FileOut
    End If
   
   'Display gCascInput.
    tempName = Trim(fileINp)
    filelen = Len(tempName)
    If filelen > 4 And InStr(1, fileINp, ".") > 0 Then
         txt_inp.Text = Mid(tempName, 1, filelen - 4)
    End If
    
    ' txt_inp.Text = Trim(fileINp)
    
    tempName = Trim(fileDat)
    filelen = Len(tempName)
    If filelen > 4 And InStr(1, fileDat, ".") > 0 Then
        txt_Dat.Text = Mid(tempName, 1, filelen - 4)
    End If
    'txt_Dat.Text = Trim(fileDat)
    'txt_Dic.Text = Trim(fileDic)
    
    tempName = Trim(FileOut)
     
    txt_Out.Text = tempName
     
    
    'txt_Out.Text = Trim(FileOut)
    
    Close gFileNum

End Sub

Public Sub RunRascw()
    Dim X, Y
    If RascwCheck.Value = Unchecked Then    'Run Version 20
        If Dir(App.Path & "\rascw.exe") = "" Then
           MsgBox App.Path & "\rascw.exe" + " was not found. Please check it and try again. "
           Exit Sub
        End If
        X = Shell(App.Path & "\rascw.exe", vbHide)
    Else                                                                    'Run Version 18
         If Dir(App.Path & "\rascwO.exe") = "" Then
           MsgBox App.Path & "\rascwO.exe" + " was not found. Please check it and try again. "
           Exit Sub
        End If
       X = Shell(App.Path & "\rascwO.exe", vbHide)
    End If
    RASCWRunSuccessKey = 1
    'Y = MsgBox(" Noisy data files may have Warnings under \Tables\Warnings!", 0, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CurWindowNum = 27
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub
