VERSION 5.00
Begin VB.Form frmCASCParameter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASC Parameters"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "frmCASCParameter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check_ScaledOptimumSeq 
      Caption         =   "Input Sequence as in Scaled Optimum Sequence"
      Height          =   405
      Left            =   4140
      TabIndex        =   19
      Top             =   2580
      Width           =   2805
   End
   Begin VB.CheckBox Check_HigherOrderOutput 
      Caption         =   "Higher order window output"
      Height          =   435
      Left            =   420
      TabIndex        =   18
      Top             =   4650
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox Text_OrderofSQRT 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   2850
      TabIndex        =   17
      Top             =   1980
      Width           =   975
   End
   Begin VB.ComboBox cboFileTypeDat 
      Height          =   315
      Left            =   2850
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdCASCParameter 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2790
      TabIndex        =   13
      Top             =   2730
      Width           =   765
   End
   Begin VB.TextBox TxtPara 
      Height          =   345
      Left            =   360
      TabIndex        =   12
      Top             =   2730
      Width           =   2415
   End
   Begin VB.FileListBox FileList3 
      Height          =   675
      Left            =   360
      TabIndex        =   11
      Top             =   3090
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5610
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox Check_SquareRoot 
      Caption         =   $"frmCASCParameter.frx":0442
      Height          =   375
      Left            =   4140
      TabIndex        =   8
      Top             =   1980
      Width           =   2805
   End
   Begin VB.CheckBox Check_MeanDifference 
      Caption         =   "Mean difference correction"
      Height          =   345
      Left            =   390
      TabIndex        =   7
      Top             =   5100
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CheckBox Check_RemoveLarge 
      Caption         =   "Remove large depth differences"
      Height          =   345
      Left            =   420
      TabIndex        =   6
      Top             =   4260
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox Text_MaxiSequence 
      Height          =   375
      Left            =   2850
      TabIndex        =   5
      Top             =   1380
      Width           =   975
   End
   Begin VB.TextBox Text_MaxiOrder 
      Height          =   375
      Left            =   2850
      TabIndex        =   4
      Top             =   810
      Width           =   975
   End
   Begin VB.TextBox Text_Width 
      Height          =   345
      Left            =   2850
      TabIndex        =   3
      Top             =   210
      Width           =   975
   End
   Begin VB.Label Label_MaxiSequence2 
      Caption         =   $"frmCASCParameter.frx":0462
      Height          =   285
      Left            =   360
      TabIndex        =   21
      Top             =   1530
      Width           =   2445
   End
   Begin VB.Label Label_Width2 
      Caption         =   "(No Square Root Transformation)"
      Height          =   315
      Left            =   360
      TabIndex        =   20
      Top             =   390
      Width           =   2445
   End
   Begin VB.Label Label_OrderofSQRT 
      Caption         =   "Order of Square Root Display"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   1980
      Width           =   2445
   End
   Begin VB.Label LabelParameter 
      Caption         =   "Parameter File Save as :"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   2385
   End
   Begin VB.Label Label_MaxiSequence 
      Caption         =   "Maximum Difference between"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2445
   End
   Begin VB.Label Label_MaxiOrder 
      Caption         =   "Order (>1) of Window"
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2445
   End
   Begin VB.Label Label_Width 
      Caption         =   "Width of Depth Filter"
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   2445
   End
End
Attribute VB_Name = "frmCASCParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormStatus As Integer
Dim kcrito As Integer, icnt As Integer, imean As Integer, isqrt As Integer, iendo As Integer, kcrit As Integer, irascs As Integer, io As Integer
Dim decrito As Double

Private Sub Check_HigherOrderOutput_Click()
    If Check_HigherOrderOutput.Value = 0 Then
         Text_MaxiOrder.BackColor = &H8000000B
    Else
         Text_MaxiOrder.BackColor = &H80000009
    End If
End Sub

Private Sub Check_SquareRoot_Click()
    If Check_SquareRoot.Value = Unchecked Then
         Text_OrderofSQRT.BackColor = &H8000000B
    Else
         Text_OrderofSQRT.BackColor = &H80000009
'         If Val(Text_OrderofSQRT.Text) <= 0 Then
'              Text_OrderofSQRT.Text = "1"
'         End If
    End If
End Sub

Private Sub cmdCancel_Click()
    'Don't save any change and recover former name
    If frmCascinput.TxtPara.Text <> CurCASCParaFile Then
        frmCascinput.TxtPara.Text = CurCASCParaFile
    End If
    Me.Hide
    Unload frmCASCParameter
End Sub

Private Sub cmdParaEdit_Click()

End Sub

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

Private Sub cmdCASCParameter_Click()
    cboFileTypeDat.AddItem "Dat files (*.Dat)"
    cboFileTypeDat.AddItem "i.out files (*i.out)"
    cboFileTypeDat.AddItem "Par files (*.Inc)"
    cboFileTypeDat.ListIndex = 2
        
    FileList3.Pattern = "*.*"
    FileList3.Pattern = "*.Inc"


    FileList3.Visible = True

End Sub

Private Sub FileList3_Click()
    Dim filelen As Integer
    Dim tempName As String
    tempName = Trim(FileList3.Filename)
    filelen = Len(tempName)
    If filelen > 4 Then
    TxtPara.Text = Mid(tempName, 1, filelen - 4)
    End If
    'txtDep.Text = FileList2.Filename
    FileList3.Visible = False

End Sub

Private Sub cmdSave_Click()
        Dim gFileNum As Integer
        Dim Temp As String
        Dim response
   'Check current parameters
        kcrito = Val(Text_MaxiSequence.Text)
        If kcrito < 0 Then
           MsgBox ("The value of Maximum sequence# difference should be > = 0.  ")
           Exit Sub
        End If
        decrito = Val(Text_Width.Text)
        If kcrit < 0 Then
           MsgBox ("The value of Maximum order of window should be >= 0.  ")
           Exit Sub
        End If
        If Check_RemoveLarge.Value = 1 Then
              icnt = 1
        Else
              icnt = 0
        End If
        If Check_MeanDifference.Value = 1 Then
              imean = 1
        Else
              imean = 0
        End If
        If Check_SquareRoot.Value = Checked Then
              io = Val(Text_OrderofSQRT.Text)
              isqrt = 1
              If io <= 0 Then
                 MsgBox ("The value of Order of SQRT should be > 0.  ")
                 Exit Sub
              End If
        Else
              isqrt = 0
        End If
        If Check_HigherOrderOutput.Value = 1 Then
           iendo = 1
        Else
           iendo = 0
        End If
        If Check_ScaledOptimumSeq.Value = 1 Then
              irascs = 1
        Else
              irascs = 0
        End If
        kcrit = Val(Text_MaxiOrder.Text)

        
Dim Y
        If Trim(TxtPara.Text) = "" Then
            Y = MsgBox("The CASC Parameter file name is empty, please Select or Give a file name and try again.", 0, "")
            Exit Sub
        End If
        txt_file_inc = Trim(TxtPara.Text)
     
    'To save current parameters
        gFileNum = FreeFile
        Open CurDir & "\" + txt_file_inc + ".inc" For Output As gFileNum
        'rem 1
        Temp = Format(kcrito, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 2
        Temp = Format(decrito, "####0.00")
        If Len(Temp) < 8 Then
            Print #gFileNum, Spc(8 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 3
         Temp = Format(icnt, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 4
          Temp = Format(imean, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 5
         Temp = Format(isqrt, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 6
         Temp = Format(iendo, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 7
         Temp = Format(kcrit, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 8
         Temp = Format(irascs, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
        'rem 9
         Temp = Format(io, "#0")
        If Len(Temp) < 2 Then
            Print #gFileNum, Spc(2 - Len(Temp)); Temp;
        Else
            Print #gFileNum, Temp;
        End If
      
        
        Close gFileNum

        frmCascinput.TxtPara.Text = txt_file_inc
        CurCASCParaFile = frmCascinput.TxtPara.Text
    Me.Hide
    Unload frmCASCParameter
End Sub

Private Sub Form_Activate()
   If FormStatus = 0 Then
       FormStatus = 1
       Me.Hide
       Unload frmCASCParameter
   End If

End Sub

Private Sub Form_Load()

   FormStatus = 1
   'Text_OrderofSQRT.BackColor = &H8000000B
   
   TxtPara.Text = txt_file_inc
   readCASCparameter

End Sub

Private Sub readCASCparameter()
        Dim gFileNum As Integer
        Dim Temp As String
        Dim response, Y
        
         If Dir(txt_file_inc + ".inc") = "" Then
             response = MsgBox("The CASC Parameter File (" + txt_file_inc + ".inc) does not exist in current directory. " + Chr$(13) + "If  OK button is pressed, this parameter file could be created." + Chr$(13), vbOKCancel)
             If response = vbOK Then   ' 用户按下“OK”。
                 FormStatus = 1   '将可保存新的参数文件
                 ' 设置参数初始值
                 Text_Width.Text = "500.0"
                 Text_MaxiOrder.Text = "1"
                 Text_MaxiOrder.BackColor = &H8000000B
                 Text_MaxiSequence.Text = "1"
                 Text_OrderofSQRT.Text = "1"
                 Text_OrderofSQRT.BackColor = &H8000000B
                 Check_RemoveLarge.Value = 0
                 Check_MeanDifference.Value = 0
                 Check_SquareRoot.Value = 0
                 Check_HigherOrderOutput.Value = 0
                 Check_ScaledOptimumSeq.Value = 0
                 iendo = 0
                 irascs = 0
                 io = 0
                 Exit Sub
             Else   ' 用户按下“否”。
                 FormStatus = 0  ' 返回。
                 Me.Hide
                 'Unload frmCASCParameter
                 Exit Sub
            End If
         End If
  
        gFileNum = FreeFile
        Open CurDir & "\" + txt_file_inc + ".inc" For Input As gFileNum
        
        'Get the  records.
        
        If EOF(gFileNum) = False Then
              Line Input #gFileNum, Temp
                'First line standard length = 24
                If Len(Temp) = 24 Then
                        kcrito = Val(Mid(Temp, 1, 2))
                        decrito = Val(Mid(Temp, 3, 8))
                        icnt = Val(Mid(Temp, 11, 2))
                        imean = Val(Mid(Temp, 13, 2))
                        isqrt = Val(Mid(Temp, 15, 2))
                        iendo = Val(Mid(Temp, 17, 2))
                        kcrit = Val(Mid(Temp, 19, 2))
                        irascs = Val(Mid(Temp, 21, 2))
                        io = Val(Mid(Temp, 23, 2))
                 Else
                       'judge whether the INC file is a new file or corupt file
                       If Len(Trim(Temp)) = 0 Or Len(Trim(Temp)) > 24 Then
                                MsgBox "The first line of current INC file is empty or the length exceed 25, default value will be given."
                                
                                FormStatus = 1   '将可保存新的参数文件
                                ' 设置参数初始值
                                Text_Width.Text = "500.0"
                                Text_MaxiOrder.Text = "1"
                                Text_MaxiSequence.Text = "1"
                                Text_OrderofSQRT.Text = "1"
                                Check_RemoveLarge.Value = 0
                                Check_MeanDifference.Value = 0
                                Check_SquareRoot.Value = 0
                                Check_HigherOrderOutput.Value = 0
                                Check_ScaledOptimumSeq.Value = 0
                                iendo = 0
                                irascs = 0
                                io = 0
                                
                                Close gFileNum
                                Exit Sub
                       Else
                                Select Case Len(Mid(Temp, 1, 24))
                                Case 1 To 2
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = 0
                                        icnt = 0
                                        imean = 0
                                        isqrt = 0
                                        iendo = 0
                                        kcrit = 0
                                        irascs = 0
                                        io = 0
                                Case 3 To 10
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = 0
                                        imean = 0
                                        isqrt = 0
                                        iendo = 0
                                        kcrit = 0
                                        irascs = 0
                                        io = 0
                                 Case 11 To 12
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = 0
                                        isqrt = 0
                                        iendo = 0
                                        kcrit = 0
                                        irascs = 0
                                        io = 0
                               Case 13 To 14
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = Val(Mid(Temp, 13, 2))
                                        isqrt = 0
                                        iendo = 0
                                        kcrit = 0
                                        irascs = 0
                                        io = 0
                               Case 15 To 16
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = Val(Mid(Temp, 13, 2))
                                        isqrt = Val(Mid(Temp, 15, 2))
                                        iendo = 0
                                        kcrit = 0
                                        irascs = 0
                                        io = 0
                             Case 17 To 18
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = Val(Mid(Temp, 13, 2))
                                        isqrt = Val(Mid(Temp, 15, 2))
                                        iendo = Val(Mid(Temp, 17, 2))
                                        kcrit = 0
                                        irascs = 0
                                        io = 0
                               Case 19 To 20
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = Val(Mid(Temp, 13, 2))
                                        isqrt = Val(Mid(Temp, 15, 2))
                                        iendo = Val(Mid(Temp, 17, 2))
                                        kcrit = Val(Mid(Temp, 19, 2))
                                        irascs = 0
                                        io = 0
                                Case 21 To 22
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = Val(Mid(Temp, 13, 2))
                                        isqrt = Val(Mid(Temp, 15, 2))
                                        iendo = Val(Mid(Temp, 17, 2))
                                        kcrit = Val(Mid(Temp, 19, 2))
                                        irascs = Val(Mid(Temp, 21, 2))
                                        io = 0
                                Case 23 To 24
                                        kcrito = Val(Mid(Temp, 1, 2))
                                        decrito = Val(Mid(Temp, 3, 8))
                                        icnt = Val(Mid(Temp, 11, 2))
                                        imean = Val(Mid(Temp, 13, 2))
                                        isqrt = Val(Mid(Temp, 15, 2))
                                        iendo = Val(Mid(Temp, 17, 2))
                                        kcrit = Val(Mid(Temp, 19, 2))
                                        irascs = Val(Mid(Temp, 21, 2))
                                        io = Val(Mid(Temp, 23, 2))
                                End Select
                       End If
                 End If
        End If
        
        'Display the records
        Text_Width.Text = Trim(decrito)
        Text_MaxiOrder.Text = Trim(kcrit)
        Text_MaxiSequence.Text = Trim(kcrito)
        Text_OrderofSQRT.Text = Trim(io)
        If icnt = 1 Then
              Check_RemoveLarge.Value = 1
        Else
              Check_RemoveLarge.Value = 0
        End If
        If imean = 1 Then
              Check_MeanDifference.Value = 1
        Else
              Check_MeanDifference.Value = 0
        End If
        If isqrt = 1 Then
              Check_SquareRoot.Value = 1
              Text_OrderofSQRT.BackColor = &H80000009
        Else
              Check_SquareRoot.Value = 0
              Text_OrderofSQRT.BackColor = &H8000000B
        End If
        If irascs = 1 Then
              Check_ScaledOptimumSeq.Value = 1
        Else
              Check_ScaledOptimumSeq.Value = 0
        End If
        If iendo = 1 Then
              Check_HigherOrderOutput.Value = 1
              Text_MaxiOrder.BackColor = &H80000009
        Else
              Check_HigherOrderOutput.Value = 0
              Text_MaxiOrder.BackColor = &H8000000B
        End If
        
        'Check_ScaledOptimumSeq is not used at present
        FormStatus = 1
        Close gFileNum
        
End Sub

