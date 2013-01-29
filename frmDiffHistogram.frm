VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmDiffHistogram 
   Caption         =   "Depth Differences Frequency Histogram"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   12840
   Icon            =   "frmDiffHistogram.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmd_WidthFilter 
      Caption         =   ">>"
      Height          =   285
      Left            =   9330
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox Text_Width 
      Height          =   255
      Left            =   9780
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CheckBox Check_RemoveLarge 
      Caption         =   "Remove large depth difference"
      Height          =   345
      Left            =   10620
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List_ClassifyScale 
      Height          =   255
      ItemData        =   "frmDiffHistogram.frx":0442
      Left            =   6300
      List            =   "frmDiffHistogram.frx":046D
      TabIndex        =   6
      Top             =   150
      Width           =   705
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Data"
      Height          =   252
      Left            =   3660
      TabIndex        =   5
      Top             =   150
      Width           =   1092
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   252
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin OlectraChart2D.Chart2D ChartBar 
      Height          =   8100
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   12345
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   21775
      _ExtentY        =   14287
      _StockProps     =   0
      ControlProperties=   "frmDiffHistogram.frx":04A3
   End
   Begin VB.CheckBox Check_UseWidthFilter 
      Caption         =   "Use Width of Depth Filter"
      Height          =   315
      Left            =   7140
      TabIndex        =   10
      Top             =   150
      Width           =   2205
   End
   Begin VB.Label LabeClassifyScale 
      Caption         =   "Select an interval :"
      Height          =   255
      Left            =   4890
      TabIndex        =   7
      Top             =   150
      Width           =   1455
   End
End
Attribute VB_Name = "frmDiffHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UnitType As String, OpenFileKey As Integer
Dim OrderDataWell() As Integer, OrderDataWellEventSeqNum() As Integer, OrderData() As Double
Dim HistoDataSeqnum() As Integer, HistoData() As Double, OrderNumber As Integer
Dim HistoDataNum As Integer, OrderDataNum As Integer
Dim QQplotDataMin As Double, QQplotDataMax As Double
Dim HistoDataMin As Double, HistoDataMax As Double
Dim HistoClassifyCountMax As Integer, DepthOrderNumTemp As Integer
Dim ClassifyScale As Integer
Dim StartValue As Double, EndValue As Double, StartTemp As Double, EndTemp As Double, StartKey As Integer, EndKey As Integer
Dim ValueStep() As Double, ValueTemp() As Double, ValueCount()  As Integer, StepCount As Integer

Dim kcrito As Integer, icnt As Integer, imean As Integer, isqrt As Integer, iendo As Integer, kcrit As Integer, irascs As Integer, io As Integer
Dim decrito As Double
Dim CurIsqrt  As Integer, CurIendo As Integer, CurKcrit As Integer
Dim Unitsflag As Integer  '=0 indicates meters; =1 indicates feet



Private Sub Check_RemoveLarge_Click()
    If OpenFileKey = 1 Then
          DataClassify
          ShowHistogram
    End If
End Sub

Private Sub Check_UseWidthFilter_Click()
   If Check_UseWidthFilter.Value = 1 Then
       Text_Width.Visible = True
       Check_RemoveLarge.Visible = True
       Cmd_WidthFilter.Visible = True
       If txt_file_inc <> "" Then
            Text_Width.Text = Trim(decrito)
             If icnt = 1 Then
                   Check_RemoveLarge.Value = 1
             Else
                   Check_RemoveLarge.Value = 0
             End If
       End If
   Else
       Text_Width.Visible = False
       Check_RemoveLarge.Visible = False
       Cmd_WidthFilter.Visible = False
        If OpenFileKey = 1 Then
              DataClassify
              ShowHistogram
        End If
   End If

End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        ChartTableLoader = 2
        frmChartTable.Show
    Else
        frmChartTable.Hide
    End If
End Sub

Private Sub Cmd_WidthFilter_Click()
    If OpenFileKey = 1 Then
          DataClassify
          ShowHistogram
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload frmDiffHistogram
End Sub

Private Sub cmdOpen_Click()
        txt_file_df1 = ""
        frmOpenDF1.Show 1
        'cmdOpen.Enabled = False
        If txt_file_df1 = "" Then
           Exit Sub
        End If
        If Dir(CurDir + "\" + txt_file_df1) = "" Then
            Beep
            MsgBox "Input file does not exist, please try again"
            'OpenFileKey = 0
            Exit Sub
        End If
        If filelen(CurDir + "\" + txt_file_df1) = 0 Then
            Beep
            MsgBox "Input file is empty, please try again"
            'OpenFileKey = 0
            Exit Sub
        End If
        startup
        OpenFile
        If OpenFileKey = 1 Then
          DataClassify
          ShowHistogram
        End If

End Sub

Private Sub cmdPrint_Click()
    ChartBar.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub cmdSave_Click()

     Set CurChartSaveObj = ChartBar
     ChartSaveAs

'        Dim ImageName As String
'        ImageName = InputBox("Please give a file name without extension", "Save chart as an image (JPG)", 1)
'       If ImageName <> "" Then
'             ChartBar.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'       End If
End Sub


Private Sub Form_Activate()
    Set CurGraphicOBJ = ChartBar
        CurWindowNum = 15
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    txt_file_df1 = ""
    cmdOpen.Enabled = True
    OpenFileKey = 0
    'Interval of Statistic
    List_ClassifyScale.ListIndex = 9
    ClassifyScale = Val(List_ClassifyScale.Text)
    ' Set to Meters for the time being
    Unitsflag = 0
    If Unitsflag = 0 Then
       UnitType = "  (m)"
    Else
       UnitType = "  (feet)"
    End If
    CurIsqrt = 0
    CurIendo = 0
    CurKcrit = 0

    DepthOrderNum = 1
    DepthOrderNumTemp = 1
    
    
    'Read CASC parameter file
'    DepthOrderNum = 1
'    DepthOrderNumTemp = 1
'    If Dir(txt_file_inc + ".inc") <> "" Then
'        ReadCASCPara
'        If decrito > 0 Then
'            Check_UseWidthFilter.Value = 1
'            Cmd_WidthFilter.Visible = 1
'            Text_Width.Text = Trim(decrito)
'            If icnt = 1 Then
'                  Check_RemoveLarge.Value = 1
'            Else
'                  Check_RemoveLarge.Value = 0
'            End If
'        End If
'    End If

    
    CurWindowNum = 15
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd


End Sub

Private Sub ReadCASCPara()
        Dim gFileNum As Integer
        Dim Temp As String
        Dim Y
    
    If Dir(CurDir & "\" + txt_file_inc + ".inc") <> "" Then
        gFileNum = FreeFile
        Open CurDir & "\" + txt_file_inc + ".inc" For Input As gFileNum
        
        'Get the three records.
        
        If EOF(gFileNum) = False Then
              Line Input #gFileNum, Temp
        End If
        If Len(Temp) < 24 Then
            Y = MsgBox("The CASC Parameter file Format  is not correct, please check it and try again.", 0, "")
            'FormStatus = 0  ' ·µ»Ø¡£
            'Me.Hide
            Exit Sub
        End If
        kcrito = Val(Mid(Temp, 1, 2))
        decrito = Val(Mid(Temp, 3, 8))
        icnt = Val(Mid(Temp, 11, 2))
        imean = Val(Mid(Temp, 13, 2))
        isqrt = Val(Mid(Temp, 15, 2))
        iendo = Val(Mid(Temp, 17, 2))
        kcrit = Val(Mid(Temp, 19, 2))
        irascs = Val(Mid(Temp, 21, 2))
        io = Val(Mid(Temp, 23, 2))
'        Text_Width.Text = Trim(decrito)
'        Text_MaxiOrder.Text = Trim(kcrit)
'        Text_MaxiSequence.Text = Trim(kcrito)
'        Text_OrderofSQRT.Text = Trim(io)
'        If icnt = 1 Then
'              Check_RemoveLarge.Value = 1
'        Else
'              Check_RemoveLarge.Value = 0
'        End If
'        If imean = 1 Then
'              Check_MeanDifference.Value = 1
'        Else
'              Check_MeanDifference.Value = 0
'        End If
'        If isqrt = 1 Then
'              Check_SquareRoot.Value = 1
'        Else
'              Check_SquareRoot.Value = 0
'        End If
        
        Close gFileNum

    
    End If

End Sub


Private Sub Form_Resize()
        Dim size As Integer
        If (frmDiffHistogram.Height > 360) And (frmDiffHistogram.Height - 360 < frmDiffHistogram.Width / 1.5) Then
           size = frmDiffHistogram.Height - 360
        ElseIf (frmDiffHistogram.Height > 360) And (frmDiffHistogram.Height - 360 > frmDiffHistogram.Width / 1.5) Then
           size = frmDiffHistogram.Width / 1.5
        Else
           size = 0
        End If
        
        ChartBar.Height = size
        ChartBar.Width = 1.5 * size
End Sub


Public Sub startup()
        Dim I, J As Integer
        ChartBar.Visible = False
        
        ChartBar.ChartGroups(1).ChartType = oc2dTypeBar
        With ChartBar.ChartGroups(1).Data
            .Layout = oc2dDataArray
            .IsBatched = True
            .NumSeries = 1  '**** changed to 2 from 1
         '   .NumPoints(1) = 150
           
        End With
        
        ChartBar.ChartGroups(1).Data.IsBatched = False
        
End Sub


Public Sub OpenFile()
        Dim FileNum As Integer
        Dim Filename As String
        
        Dim Temp As String
        Dim tempfrom As Integer
        Dim tempto As Integer
        Dim tempstr As String
        Dim tempstr1 As String
        Dim minx, maxx As Double
        Dim I, J As Integer
        Dim title_str As String
        Dim OrderNumberCalcuKey As Integer
        
        Dim ArrayNum As Integer
        Dim tnum
        
        FileNum = FreeFile
        If txt_file_df1 <> "" Then
            Open CurDir + "\" + txt_file_df1 For Input As FileNum
            OpenFileKey = 1
        Else
            MsgBox "The input file  " + CurDir + "\" + txt_file_df1 + " does not exist, " + Chr$(13) + "Please check it and try again."
            OpenFileKey = 0
            Exit Sub
        End If
            
                
        HistoDataNum = 0
        OrderDataNum = 0
        OrderNumberCalcuKey = 0
        
        'count the number of two type datasets from DF1 file
        While Not (EOF(FileNum))
            Line Input #FileNum, Temp
            If InStr(1, UCase(Temp), UCase("Graph parameters:")) = 0 Then
                    If Len(Temp) = 28 Then
                       HistoDataNum = HistoDataNum + 1
                    Else
                       OrderDataNum = OrderDataNum + 1
                       If OrderNumberCalcuKey = 0 Then
                          OrderNumber = (Len(Temp) - 8#) / 12#
                          OrderNumberCalcuKey = 1
                       End If
                    End If
            End If
        Wend
        '              Read Order Number
        '                If InStr(1, Temp, "=") > 0 Then
        '                   LineSlope = Val(Mid(Temp, 20, 10))
        '                   If InStr(1, Temp, "Order") > 0 Then
        '                      DepthOrderNumTemp = Val(Mid(Temp, 40, 2))
        '                   Else
        '                     DepthOrderNumTemp = 1
        '                   End If
        '                End If
        Close #FileNum
        
        'read the data into two dimensions
        HistoDataMin = 0
        HistoDataMax = 0
        QQplotDataMin = 0
        QQplotDataMax = 0
        
        If OrderDataNum = 0 Or HistoDataNum = 0 Then
            MsgBox "The input file " + CurDir + "\" + txt_file_df1 + " is empty or the format is  not correct, " + Chr$(13) + "Please check it and try again." + Chr$(13)
            OpenFileKey = 0
            Exit Sub
        End If
        
        
        ReDim OrderDataWell(1 To OrderDataNum) As Integer, OrderDataWellEventSeqNum(1 To OrderDataNum) As Integer, OrderData(1 To OrderDataNum, 1 To OrderNumber) As Double
        ReDim HistoDataSeqnum(1 To HistoDataNum) As Integer, HistoData(1 To HistoDataNum, 1 To 2) As Double
        Open CurDir + "\" + txt_file_df1 For Input As FileNum
        For I = 1 To OrderDataNum
            Line Input #FileNum, Temp
            OrderDataWell(I) = Mid(Temp, 1, 4)
            OrderDataWellEventSeqNum(I) = Mid(Temp, 5, 4)
            For J = 1 To OrderNumber
                OrderData(I, J) = Mid(Temp, 9 + (J - 1) * 12, 12)
            Next J
            
'            If I = 1 Then
'              MsgBox Temp + "   " + Str(OrderDataWell(I)) + "  " + Str(OrderDataWellEventSeqNum(I)) + "   " + Str(OrderData(I, 1)) + "   " + Str(OrderData(I, 2)) + "   " + Str(OrderData(I, 3)) + "   " + Str(OrderData(I, 4)) + "   " + Str(OrderData(I, 5))
'           End If
        Next I
        For I = 1 To HistoDataNum
           Line Input #FileNum, Temp
           HistoDataSeqnum(I) = Mid(Temp, 1, 4)
           HistoData(I, 1) = Mid(Temp, 5, 12)
           HistoData(I, 2) = Mid(Temp, 17, 12)
           
          'Compare QQplotData Value
           If HistoData(I, 1) > QQplotDataMax Then
              QQplotDataMax = HistoData(I, 1)
           End If
           If HistoData(I, 1) < QQplotDataMin Then
              QQplotDataMin = HistoData(I, 1)
           End If
           'Compare HistoData Value of 1st Order
           If HistoData(I, 2) > HistoDataMax Then
              HistoDataMax = HistoData(I, 2)
           End If
           If HistoData(I, 2) < HistoDataMin Then
              HistoDataMin = HistoData(I, 2)
           End If
           
'           If I = 1 Then
'              MsgBox Temp + "   " + Str(HistoDataSeqnum(I)) + "  " + Str(HistoData(I, 1)) + "   " + Str(HistoData(I, 2))
'           End If
        Next I
        
'Read Graph parameters: isqrt    iendo(irascs?)   kcrit
            If Not (EOF(FileNum)) Then
                Line Input #FileNum, Temp
                If InStr(1, UCase(Temp), UCase("Graph parameters:")) > 0 Then
                   CurIsqrt = Val(Mid(Temp, 20, 2))
                   CurIendo = Val(Mid(Temp, 22, 2))
                   CurKcrit = Val(Mid(Temp, 24, 2))
                Else
                    'MsgBox "The graphic parameters were not included in input file: " + CurDir + "\" + txt_file_di1 + Chr$(13) + "Current CASCW.EXE may be an old version file, please check it."
                    'OpenFileKey = 0
                    'Close #FileNum
                    'Exit Sub
                    CurIsqrt = 0
                    CurIendo = 0
                    CurKcrit = 0
                End If
            Else
                'MsgBox "The graphic parameters were not included in input file: " + CurDir + "\" + txt_file_di1 + Chr$(13) + "Current CASCW.EXE may be an old version file, please check it."
                'OpenFileKey = 0
                'Close #FileNum
                'Exit Sub
                CurIsqrt = 0
                CurIendo = 0
                CurKcrit = 0

            End If
        
        Close #FileNum
           
        
        'UnitFlag:  Meter or Feet
        'Unitsflag = Val(Right(Temp, 1))
        ' MsgBox temp
        
        'MsgBox "The actural num of QQpotNUM and HistoDataNum are  " + Str(OrderDataNum) + ", " + Str(HistoDataNum)
        
End Sub

Private Sub DataClassify()
        Dim I, J, k
        Dim WTemp As Double
        
'        ClassifyScale = 100
        
        StartValue = 0
        EndValue = 0
        If HistoDataMin < 0 Then
            StartValue = -(Abs(HistoDataMin) + ClassifyScale - (Int(Abs(HistoDataMin)) Mod ClassifyScale))
        End If
        If HistoDataMin >= 0 Then
            StartValue = Abs(HistoDataMin) - (Int(Abs(HistoDataMin)) Mod ClassifyScale)
        End If
        
        If HistoDataMax < 0 Then
            EndValue = -(Abs(HistoDataMax) - (Int(Abs(HistoDataMax)) Mod ClassifyScale))
        End If
        If HistoDataMax >= 0 Then
            EndValue = Abs(HistoDataMax) + ClassifyScale - (Int(Abs(HistoDataMax)) Mod ClassifyScale)
        End If
        
        
'        If HistoDataMax >= 0 Then
'            If HistoDataMax >= Abs(HistoDataMin) Then
'                temp = HistoDataMax
'            Else
'                temp = Abs(HistoDataMin)
'            End If
'        Else
'            temp = Abs(HistoDataMin)
'        End If
        'The DepthWidth and RemoveOutsideKey are two CASC parameters about Depth Differences given by the User
'        If DepthWidth > temp Then
'            temp = DepthWidth
'        End If
        
          
        StepCount = 0
        For I = StartValue To EndValue Step ClassifyScale
             StepCount = StepCount + 1
        Next I
        
       ReDim ValueTemp(0 To StepCount), ValueCount(0 To StepCount), ValueStep(0 To StepCount)
       StepCount = 0
       ValueCount(StepCount) = 0
       ValueStep(StepCount) = 0
       For I = StartValue To EndValue Step ClassifyScale
             StepCount = StepCount + 1
             ValueTemp(StepCount) = I
             ValueCount(StepCount) = 0
             ValueStep(StepCount) = 0
        Next I
 
        'The DepthWidth and RemoveOutsideKey are two CASC parameters about Depth Differences given by the User
        'Adjust Step and Value according to the User
        ReDim HistoDataTemp(1 To HistoDataNum) As Double
        Dim CountTemp As Integer
        WTemp = Val(Text_Width.Text)
        If Check_UseWidthFilter.Value = 1 And WTemp > 0 Then
            If WTemp < HistoDataMax Or Abs(HistoDataMin) > WTemp Then
                'Special processing
                'Change statistic value
                For J = 1 To HistoDataNum
                          If Abs(HistoData(J, 2)) > WTemp Then
                               If HistoData(J, 2) < 0 Then
                                    If Check_RemoveLarge.Value = 1 And Abs(HistoData(J, 2)) >= WTemp + ClassifyScale Then
                                        HistoDataTemp(J) = StartValue - 1
                                    Else
                                        HistoDataTemp(J) = -WTemp - 0.001   '  - ClassifyScale
                                    End If
                                Else
                                    If Check_RemoveLarge.Value = 1 Then
                                        HistoDataTemp(J) = EndValue + 1
                                    Else
                                        HistoDataTemp(J) = WTemp + ClassifyScale - 0.001
                                    End If
                                End If
                          Else
                               HistoDataTemp(J) = HistoData(J, 2)
                          End If
                Next J
                For J = 1 To HistoDataNum
                     For k = 1 To StepCount - 1
                          If (k = 1 And HistoDataTemp(J) = ValueTemp(k)) Or (HistoDataTemp(J) < 0 And HistoDataTemp(J) = -WTemp And HistoDataTemp(J) = ValueTemp(k)) Then
                               'StartValue processing
                               ValueStep(0) = ValueTemp(k)
                               ValueCount(0) = ValueCount(0) + 1
                          Else
                                ValueStep(k) = ValueTemp(k + 1)
                                If ((HistoDataTemp(J) > ValueTemp(k)) And (HistoDataTemp(J) <= ValueTemp(k + 1))) Then
                                     ValueCount(k) = ValueCount(k) + 1
                                End If
                          End If
                    Next k
                Next J
            Else
                'Normal processing
                For J = 1 To HistoDataNum
                     For k = 1 To StepCount - 1
                          If (k = 1 And HistoData(J, 2) = ValueTemp(k)) Then
                               'StartValue processing
                               ValueStep(k - 1) = ValueTemp(k)
                               ValueCount(k - 1) = ValueCount(0) + 1
                          Else
                                ValueStep(k) = ValueTemp(k + 1)
                                If ((HistoData(J, 2) > ValueTemp(k)) And (HistoData(J, 2) <= ValueTemp(k + 1))) Then
                                     ValueCount(k) = ValueCount(k) + 1
                                End If
                          End If
                    Next k
                Next J
            End If
        Else
            'Normal processing
            For J = 1 To HistoDataNum
                 For k = 1 To StepCount - 1
                          If k = 1 And HistoData(J, 2) = ValueTemp(k) Then
                               'StartValue processing
                               ValueStep(0) = ValueTemp(k)
                               ValueCount(0) = ValueCount(0) + 1
                          Else
                                ValueStep(k) = ValueTemp(k + 1)
                                If ((HistoData(J, 2) > ValueTemp(k)) And (HistoData(J, 2) <= ValueTemp(k + 1))) Then
                                     ValueCount(k) = ValueCount(k) + 1
                                End If
                          End If
                Next k
            Next J
        End If


End Sub
        
Private Sub ShowHistogram()
        Dim gridspace As Double
        Dim I, J As Integer
        Dim MaxCount As Integer, SumCount As Integer
        
        ChartBar.Visible = False
        
        'ChartBar.ChartGroups(1).Styles(1).Symbol.size = 0
        'ChartBar.ChartGroups(1).Styles(1).Line.Width = 2
        'ChartBar.ChartGroups(1).Styles(1).Line.Color = 0
        ChartBar.ChartGroups(1).ChartType = oc2dTypeBar
        
        'Clear Old Data
        For I = 1 To ChartBar.ChartGroups(1).Data.NumPoints(1)
             ChartBar.ChartGroups(1).Data.X(1, I) = 0
             ChartBar.ChartGroups(1).Data.Y(1, I) = 0
        Next I
        
'        If file_type = "DE1" Then
'            title_str = "Ranked"
'        ElseIf file_type = "DE2" Then
'            title_str = "Scaled"
'        End If
        
'        If Unitsflag = 0 Then
'           ChartBar.ChartArea.Axes("Y").Title.Text = "Depth (meters)"
'        Else
'           ChartBar.ChartArea.Axes("Y").Title.Text = "Depth (feet)"
'        End If


        MaxCount = 0
        SumCount = 0
        With ChartBar.ChartGroups(1).Data
            .Layout = oc2dDataArray
            .IsBatched = True
            .NumSeries = 1
            .NumPoints(1) = StepCount
             frmChartTable.listOfData.Clear
             frmChartTable.listOfData.AddItem "    No. " + "  Depth Difference  " + "    Frequency"
             For I = 0 To StepCount - 1
                'For the ValueCount(0)=0 processing
                If I = 0 And ValueCount(I) = 0 Then
                     'Nothing....
                     .X(1, I + 1) = 0
                     .Y(1, I + 1) = 0
                Else
                     .X(1, I + 1) = ValueStep(I)
                     .Y(1, I + 1) = ValueCount(I)
                     SumCount = SumCount + ValueCount(I)
                
                      frmChartTable.listOfData.AddItem Space(5 - Len(Str(I))) + Str(I) _
                            + Space(12 - Len(Format(ValueStep(I) - ClassifyScale, "0.0"))) _
                            + Format(ValueStep(I) - ClassifyScale, "0.0") + " ~ " + Space(8 - Len(Format(ValueStep(I), "0.0"))) + Format(ValueStep(I), "0.0") _
                            + Space(10 - Len(Str(ValueCount(I)))) _
                            + Str(ValueCount(I))
                
                        ' Get the Max Count Value of Value
                        If ValueCount(I) > MaxCount Then
                             MaxCount = ValueCount(I)
                        End If
                End If
             Next I
             frmChartTable.listOfData.AddItem Space(5)
             frmChartTable.listOfData.AddItem "Total Points: " + Str(HistoDataNum)
             If Check_UseWidthFilter.Value = 1 And Val(Text_Width.Text) > 0 Then
                  frmChartTable.listOfData.AddItem "Result Points: " + Str(SumCount) + "   with Width of Depth Filter: " + Text_Width.Text
             End If
        End With
        

        ChartBar.ChartArea.Axes("Y").Min.Value = 0
        ChartBar.ChartArea.Axes("Y").Max.Value = 1.1 * MaxCount
        
        
        'to make two groups of chart with the same scale
'        ChartBar.ChartArea.Axes("y2").Min.Value = ChartBar.ChartArea.Axes("y").Min.Value
'        ChartBar.ChartArea.Axes("y2").Max.Value = ChartBar.ChartArea.Axes("Y").Max.Value
          
        
        ChartBar.ChartArea.Axes("y").Origin.Value = 0
        
        If ValueCount(0) > 0 Then
             ChartBar.ChartArea.Axes("x").Min.Value = StartValue - ClassifyScale
             ChartBar.ChartArea.Axes("x").Origin.Value = StartValue - ClassifyScale
        Else
             ChartBar.ChartArea.Axes("x").Min.Value = StartValue
             ChartBar.ChartArea.Axes("x").Origin.Value = StartValue
        End If
        If ValueCount(StepCount - 1) > 0 Then
            ChartBar.ChartArea.Axes("x").Max.Value = EndValue + ClassifyScale
        Else
            ChartBar.ChartArea.Axes("x").Max.Value = EndValue
        End If
'         Gridspace = ChartBar.ChartArea.Axes("Y").Max.Value / 10
'         ChartBar.ChartArea.Axes("x").MajorGrid.Spacing = Gridspace

'*df1 (2 times: histogram and Q-Q plot): on horizontal scale "m" replaced
'by "m^0.5" if isqrt=1
        If CurIsqrt = 1 Then
             'Chartden.ChartArea.Axes("x").Title.Text = "Interevent distance"
            If Unitsflag = 0 Then
               ChartBar.ChartArea.Axes("X").Title.Text = "Depth Difference (m^0.5)"
            Else
               ChartBar.ChartArea.Axes("X").Title.Text = "Depth Difference (feet^0.5)"
            End If
        Else
             'Chartden.ChartArea.Axes("x").Title.Text = "Interevent distance (m)"
            If Unitsflag = 0 Then
               ChartBar.ChartArea.Axes("X").Title.Text = "Depth Difference (m)"
            Else
               ChartBar.ChartArea.Axes("X").Title.Text = "Depth Difference (feet)"
            End If
        End If
'        If Unitsflag = 0 Then
'           ChartBar.ChartArea.Axes("X").Title.Text = "Depth Difference (m)"
'        Else
'           ChartBar.ChartArea.Axes("X").Title.Text = "Depth Difference (feet)"
'        End If
    
        ChartBar.ChartArea.Axes("Y").Title.Text = "Frequency"
        
        ChartBar.Header.Text = "Histogram of Depth Differences; Order=" + Trim(Str(DepthOrderNumTemp))

        
        ChartBar.ChartLabels(1).Text.Text = txt_file_df1
        ChartBar.ChartLabels(1).Anchor = oc2dAnchorSouthWest
        ChartBar.ChartLabels(1).AttachDataCoord.X = ChartBar.ChartArea.Axes("x").Max.Value
        ChartBar.ChartLabels(1).AttachDataCoord.Y = ChartBar.ChartArea.Axes("Y").Max.Value
        
         ChartBar.ChartGroups(1).Data.IsBatched = False
        
         ChartBar.Visible = True


End Sub


Private Sub List_Well_Click()
'        Dim gridspace As Double
'        Dim I, J As Integer
'        Dim well As Integer
'        Dim Emin, Emax, E_temp As Double
'        Dim Name_temp As String
'        Dim title_str As String
'        Dim tempfrom As Integer
'        ChartBar.Visible = False
'
'        well = List_Well.ListIndex + 1
'        'MsgBox "Well No. == " + Str(well)
'        'LabelWell.Visible = False
'        'List_Well.Visible = False
'
'
'        ChartBar.ChartGroups(1).Styles(1).Symbol.size = 0
'        'ChartBar.ChartGroups(1).Styles(1).Line.Width = 2
'        'ChartBar.ChartGroups(1).Styles(1).Line.Color = 0
'        ChartBar.ChartGroups(1).ChartType = oc2dTypePlot
'
'        If file_type = "DE1" Then
'            title_str = "Ranked"
'        ElseIf file_type = "DE2" Then
'            title_str = "Scaled"
'        End If
'
'        'ChartBar.Header.Text.Text = "Scattergram   " + txt_file_DF1 + "  (Well No. " + Str(well) + ")"
'        ChartBar.ChartArea.Axes("x").Title.Text = title_str + " Optimum Sequence of Events"
'
'        ChartBar.ChartLabels(1).Text.Text = nameOfWell(well)
'        With ChartBar.ChartGroups(1).Data
'            .Layout = oc2dDataGeneral
'            .IsBatched = True
'            .NumSeries = 2
'            .NumPoints(1) = NumOfEvent(well)
'             tempfrom = 0
'                 For J = 1 To well - 1
'                 tempfrom = tempfrom + NumOfEvent(J)
'                 Next J
'            Emax = 0
'            Emin = 0
'         'sort data according to x values
'               For I = tempfrom + 1 To tempfrom + NumOfEvent(well) - 1
'                For J = I + 1 To tempfrom + NumOfEvent(well)
'                    If x_Value(I) > x_Value(J) Then
'                       E_temp = x_Value(I)
'                       x_Value(I) = x_Value(J)
'                       x_Value(J) = E_temp
'                       E_temp = E_Value(I)
'                       E_Value(I) = E_Value(J)
'                       E_Value(J) = E_temp
'                       E_temp = y_value(I)
'                       y_value(I) = y_value(J)
'                       y_value(J) = E_temp
'                       Name_temp = NameOfEvent(I)
'                       NameOfEvent(I) = NameOfEvent(J)
'                       NameOfEvent(J) = Name_temp
'
'                     End If
'                Next J
'                Next I
'
'                For I = 1 To NumOfEvent(well)
'                    If y_value(tempfrom + I) < Emin Then
'                       Emin = y_value(tempfrom + I)
'                    End If
'                    If y_value(tempfrom + I) > Emax Then
'                       Emax = y_value(tempfrom + I)
'                    End If
'                Next I
'
'                frmChartTable.listOfData.Clear
'                frmChartTable.listOfData.AddItem "    I" + "         X" + "       Depth" + "   No" + "  Event Name"
'
'               For I = 1 To NumOfEvent(well)
'                .X(1, I) = x_Value(tempfrom + I)
'                .Y(1, I) = E_Value(tempfrom + I)
'
'                Next I
'
'        End With
'
'        ChartBar.ChartArea.Axes("y").Min.Value = Int(Emin) + 1
'        ChartBar.ChartArea.Axes("Y").Max.Value = Int(Emax) + 1
'        'ChartBar.ChartArea.Axes("y2").Min.Value = Int(Emin) + 1
'        'ChartBar.ChartArea.Axes("y2").Max.Value = Int(Emax) + 1
'
'        'to make two groups of chart with the same scale
'        ChartBar.ChartArea.Axes("y2").Min.Value = ChartBar.ChartArea.Axes("y").Min.Value
'        ChartBar.ChartArea.Axes("y2").Max.Value = ChartBar.ChartArea.Axes("Y").Max.Value
'
'        ChartBar.ChartArea.Axes("x").Origin.Value = 0
'        ChartBar.ChartArea.Axes("y").Origin.Value = Int(Emax) + 1
'
'        ChartBar.ChartArea.Axes("x").Min.Value = 0
'        ChartBar.ChartArea.Axes("x").Max.Value = x_Value(tempfrom + NumOfEvent(well)) + 1
'
'        If file_type = "DE2" Then
'         gridspace = x_Value(tempfrom + NumOfEvent(well)) / 10
'         ChartBar.ChartArea.Axes("x").MajorGrid.Spacing = gridspace
'        End If
'
'        'ChartBar.ChartArea.Axes("y2").Origin.Value = Int(Emin) - 1
'
'        ChartBar.ChartGroups(1).Data.IsBatched = False
'
'
'        ChartBar.ChartGroups(2).Styles(1).Symbol.size = 5
'        ChartBar.ChartGroups(2).ChartType = oc2dTypePlot
'
'
'        With ChartBar.ChartGroups(1).Data
'            .Layout = oc2dDataGeneral
'            .IsBatched = True
'            '.NumSeries = 1
'            .NumPoints(2) = NumOfEvent(well)
'               For I = 1 To NumOfEvent(well)
'                .X(2, I) = x_Value(tempfrom + I)
'                .Y(2, I) = y_value(tempfrom + I)
'
'                frmChartTable.listOfData.AddItem Space(5 - Len(Str(I))) + Str(I) _
'                            + Space(10 - Len(Str(x_Value(tempfrom + I)))) _
'                            + Str(x_Value(tempfrom + I)) _
'                            + Space(12 - Len(Str(y_value(tempfrom + I)))) _
'                            + Str(y_value(tempfrom + I)) _
'                            + Space(5 - Len(Str(I))) + NameOfEvent(tempfrom + I)
'
'              Next I
'        End With
'
'        ChartBar.ChartGroups(1).Data.IsBatched = False
'        'ChartBar.Visible = True
'        ChartBar.Visible = True
'        If Unitsflag = 0 Then
'           ChartBar.ChartArea.Axes("Y").Title.Text = "Depth (meters)"
'        Else
'           ChartBar.ChartArea.Axes("Y").Title.Text = "Depth (feet)"
'        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CurGraphicOBJ = Nothing
    CurWindowNum = 15
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1
End Sub

Private Sub List_ClassifyScale_Click()
    ClassifyScale = Val(List_ClassifyScale.Text)
    If OpenFileKey = 1 Then
          DataClassify
          ShowHistogram
    End If
End Sub


