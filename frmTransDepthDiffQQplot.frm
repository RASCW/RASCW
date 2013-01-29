VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmTransDepthDiffQQplot 
   Caption         =   "Transformed Depth Difference Plot"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   12690
   Icon            =   "frmTransDepthDiffQQplot.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "View Data"
      Height          =   252
      Left            =   3540
      TabIndex        =   5
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Default         =   -1  'True
      Height          =   252
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1110
      TabIndex        =   2
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2670
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   252
      Left            =   1950
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin OlectraChart2D.Chart2D ChartTransQQPlot 
      Height          =   7920
      Left            =   120
      TabIndex        =   4
      Top             =   510
      Visible         =   0   'False
      Width           =   11955
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   21087
      _ExtentY        =   13970
      _StockProps     =   0
      ControlProperties=   "frmTransDepthDiffQQplot.frx":0442
   End
End
Attribute VB_Name = "frmTransDepthDiffQQplot"
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
Dim ClassifyScale As Integer, LineSlope As Double
Dim FitLineData() As Double, LineDataNum As Integer
Dim StartValue As Double, EndValue As Double, StartTemp As Double, EndTemp As Double, StartKey As Integer, EndKey As Integer
Dim ValueStep() As Double, ValueTemp() As Double, ValueCount()  As Integer, StepCount As Integer

Dim kcrito As Integer, icnt As Integer, imean As Integer, isqrt As Integer, iendo As Integer, kcrit As Integer, irascs As Integer, io As Integer
Dim decrito As Double
Dim CurIsqrt  As Integer, CurIendo As Integer, CurKcrit As Integer
Dim Unitsflag As Integer  '=0 indicates meters; =1 indicates feet


'Private Sub Check_RemoveLarge_Click()
'    If OpenFileKey = 1 Then
'          DataClassify
'          ShowQQPlot
'    End If
'End Sub
'
'Private Sub Check_UseWidthFilter_Click()
'   If Check_UseWidthFilter.Value = 1 Then
'       Text_Width.Visible = True
'       Check_RemoveLarge.Visible = True
'       Cmd_WidthFilter.Visible = True
'       If txt_file_par <> "" Then
'            Text_Width.Text = Trim(decrito)
'             If icnt = 1 Then
'                   Check_RemoveLarge.Value = 1
'             Else
'                   Check_RemoveLarge.Value = 0
'             End If
'       End If
'   Else
'       Text_Width.Visible = False
'       Check_RemoveLarge.Visible = False
'       Cmd_WidthFilter.Visible = False
'        If OpenFileKey = 1 Then
'              DataClassify
'              ShowQQPlot
'        End If
'   End If
'
'End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        ChartTableLoader = 6
        frmChartTable.Show
    Else
        frmChartTable.Hide
    End If
End Sub

'Private Sub Cmd_WidthFilter_Click()
'    If OpenFileKey = 1 Then
'          DataClassify
'          ShowQQPlot
'    End If
'End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
        txt_file_df2 = ""
        frmOpenDF2.Show 1
        'cmdOpen.Enabled = False
        If txt_file_df2 = "" Then
           Exit Sub
        End If
        If Dir(CurDir + "\" + txt_file_df2) = "" Then
            Beep
            MsgBox "Input file does not exist, please try again"
            'OpenFileKey = 0
            Exit Sub
        End If
        If filelen(CurDir + "\" + txt_file_df2) = 0 Then
            Beep
            MsgBox "Input file is empty, please try again"
            'OpenFileKey = 0
            Exit Sub
        End If
        startup
        OpenFile
        If OpenFileKey = 1 Then
          DataClassify
          ShowQQPlot
        End If

End Sub

Private Sub cmdPrint_Click()
    ChartTransQQPlot.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub cmdSave_Click()

     Set CurChartSaveObj = ChartTransQQPlot
     ChartSaveAs


'        Dim ImageName As String
'        ImageName = InputBox("Please give a file name without extension", "Save chart as an image (JPG)", 1)
'       If ImageName <> "" Then
'             ChartTransQQPlot.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'       End If
End Sub


Private Sub Form_Activate()
    Set CurGraphicOBJ = ChartTransQQPlot
        CurWindowNum = 17
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CurGraphicOBJ = Nothing
    
    CurWindowNum = 17
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub

Private Sub Form_Load()
    txt_file_df2 = ""
    cmdOpen.Enabled = True
    OpenFileKey = 0
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
    
    CurWindowNum = 17
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

    'Read CASC parameter file
'    If Dir(txt_file_par + ".par") <> "" Then
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
        If (Me.Height > 360) And (Me.Height - 360 < Me.Width / 1.5) Then
           size = Me.Height - 500
        ElseIf (Me.Height > 360) And (Me.Height - 360 > Me.Width / 1.5) Then
           size = Me.Width / 1.5
        Else
           size = 0
        End If
        
        ChartTransQQPlot.Height = size
        ChartTransQQPlot.Width = 1.5 * size
End Sub


Public Sub startup()
        Dim I, J As Integer
        
        ChartTransQQPlot.Visible = False
        
        ChartTransQQPlot.ChartGroups(1).ChartType = oc2dTypePlot
        
        With ChartTransQQPlot.ChartGroups(1).Data
            .Layout = oc2dDataGeneral
            .IsBatched = True
            .NumSeries = 1
        End With
        
        ChartTransQQPlot.ChartGroups(1).Data.IsBatched = False
        
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
        If txt_file_df2 <> "" Then
            Open CurDir + "\" + txt_file_df2 For Input As FileNum
            OpenFileKey = 1
        Else
            MsgBox "The input file does not exist, please try again"
            OpenFileKey = 0
            Exit Sub
        End If
            
                
        HistoDataNum = 0
        OrderDataNum = 0
        OrderNumberCalcuKey = 0
        CurIsqrt = 0
        CurIendo = 0
        CurKcrit = 0
        
        'count the number of two type datasets from DF2 file
        While Not (EOF(FileNum))
            Line Input #FileNum, Temp
            If InStr(1, Temp, "=") > 0 Then
               LineSlope = Val(Mid(Temp, 20, 10))
               If InStr(1, Temp, "Order") > 0 Then
                  DepthOrderNumTemp = Val(Mid(Temp, 40, 2))
               Else
                 DepthOrderNumTemp = 1
               End If
            Else
                If Len(Temp) = 28 Then
                   HistoDataNum = HistoDataNum + 1
                End If
'Read Graph parameters: isqrt    iendo(irascs?)   kcrit
                If InStr(1, UCase(Temp), UCase("Graph parameters:")) > 0 Then
                   CurIsqrt = Val(Mid(Temp, 20, 2))
                   CurIendo = Val(Mid(Temp, 22, 2))
                   CurKcrit = Val(Mid(Temp, 24, 2))
                Else
                    If Len(Trim(Temp)) > 0 And Len(Temp) <> 28 Then
                       OrderDataNum = OrderDataNum + 1
                    End If
                    If OrderNumberCalcuKey = 0 Then
                       OrderNumber = (Len(Temp) - 8#) / 12#
                       OrderNumberCalcuKey = 1
                    End If
                End If
            End If
        Wend
        Close #FileNum
        
        'read the data into two dimensions
        HistoDataMin = 0
        HistoDataMax = 0
        QQplotDataMin = 0
        QQplotDataMax = 0
                
         If OrderDataNum = 0 Or HistoDataNum = 0 Then
            MsgBox "The input file " + CurDir + "\" + txt_file_df2 + " is empty or the format is  not correct, " + Chr$(13) + "Please check it and try again." + Chr$(13)
            OpenFileKey = 0
            Exit Sub
        End If

        
        ReDim OrderDataWell(1 To OrderDataNum) As Integer, OrderDataWellEventSeqNum(1 To OrderDataNum) As Integer, OrderData(1 To OrderDataNum, 1 To OrderNumber) As Double
        ReDim HistoDataSeqnum(1 To HistoDataNum) As Integer, HistoData(1 To HistoDataNum, 1 To 2) As Double
        Open CurDir + "\" + txt_file_df2 For Input As FileNum
        For I = 1 To OrderDataNum
            Line Input #FileNum, Temp
            OrderDataWell(I) = Val(Mid(Temp, 1, 4))
            OrderDataWellEventSeqNum(I) = Val(Mid(Temp, 5, 4))
            For J = 1 To OrderNumber
                OrderData(I, J) = Mid(Temp, 9 + (J - 1) * 12, 12)
            Next J
            
'            If I = 1 Then
'              MsgBox Temp + "   " + Str(OrderDataWell(I)) + "  " + Str(OrderDataWellEventSeqNum(I)) + "   " + Str(OrderData(I, 1)) + "   " + Str(OrderData(I, 2)) + "   " + Str(OrderData(I, 3)) + "   " + Str(OrderData(I, 4)) + "   " + Str(OrderData(I, 5))
'           End If

        Next I
        For I = 1 To HistoDataNum
           Line Input #FileNum, Temp
           HistoDataSeqnum(I) = Val(Mid(Temp, 1, 4))
           HistoData(I, 1) = Val(Mid(Temp, 5, 12))
           HistoData(I, 2) = Val(Mid(Temp, 17, 12))
           
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
        
        If HistoDataMin >= 0 Then
             StartValue = 0.9 * HistoDataMin
        Else
             StartValue = 1.1 * HistoDataMin
        End If
        If HistoDataMax >= 0 Then
             EndValue = 1.1 * HistoDataMax
        Else
            EndValue = 0.9 * HistoDataMax
        End If
        
        'Calculate Best-Fit Line data
        LineDataNum = 2
        ReDim FitLineData(1 To LineDataNum, 1 To 2) As Double
        'X value
        FitLineData(1, 1) = HistoDataMin
        FitLineData(1, 2) = HistoDataMax
        'Y value
        FitLineData(2, 1) = LineSlope * HistoDataMin
        FitLineData(2, 2) = LineSlope * HistoDataMax
       
        
        
'        If HistoDataMin < 0 Then
'            StartValue = -(Abs(HistoDataMin) + ClassifyScale - (Int(Abs(HistoDataMin)) Mod ClassifyScale))
'        End If
'        If HistoDataMin > 0 Then
'            StartValue = Abs(HistoDataMin) + ClassifyScale - (Int(Abs(HistoDataMin)) Mod ClassifyScale)
'        End If
'
'        If HistoDataMax < 0 Then
'            EndValue = -(Abs(HistoDataMax) - (Int(Abs(HistoDataMax)) Mod ClassifyScale))
'        End If
'        If HistoDataMax > 0 Then
'            EndValue = Abs(HistoDataMax) + ClassifyScale - (Int(Abs(HistoDataMax)) Mod ClassifyScale)
'        End If
        
        
''        If HistoDataMax >= 0 Then
''            If HistoDataMax >= Abs(HistoDataMin) Then
''                temp = HistoDataMax
''            Else
''                temp = Abs(HistoDataMin)
''            End If
''        Else
''            temp = Abs(HistoDataMin)
''        End If
'        'The DepthWidth and RemoveOutsideKey are two CASC parameters about Depth Differences given by the User
''        If DepthWidth > temp Then
''            temp = DepthWidth
''        End If
'
'
'        StepCount = 0
'        For I = StartValue To EndValue Step ClassifyScale
'             StepCount = StepCount + 1
'        Next I
'
'       ReDim ValueTemp(0 To StepCount), ValueCount(0 To StepCount), ValueStep(0 To StepCount)
'       StepCount = 0
'       ValueCount(StepCount) = 0
'       ValueStep(StepCount) = 0
'       For I = StartValue To EndValue Step ClassifyScale
'             StepCount = StepCount + 1
'             ValueTemp(StepCount) = I
'             ValueCount(StepCount) = 0
'             ValueStep(StepCount) = 0
'        Next I
'
'        'The DepthWidth and RemoveOutsideKey are two CASC parameters about Depth Differences given by the User
'        'Adjust Step and Value according to the User
'        ReDim HistoDataTemp(1 To HistoDataNum) As Double
'        Dim CountTemp As Integer
'        WTemp = Val(Text_Width.Text)
'        If Check_UseWidthFilter.Value = 1 And WTemp > 0 Then
'            If WTemp < HistoDataMax Or Abs(HistoDataMin) > WTemp Then
'                'Special processing
'                'Change statistic value
'                For J = 1 To HistoDataNum
'                          If Abs(HistoData(J, 2)) > WTemp Then
'                               If HistoData(J, 2) < 0 Then
'                                    If Check_RemoveLarge.Value = 1 Then
'                                        HistoDataTemp(J) = StartValue - 1
'                                    Else
'                                        HistoDataTemp(J) = -WTemp - ClassifyScale - 0.001
'                                    End If
'                                Else
'                                    If Check_RemoveLarge.Value = 1 Then
'                                        HistoDataTemp(J) = EndValue + 1
'                                    Else
'                                        HistoDataTemp(J) = WTemp + ClassifyScale - 0.001
'                                    End If
'                                End If
'                          Else
'                               HistoDataTemp(J) = HistoData(J, 2)
'                          End If
'                Next J
'                For J = 1 To HistoDataNum
'                     For K = 1 To StepCount - 1
'                          If (K = 1 And HistoDataTemp(J) = ValueTemp(K)) Or (HistoDataTemp(J) = WTemp And HistoDataTemp(J) = ValueTemp(K)) Then
'                               'StartValue processing
'                               ValueStep(0) = ValueTemp(K)
'                               ValueCount(0) = ValueCount(0) + 1
'                          Else
'                                ValueStep(K) = ValueTemp(K + 1)
'                                If ((HistoDataTemp(J) > ValueTemp(K)) And (HistoDataTemp(J) <= ValueTemp(K + 1))) Then
'                                     ValueCount(K) = ValueCount(K) + 1
'                                End If
'                          End If
'                    Next K
'                Next J
'            Else
'                'Normal processing
'                For J = 1 To HistoDataNum
'                     For K = 1 To StepCount - 1
'                          If (K = 1 And HistoData(J, 2) = ValueTemp(K)) Then
'                               'StartValue processing
'                               ValueStep(K - 1) = ValueTemp(K)
'                               ValueCount(K - 1) = ValueCount(0) + 1
'                          Else
'                                ValueStep(K) = ValueTemp(K + 1)
'                                If ((HistoData(J, 2) > ValueTemp(K)) And (HistoData(J, 2) <= ValueTemp(K + 1))) Then
'                                     ValueCount(K) = ValueCount(K) + 1
'                                End If
'                          End If
'                    Next K
'                Next J
'            End If
'        Else
'            'Normal processing
'            For J = 1 To HistoDataNum
'                 For K = 1 To StepCount - 1
'                          If K = 1 And HistoData(J, 2) = ValueTemp(K) Then
'                               'StartValue processing
'                               ValueStep(0) = ValueTemp(K)
'                               ValueCount(0) = ValueCount(0) + 1
'                          Else
'                                ValueStep(K) = ValueTemp(K + 1)
'                                If ((HistoData(J, 2) > ValueTemp(K)) And (HistoData(J, 2) <= ValueTemp(K + 1))) Then
'                                     ValueCount(K) = ValueCount(K) + 1
'                                End If
'                          End If
'                Next K
'            Next J
'        End If


End Sub
        
Private Sub ShowQQPlot()
        Dim gridspace As Double
        Dim I, J As Integer
        Dim MaxCount As Integer, SumCount As Integer
        
        ChartTransQQPlot.Visible = False
        
        'ChartTransQQPlot.ChartGroups(1).Styles(1).Symbol.size = 0
        'ChartTransQQPlot.ChartGroups(1).Styles(1).Line.Width = 2
        'ChartTransQQPlot.ChartGroups(1).Styles(1).Line.Color = 0
        
        ChartTransQQPlot.ChartGroups(1).ChartType = oc2dTypePlot
        
'        If file_type = "DE1" Then
'            title_str = "Ranked"
'        ElseIf file_type = "DE2" Then
'            title_str = "Scaled"
'        End If
        
'        If Unitsflag = 0 Then
'           ChartTransQQPlot.ChartArea.Axes("Y").Title.Text = "Depth (meters)"
'        Else
'           ChartTransQQPlot.ChartArea.Axes("Y").Title.Text = "Depth (feet)"
'        End If

'        If Unitsflag = 0 Then
'           ChartTransQQPlot.ChartArea.Axes("X").Title.Text = "Depth Difference (m)"
'        Else
'           ChartTransQQPlot.ChartArea.Axes("X").Title.Text = "Depth Difference (feet)"
'        End If

        ChartTransQQPlot.ChartArea.Axes("X").Title.Text = ""   '"SQRT & Ave-Transformed Depth Differences"
        ChartTransQQPlot.ChartArea.Axes("Y").Title.Text = ""   '"Normal Quantile"
        
        MaxCount = 0
        SumCount = 0
        With ChartTransQQPlot.ChartGroups(1).Data
            .Layout = oc2dDataGeneral
            .IsBatched = True
            .NumSeries = 1
            .NumPoints(1) = HistoDataNum
            
             frmChartTable.listOfData.Clear
             frmChartTable.listOfData.AddItem "   No.  " + "Transformed Depth Difference" + "  Normal Quantile"
             
             For I = 1 To HistoDataNum
                .X(1, I) = HistoData(I, 2)
                .Y(1, I) = HistoData(I, 1)
                
                frmChartTable.listOfData.AddItem Space(5 - Len(Str(I))) + Str(I) _
                            + Space(23 - Len(Format(HistoData(I, 2), "0.000"))) _
                            + Format(HistoData(I, 2), "0.000") _
                            + Space(20 - Len(Format(HistoData(I, 1), "0.000"))) _
                            + Format(HistoData(I, 1), "0.000")
                
             Next I
             
             frmChartTable.listOfData.AddItem Space(5)
             frmChartTable.listOfData.AddItem "Total Points: " + Str(HistoDataNum)
             frmChartTable.listOfData.AddItem "Slope of Best-Fit: " + Format(LineSlope, "###0.0000") + "; Order = " + Trim(Str(DepthOrderNumTemp))
             
        End With
        
       With ChartTransQQPlot.ChartGroups(1).Data
            .Layout = oc2dDataGeneral
            .IsBatched = True
            .NumSeries = 2
            .NumPoints(2) = LineDataNum
             For I = 1 To LineDataNum
                .X(2, I) = FitLineData(1, I)
                .Y(2, I) = FitLineData(2, I)
             Next I
        End With
        
         'Verify Y Min and Max value according to the slope data
'        'X value
'        FitLineData(1, 1) = HistoDataMin
'        FitLineData(1, 2) = HistoDataMax
        'Y value
        If FitLineData(2, 1) <= QQplotDataMin Then
            QQplotDataMin = FitLineData(2, 1)
        End If
        If FitLineData(2, 2) > QQplotDataMax Then
            QQplotDataMax = FitLineData(2, 2)
        End If
       
        If QQplotDataMin >= 0 Then
            ChartTransQQPlot.ChartArea.Axes("Y").Min.Value = 0.9 * QQplotDataMin
        Else
            ChartTransQQPlot.ChartArea.Axes("Y").Min.Value = 1.1 * QQplotDataMin
        End If
        If QQplotDataMax >= 0 Then
            ChartTransQQPlot.ChartArea.Axes("Y").Max.Value = 1.1 * QQplotDataMax
        Else
            ChartTransQQPlot.ChartArea.Axes("Y").Max.Value = 0.9 * QQplotDataMax
        End If
        
        
        ChartTransQQPlot.ChartArea.Axes("x").Min.Value = StartValue
        ChartTransQQPlot.ChartArea.Axes("x").Max.Value = EndValue
        
        'to make two groups of chart with the same scale
'        ChartTransQQPlot.ChartArea.Axes("y2").Min.Value = ChartTransQQPlot.ChartArea.Axes("y").Min.Value
'        ChartTransQQPlot.ChartArea.Axes("y2").Max.Value = ChartTransQQPlot.ChartArea.Axes("Y").Max.Value
          
        ChartTransQQPlot.ChartArea.Axes("x").Origin.Value = 0
        ChartTransQQPlot.ChartArea.Axes("y").Origin.Value = 0
        
        
'         Gridspace = ChartTransQQPlot.ChartArea.Axes("Y").Max.Value / 10
'         ChartTransQQPlot.ChartArea.Axes("x").MajorGrid.Spacing = Gridspace
        
        ChartTransQQPlot.ChartLabels(1).Text.Text = txt_file_df2
        ChartTransQQPlot.ChartLabels(1).Anchor = oc2dAnchorNorthWest
        ChartTransQQPlot.ChartLabels(1).AttachDataCoord.X = ChartTransQQPlot.ChartArea.Axes("x").Max.Value
        ChartTransQQPlot.ChartLabels(1).AttachDataCoord.Y = ChartTransQQPlot.ChartArea.Axes("Y").Min.Value
        ChartTransQQPlot.ChartLabels(1).Offset = 0
        
'        *df2: "add (m^0.5)" if isqrt=1 on horizontal scale
        If CurIsqrt = 1 Then
            If Unitsflag = 0 Then
               ChartTransQQPlot.ChartLabels(2).Text.Text = "SQRT & Ave-Transformed Depth Difference (m^0.5)"
            Else
               ChartTransQQPlot.ChartLabels(2).Text.Text = "SQRT & Ave-Transformed Depth Difference (feet^0.5)"
            End If
        Else
            If Unitsflag = 0 Then
               ChartTransQQPlot.ChartLabels(2).Text.Text = "SQRT & Ave-Transformed Depth Difference (m)"
            Else
               ChartTransQQPlot.ChartLabels(2).Text.Text = "SQRT & Ave-Transformed Depth Difference (feet)"
            End If
        End If

        'ChartTransQQPlot.ChartLabels(2).Text.Text = "SQRT & Ave-Transformed Depth Difference"
        ChartTransQQPlot.ChartLabels(2).Anchor = oc2dAnchorSouth
        ChartTransQQPlot.ChartLabels(2).AttachDataCoord.X = 0
        ChartTransQQPlot.ChartLabels(2).AttachDataCoord.Y = ChartTransQQPlot.ChartArea.Axes("Y").Min.Value
        ChartTransQQPlot.ChartLabels(2).Offset = 10
        
        ChartTransQQPlot.ChartLabels(3).Text.Text = "Normal Quantile"
        ChartTransQQPlot.ChartLabels(3).Anchor = oc2dAnchorWest
        ChartTransQQPlot.ChartLabels(3).Rotation = oc2dRotate90Degrees
        ChartTransQQPlot.ChartLabels(3).AttachDataCoord.X = ChartTransQQPlot.ChartArea.Axes("x").Min.Value
        ChartTransQQPlot.ChartLabels(3).AttachDataCoord.Y = 0
        ChartTransQQPlot.ChartLabels(3).Offset = 10
        
        ChartTransQQPlot.ChartLabels(4).Text.Text = "y=" + Format(LineSlope, "###0.0000") + "x"
        ChartTransQQPlot.ChartLabels(4).Anchor = oc2dAnchorWest
        ChartTransQQPlot.ChartLabels(4).Rotation = oc2dRotateNone
        ChartTransQQPlot.ChartLabels(4).AttachDataCoord.X = FitLineData(1, LineDataNum)
        ChartTransQQPlot.ChartLabels(4).AttachDataCoord.Y = FitLineData(2, LineDataNum)
        ChartTransQQPlot.ChartLabels(4).Offset = 25
       
        ChartTransQQPlot.Header.Text = "Transformed Depth Differences; Order=" + Trim(Str(DepthOrderNumTemp))
        
        
        ChartTransQQPlot.ChartGroups(1).Data.IsBatched = False
        
        ChartTransQQPlot.Visible = True


End Sub


'Private Sub List_Well_Click()
'        Dim gridspace As Double
'        Dim I, J As Integer
'        Dim well As Integer
'        Dim Emin, Emax, E_temp As Double
'        Dim Name_temp As String
'        Dim title_str As String
'        Dim tempfrom As Integer
'        ChartTransQQPlot.Visible = False
'
'        well = List_Well.ListIndex + 1
'        'MsgBox "Well No. == " + Str(well)
'        'LabelWell.Visible = False
'        'List_Well.Visible = False
'
'
'        ChartTransQQPlot.ChartGroups(1).Styles(1).Symbol.size = 0
'        'ChartTransQQPlot.ChartGroups(1).Styles(1).Line.Width = 2
'        'ChartTransQQPlot.ChartGroups(1).Styles(1).Line.Color = 0
'        ChartTransQQPlot.ChartGroups(1).ChartType = oc2dTypePlot
'
'        If file_type = "DE1" Then
'            title_str = "Ranked"
'        ElseIf file_type = "DE2" Then
'            title_str = "Scaled"
'        End If
'
'        'ChartTransQQPlot.Header.Text.Text = "Scattergram   " + txt_file_DF2 + "  (Well No. " + Str(well) + ")"
'        ChartTransQQPlot.ChartArea.Axes("x").Title.Text = title_str + " Optimum Sequence of Events"
'
'        ChartTransQQPlot.ChartLabels(1).Text.Text = nameOfWell(well)
'        With ChartTransQQPlot.ChartGroups(1).Data
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
'        ChartTransQQPlot.ChartArea.Axes("y").Min.Value = Int(Emin) + 1
'        ChartTransQQPlot.ChartArea.Axes("Y").Max.Value = Int(Emax) + 1
'        'ChartTransQQPlot.ChartArea.Axes("y2").Min.Value = Int(Emin) + 1
'        'ChartTransQQPlot.ChartArea.Axes("y2").Max.Value = Int(Emax) + 1
'
'        'to make two groups of chart with the same scale
'        ChartTransQQPlot.ChartArea.Axes("y2").Min.Value = ChartTransQQPlot.ChartArea.Axes("y").Min.Value
'        ChartTransQQPlot.ChartArea.Axes("y2").Max.Value = ChartTransQQPlot.ChartArea.Axes("Y").Max.Value
'
'        ChartTransQQPlot.ChartArea.Axes("x").Origin.Value = 0
'        ChartTransQQPlot.ChartArea.Axes("y").Origin.Value = Int(Emax) + 1
'
'        ChartTransQQPlot.ChartArea.Axes("x").Min.Value = 0
'        ChartTransQQPlot.ChartArea.Axes("x").Max.Value = x_Value(tempfrom + NumOfEvent(well)) + 1
'
'        If file_type = "DE2" Then
'         gridspace = x_Value(tempfrom + NumOfEvent(well)) / 10
'         ChartTransQQPlot.ChartArea.Axes("x").MajorGrid.Spacing = gridspace
'        End If
'
'        'ChartTransQQPlot.ChartArea.Axes("y2").Origin.Value = Int(Emin) - 1
'
'        ChartTransQQPlot.ChartGroups(1).Data.IsBatched = False
'
'
'        ChartTransQQPlot.ChartGroups(2).Styles(1).Symbol.size = 5
'        ChartTransQQPlot.ChartGroups(2).ChartType = oc2dTypePlot
'
'
'        With ChartTransQQPlot.ChartGroups(1).Data
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
'        ChartTransQQPlot.ChartGroups(1).Data.IsBatched = False
'        'ChartTransQQPlot.Visible = True
'        ChartTransQQPlot.Visible = True
'        If Unitsflag = 0 Then
'           ChartTransQQPlot.ChartArea.Axes("Y").Title.Text = "Depth (meters)"
'        Else
'           ChartTransQQPlot.ChartArea.Axes("Y").Title.Text = "Depth (feet)"
'        End If
'End Sub

'Private Sub List_ClassifyScale_Click()
'    ClassifyScale = Val(List_ClassifyScale.Text)
'    If OpenFileKey = 1 Then
'          DataClassify
'          ShowQQPlot
'    End If
'End Sub
