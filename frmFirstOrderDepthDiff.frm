VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmFirstOrderDepthDiff 
   Caption         =   "Depth Difference Optimum Sequence"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   12690
   Icon            =   "frmFirstOrderDepthDiff.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "View Data"
      Height          =   252
      Left            =   3660
      TabIndex        =   5
      Top             =   90
      Width           =   1092
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Default         =   -1  'True
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1080
      TabIndex        =   2
      Top             =   90
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2640
      TabIndex        =   1
      Top             =   90
      Width           =   732
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   252
      Left            =   1920
      TabIndex        =   0
      Top             =   90
      Width           =   612
   End
   Begin OlectraChart2D.Chart2D ChartDDPlot 
      Height          =   7920
      Left            =   60
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   11955
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   21087
      _ExtentY        =   13970
      _StockProps     =   0
      ControlProperties=   "frmFirstOrderDepthDiff.frx":0442
   End
End
Attribute VB_Name = "frmFirstOrderDepthDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UnitType As String, Unitsflag As Integer, OpenFileKey As Integer

Dim OPSeqData() As Integer, EventNum() As Integer, DDData() As Double
Dim DataNum As Integer, OrderDataNum As Integer
Dim DDDataMin As Double, DDDataMax As Double
Dim ESeqDataMin As Double, ESeqDataMax As Double

Dim kcrito As Integer, icnt As Integer, imean As Integer, isqrt As Integer, iendo As Integer, kcrit As Integer, irascs As Integer
Dim decrito As Double
Dim CurIsqrt  As Integer, CurIendo As Integer, CurKcrit As Integer


Private Sub Check1_Click()
    If Check1.Value = 1 Then
       ChartTableLoader = 5
       frmChartTable.Show
    Else
        frmChartTable.Hide
    End If
End Sub

'Private Sub Cmd_WidthFilter_Click()
'    If OpenFileKey = 1 Then
'          DataClassify
'          ShowDDPlot
'    End If
'End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
       txt_file_di1 = ""
        frmOpenDI1.Show 1
        'cmdOpen.Enabled = False
        
        If txt_file_di1 = "" Then
             Exit Sub
        End If
        If Dir(CurDir + "\" + txt_file_di1) = "" Then
            Beep
            MsgBox "Input file does not exist, please try again"
            'OpenFileKey = 0
            Exit Sub
        End If
        If filelen(CurDir + "\" + txt_file_di1) = 0 Then
            Beep
            MsgBox "Input file is empty, please try again"
            'OpenFileKey = 0
            Exit Sub
        End If
        startup
        OpenFile
        If OpenFileKey = 1 Then
'          DataClassify
          ShowDDPlot
        End If

End Sub

Private Sub cmdPrint_Click()
      ChartDDPlot.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
          'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub cmdSave_Click()
     
     Set CurChartSaveObj = ChartDDPlot
     ChartSaveAs

'        Dim ImageName As String
'        ImageName = InputBox("Please give a file name without extension", "Save chart as an image (JPG)", 1)
'       If ImageName <> "" Then
'             ChartDDPlot.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'       End If
End Sub


Private Sub Form_Activate()
    Set CurGraphicOBJ = ChartDDPlot
        CurWindowNum = 14
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    txt_file_di1 = ""
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

    
    CurWindowNum = 14
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

    'Read CASC parameter file
'    If Dir(txt_file_di1 + ".di1") <> "" Then
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
    
    If Dir(txt_file_di1 + ".di1" <> 0) Then
        gFileNum = FreeFile
        Open CurDir & "\" + txt_file_di1 + ".di1" For Input As gFileNum
        
        'Get the three records.
        
        If EOF(gFileNum) = False Then
              Line Input #gFileNum, Temp
        End If
        kcrito = Mid(Temp, 1, 2)
        decrito = Mid(Temp, 3, 8)
        icnt = Mid(Temp, 11, 2)
        imean = Mid(Temp, 13, 2)
        isqrt = Mid(Temp, 15, 2)
        iendo = Mid(Temp, 17, 2)
        kcrit = Mid(Temp, 19, 2)
        irascs = Mid(Temp, 21, 2)
'        Text_Width.Text = Trim(decrito)
'        Text_MaxiOrder.Text = Trim(kcrit)
'        Text_MaxiSequence.Text = Trim(kcrito)
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
           size = Me.Height - 360
        ElseIf (Me.Height > 360) And (Me.Height - 360 > Me.Width / 1.5) Then
           size = Me.Width / 1.5
        Else
           size = 0
        End If
        
        ChartDDPlot.Height = size
        ChartDDPlot.Width = 1.5 * size
End Sub


Public Sub startup()
        Dim I, J As Integer
        
        ChartDDPlot.Visible = False
        
        ChartDDPlot.ChartGroups(1).ChartType = oc2dTypePlot
        
        With ChartDDPlot.ChartGroups(1).Data
            .Layout = oc2dDataGeneral
            .IsBatched = True
            .NumSeries = 1  '**** changed to 2 from 1
         '   .NumPoints(1) = 150
           
        End With
        
        ChartDDPlot.ChartGroups(1).Data.IsBatched = False
        
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
        
        Dim ArrayNum As Integer
        Dim tnum
        
        FileNum = FreeFile
        If txt_file_di1 <> "" Then
            Open CurDir + "\" + txt_file_di1 For Input As FileNum
            OpenFileKey = 1
        Else
            MsgBox "The input file: " + CurDir + "\" + txt_file_di1 + " does not exist, please check it and try again."
            OpenFileKey = 0
            Exit Sub
        End If
            
        CurIsqrt = 0
        CurIendo = 0
        CurKcrit = 0
                    
        DataNum = 0
        
        'count the number of  datasets from DI1 file
        While Not (EOF(FileNum))
            Line Input #FileNum, Temp
            If Len(Temp) > 0 And InStr(1, UCase(Temp), UCase("Graph parameters:")) = 0 Then
               DataNum = DataNum + 1
            End If
        Wend
        Close #FileNum
        
        'read the data into three dimensions
        ESeqDataMin = 0
        ESeqDataMax = 0
        DDDataMin = 0
        DDDataMax = 0
        
        If DataNum = 0 Then
            MsgBox "The input file: " + CurDir + "\" + txt_file_di1 + " is empty, please check it."
            OpenFileKey = 0
            Exit Sub
        End If
        
        ReDim OPSeqData(1 To DataNum) As Integer, EventNum(1 To DataNum) As Integer, DDData(1 To DataNum) As Double
        
        Open CurDir + "\" + txt_file_di1 For Input As FileNum
        For I = 1 To DataNum
            Line Input #FileNum, Temp
            OPSeqData(I) = Mid(Temp, 1, 4)
            EventNum(I) = Mid(Temp, 5, 4)
            DDData(I) = Mid(Temp, 9, 10)
          'Compare DDplotData Value
           If OPSeqData(I) > ESeqDataMax Then
              ESeqDataMax = OPSeqData(I)
           End If
           If OPSeqData(I) < ESeqDataMin Then
              ESeqDataMin = OPSeqData(I)
           End If
           'Compare HistoData Value of 1st Order
           If DDData(I) > DDDataMax Then
              DDDataMax = DDData(I)
           End If
           If DDData(I) < DDDataMin Then
              DDDataMin = DDData(I)
           End If
           
'           If I = 1 Then
'              MsgBox Temp + "   " + Str(HistoDataSeqnum(I)) + "  " + Str(HistoData(I, 1)) + "   " + Str(HistoData(I, 2))
'           End If
            
'            If I = 1 Then
'              MsgBox Temp + "   " + Str(OrderDataWell(I)) + "  " + Str(OrderDataWellEventSeqNum(I)) + "   " + Str(OrderData(I, 1)) + "   " + Str(OrderData(I, 2)) + "   " + Str(OrderData(I, 3)) + "   " + Str(OrderData(I, 4)) + "   " + Str(OrderData(I, 5))
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
                MsgBox "The graphic parameters were not included in input file: " + CurDir + "\" + txt_file_di1 + Chr$(13) + "Current CASCW.EXE may be an old version file, please check it."
                'OpenFileKey = 0
                Close #FileNum
                Exit Sub
            End If
        Else
            MsgBox "The graphic parameters were not included in input file: " + CurDir + "\" + txt_file_di1 + Chr$(13) + "Current CASCW.EXE may be an old version file, please check it."
            'OpenFileKey = 0
            Close #FileNum
            Exit Sub
        End If
        Close #FileNum
           
        
        'UnitFlag:  Meter or Feet
        'Unitsflag = Val(Right(Temp, 1))
        ' MsgBox temp
        
        'MsgBox "The actural num of DDpotNUM and HistoDataNum are  " + Str(OrderDataNum) + ", " + Str(HistoDataNum)
        
End Sub

        
Private Sub ShowDDPlot()
        Dim gridspace As Double
        Dim I, J As Integer
        Dim MaxCount As Integer, SumCount As Integer
        
        ChartDDPlot.Visible = False
        
        'ChartDDPlot.ChartGroups(1).Styles(1).Symbol.size = 0
        'ChartDDPlot.ChartGroups(1).Styles(1).Line.Width = 2
        'ChartDDPlot.ChartGroups(1).Styles(1).Line.Color = 0
        
        ChartDDPlot.ChartGroups(1).ChartType = oc2dTypePlot
        
'        If file_type = "DE1" Then
'            title_str = "Ranked"
'        ElseIf file_type = "DE2" Then
'            title_str = "Scaled"
'        End If
        
'        If Unitsflag = 0 Then
'           ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Depth (meters)"
'        Else
'           ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Depth (feet)"
'        End If

'        If Unitsflag = 0 Then
'           ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Cumulative Depth Difference (m)"
'        Else
'           ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Cumulative Depth Difference (feet)"
'        End If
'
'        ChartDDPlot.ChartArea.Axes("X").Title.Text = "Ranked Optimum Sequence of Events"
        
        MaxCount = 0
        SumCount = 0
        With ChartDDPlot.ChartGroups(1).Data
            .Layout = oc2dDataGeneral
            .IsBatched = True
            .NumSeries = 1
            .NumPoints(1) = DataNum
            
             frmChartTable.listOfData.Clear
             frmChartTable.listOfData.AddItem " Op. Sequence  " + " Events Number " + "  Cumu. Depth Difference"
             
             For I = 1 To DataNum
                .X(1, I) = OPSeqData(I)
                .Y(1, I) = DDData(I)
                
                frmChartTable.listOfData.AddItem Space(10 - Len(Str(I))) + Str(I) _
                            + Space(15 - Len(Str(EventNum(I)))) _
                            + Str(EventNum(I)) _
                            + Space(20 - Len(Format(DDData(I), "0.0000"))) _
                            + Format(DDData(I), "0.0000")
                
             Next I
             
             frmChartTable.listOfData.AddItem Space(5)
             frmChartTable.listOfData.AddItem "Total Points: " + Str(DataNum)
             
        End With
        
'*.di1 and *di2: on vertical scale "m" replaced by "m^0.5" if isqrt=1; on
'horizontal scale "ranked" replaced by "scaled" if iendo(irascs?)=1
        If CurIsqrt = 1 Then
            If Unitsflag = 0 Then
               ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Cumulative Depth Difference (m^0.5)"
            Else
               ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Cumulative Depth Difference (feet^0.5)"
            End If
        Else
            If Unitsflag = 0 Then
               ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Cumulative Depth Difference (m)"
            Else
               ChartDDPlot.ChartArea.Axes("Y").Title.Text = "Cumulative Depth Difference (feet)"
            End If
        End If
        If CurIendo = 1 Then
            ChartDDPlot.ChartArea.Axes("X").Title.Text = "Scaled Optimum Sequence of Events"
        Else
            ChartDDPlot.ChartArea.Axes("X").Title.Text = "Ranked Optimum Sequence of Events"
        End If
'At top of .di2 graph: Replace "First" by "Higher" and insert "Window";
'after "Higher Order Window Depth Differences" add: (Maximum order =
'kcrit)
        If InStr(1, UCase(txt_file_di1), UCase(".di1")) > 0 Then
           ChartDDPlot.Header.Text = "First Order Depth Differences"
        Else
           ChartDDPlot.Header.Text = "Higher Order Window Depth Differences (Maximum Order = " + Trim(Str(CurKcrit)) + ")"
        End If
        
        
        
'        If DDplotDataMin >= 0 Then
'            ChartDDPlot.ChartArea.Axes("Y").Min.Value = 0.9 * DDplotDataMin
'        Else
'            ChartDDPlot.ChartArea.Axes("Y").Min.Value = 1.1 * DDplotDataMin
'        End If
'        If DDplotDataMax >= 0 Then
'            ChartDDPlot.ChartArea.Axes("Y").Max.Value = 1.1 * DDplotDataMax
'        Else
'            ChartDDPlot.ChartArea.Axes("Y").Max.Value = 0.9 * DDplotDataMax
'        End If

         ChartDDPlot.ChartArea.Axes("Y").Min.Value = 0
         ChartDDPlot.ChartArea.Axes("Y").Max.Value = 1.1 * DDDataMax
         ChartDDPlot.ChartArea.Axes("Y").IsReversed = True
        
        ChartDDPlot.ChartArea.Axes("x").Min.Value = 0
        ChartDDPlot.ChartArea.Axes("x").Max.Value = 1.1 * ESeqDataMax
        
        'to make two groups of chart with the same scale
'        ChartDDPlot.ChartArea.Axes("y2").Min.Value = ChartDDPlot.ChartArea.Axes("y").Min.Value
'        ChartDDPlot.ChartArea.Axes("y2").Max.Value = ChartDDPlot.ChartArea.Axes("Y").Max.Value
          
        ChartDDPlot.ChartArea.Axes("x").Origin.Value = ChartDDPlot.ChartArea.Axes("x").Min.Value
        ChartDDPlot.ChartArea.Axes("y").Origin.Value = ChartDDPlot.ChartArea.Axes("Y").Max.Value
        
        
'         Gridspace = ChartDDPlot.ChartArea.Axes("Y").Max.Value / 10
'         ChartDDPlot.ChartArea.Axes("x").MajorGrid.Spacing = Gridspace
        
        ChartDDPlot.ChartLabels(1).Text.Text = txt_file_di1
        ChartDDPlot.ChartLabels(1).Anchor = oc2dAnchorSouthWest
        ChartDDPlot.ChartLabels(1).AttachDataCoord.X = ChartDDPlot.ChartArea.Axes("x").Max.Value
        ChartDDPlot.ChartLabels(1).AttachDataCoord.Y = ChartDDPlot.ChartArea.Axes("Y").Min.Value
        
         ChartDDPlot.ChartGroups(1).Data.IsBatched = False
        
         ChartDDPlot.Visible = True


End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set CurGraphicOBJ = Nothing
    
    CurWindowNum = 14
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub
