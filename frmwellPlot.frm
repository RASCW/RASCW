VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmwellPlot1 
   Caption         =   " CASC Correlation - Ranking"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   Icon            =   "frmwellPlot.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   1620
      TabIndex        =   10
      Top             =   120
      Width           =   732
   End
   Begin VB.CheckBox chkTable 
      Caption         =   "Data Table"
      Height          =   315
      Left            =   3900
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox list_flattening 
      Height          =   255
      Left            =   7980
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   252
      Left            =   3060
      TabIndex        =   6
      Top             =   120
      Width           =   732
   End
   Begin VB.ListBox List_well_width 
      Height          =   255
      Left            =   6060
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2340
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   900
      TabIndex        =   2
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   252
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin OlectraChart2D.Chart2D ChartWell 
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11745
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   20722
      _ExtentY        =   14203
      _StockProps     =   0
      ControlProperties=   "frmwellPlot.frx":0442
   End
   Begin VB.Label lblflattening 
      Caption         =   "Flattening"
      Height          =   255
      Left            =   7140
      TabIndex        =   8
      Top             =   150
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Well width"
      Height          =   255
      Left            =   5220
      TabIndex        =   5
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmwellPlot1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileNum As Integer
Dim Filename As String
Dim Temp As String
Dim tempfrom As Integer
Dim tempto As Integer
Dim tempstr As String
Dim tempstr1 As String
Dim AllmaxDepth As Integer
Dim allmindepth As Integer
Dim ObsDepth() As Integer
Dim MinObs() As Integer
Dim MaxObs() As Integer
Dim probDepth() As Integer
Dim MinDepth() As Integer
Dim maxDepth() As Integer
Dim well_width As Double
Dim newwell() As Integer
Dim ratio() As Double
Dim AVESD As Double
Dim flattening_num As Integer
Dim Unitsflag As Integer


Public Sub startup()
Dim I, J As Integer

    ChartWell.ChartGroups(1).ChartType = oc2dTypeBar
    
    ChartWell.ChartArea.Bar.ClusterWidth = well_width * 100   'old value =2 * well_width * 100
    With ChartWell.ChartGroups(1).Data
        .Layout = oc2dDataArray
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
        .NumPoints(1) = max_well_num
       
    End With


End Sub

Private Sub chkTable_Click()
    Dim I, J As Integer
    If Not (frmDataTable Is Nothing) Then
       Unload frmDataTable
    End If
    If chkTable.Value = 1 Then
         'get values for data table()
         For J = 1 To max_event_num 'ratio values for data table
             TableRatio(J) = ratio(J)
         Next J
         ' data table values
         For I = 1 To max_well_num
            For J = 1 To max_event_num
                TableObs(I, J) = ObsDepth(I, J)
                TableMinObs(I, J) = MinObs(I, J)
                TableMaxObs(I, J) = MaxObs(I, J)
                TableProb(I, J) = probDepth(I, J)
                TableMinProb(I, J) = MinDepth(I, J)
                TableMaxProb(I, J) = maxDepth(I, J)
            Next J
        Next I
        frmDataTable.Show
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmwellPlot1
End Sub

Private Sub cmdModify_Click()
    
    modifyflag = 1
    
    Unload frmwellPlot1
    Load frmwellPlot1
    'ChartWell.Visible = False
    
    frmWells.Show 1
    well_width = 0.15
    ' set initial for flattening number = 0
    flattening_num = 0
    
    startup
    OpenFile
    ' add items to list
    
    For I = 1 To 6
    List_well_width.AddItem Str(I * 5) + "%"
    Next I
    List_well_width.Visible = True
    Label1.Visible = True
    
    'put events into flattening list box
     list_flattening.Clear
    For J = 1 To max_event_num
      list_flattening.AddItem Space(2 * (4 - Len(Str(event_numbers(J))))) + Str(event_numbers(J)) + Space(5) + event_names(J)
    Next J
    lblflattening.Visible = True
    list_flattening.Visible = True
    chkTable.Visible = True
End Sub

Private Sub cmdOpen_Click()
    Dim J, I As Integer
    
    modifyflag = 0 ' for modifying chart
    
    file_type = "Ca1"
    frmOpenRan1.Show 1
    'cmdOpen.Enabled = False
    
    If txt_file_ran1 = "" Then
      Beep
      Exit Sub
    End If
    
    frmWells.Show 1
    
    If txt_file_ran1 = "" Or testing_flag = 0 Then
      Beep
    Exit Sub
    End If
    ' set initial for flattening number = 0
    flattening_num = 0
    
    startup
    OpenFile
    ' add items to list
    
    For I = 1 To 6
    List_well_width.AddItem Str(I * 5) + "%"
    Next I
    List_well_width.Visible = True
    Label1.Visible = True
    
    
     'put events into flattening list box
     list_flattening.Clear
    For J = 1 To max_event_num
      list_flattening.AddItem Space(2 * (4 - Len(Str(event_numbers(J))))) + Str(event_numbers(J)) + Space(5) + event_names(J)
    Next J
    lblflattening.Visible = True
    list_flattening.Visible = True
    chkTable.Visible = True
        
End Sub

Private Sub cmdPrint_Click()
    ChartWell.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
    'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub Command1_Click()

     Set CurChartSaveObj = ChartWell
     ChartSaveAs


'    Dim ImageName As String
'    ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'    If ImageName <> "" Then
'        ChartWell.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'    End If
End Sub

Private Sub Form_Activate()
    Set CurGraphicOBJ = ChartWell
        CurWindowNum = 12
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
        List_well_width.Visible = False
        cmdModify.Enabled = False
        Label1.Visible = False
        well_width = 0.15
        ChartWell.Visible = False
        lblflattening.Visible = False
        list_flattening.Visible = False
        
        startup
        
       CurWindowNum = 12
       Call MDIWindowsMenuAdd(CurWindowNum)
       WindowsHwnd(CurWindowNum) = Me.hwnd

End Sub


Public Sub OpenFile()
Dim I, J, k As Integer
Dim title_str As String
Dim act_Num_well As Integer
Dim act_Num_event As Integer
Dim III As Integer
Dim test_num As Integer
Dim test_value As Double
Dim new_I As Integer

 ' test the ordered wells
 '   For I = 1 To max_well_num
 '       MsgBox Str(I) + " original order" + Str(orders(I)) + " well " + Str(well_numbers(I))
 '  Next I
    
FileNum = FreeFile
If txt_file_ran1 <> "" Then
    Open CurDir + "\" + txt_file_ran1 For Input As FileNum
Else
    MsgBox "Input file is not exist, please try again"
Exit Sub
End If
    
ReDim orders(1 To max_well_num)
'sort well_numbers(I)  max_well_num is from use selected number in "frmwells.frm"
For I = 1 To max_well_num    ' initialize the orders array
     orders(I) = I
Next

' sort well_numbers by linear sorting
   If max_well_num > 1 Then
     For I = 1 To max_well_num
         For J = 1 To max_well_num - 1
           If well_numbers(J) > well_numbers(J + 1) Then
                test_num = well_numbers(J)
                well_numbers(J) = well_numbers(J + 1)
                well_numbers(J + 1) = test_num
                test_num = orders(J)
                orders(J) = orders(J + 1)
                orders(J + 1) = test_num
            End If
         Next J
    Next I
   End If
    
     
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
'    MsgBox temp
'act_Num_well is the total number of wells
act_Num_well = Mid(Temp, 24, 4)
AVESD = Mid(Temp, 43, 7)
' MsgBox Str(avesd)
Input #FileNum, Temp

Input #FileNum, Temp
Unitsflag = Val(right(Temp, 1))
act_Num_event = Mid(Temp, 25, 4)

'MsgBox Str(act_Num_event)

ReDim ObsDepth(1 To max_well_num, 1 To max_event_num) As Integer
ReDim MinObs(1 To max_well_num, 1 To max_event_num) As Integer
ReDim MaxObs(1 To max_well_num, 1 To max_event_num) As Integer
ReDim probDepth(1 To max_well_num, 1 To max_event_num) As Integer
ReDim MinDepth(1 To max_well_num, 1 To max_event_num) As Integer
ReDim maxDepth(1 To max_well_num, 1 To max_event_num) As Integer
ReDim ratio(1 To max_event_num) As Double

For I = 1 To max_well_num
   For J = 1 To max_event_num
      ObsDepth(I, J) = -1
      MinObs(I, J) = -1
      MaxObs(I, J) = -1
      probDepth(I, J) = -1
      MinDepth(I, J) = -1
      maxDepth(I, J) = -1
      ratio(J) = -1
   Next J
Next I


'read number of events and wells
  
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp

'Initial Value of the MaxDepth and MinDepth
AllmaxDepth = 0
allmindepth = 1000

III = 0
For I = 1 To max_well_num       'every wells
     If (well_numbers(I) - III) > 1 Then
        ' (act_Num_event + 8)  ---->  data lines of each well in datafile
        For J = 1 To (well_numbers(I) - III - 1) * (act_Num_event + 8)
           Input #FileNum, Temp
        Next J
      End If
      III = well_numbers(I)
      
      ' read well number and name
      Line Input #FileNum, Temp
        test_num = Val(Mid(Temp, 1, 3))
        If test_num <> well_numbers(I) Then
           MsgBox "the well number and actural number does not match"
        End If
       'skip four lines
       
       Input #FileNum, Temp
       Input #FileNum, Temp
       Input #FileNum, Temp
       Input #FileNum, Temp
       
       ' input events values
    
        ' read the data into the original order
        'MsgBox "orders (I) is " + Str(orders(I))
         new_I = orders(I)
         For J = 1 To act_Num_event
          Line Input #FileNum, Temp   'Old:  Input #FileNum, Temp
          If Mid(Temp, 2, 1) <> "*" Then
              test_num = Mid(Temp, 42, 4)
                For k = 1 To max_event_num
                   If event_numbers(k) = test_num Then
                        If Mid(Temp, 5, 1) = "." Then
                            If Mid(Temp, 20, 1) <> "*" Then
                                     ' MsgBox temp
                                     ' for multplying 10 to the data from feet
                                     'since they were reduced by 10 factor when generated from Fortan cascw
                                      
                                    ObsDepth(new_I, k) = Mid(Temp, 18, 5)
                   
                                    ratio(k) = Mid(Temp, 24, 6)
                                    
                                    MinObs(new_I, k) = Mid(Temp, 30, 5)
                  
                                    MaxObs(new_I, k) = Mid(Temp, 36, 5)
                                    
                                    If Unitsflag = 1 Then
                                        ObsDepth(new_I, k) = ObsDepth(new_I, k) * 10
                                        MinObs(new_I, k) = MinObs(new_I, k) * 10
                                        MaxObs(new_I, k) = MaxObs(new_I, k) * 10
                                    End If
                                  '  MsgBox Temp + Chr$(13) + "minobs + maxObs + ObsDepth " + Str(MinObs(I, k)) + _
                                           Str(MaxObs(I, k)) + Str(ObsDepth(I, k))
                            End If
                            
                            If Mid(Temp, 2, 1) <> "*" Then
                                 probDepth(new_I, k) = Mid(Temp, 1, 5)
                                 MinDepth(new_I, k) = Mid(Temp, 6, 5)
                                 maxDepth(new_I, k) = Mid(Temp, 12, 5)
                                  
                                 If Unitsflag = 1 Then
                                      probDepth(new_I, k) = probDepth(new_I, k) * 10
                                      MinDepth(new_I, k) = MinDepth(new_I, k) * 10
                                      maxDepth(new_I, k) = maxDepth(new_I, k) * 10
                                 End If
                                 '  MsgBox Temp + Chr$(13) + "mindepth + maxdepth + probDepth " + Str(MinDepth(I, k)) + _
                                    Str(maxDepth(I, k)) + Str(probDepth(I, k))
                            End If
                            If maxDepth(new_I, k) > AllmaxDepth Then
                                AllmaxDepth = maxDepth(new_I, k)
                            End If
                            If MinDepth(new_I, k) < allmindepth And MinDepth(new_I, k) > -1 Then
                                allmindepth = MinDepth(new_I, k)
                            End If
                               
                            If MaxObs(new_I, k) > AllmaxDepth Then
                                    AllmaxDepth = MaxObs(new_I, k)
                            End If
                                
                            If MinObs(new_I, k) < allmindepth And MinObs(new_I, k) > -1 Then
                                allmindepth = MinObs(new_I, k)
                            End If
                        Else
                                If Mid(Temp, 21, 1) <> "*" Then
                                    ObsDepth(new_I, k) = Mid(Temp, 19, 5)
                                    ratio(k) = Mid(Temp, 24, 6)
                                    MinObs(new_I, k) = Mid(Temp, 31, 5)
                                    MaxObs(new_I, k) = Mid(Temp, 37, 5)
                                      
                                      If Unitsflag = 1 Then
                                        ObsDepth(new_I, k) = ObsDepth(new_I, k) * 10
                                        MinObs(new_I, k) = MinObs(new_I, k) * 10
                                        MaxObs(new_I, k) = MaxObs(new_I, k) * 10
                                      End If
                                End If
                            
                                If Mid(Temp, 2, 1) <> "*" Then
                                     probDepth(new_I, k) = Mid(Temp, 2, 4)
                                     MinDepth(new_I, k) = Mid(Temp, 7, 5)
                                     maxDepth(new_I, k) = Mid(Temp, 13, 5)
                                                     
                                        If Unitsflag = 1 Then
                                          probDepth(new_I, k) = probDepth(new_I, k) * 10
                                          MinDepth(new_I, k) = MinDepth(new_I, k) * 10
                                          maxDepth(new_I, k) = maxDepth(new_I, k) * 10
                                        End If
                                                     
                                 End If
                                 If maxDepth(new_I, k) > AllmaxDepth Then
                                     AllmaxDepth = maxDepth(new_I, k)
                                 End If
                                 If MinDepth(new_I, k) < allmindepth And MinDepth(new_I, k) > -1 Then
                                     allmindepth = MinDepth(new_I, k)
                                 End If
                                 If MaxObs(new_I, k) > AllmaxDepth Then
                                      AllmaxDepth = MaxObs(new_I, k)
                                 End If
                                 If MinObs(new_I, k) < allmindepth And MinObs(new_I, k) > -1 Then
                                      allmindepth = MinObs(new_I, k)
                                 End If
                               
                        End If
                    End If
                 Next k
          ElseIf Len(Temp) > 6 Then       'old:  Len(Temp) > 5
             test_num = Mid(Temp, 42, 4)
                For k = 1 To max_event_num
                   If event_numbers(k) = test_num Then
                        '       MsgBox temp
                        ' add to flattening list box the number and name of events
                        
                         ObsDepth(new_I, k) = Mid(Temp, 18, 5)
                         ratio(k) = Mid(Temp, 24, 6)
                         
                         If Unitsflag = 1 Then
                          ObsDepth(new_I, k) = ObsDepth(new_I, k) * 10
                          'MinObs(new_I, k) = MinObs(new_I, k) * 10
                          'MaxObs(new_I, k) = MaxObs(new_I, k) * 10
                         End If
                         If ObsDepth(new_I, k) > AllmaxDepth Then
                             AllmaxDepth = ObsDepth(new_I, k)
                         End If
                         If ObsDepth(new_I, k) < allmindepth And ObsDepth(new_I, k) > -1 Then
                             allmindepth = ObsDepth(new_I, k)
                         End If
                                    
                    End If
                Next k
          End If
      Next J
      For J = 1 To 3
       If Not EOF(FileNum) Then
       Input #FileNum, Temp
       End If
      Next J
 Next I
   
Close FileNum


' back to original orders of well numbers
' call function to draw diagram

List_Well

End Sub


Private Sub List_Well()
Dim I, J, acumIndex As Integer
Dim well As Integer
Dim Emin, Emax, E_temp As Double
Dim tempfrom As Integer
Dim title_str As String
Dim column, row   As Integer
Dim num_well_event As Integer
Dim tempPost As Integer
Dim tempTitle As String
Dim numOfflattening As Integer
Dim PlotMinObs(1 To 20, 1 To 20) As Double
Dim PlotMaxObs(1 To 20, 1 To 20) As Double
Dim PlotObsdepth(1 To 20, 1 To 20) As Double
Dim PlotMindepth(1 To 20, 1 To 20) As Double
Dim PlotMaxdepth(1 To 20, 1 To 20) As Double
Dim Plotprobdepth(1 To 20, 1 To 20) As Double
Dim TempProbDepth As String
Dim TempObsDepth As String
Dim new_I As Integer

'find the values of flattening event on each wells
numOfflattening = 0

If flattening_num <> 0 Then

  For J = 1 To max_event_num
     If event_numbers(J) = flattening_num Then
         numOfflattening = J     'index of flattening event
     End If
  Next J
    If numOfflattening = 0 Then
        MsgBox "Flattening event does not occur!"
    End If
End If

'find the new values of events after flattened
' find min and max depth of new plot after flattened
If numOfflattening <> 0 And flattening_num <> 0 Then
    allmindepth = 1000
    AllmaxDepth = 0
   
 For I = 1 To max_well_num
    If prob_opt = True Then
        If probDepth(I, numOfflattening) = -1 Then
            If ObsDepth(I, numOfflattening) = -1 Then
               TempProbDepth = InputBox("Probable and observed depths do not exist for " + well_label_used(I), _
                                "Please enter a value for probable depth", Default)
                While TempProbDepth = ""
                Beep
                TempProbDepth = InputBox("Probable and observed depths do not exist for " + well_label_used(I), _
                                "Please enter a value for probable depth", Default)
                Wend
            Else
              TempProbDepth = InputBox("Probable depth does not exist for " + well_label_used(I) + Chr(13) + _
                    Chr(10) + "Observed depth = " + Str(ObsDepth(I, numOfflattening)), _
                            "Please enter a value for probable depth", Default)
               While TempProbDepth = ""
                Beep
                TempProbDepth = InputBox("Probable depth does not exist for " + well_label_used(I) + Chr(13) + _
                    Chr(10) + "Observed depth = " + Str(ObsDepth(I, numOfflattening)), _
                            "Please enter a value for probable depth", Default)
               Wend
            End If
        Else
              TempProbDepth = probDepth(I, numOfflattening)
        End If
        
         For J = 1 To max_event_num
            Plotprobdepth(I, J) = probDepth(I, J) - CDbl(TempProbDepth)
            PlotMaxdepth(I, J) = maxDepth(I, J) - CDbl(TempProbDepth)
            PlotMindepth(I, J) = MinDepth(I, J) - CDbl(TempProbDepth)
        Next J
     
    ElseIf norm_opt = True Or obs_opt = True Then
      
        If ObsDepth(I, numOfflattening) = -1 Then
            If probDepth(I, numOfflattening) = -1 Then
              TempObsDepth = InputBox("Probable and observed depths do not exist for " + well_label_used(I), _
                                "Please enter a value for observed depth", Default)
              While TempObsDepth = ""
               Beep
               TempObsDepth = InputBox("Probable and observed depths do not exist for " + well_label_used(I), _
                                "Please enter a value for observed depth", Default)
              Wend
                                
            Else
              TempObsDepth = InputBox("Observed depth does not exist for " + well_label_used(I) + Chr(13) + _
                    Chr(10) + "Probable depth = " + Str(probDepth(I, numOfflattening)), _
                            "Please enter a value for observed depth", Default)
            While TempObsDepth = ""
               Beep
               TempObsDepth = InputBox("Observed depth does not exist for " + well_label_used(I) + Chr(13) + _
                    Chr(10) + "Probable depth = " + Str(probDepth(I, numOfflattening)), _
                            "Please enter a value for observed depth", Default)
            Wend
            
            End If
        Else
            TempObsDepth = ObsDepth(I, numOfflattening)
        End If
        For J = 1 To max_event_num
            PlotObsdepth(I, J) = ObsDepth(I, J) - CDbl(TempObsDepth)
            PlotMaxObs(I, J) = MaxObs(I, J) - CDbl(TempObsDepth)
            PlotMinObs(I, J) = MinObs(I, J) - CDbl(TempObsDepth)
        Next J
    End If
         
     'find max depth
    For J = 1 To max_event_num
     If probDepth(I, J) > -1 And Plotprobdepth(I, J) > AllmaxDepth Then
        AllmaxDepth = Plotprobdepth(I, J)
     End If
     If maxDepth(I, J) > -1 And PlotMaxdepth(I, J) > AllmaxDepth Then
        AllmaxDepth = PlotMaxdepth(I, J)
     End If
     If MinDepth(I, J) > -1 And PlotMindepth(I, J) > AllmaxDepth Then
        AllmaxDepth = PlotMindepth(I, J)
     End If
     If MinObs(I, J) > -1 And PlotMinObs(I, J) > AllmaxDepth Then
        AllmaxDepth = PlotMinObs(I, J)
     End If
     If MaxObs(I, J) > -1 And PlotMaxObs(I, J) > AllmaxDepth Then
        AllmaxDepth = PlotMaxObs(I, J)
     End If
     If ObsDepth(I, J) > -1 And PlotObsdepth(I, J) > AllmaxDepth Then
        AllmaxDepth = PlotObsdepth(I, J)
     End If
     
     'find min depth
     
     If probDepth(I, J) > -1 And Plotprobdepth(I, J) < allmindepth Then
        allmindepth = Plotprobdepth(I, J)
     End If
     If maxDepth(I, J) > -1 And PlotMaxdepth(I, J) < allmindepth Then
        allmindepth = PlotMaxdepth(I, J)
     End If
     If MinDepth(I, J) > -1 And PlotMindepth(I, J) < allmindepth Then
        allmindepth = PlotMindepth(I, J)
     End If
     If MinObs(I, J) > -1 And PlotMinObs(I, J) < allmindepth Then
        allmindepth = PlotMinObs(I, J)
     End If
     If MaxObs(I, J) > -1 And PlotMaxObs(I, J) < allmindepth Then
        allmindepth = PlotMaxObs(I, J)
     End If
     If ObsDepth(I, J) > -1 And PlotObsdepth(I, J) < allmindepth Then
        allmindepth = PlotObsdepth(I, J)
     End If
   Next J
  Next I
End If
     
column = 0
row = 0

'Set a extension scale for Y Axe for a good display
Dim YExtendLength As Double
Select Case (AllmaxDepth - allmindepth)
Case Is < 50
   YExtendLength = (AllmaxDepth - allmindepth) * 0.4
Case 50 To 100
   YExtendLength = (AllmaxDepth - allmindepth) * 0.3
Case 101 To 800
   YExtendLength = (AllmaxDepth - allmindepth) * 0.2
Case 801 To 2000
   YExtendLength = (AllmaxDepth - allmindepth) * 0.1
Case 2001 To 5000
   YExtendLength = (AllmaxDepth - allmindepth) * 0.05
Case Else
   YExtendLength = (AllmaxDepth - allmindepth) * 0.025
End Select

'MsgBox "Well No. == " + Str(well)
'LabelWell.Visible = False
'List_Well.Visible = False
'ReDim newwell(1 To max_well_num) ' back to the original order for labelling
'For I = 1 To max_well_num
'new_I = orders(I)
'    newwell(I) = well_numbers(new_I)
'Next I
'MsgBox "well  width " + Str(well_width)

ChartWell.ChartGroups(1).ChartType = oc2dTypeBar

ChartWell.ChartArea.Axes("x").Min.Value = -well_width
ChartWell.ChartArea.Axes("x").Max.Value = max_well_num - 1 + well_width
'Define the label spacing and precision of Y axe
ChartWell.ChartArea.Axes("y").NumSpacing.Value = 500
ChartWell.ChartArea.Axes("y").Precision = 0

' design plot area
  If max_well_num < 6 Then
      ChartWell.ChartArea.PlotArea.left.Value = 90 * (6 - max_well_num)
      ChartWell.ChartArea.PlotArea.right.Value = 90 * (6 - max_well_num)
  End If

'desgin line types


    For I = 1 To 2 * max_event_num * max_well_num + 2 * max_event_num
      ChartWell.ChartGroups(2).Styles.Add(1).Line.Pattern = oc2dLineSolid
      ChartWell.ChartGroups(2).Styles(1).Line.Width = 2
      ChartWell.ChartGroups(2).Styles(1).Line.Color = RGB(255, 0, 0)
      ChartWell.ChartGroups(2).Styles(1).Symbol.Shape = oc2dShapeNone
      ChartWell.ChartGroups(2).Styles(1).Symbol.size = 17
      ChartWell.ChartGroups(2).Styles(1).Symbol.Color = RGB(255, 255, 0)
      
    Next I
 
 'Define the label spacing of Y2 axe
 ChartWell.ChartArea.Axes("y2").NumSpacing.Value = 500
 ChartWell.ChartArea.Axes("y2").Precision = 0

 
 'add chartlabels well no. 1 ....
'add chartlabels for event numbers
' add event numbers to the the well at the middle
 'middle = Int(max_well_num / 2) ' get the middle well
 
num_well_event = max_well_num + max_event_num * max_well_num

For I = 1 To num_well_event  'max_well_num + max_event_num
        ChartWell.ChartLabels.Add
  '    ChartWell.ChartLabels(I).Anchor = oc2dAnchorNorth
       ChartWell.ChartLabels(I).AttachMethod = oc2dAttachDataCoord
       ChartWell.ChartLabels(I).Font.size = 10
       ChartWell.ChartLabels(I).Font.name = Arial
   '     ChartWell.ChartLabels(I).Offset = 5
 If I <= max_well_num Then
    ChartWell.ChartLabels(I).Anchor = oc2dAnchorNorth
    If I <> max_well_num Then
     ChartWell.ChartLabels(I).AttachDataCoord.X = orders(I) - 1
    Else
    ChartWell.ChartLabels(I).AttachDataCoord.X = orders(I) - 1
    End If
     ' Label well number or name
     ChartWell.ChartLabels(I).AttachDataCoord.Y = allmindepth - YExtendLength
     ChartWell.ChartLabels(I).Text.Text = well_label_used(I)
     ChartWell.ChartLabels(I).Offset = 10
 Else
  
        'middle > 0 Then
         If max_well_num = 1 Then
           column = 1
         Else
            column = I Mod max_well_num 'reminder of number of well
         End If
         
         If column = 1 Then
             row = row + 1
         End If
            
        'set well_width coeffiecient
        Dim width_ceo
        width_ceo = 0.6
        
        ' add labels for the first well
                
         If column = 1 Then
            If prob_opt = True And probDepth(column, row) > -1 Then
                ChartWell.ChartLabels(I).Anchor = oc2dAnchorEast
                ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 + well_width * width_ceo
            '    If probDepth(column, row) < probDepth(column + 1, row) Then
                                
                If numOfflattening <> 0 Then
                     ChartWell.ChartLabels(I).AttachDataCoord.Y = Plotprobdepth(column, row)
                Else
                     ChartWell.ChartLabels(I).AttachDataCoord.Y = probDepth(column, row)
                End If
            '    Else
            '    ChartWell.ChartLabels(i).AttachDataCoord.y = probDepth(column, row)
            '    End If
                 ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
             End If
             
             If norm_opt = True And ObsDepth(column, row) > -1 Then
                        'And MinObs(column, row) > -1 And MaxObs(column, row) > -1 Then
                ChartWell.ChartLabels(I).Anchor = oc2dAnchorEast
                ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 + well_width * width_ceo
            '    If ObsDepth(column, row) < ObsDepth(column + 1, row) Then
                If numOfflattening <> 0 Then
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = PlotObsdepth(column, row)
                Else
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = ObsDepth(column, row)
                End If
             '
             '   Else
              '  ChartWell.ChartLabels(i).AttachDataCoord.y = ObsDepth(column, row)
              '  End If
                ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
             End If
               
             If obs_opt = True And ObsDepth(column, row) > -1 Then
                ChartWell.ChartLabels(I).Anchor = oc2dAnchorEast
                ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 + well_width * width_ceo
            '    If ObsDepth(column, row) < ObsDepth(column + 1, row) Then
                
                If numOfflattening <> 0 Then
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = PlotObsdepth(column, row)
                Else
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = ObsDepth(column, row)
                End If
             '
             '   Else
              '  ChartWell.ChartLabels(i).AttachDataCoord.y = ObsDepth(column, row)
              '  End If
                ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
             End If
               
         End If
         
        ' add labels to the last well
        
        If column = 0 Then
               column = max_well_num
            If prob_opt = True And probDepth(column, row) > -1 Then
                ChartWell.ChartLabels(I).Anchor = oc2dAnchorWest
                ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 - well_width * width_ceo
               ' If probDepth(column, row) < probDepth(column - 1, row) Then
                If numOfflattening <> 0 Then
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = Plotprobdepth(column, row)
                Else
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = probDepth(column, row)
                End If
               ' Else
               ' ChartWell.ChartLabels(i).AttachDataCoord.y = probDepth(column, row)
               ' End If
                ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
             End If
             
             If norm_opt = True And ObsDepth(column, row) > -1 Then
                   'And MinObs(column, row) > -1 And MaxObs(column, row) > -1 Then
                ChartWell.ChartLabels(I).Anchor = oc2dAnchorWest
                ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 - well_width * width_ceo
                'If ObsDepth(column, row) < ObsDepth(column - 1, row) Then
                 If numOfflattening <> 0 Then
                       ChartWell.ChartLabels(I).AttachDataCoord.Y = PlotObsdepth(column, row)
                 Else
                       ChartWell.ChartLabels(I).AttachDataCoord.Y = ObsDepth(column, row)
                 End If
               ' Else
               ' ChartWell.ChartLabels(i).AttachDataCoord.y = ObsDepth(column, row)
               ' End If
                ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
             End If
               
             If obs_opt = True And ObsDepth(column, row) > -1 Then
                ChartWell.ChartLabels(I).Anchor = oc2dAnchorWest
                ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 - well_width * width_ceo
                'If ObsDepth(column, row) < ObsDepth(column - 1, row) Then
                If numOfflattening <> 0 Then
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = PlotObsdepth(column, row)
                Else
                      ChartWell.ChartLabels(I).AttachDataCoord.Y = ObsDepth(column, row)
                End If
               ' Else
               ' ChartWell.ChartLabels(i).AttachDataCoord.y = ObsDepth(column, row)
               ' End If
                ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
             End If
             
                   
         End If
         
         'add labels to the wells in middle
          
          If column > 1 And column < max_well_num Then
            If prob_opt = True And probDepth(column, row) > -1 Then
            
'                If probDepth(column + 1, row) = -1 Or probDepth(column - 1, row) = -1 Then
'                    If probDepth(column + 1, row) = -1 Then
                       ChartWell.ChartLabels(I).Anchor = oc2dAnchorEast
                       ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 + well_width * width_ceo
'                    Else
'                       ChartWell.ChartLabels(I).Anchor = oc2dAnchorWest
'                       ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 - well_width * width_ceo
'                    End If
                    If numOfflattening <> 0 Then
                       ChartWell.ChartLabels(I).AttachDataCoord.Y = Plotprobdepth(column, row)
                    Else
                       ChartWell.ChartLabels(I).AttachDataCoord.Y = probDepth(column, row)
                    End If
                    ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
'                End If
             End If
             
             If norm_opt = True And ObsDepth(column, row) > -1 Then
                 'And MinObs(column, row) > -1 And MaxObs(column, row) > -1 Then
            
'                If ObsDepth(column + 1, row) = -1 Or ObsDepth(column - 1, row) = -1 Then
'                    If ObsDepth(column + 1, row) = -1 Then
                       ChartWell.ChartLabels(I).Anchor = oc2dAnchorEast
                       ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 + well_width * width_ceo
'                    Else
'                       ChartWell.ChartLabels(I).Anchor = oc2dAnchorWest
'                       ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 - well_width * width_ceo
'                    End If
                    If numOfflattening <> 0 Then
                         ChartWell.ChartLabels(I).AttachDataCoord.Y = PlotObsdepth(column, row)
                    Else
                         ChartWell.ChartLabels(I).AttachDataCoord.Y = ObsDepth(column, row)
                     End If
                       ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
'                End If
             End If
               
             If obs_opt = True And ObsDepth(column, row) > -1 Then
            
'                If ObsDepth(column + 1, row) = -1 Or ObsDepth(column - 1, row) = -1 Then
'                    If ObsDepth(column + 1, row) = -1 Then
                       ChartWell.ChartLabels(I).Anchor = oc2dAnchorEast
                       ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 + well_width * width_ceo
'                    Else
'                       ChartWell.ChartLabels(I).Anchor = oc2dAnchorWest
'                       ChartWell.ChartLabels(I).AttachDataCoord.X = column - 1 - well_width * width_ceo
'                    End If
                    If numOfflattening <> 0 Then
                         ChartWell.ChartLabels(I).AttachDataCoord.Y = PlotObsdepth(column, row)
                    Else
                         ChartWell.ChartLabels(I).AttachDataCoord.Y = ObsDepth(column, row)
                    End If
                       ChartWell.ChartLabels(I).Text.Text = Str(event_numbers(row))
'                End If
             End If
      
 '             ChartWell.ChartLabels(I).Offset = 5
         End If
   End If
 Next I
 
 'add chartlabels for event numbers
 ' add event numbers to the the well at the middle

    
' add title to the chart

If file_type = "Ca1" Then
    title_str = "Ranking"
ElseIf file_type = "Ca2" Then
    title_str = "Scaling"
End If
tempPost = InStr(txt_file_ran1, ".")
tempTitle = Mid(txt_file_ran1, 1, tempPost - 1)


   If prob_opt = True Then
        ChartWell.Header.Text.Text = "Probable Depths with Error Bars (" + title_str + "; " + tempTitle + ")"
        Table_Title = title_str + " - File: " + tempTitle
   ElseIf norm_opt = True Then
        ChartWell.Header.Text.Text = "Observed Depths with Error Bars (" + title_str + "; " + tempTitle + ")"
        Table_Title = title_str + " - File: " + tempTitle
   ElseIf obs_opt = True Then
        ChartWell.Header.Text.Text = "Observed Depths (" + title_str + "; " + tempTitle + ")"
        Table_Title = title_str + " - File: " + tempTitle
    End If
    
    
'ChartWell.ChartArea.Axes("y").Title.Text = "Depth"


With ChartWell.ChartGroups(1).Data
'ChartWell.ChartArea.Bar.ClusterWidth = well_width * 100
    .Layout = oc2dDataArray
    .IsBatched = True
    .NumSeries = 1
    .NumPoints(1) = max_well_num
        For J = 1 To max_well_num
        .X(1, J) = orders(J) - 1
        .Y(1, J) = AllmaxDepth + YExtendLength
        Next J
        
End With


ChartWell.ChartArea.Axes("y").Min.Value = allmindepth - YExtendLength
ChartWell.ChartArea.Axes("y").Max.Value = AllmaxDepth + YExtendLength
ChartWell.ChartArea.Axes("x").Origin.Value = -well_width
ChartWell.ChartArea.Axes("y").Origin.Value = allmindepth - YExtendLength
ChartWell.ChartArea.Axes("y").DataMax.Value = AllmaxDepth + YExtendLength
ChartWell.ChartArea.Axes("y").DataMin.Value = allmindepth - YExtendLength
ChartWell.ChartArea.Axes("x").Max.Value = max_well_num - 1 + well_width
ChartWell.ChartArea.Axes("x").Min.Value = -well_width
ChartWell.ChartArea.Axes("x").DataMax.Value = max_well_num - 1
ChartWell.ChartArea.Axes("x").DataMin.Value = 0
ChartWell.ChartArea.Axes("y2").Max.Value = AllmaxDepth + YExtendLength
ChartWell.ChartArea.Axes("y2").Min.Value = allmindepth - YExtendLength
ChartWell.ChartArea.Axes("y2").DataMax.Value = AllmaxDepth + YExtendLength
ChartWell.ChartArea.Axes("y2").DataMin.Value = allmindepth - YExtendLength
     
'Setting chart label number spacing for Y axes
Select Case ((AllmaxDepth + YExtendLength) - (allmindepth - YExtendLength))
Case Is < 50
   ChartYAxeNuberSpace = 5
Case 50 To 100
   ChartYAxeNuberSpace = 10
Case 101 To 500
   ChartYAxeNuberSpace = 50
Case 501 To 1000
   ChartYAxeNuberSpace = 100
Case 1001 To 3000
   ChartYAxeNuberSpace = 200
Case 3001 To 5000
   ChartYAxeNuberSpace = 250
Case Else
   ChartYAxeNuberSpace = 500
End Select
ChartWell.ChartArea.Axes("y").NumSpacing = ChartYAxeNuberSpace
ChartWell.ChartArea.Axes("y2").NumSpacing = ChartYAxeNuberSpace

ChartWell.ChartGroups(2).Data.IsBatched = False
ChartWell.ChartGroups(2).ChartType = oc2dTypePlot
With ChartWell.ChartGroups(2).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 2 * max_event_num * max_well_num + 2 * max_event_num
   
   ' inilitialize the points per series
   
      For I = 1 To 3 * max_event_num
       .NumPoints(I) = max_well_num
      Next I
      
      For I = 1 To 2 * max_event_num * max_well_num - max_event_num
       .NumPoints(I + 3 * max_event_num) = 2
      Next I
      
   'provide values for x and y
   
   ' plot obs depth or prob depth as points
   
    For I = 1 To max_event_num
        ChartWell.ChartGroups(2).Styles.Add(1).Symbol.Shape = oc2dShapeHorizontalLine
        ChartWell.ChartGroups(2).Styles(1).Line.Pattern = oc2dLineNone
        ChartWell.ChartGroups(2).Styles(1).Symbol.size = 15
            
'        ChartWell.ChartArea.Axes("y").Min.Value = allmindepth - YExtendLength
'        ChartWell.ChartArea.Axes("Y2").Max.Value = AllmaxDepth + YExtendLength
'        ChartWell.ChartArea.Axes("x").Origin.Value = -well_width
'        ChartWell.ChartArea.Axes("y").Origin.Value = allmindepth - YExtendLength
'        ChartWell.ChartArea.Axes("y2").DataMax.Value = AllmaxDepth + YExtendLength
'        ChartWell.ChartArea.Axes("y2").DataMin.Value = allmindepth - YExtendLength
'        ChartWell.ChartArea.Axes("x").Max.Value = max_well_num - 1 + well_width
'        ChartWell.ChartArea.Axes("x").Min.Value = -well_width
'        ChartWell.ChartArea.Axes("y2").Min.Value = allmindepth - YExtendLength

           
          If prob_opt = True Then
                 For J = 1 To max_well_num
                    If probDepth(J, I) > -1 Then
                      .X(I, J) = J - 1
                         If numOfflattening <> 0 Then
                         .Y(I, J) = Plotprobdepth(J, I)
                         Else
                         .Y(I, J) = probDepth(J, I)
                         End If
                     End If
                  Next J
                  
           ElseIf norm_opt = True Then
                  For J = 1 To max_well_num
                       If ObsDepth(J, I) > -1 Then
                        .X(I, J) = J - 1
                         If numOfflattening <> 0 Then
                         .Y(I, J) = PlotObsdepth(J, I)
                         Else
                           .Y(I, J) = ObsDepth(J, I)
                         End If
                       End If
                    Next J
                    
           ElseIf obs_opt = True Then
                   For J = 1 To max_well_num
                     If ObsDepth(J, I) > -1 Then
                        .X(I, J) = J - 1
                        If numOfflattening <> 0 Then
                        .Y(I, J) = PlotObsdepth(J, I)
                        Else
                        .Y(I, J) = ObsDepth(J, I)
                        End If
                      End If
                    Next J
            End If
      Next I
      
   'plot min depth as bars
   
      For I = 1 To max_event_num
          ChartWell.ChartGroups(2).Styles.Add(1).Symbol.Shape = oc2dShapeHorizontalLine
            ChartWell.ChartGroups(2).Styles(1).Symbol.size = 15
               ChartWell.ChartGroups(2).Styles(1).Line.Pattern = oc2dLineNone

           If prob_opt = True Then
                    For J = 1 To max_well_num
                        If MinDepth(J, I) > -1 Then
                            .X(I + max_event_num, J) = J - 1
                                If numOfflattening <> 0 Then
                                 .Y(I + max_event_num, J) = PlotMindepth(J, I)
                                 Else
                                .Y(I + max_event_num, J) = MinDepth(J, I)
                                End If
                         End If
                     Next J
                     
             ElseIf norm_opt = True Then
                     For J = 1 To max_well_num
                         If MinObs(J, I) > -1 Then
                            .X(I + max_event_num, J) = J - 1
                               If numOfflattening <> 0 Then
                                .Y(I + max_event_num, J) = PlotMinObs(J, I)
                                Else
                                .Y(I + max_event_num, J) = MinObs(J, I)
                                End If
                           End If
                      Next J
              End If
       Next I
       
       'plot mas depth as bar
       
       For I = 1 To max_event_num
           
               ChartWell.ChartGroups(2).Styles.Add(1).Symbol.Shape = oc2dShapeHorizontalLine
            ChartWell.ChartGroups(2).Styles(1).Symbol.size = 13
            ChartWell.ChartGroups(2).Styles(1).Symbol.Color = RGB(255, 255, 0)
                    
            ChartWell.ChartGroups(2).Styles(1).Line.Pattern = oc2dLineNone
              If prob_opt = True Then
                    For J = 1 To max_well_num
                        If maxDepth(J, I) > -1 Then
                           .X(I + 2 * max_event_num, J) = J - 1
                           If numOfflattening <> 0 Then
                           .Y(I + 2 * max_event_num, J) = PlotMaxdepth(J, I)
                            Else
                            .Y(I + 2 * max_event_num, J) = maxDepth(J, I)
                           End If
                         End If
                     Next J
                     
               ElseIf norm_opt = True Then
                    For J = 1 To max_well_num
                         If MaxObs(J, I) > -1 Then
                            .X(I + 2 * max_event_num, J) = J - 1
                            If numOfflattening <> 0 Then
                            .Y(I + 2 * max_event_num, J) = PlotMaxObs(J, I)
                            Else
                            .Y(I + 2 * max_event_num, J) = MaxObs(J, I)
                            End If
                         End If
                    Next J
                End If
       Next I
     
    'add vertical lines
    acumIndex = 0
     For I = 1 To max_event_num
         For J = 1 To max_well_num
            'skip no data points
              If prob_opt = True And MinDepth(J, I) > -1 And maxDepth(J, I) > -1 Then
                 acumIndex = acumIndex + 1
                .X(acumIndex + 3 * max_event_num, 1) = J - 1
                .X(acumIndex + 3 * max_event_num, 2) = J - 1
                    If numOfflattening <> 0 Then
                    .Y(acumIndex + 3 * max_event_num, 1) = PlotMindepth(J, I)
                    .Y(acumIndex + 3 * max_event_num, 2) = PlotMaxdepth(J, I)
                    Else
                    .Y(acumIndex + 3 * max_event_num, 1) = MinDepth(J, I)
                    .Y(acumIndex + 3 * max_event_num, 2) = maxDepth(J, I)
                    End If
               End If
              If norm_opt = True And ObsDepth(J, I) > -1 And MinObs(J, I) > -1 And MaxObs(J, I) > -1 Then
                acumIndex = acumIndex + 1
                .X(acumIndex + 3 * max_event_num, 1) = J - 1
                .X(acumIndex + 3 * max_event_num, 2) = J - 1
                    If numOfflattening <> 0 Then
                     .Y(acumIndex + 3 * max_event_num, 1) = PlotMinObs(J, I)
                     .Y(acumIndex + 3 * max_event_num, 2) = PlotMaxObs(J, I)
                    Else
                    .Y(acumIndex + 3 * max_event_num, 1) = MinObs(J, I)
                    .Y(acumIndex + 3 * max_event_num, 2) = MaxObs(J, I)
                    End If
             End If
       Next J
     Next I
    
    'add connect lines to obsdepth or probdepth
    acumIndex = 0
    For I = 1 To max_event_num
 
         For J = 1 To max_well_num - 1
   
         'skip no data points
         If prob_opt = True And probDepth(J, I) > -1 And probDepth(J + 1, I) > -1 Then
         acumIndex = acumIndex + 1
        .X(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 1) = J - 1 + well_width * 0.5
        .X(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 2) = J - well_width * 0.5
           If numOfflattening <> 0 Then
            .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 1) = Plotprobdepth(J, I)
            .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 2) = Plotprobdepth(J + 1, I)
           Else
            .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 1) = probDepth(J, I)
            .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 2) = probDepth(J + 1, I)
           End If
         End If
         If prob_opt = False And ObsDepth(J, I) > -1 And ObsDepth(J + 1, I) > -1 Then
              acumIndex = acumIndex + 1
        .X(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 1) = J - 1 + well_width * 0.5
        .X(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 2) = J - well_width * 0.5
           If numOfflattening <> 0 Then
             .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 1) = PlotObsdepth(J, I)
             .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 2) = PlotObsdepth(J + 1, I)
           Else
             .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 1) = ObsDepth(J, I)
             .Y(acumIndex + 3 * max_event_num + max_event_num * max_well_num, 2) = ObsDepth(J + 1, I)
            End If
         End If
       Next J
     Next I
     
End With

ChartWell.ChartGroups(2).Data.IsBatched = False
'ChartScatter.Visible = True

If Unitsflag = 0 Then
   ChartWell.ChartArea.Axes("Y").Title.Text = "Depth (meters)"
Else
   ChartWell.ChartArea.Axes("Y").Title.Text = "Depth (feet)"
End If

ChartWell.Visible = True
cmdModify.Enabled = True
'????????????????????????
'For I = 1 To 3 * max_event_num + max_event_num * max_well_num
 ' ChartWell.ChartGroups(2).Styles(1).Line.Width = 2
  '    ChartWell.ChartGroups(2).Styles(1).Line.Color = RGB(255, 0, 0)
'Next I

End Sub

Private Sub Form_Resize()
Dim size As Integer
If (frmwellPlot1.Height > 360) And (frmwellPlot1.Height - 360 < frmwellPlot1.Width / 1.5) Then
   size = frmwellPlot1.Height - 360
ElseIf (frmwellPlot1.Height > 360) And (frmwellPlot1.Height - 360 > frmwellPlot1.Width / 1.5) Then
    size = frmwellPlot1.Width / 1.5
Else
    size = 0
End If

ChartWell.Height = size
ChartWell.Width = 1.5 * size
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CurGraphicOBJ = Nothing
    
   CurWindowNum = 12
   Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub

Private Sub list_flattening_Click()

ChartWell.Visible = False
flattening_num = event_numbers(list_flattening.ListIndex + 1)
startup
List_Well
End Sub

Private Sub List_well_width_Click()
ChartWell.Visible = False
well_width = (List_well_width.ListIndex + 1) * 0.05
startup
List_Well
End Sub

Private Sub List1_Click()

ChartWell.Visible = False
End Sub
