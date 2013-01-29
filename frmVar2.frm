VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmVar2 
   Caption         =   "Variance Analysis - Scaling"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "frmVar2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   252
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   612
   End
   Begin VB.ListBox List_Event 
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5385
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2340
      TabIndex        =   4
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1020
      TabIndex        =   3
      Top             =   120
      Width           =   612
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D Chartvar2 
      Height          =   3810
      Left            =   120
      TabIndex        =   1
      Top             =   4350
      Visible         =   0   'False
      Width           =   11505
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   20285
      _ExtentY        =   6720
      _StockProps     =   0
      ControlProperties=   "frmVar2.frx":0442
   End
   Begin OlectraChart2D.Chart2D ChartVar1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Visible         =   0   'False
      Width           =   11505
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   20285
      _ExtentY        =   6800
      _StockProps     =   0
      ControlProperties=   "frmVar2.frx":0B28
   End
   Begin VB.Label labelEvent 
      Caption         =   "Select an event"
      Height          =   255
      Left            =   3270
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmVar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Dev_Value() As Double
Dim Well_Num() As Integer
Dim frequency(1 To 10) As Integer
Dim classfrom(1 To 10), classto(1 To 10) As Double

Dim MaxNumWell As Integer
Dim MaxNumEvent As Integer
Dim NumOfEvent() As Integer
Dim EventName() As String * 50
Dim NumOfWell() As Integer
Dim EventNum As Integer
Dim MaxArray As Integer

Dim Temp As String
Dim FileNum As Integer
Dim Filename As String
Dim sample_size As Integer
Dim Std_sample As Double



Private Sub ChartVar1_Click()
    Set CurGraphicOBJ = ChartVar1
End Sub

Private Sub Chartvar2_Click()
    Set CurGraphicOBJ = Chartvar2
End Sub

Private Sub cmdCancel_Click()
Unload frmVar2
End Sub

Private Sub cmdOpen_Click()
file_type = "Va2"
frmOpenRan1.Show 1
'cmdOpen.Enabled = False

If txt_file_ran1 = "" Then
  Beep
  Exit Sub
End If

startup
OpenFile

End Sub

Private Sub cmdPrint_Click()
    ChartVar1.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
    Chartvar2.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
    'ActiveForm.PrintForm
    'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub cmdSave_Click()

     Set CurChartSaveObj = ChartVar1
     ChartSaveAs

     Set CurChartSaveObj = Chartvar2
     ChartSaveAs


' Dim ImageName As String
'ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'If ImageName <> "" Then
'    ChartVar1.SaveImageAsJpeg ImageName & "1.jpg", 100, False, False, False
'    Chartvar2.SaveImageAsJpeg ImageName & "2.jpg", 100, False, False, False
'End If
End Sub

Private Sub Form_Activate()
    Set CurGraphicOBJ = ChartVar1
        CurWindowNum = 7
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

'Private Sub Command1_Click()  this was for copy to clipboard
'Dim response As Boolean
'response = ChartVar1.CopyToClipboard(oc2dFormatBitmap)
'If response = False Then
'    MsgBox "It is not successful! Try it again!"
'End If

'End Sub

Private Sub Form_Load()
    txt_file_ran1 = ""
    cmdOpen.Enabled = True
    
    CurWindowNum = 7
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

End Sub


Public Sub startup()
    Dim I, J As Integer
    ChartVar1.Visible = False
    
    ChartVar1.ChartGroups(1).ChartType = oc2dTypeBar
    With ChartVar1.ChartGroups(1).Data
        .Layout = oc2dDataArray
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
     '   .NumPoints(1) = 150
       
    End With
    
    ChartVar1.ChartGroups(1).Data.IsBatched = False
    
    Chartvar2.Visible = False
    
    Chartvar2.ChartGroups(1).ChartType = oc2dTypeBar
    With Chartvar2.ChartGroups(1).Data
        .Layout = oc2dDataArray
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
     '   .NumPoints(1) = 150
       
    End With
    
    Chartvar2.ChartGroups(1).Data.IsBatched = False
    
End Sub


Public Sub OpenFile()
    Dim tempfrom As Integer
    Dim tempto As Integer
    Dim I, J As Integer
    Dim title_str As String
    
    Dim ArrayNum As Integer
    Dim tnum
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If txt_file_ran1 <> "" Then
    Open CurDir + "\" + txt_file_ran1 For Input As FileNum
    Else
        MsgBox "Input file is not exist, please try again"
    Exit Sub
    End If
        
    Input #FileNum, Temp
    Input #FileNum, Temp
    Input #FileNum, Temp
    
    MaxNumEvent = Mid(Temp, 25, 4)
    
    Input #FileNum, Temp  'empty line
    Input #FileNum, Temp  'number of well
    'MsgBox temp
    MaxNumWell = Mid(Temp, 24, 4)
    'MsgBox Str(MaxNumWell)
    
    Input #FileNum, Temp 'three empty lines
    Input #FileNum, Temp
    Input #FileNum, Temp
    
    ReDim NumOfEvent(1 To MaxNumEvent)
    ReDim NumOfWell(1 To MaxNumEvent)
    ReDim EventName(1 To MaxNumEvent)
    
    'read events, well and event name
    
       For I = 1 To MaxNumEvent
           Input #FileNum, Temp
           NumOfEvent(I) = Mid(Temp, 6, 4)
           NumOfWell(I) = Mid(Temp, 20, 3)
           EventName(I) = Mid(Temp, 30, 50)
           
     'MsgBox "Events   : " + Str(NumOfEvent(I)) + " in Wells " + Str(NumOfWell(I))
      
       Next I
    
    Close FileNum
    
    'write into list box
    List_Event.Visible = True
    labelEvent.Visible = True
    List_Event.Clear
        
    
    For I = 1 To MaxNumEvent
        List_Event.AddItem "Event  " + Str(NumOfEvent(I)) + "  Occurs in  " + Str(NumOfWell(I)) + " Wells: " + EventName(I)
    Next I
End Sub


Private Sub Form_Resize()
    Dim size As Integer
    If (frmVar2.Height > 600) And (frmVar2.Height - 600 < frmVar2.Width / 1.5) Then
       size = frmVar2.Height - 600
    ElseIf (frmVar2.Height > 600) And (frmVar2.Height - 600 > frmVar2.Width / 1.5) Then
       size = frmVar2.Width / 1.5
    Else
       size = 0
    End If
    
    ChartVar1.Height = size / 2 - 100
    ChartVar1.Width = 1.5 * size
    Chartvar2.top = 600 + ChartVar1.Height
    Chartvar2.Height = ChartVar1.Height
    Chartvar2.Width = ChartVar1.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CurGraphicOBJ = Nothing
    
    CurWindowNum = 7
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub

Private Sub List_Event_Click()
    Dim I, J As Integer
    Dim well As Integer
    Dim Emin, Emax, E_temp As Double
    Dim title_str As String
    Dim RecordFrom, RecordTo As Integer
    Dim skipRecord As Integer
    Dim checking As Integer
    Dim FileNum As Integer
    Dim minSDV As Double
    Dim maxSdv As Double
    Dim tempvalue As Double
    ChartVar1.Visible = False
    Chartvar2.Visible = False
    
    
    EventNum = List_Event.ListIndex + 1
    
    'MsgBox "index of Event No. == " + Str(EventNum)
    'LabelEvent.Visible = False
    'List_Event.Visible = False
    
    FileNum = FreeFile
    
    'read data from file
    Open CurDir + "\" + txt_file_ran1 For Input As FileNum
    
    skipRecord = 12 + MaxNumEvent ' skip to the begining of indvisual event data
    
    For I = 1 To skipRecord
       Input #FileNum, Temp
    Next I
    
    'skip to the begining of data
    
    RecordFrom = 0
    If EventNum > 1 Then
      For I = 2 To EventNum
                    RecordFrom = RecordFrom + NumOfWell(I - 1) + 27
        Next I
    End If
    ' skip to the begining of data
      If RecordFrom > 1 Then
        For I = 1 To RecordFrom
          Input #FileNum, Temp
        Next I
      End If
      
     'testing the skipping is correct
     Input #FileNum, Temp
     
     checking = Mid(Temp, 10, 5)
       If checking <> NumOfEvent(EventNum) Then
          MsgBox "The event given in list box <> event read from file"
       Exit Sub
       End If
    ' skip three empty lines
      Input #FileNum, Temp
      Input #FileNum, Temp
      Input #FileNum, Temp
    'read deviation value data into
    ReDim Well_Num(1 To NumOfWell(EventNum))
    ReDim Dev_Value(1 To NumOfWell(EventNum))
    minSDV = 0
    maxSdv = 0
       For I = 1 To NumOfWell(EventNum)
       Input #FileNum, Temp
        Well_Num(I) = Mid(Temp, 1, 2)
        Dev_Value(I) = Mid(Temp, 4, 10)
    If minSDV > Dev_Value(I) Then
       minSDV = Dev_Value(I)
    End If
    If maxSdv < Dev_Value(I) Then
       maxSdv = Dev_Value(I)
    End If
    'MsgBox "Well Number  " + Str(Well_Num(I)) + "  Dev Value  " + Str(Dev_Value(I))
       Next I
       
     'skip 9 empty lines
     
     Input #FileNum, Temp
     Input #FileNum, Temp
        sample_size = Mid(Temp, 14, 5)
     Input #FileNum, Temp
         Std_sample = Mid(Temp, 14, 8)
         
     For I = 1 To 6
         Input #FileNum, Temp
     Next I
     
     'read frequencey data
     For I = 1 To 10
        Input #FileNum, Temp
           If I < 10 Then
             classfrom(I) = Mid(Temp, 2, 7)
             classto(I) = Mid(Temp, 12, 6)
             frequency(I) = Mid(Temp, 20, 8)
            Else
             classfrom(I) = Mid(Temp, 3, 7)
             classto(I) = Mid(Temp, 13, 6)
             frequency(I) = Mid(Temp, 21, 8)
            End If
    '     MsgBox "Values    " + Str(classfrom(I)) + "  " + Str(classto(I)) + "  " + Str(frequency(I))
         
         Next I
    Close FileNum
        
    ' Setup the chart so it will display correctly
    ChartVar1.ChartGroups(1).Styles(1).Symbol.size = 0
        
        
    ChartVar1.ChartGroups(1).ChartType = oc2dTypeBar
    
    
    If file_type = "Va1" Then
        title_str = "(Ranking Solution)"
    ElseIf file_type = "Va2" Then
        title_str = "(Scaling Solution)"
    End If
        
    
    'ChartVar1.Header.Text = "Deviation from Expected Stratigraphic Position for " + _
    '                            "Event: " + EventName(EventNum)
    
    
    'Label of chart
    
    ChartVar1.ChartLabels(3).Text = EventName(EventNum)
    ChartVar1.ChartLabels(3).AttachDataCoord.X = 0
    
    ' ChartVar1.ChartLabels(3).AttachDataCoord.y = 58
    'ChartVar1.ChartLabels(3).Font.size = 14
    tempvalue = minSDV
    minSDV = -maxSdv
    maxSdv = -tempvalue
    
    If Int(minSDV) < minSDV Then
    ChartVar1.ChartArea.Axes("y").Min = Int(minSDV)
    ChartVar1.ChartLabels(2).AttachDataCoord.Y = Int(minSDV)
    Else
    ChartVar1.ChartArea.Axes("y").Min = Int(minSDV) - 1
    ChartVar1.ChartLabels(2).AttachDataCoord.Y = Int(minSDV) - 1
    End If
    If Int(maxSdv) > maxSdv Then
    ChartVar1.ChartArea.Axes("y").Max = Int(maxSdv)
    ChartVar1.ChartLabels(1).AttachDataCoord.X = NumOfWell(EventNum)
    ChartVar1.ChartLabels(1).AttachDataCoord.Y = Int(maxSdv)
    ChartVar1.ChartLabels(3).AttachDataCoord.Y = Int(maxSdv)
    Else
    ChartVar1.ChartArea.Axes("y").Max = Int(maxSdv) + 1
    ChartVar1.ChartLabels(1).AttachDataCoord.X = NumOfWell(EventNum)
    ChartVar1.ChartLabels(1).AttachDataCoord.Y = Int(maxSdv) + 1
    ChartVar1.ChartLabels(3).AttachDataCoord.Y = Int(maxSdv) + 1
    End If
    
    'Label of sample size and Std
    
    ChartVar1.ChartLabels(1).Text = "Sample Size =  " + Str(sample_size) + "    " + "Standard Deviation =  " + Format(Std_sample, "##0.000")
    ChartVar1.ChartLabels(1).AttachCoord.X = 600
    ChartVar1.ChartLabels(1).AttachCoord.Y = 58
    ChartVar1.ChartLabels(1).Font.size = 10
    
    Chartvar2.ChartLabels(1).Text = title_str
    
    
    'ChartVar1.ChartArea.Axes("x").Title.Text = "Well No."
        With ChartVar1.ChartArea.Axes("x")
            .AnnotationMethod = oc2dAnnotateValueLabels
            With .valuelabels
            .RemoveAll
           End With
        End With
    
        With ChartVar1.ChartArea.Axes("x")
            .AnnotationMethod = oc2dAnnotateValueLabels
            With .valuelabels
              For I = 1 To NumOfWell(EventNum)
     
                .Add I, Well_Num(I)
              Next I
           End With
        End With
    
    With ChartVar1.ChartGroups(1).Data
        .Layout = oc2dDataArray
        .IsBatched = True
        .NumSeries = 1
        .NumPoints(1) = NumOfWell(EventNum)
                   
           For I = 1 To NumOfWell(EventNum)
            .X(1, I) = I
         'change the sign of the value
            .Y(1, I) = -Dev_Value(I)
            Next I
       
    End With
    
    ChartVar1.ChartGroups(1).Data.IsBatched = False
    ChartVar1.Visible = True
    
    Chartvar2.ChartGroups(1).ChartType = oc2dTypeBar
    'Chartvar2.Header.Text.Text = "  Event No. " + Str(NumOfEvent(EventNum)) + ":  " + EventName(EventNum)
     Chartvar2.ChartArea.Axes("x").Title.Text = "Classes"
    
    With Chartvar2.ChartGroups(1).Data
        .Layout = oc2dDataArray
        .IsBatched = True
        .NumSeries = 1
        .NumPoints(1) = 10
           For I = 1 To 10
            .X(1, I) = I
            .Y(1, I) = frequency(I)
          Next I
    End With
    
    Chartvar2.ChartGroups(1).Data.IsBatched = False
    'ChartVar2.Visible = True
    Chartvar2.Visible = True
    
End Sub


