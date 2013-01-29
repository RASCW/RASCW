VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmDen 
   Caption         =   "Scaled Optimum Sequence Dendrogram"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "frmDen.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   252
      Left            =   2940
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   252
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open File"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D Chartden 
      Height          =   7404
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   8172
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   14414
      _ExtentY        =   13060
      _StockProps     =   0
      ControlProperties=   "frmDen.frx":0442
   End
End
Attribute VB_Name = "frmDen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileNameCurrent As String


Private Sub cmdCancel_Click()
Unload frmDen
End Sub

Private Sub cmdOpen_Click()
        frmOpenDen.Show 1
        'cmdOpen.Enabled = False
        If CancelKeyPress = 1 Then
            CancelKeyPress = 0
            Exit Sub
        End If
        If txt_file_den = "" Or Dir(CurDir + "\" + txt_file_den) = "" Then
            Beep
            MsgBox "Input file does not exist, please try again"
            OpenFileKey = 0
            Exit Sub
        End If
        If OpenFileKey = 1 Then
            startup
            OpenFile
            FileNameCurrent = CurDir + "\" + txt_file_den
        End If
End Sub

Private Sub cmdPrint_Click()

    Chartden.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
    'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

 
Private Sub Command1_Click()

     Set CurChartSaveObj = Chartden
     ChartSaveAs

'Dim ImageName As String
'ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'If ImageName <> "" Then
'    Chartden.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'End If
End Sub
 

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
  cmdOpen.Enabled = True
  
  OpenFileKey = 0
  CancelKeyPress = 0
  FileNameCurrent = ""
'startup
'openFile
    CurWindowNum = 3
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

End Sub


Public Sub startup()
    Dim I, J As Integer
    
    'Try to clear all Y annotation data on current chartarea
    Chartden.Visible = False
    Chartden.ChartArea.Axes("Y").valuelabels.RemoveAll


    Chartden.ChartGroups(1).Styles(1).Symbol.size = 0
    With Chartden.ChartGroups(1).Data
        .Layout = oc2dDataGeneral
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
        .NumPoints(1) = 3
        
    End With
    Chartden.ChartGroups(1).Data.IsBatched = False
End Sub



Public Function add_Fit(ByVal m As Integer, ByVal x1, x2, x3 As Double, ByVal y1, y2, y3 As Integer)
Dim I, J As Integer

Chartden.ChartGroups(1).Styles(m).Symbol.size = 0
Chartden.ChartGroups(1).Styles(m).Line.Width = 1.5
Chartden.ChartGroups(1).Styles(m).Line.Color = 0

With Chartden.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = m
    .NumPoints(m) = 3
    
        .X(m, 1) = x1
        .X(m, 2) = x2
        .X(m, 3) = x3
        .Y(m, 1) = y1
        .Y(m, 2) = y2
        .Y(m, 3) = y3
  
End With

Chartden.ChartGroups(1).Data.IsBatched = False


End Function


Public Function valuelabels(ByVal k As Double, ByVal ch As String)
With Chartden.ChartArea.Axes("Y")
        .AnnotationMethod = oc2dAnnotateValueLabels
        With .valuelabels
            .Add k, ch
       End With
    End With

'The ValueLabels collection can be indexed either by subscript or by value:'

' this retrieves the label for the second Value-label
'    Value = Chart2D1.ChartArea.Axes("X").ValueLabels(2).Value;
'    ' this retrieves the label at chart coordinate 2.0
'    Value = Chart2D1.ChartArea.Axes("X").ValueLabels(2.0).Value;
End Function

Public Sub OpenFile()
Dim FileNum As Integer
Dim Filename As String
Dim Temp1 As String * 20
Dim Temp As String
Dim NumOFEvents1, NumOFEvents2, NumOfEvents3, NumOfEvents As Integer
Dim numEvent() As Integer
Dim correlation() As Double
Dim EventName() As String * 50   'to include the , number as part of the event name
Dim MaxCorr, minCorr As Double
Dim distance() As Integer
Dim I As Integer
Dim J As Integer
Dim k As Integer
Dim Temp_Num As Double
Dim pass
Dim mystring
Dim final_second As Integer
Dim skipline As Integer

FileNum = FreeFile
If txt_file_den = "" Or Dir(CurDir + "\" + txt_file_den) = "" Then
    Beep
    MsgBox "Input file does not exist, please try again"
    Exit Sub

Else

        Open CurDir + "\" + txt_file_den For Input As FileNum
        
        Input #FileNum, Temp
        Input #FileNum, NumOfEvents
        
        'final_second = Len(Temp)
        ''MsgBox Str(final_second)
        'NumOFEvents1 = Mid(Temp, 19, 4)
        'NumOFEvents2 = Mid(Temp, 23, 4)
        'NumOfEvents = NumOFEvents2
        'MsgBox Str(NumOFEvents1) + "     " + Str(NumOFEvents2)
        
        NumOFEvents1 = NumOfEvents
        NumOFEvents2 = NumOfEvents
        
        
        'skipline = 3 + NumOFEvents1 + 2
        
        'If final_second > 26 Then
        'NumOfEvents3 = Mid(Temp, 27, 4)
        ''MsgBox Str(NumOfEvents3)
        'NumOfEvents = NumOfEvents3
        'skipline = 3 + NumOFEvents1 + 3 + NumOFEvents2 + 2
        'End If
        
        'For I = 1 To skipline
        'Input #FileNum, Temp
        'Next I
        
        
        'define arrays
        
        ReDim numEvent(1 To NumOfEvents)
        ReDim correlation(1 To NumOfEvents)
        ReDim EventName(1 To NumOfEvents)
        ReDim distance(1 To NumOfEvents)
        
        
        
        'determine the final or second
        
        Input #FileNum, MaxCorr, minCorr
         For I = 1 To NumOfEvents - 1
        Input #FileNum, numEvent(I), correlation(I), EventName(I), Temp_Num
        
        'MsgBox Str(numEvent(I)) + Str(correlation(I)) + "  " + EventName(I)
        Next I
        
        Input #FileNum, numEvent(NumOfEvents), Temp_Num, EventName(NumOfEvents)
        correlation(NumOfEvents) = MaxCorr
        
        Close FileNum

End If

For I = 1 To NumOfEvents - 1
   J = I
   Do
       distance(I) = J + 1
       J = J + 1
   Loop While correlation(I) > correlation(J)
 'MsgBox Str(distance(I))
 Next I
  For I = 1 To NumOfEvents - 1
    If left(EventName(I), 2) <> "**" Then
    pass = add_Fit(I, 0, correlation(I), correlation(I), I, I, distance(I))
    End If
       
    J = 10 - Len(Str(numEvent(I)))
    If numEvent(I) < 10 Then
    mystring = Space(8)
    ElseIf numEvent(I) < 100 Then
    mystring = Space(6)
    ElseIf numEvent(I) < 1000 Then
    mystring = Space(4)
    ElseIf numEvent(I) < 10000 Then
    mystring = Space(2)
    End If
    If left(EventName(I), 2) = "**" Then
    mystring = Space(Len(mystring) - 3)
    End If
 pass = valuelabels(I, Str(numEvent(I)) + mystring + EventName(I))
' pass = valuelabels(I + 0.5, "           " + Str(correlation(I)))
 Next I
 I = NumOfEvents
     If left(EventName(I), 2) <> "**" Then
     pass = add_Fit(I, 0, correlation(I), correlation(I), I, I, I)
    End If
    If numEvent(NumOfEvents) < 10 Then
    mystring = Space(8)
    ElseIf numEvent(NumOfEvents) < 100 Then
    mystring = Space(6)
    ElseIf numEvent(NumOfEvents) < 1000 Then
    mystring = Space(4)
    ElseIf numEvent(NumOfEvents) < 10000 Then
    mystring = Space(2)
    End If
    If left(EventName(I), 2) = "**" Then
    mystring = Space(Len(mystring) - 3)
    End If
 pass = valuelabels(NumOfEvents, Str(numEvent(NumOfEvents)) + mystring + _
            EventName(NumOfEvents))
 Chartden.ChartArea.Axes("x").Max.Value = correlation(NumOfEvents)
    
Chartden.ChartArea.Axes("y").Max.Value = NumOfEvents
Chartden.Visible = True

End Sub


Private Sub Form_Resize()
    Dim size As Integer
    If (frmDen.Height > 360) And (frmDen.Height - 360 < frmDen.Width) Then
       size = frmDen.Height - 360
    ElseIf (frmDen.Height > 360) And (frmDen.Height - 360 > frmDen.Width) Then
       size = frmDen.Width
    Else
       size = 0
    End If
    
    Chartden.Height = size
    Chartden.Width = size
End Sub


Private Sub Form_Activate()
    'Set the focus to the command button to force the allowance of
    ' the <F1> key to bring up help
        Set CurGraphicOBJ = Chartden
        CurWindowNum = 3
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))

        'cmdReset.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
'End the program.
    Set CurGraphicOBJ = Nothing
    
    CurWindowNum = 3
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

    'End
End Sub


