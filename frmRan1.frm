VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmRan1 
   AutoRedraw      =   -1  'True
   Caption         =   "Event Ranges - Ranking"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12810
   Icon            =   "frmRan1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDShowfrmRAN2 
      Caption         =   "To Show Cross-over Ranges"
      Height          =   255
      Left            =   3930
      TabIndex        =   5
      Top             =   120
      Width           =   2445
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   1950
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1020
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   252
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D ChartRan1 
      Height          =   7815
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   11505
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   20294
      _ExtentY        =   13785
      _StockProps     =   0
      ControlProperties=   "frmRan1.frx":0442
   End
End
Attribute VB_Name = "frmRan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdCancel_Click()
  Unload frmRan1
End Sub

Private Sub CancelButton_Click()
  Unload frmRan2
  Unload frmRan1
End Sub

Private Sub cmdOpen_Click()
file_type = "Ra1"
frmOpenRan1.Show 1
'cmdOpen.Enabled = False

 If txt_file_ran1 = "" Then
      Beep
      Exit Sub
 End If
 Check1 = txt_file_ran1
 txt_file_range1 = txt_file_ran1


startup
OpenFile

End Sub

Private Sub cmdPrint_Click()
    ChartRan1.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
    'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub CMDShowfrmRAN2_Click()
   frmRan2.Show
   frmRan2.SetFocus
End Sub

Private Sub Command1_Click()

     Set CurChartSaveObj = ChartRan1
     ChartSaveAs

'   Dim ImageName As String
'   ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'   If ImageName <> "" Then
'       ChartRan1.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'    End If
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    txt_file_ran1 = ""
    cmdOpen.Enabled = True
    'checking = True
    Check1 = ""
    txt_file_range1 = ""
    
    CurWindowNum = 8
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

End Sub


Public Sub startup()
Dim I, J As Integer

    'Try to clear all Y annotation data on current chartarea
    ChartRan1.Visible = False
    ChartRan1.ChartArea.Axes("Y").valuelabels.RemoveAll
    'Try to clear all X annotation data on current chartarea
    ChartRan1.ChartArea.Axes("X").valuelabels.RemoveAll

ChartRan1.ChartGroups(1).Styles(1).Symbol.size = 0
ChartRan1.ChartGroups(1).ChartType = oc2dTypeHiLo
With ChartRan1.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 1  '**** changed to 2 from 1
    .NumPoints(1) = 150
   
End With

ChartRan1.ChartGroups(1).Data.IsBatched = False

End Sub



Public Sub OpenFile()
Dim FileNum As Integer
Dim Filename As String
Dim minrange() As Double
Dim Maxrange() As Double
Dim center()  As Double
Dim num() As Integer
Dim MaxNumEvent As Integer
Dim EventName() As String
Dim Temp As String
Dim tempfrom As Integer
Dim tempto As Integer
Dim tempstr As String
Dim tempstr1 As String
Dim minx, maxx As Double
Dim I As Integer
Dim title_str As String
FileNum = FreeFile
If txt_file_range1 <> "" Then
    Open CurDir + "\" + txt_file_range1 For Input As FileNum
Else
    MsgBox "Input file is not exist, please try again"
Exit Sub
End If

Input #FileNum, Temp
Input #FileNum, Temp
Input #FileNum, Temp
'    MsgBox temp
MaxNumEvent = Mid(Temp, 26, 5)
'MsgBox MaxNumEvent
ReDim minrange(1 To MaxNumEvent)
ReDim center(1 To MaxNumEvent)
ReDim Maxrange(1 To MaxNumEvent)
ReDim num(1 To MaxNumEvent)
ReDim EventName(1 To MaxNumEvent)
minx = 0
maxx = 0
Input #FileNum, Temp
For I = 1 To MaxNumEvent
   Input #FileNum, Temp
     If I < 10 Then
        tempfrom = 3
     ElseIf I < 100 Then
        tempfrom = 2
     ElseIf I < 1000 Then
         tempfrom = 1
     Else
         tempfrom = 0
     End If
     
   minrange(I) = Mid(Temp, 5 - tempfrom, 10)
   center(I) = Mid(Temp, 15 - tempfrom, 8)
   Maxrange(I) = Mid(Temp, 23 - tempfrom, 8)
   If Maxrange(I) < minx Then
          minx = Maxrange(I)
    End If
   If minrange(I) > maxx Then
         maxx = minrange(I)
    End If
    
   num(I) = Mid(Temp, 31 - tempfrom, 5)
   EventName(I) = Mid(Temp, 36 - tempfrom, 45)
'MsgBox Str(minrange(I)) + Str(center(I)) + Str(Maxrange(I)) + Str(num(I)) + eventName(I)
 Next I

Close FileNum

ChartRan1.ChartGroups(1).Styles(1).Symbol.size = 0
ChartRan1.ChartGroups(1).Styles(1).Line.Width = 2
ChartRan1.ChartGroups(1).Styles(1).Line.Color = 0
ChartRan1.ChartGroups(1).ChartType = oc2dTypeHiLo
ChartRan1.ChartArea.Axes("x").Min.Value = 0
ChartRan1.ChartArea.Axes("y").Min.Value = minx
ChartRan1.ChartArea.Axes("y2").Min.Value = minx
ChartRan1.ChartArea.Axes("Y").Max.Value = Int(maxx) + 1
ChartRan1.ChartArea.Axes("y2").Max.Value = Int(maxx) + 1

ChartRan1.ChartArea.Axes("x").Origin.Value = 0
ChartRan1.ChartArea.Axes("y").Origin.Value = minx
If file_type = "Ra1" Then
    title_str = "Ranking"
ElseIf file_type = "Ra2" Then
    title_str = "Scaling"
End If
    
ChartRan1.Header.Text.Text = "Event Ranges (" + title_str + ")"
ChartRan1.ChartLabels(1).AttachDataCoord.y = (maxx + minx) / 2
ChartRan1.ChartLabels(1).AttachDataCoord.x = MaxNumEvent
                            
 If title_str = "Scaling" Then
 ChartRan1.ChartLabels(1).Text = "Scaled Optimum Sequence of Events"
 ElseIf title_str = "Ranking" Then
 ChartRan1.ChartLabels(1).Text = "Ranked Optimum Sequence of Events"
 End If
                            
With ChartRan1.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 2
    .NumPoints(1) = MaxNumEvent
    .NumPoints(2) = MaxNumEvent
      For I = 1 To MaxNumEvent
        .x(1, I) = I
        .y(1, I) = minrange(I)
        .y(2, I) = Maxrange(I)
      
       Next I
End With

ChartRan1.ChartGroups(1).Data.IsBatched = False
 
   
With ChartRan1.ChartArea.Axes("X")
        .AnnotationMethod = oc2dAnnotateValueLabels
        With .valuelabels
            For I = 1 To MaxNumEvent
            If num(I) < 10 Then
               tempstr = Space(6)
            ElseIf num(I) < 100 Then
               tempstr = Space(4)
            Else
               tempstr = Space(2)
               End If
            If I < 10 Then
            tempstr1 = Space(6)
            ElseIf I < 100 Then
            tempstr1 = Space(4)
            ElseIf I < 1000 Then
            tempstr1 = Space(2)
            Else
            tempstr1 = Space(0)
            End If
            
            .Add I, tempstr1 + Str(I) + tempstr + Str(num(I)) + "   " + EventName(I)
            Next I
       End With
    End With
 
  
ChartRan1.ChartGroups(2).Styles(1).Symbol.size = 5
ChartRan1.ChartGroups(2).Styles(1).Line.Width = 2
ChartRan1.ChartGroups(2).Styles(1).Line.Color = 255
ChartRan1.ChartGroups(2).ChartType = oc2dTypePlot
With ChartRan1.ChartGroups(2).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 1
    .NumPoints(1) = MaxNumEvent
       For I = 1 To MaxNumEvent
        .x(1, I) = I
        .y(1, I) = center(I)
  
      Next I
End With

ChartRan1.ChartGroups(2).Data.IsBatched = False
ChartRan1.Visible = True
End Sub

Private Sub Form_Resize()
    Dim size As Integer
    If (frmRan1.Height > 360) And (frmRan1.Height - 360 < frmRan1.Width / 1.5) Then
       size = frmRan1.Height - 360
    ElseIf (frmRan1.Height > 360) And (frmRan1.Height - 360 > frmRan1.Width / 1.5) Then
       size = frmRan1.Width / 1.5
    Else
       size = 0
    End If
    
    ChartRan1.Height = size
    ChartRan1.Width = 1.5 * size
End Sub



Private Sub Form_Activate()
'Set the focus to the command button to force the allowance of
' the <F1> key to bring up help
    Set CurGraphicOBJ = ChartRan1
    CurWindowNum = 8
    CurWindowSetFocus (CheckExistWindows(CurWindowNum))
    'cmdReset.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
'End the program.
    Set CurGraphicOBJ = Nothing
    Unload frmRan2
    Unload frmRan1
    
    CurWindowNum = 8
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1
    'End
End Sub



