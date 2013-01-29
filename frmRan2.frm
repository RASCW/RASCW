VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmRan2 
   AutoRedraw      =   -1  'True
   Caption         =   "Event Ranges - Ranking"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   Icon            =   "frmRan2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   12750
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDShowfrmRan1 
      Caption         =   "To Show Event Ranges"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   3525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   1020
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   1830
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D ChartRan2 
      Height          =   7815
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   11505
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   20294
      _ExtentY        =   13785
      _StockProps     =   0
      ControlProperties=   "frmRan2.frx":0442
   End
End
Attribute VB_Name = "frmRan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload frmRan2
    Unload frmRan1
End Sub
 

Private Sub cmdPrint_Click()
    ChartRan2.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
    'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub


Public Sub startup()
Dim I, J As Integer

  'Try to clear all Y annotation data on current chartarea
  ChartRan2.Visible = False
  ChartRan2.ChartArea.Axes("Y").valuelabels.RemoveAll
  'Try to clear all X annotation data on current chartarea
  ChartRan2.ChartArea.Axes("X").valuelabels.RemoveAll

    ChartRan2.ChartGroups(1).Styles(1).Symbol.size = 0
    ChartRan2.ChartGroups(1).ChartType = oc2dTypeHiLo
    With ChartRan2.ChartGroups(1).Data
        .Layout = oc2dDataGeneral
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
        .NumPoints(1) = 150
       
    End With
    
    ChartRan2.ChartGroups(1).Data.IsBatched = False
    
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
  'skip lines to read the second set of data
  For I = 1 To MaxNumEvent
     If Not (EOF(FileNum)) Then
     Input #FileNum, Temp
     End If
  Next I
  
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

ChartRan2.ChartGroups(1).Styles(1).Symbol.size = 0
ChartRan2.ChartGroups(1).Styles(1).Line.Width = 2
ChartRan2.ChartGroups(1).Styles(1).Line.Color = 0
ChartRan2.ChartGroups(1).ChartType = oc2dTypeHiLo
ChartRan2.ChartArea.Axes("x").Min.Value = 0
ChartRan2.ChartArea.Axes("y").Min.Value = minx
ChartRan2.ChartArea.Axes("y2").Min.Value = minx
ChartRan2.ChartArea.Axes("Y").Max.Value = Int(maxx) + 1
ChartRan2.ChartArea.Axes("y2").Max.Value = Int(maxx) + 1

ChartRan2.ChartArea.Axes("x").Origin.Value = 0
ChartRan2.ChartArea.Axes("y").Origin.Value = minx
If file_type = "Ra1" Then
    title_str = "Ranking"
ElseIf file_type = "Ra2" Then
    title_str = "Scaling"
End If
    
ChartRan2.Header.Text.Text = "Cross-over Ranges (" + title_str + ")"
 ChartRan2.ChartLabels(1).AttachDataCoord.y = (maxx + minx) / 2
ChartRan2.ChartLabels(1).AttachDataCoord.x = MaxNumEvent

 If title_str = "Scaling" Then
 ChartRan2.ChartLabels(1).Text = "Scaled Optimum Sequence of Events"
 ElseIf title_str = "Ranking" Then
 ChartRan2.ChartLabels(1).Text = "Ranked Optimum Sequence of Events"
 End If

With ChartRan2.ChartGroups(1).Data
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

ChartRan2.ChartGroups(1).Data.IsBatched = False
 
   
With ChartRan2.ChartArea.Axes("X")
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
 
  
ChartRan2.ChartGroups(2).Styles(1).Symbol.size = 5
ChartRan2.ChartGroups(2).Styles(1).Line.Width = 2
ChartRan2.ChartGroups(2).Styles(1).Line.Color = 255
ChartRan2.ChartGroups(2).ChartType = oc2dTypePlot
With ChartRan2.ChartGroups(2).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 1
    .NumPoints(1) = MaxNumEvent
       For I = 1 To MaxNumEvent
        .x(1, I) = I
        .y(1, I) = center(I)
  
      Next I
End With

ChartRan2.ChartGroups(2).Data.IsBatched = False
ChartRan2.Visible = True
End Sub


Private Sub CMDShowfrmRan1_Click()
   frmRan1.Show
   frmRan1.SetFocus
     
End Sub

Private Sub Command1_Click()

     Set CurChartSaveObj = ChartRan2
     ChartSaveAs

'  Dim ImageName As String
'  ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'  If ImageName <> "" Then
'      ChartRan2.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'    End If
End Sub

'Private Sub Command1_Click()
'Dim response As Boolean
'response = ChartRan2.CopyToClipboard(oc2dFormatBitmap)
'If response = False Then
'    MsgBox "It is not successful! Try it again!"
'End If
'
'End Sub

Private Sub Form_Activate()
    If (Check1 <> "") And (txt_file_range1 <> "") Then
        startup
        OpenFile
        Set CurGraphicOBJ = ChartRan2
        'checking = False
    End If

End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

'Private Sub Form_GotFocus()
'If (txt_file_range1 <> "") And (Check1 <> "") Then
'    startup
'    OpenFile
'    'checking = False
'End If
'End Sub

Private Sub Form_Resize()
Dim size As Integer
If (frmRan2.Height > 360) And (frmRan2.Height - 360 < frmRan2.Width / 1.5) Then
   size = frmRan2.Height - 360
ElseIf (frmRan2.Height > 360) And (frmRan2.Height - 360 > frmRan2.Width / 1.5) Then
   size = frmRan2.Width / 1.5
Else
   size = 0
End If

ChartRan2.Height = size
ChartRan2.Width = 1.5 * size
End Sub


Private Sub Form_Unload(Cancel As Integer)
'End the program.
    Set CurGraphicOBJ = Nothing
    
    Unload frmRan2
    Unload frmRan1
    'End
End Sub



