VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmDenRank 
   Caption         =   "Ranked Optimum Sequence"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   Icon            =   "frmDenRank.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   12105
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   2070
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   252
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   252
      Left            =   1110
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open File"
      Default         =   -1  'True
      Height          =   252
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D Chartden2 
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
      ControlProperties=   "frmDenRank.frx":1982
   End
End
Attribute VB_Name = "frmDenRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileNameCurrent As String


Private Sub cmdCancel_Click()
  Unload frmDenRank
End Sub

Private Sub cmdOpen_Click()
   frmOpenRank.Show 1
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
    Chartden2.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

 
Private Sub Command1_Click()

     Set CurChartSaveObj = Chartden2
     ChartSaveAs

'    Dim ImageName As String
'    ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'    If ImageName <> "" Then
'        Chartden2.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'    End If
End Sub
 

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    cmdOpen.Enabled = True
    Chartden2.Visible = False

    OpenFileKey = 0
    CancelKeyPress = 0
    FileNameCurrent = ""
'startup
'openFile
    CurWindowNum = 2
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

End Sub


Public Sub startup()
    Dim I, J As Integer
    Chartden2.Visible = False
    Chartden2.ChartArea.Axes("X").valuelabels.RemoveAll
    Chartden2.ChartArea.Axes("Y").valuelabels.RemoveAll
    Chartden2.ChartArea.Axes("Y2").valuelabels.RemoveAll
    
    Chartden2.ChartGroups(1).Data.IsBatched = False
    Chartden2.ChartGroups(1).ChartType = oc2dTypeBar
    With Chartden2.ChartGroups(1).Data
        .Layout = oc2dDataGeneral
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
        .NumPoints(1) = 3
        
    End With
    Chartden2.ChartGroups(1).Data.IsBatched = False
    Chartden2.ChartGroups(2).Data.IsBatched = False
    Chartden2.Visible = False
End Sub

 

Public Sub OpenFile()
Dim FileNum As Integer
Dim Filename As String
Dim Temp1 As String
Dim Temp As String
Dim NumOFEntries As Integer
Dim numEvent As Integer
Dim SD As Double
Dim AVESD As Double
Dim labels As String
Dim I As Integer
Dim J As Integer
Dim len_label As Integer
Dim maxSD As Double
'On Error GoTo Endp
Chartden2.Visible = False

FileNum = FreeFile
If txt_file_den = "" Then
        MsgBox "Input file does not exist, please try again"
        Exit Sub

Else

        Open CurDir + "\" + txt_file_den For Input As FileNum
        
        Input #FileNum, Temp
        NumOFEntries = Mid(Temp, 21, 4)
         
        AVESD = Mid(Temp, 36, 8)
         
          For I = 1 To 6
              If Not EOF(FileNum) Then Input #FileNum, Temp
          Next I
          
          With Chartden2.ChartArea.Axes("x")
                .AnnotationMethod = oc2dAnnotateValueLabels
             End With
          
          With Chartden2.ChartGroups(1).Data
            .Layout = oc2dDataArray
            .IsBatched = True
            .NumSeries = 1
            .NumPoints(1) = NumOFEntries
           maxSD = 0
          For J = 1 To NumOFEntries
           If Not EOF(FileNum) Then Input #FileNum, Temp
                SD = right(Temp, 7)
                       
                numEvent = left(Temp, 3)
                  .X(1, J) = numEvent
                   If Mid(Temp, 6, 2) = "99" Then
                   .Y(1, J) = 0
               
                 len_label = Len(Temp)
                 Mid(Temp, len_label - 8, 9) = "         "
                 Mid(Temp, 4, 5) = "      "
                           
                With Chartden2.ChartArea.Axes("x").valuelabels
                 .Add J, Temp
                 End With
                   
                   Else
                       If SD > maxSD Then maxSD = SD
                   .Y(1, J) = -SD
                 len_label = Len(Temp)
                 Temp1 = Mid(Temp, len_label - 63, 5)
                    If Temp1 = "  0 0" Then
                        Mid(Temp, len_label - 63, 5) = "      "
                    Else
                        Mid(Temp, len_label - 63, 1) = ""
                        Mid(Temp, len_label - 58, 1) = "  "
                    End If
                  With Chartden2.ChartArea.Axes("x").valuelabels
                   .Add J, Temp
                 End With
               End If
              Next J
                
             'add the line for AVE SD
                
        End With
         Chartden2.ChartArea.Axes("x").Max = NumOFEntries + 1
          Chartden2.ChartArea.Axes("x").Min = -0.5
        Close FileNum
         
         With Chartden2.ChartGroups(2).Data
            .Layout = oc2dDataArray
            .IsBatched = True
            .NumSeries = 1
            .NumPoints(1) = 2
            .X(1, 1) = -0.5
            .Y(1, 1) = -AVESD
            .X(1, 2) = NumOFEntries + 1
            .Y(1, 2) = -AVESD
            
          End With
        Chartden2.ChartArea.Axes("y").DataMin = -Int(maxSD) - 1
        Chartden2.ChartArea.Axes("y").DataMax = 0
        Chartden2.ChartArea.Axes("y2").DataMax = 0
        Chartden2.ChartArea.Axes("y2").DataMin = -Int(maxSD) - 1
         With Chartden2.ChartArea.Axes("y2").valuelabels
                   .Add -AVESD, "'Ave' SD"
                 End With
         With Chartden2.ChartArea.Axes("x").valuelabels
                   .Add -0.5, " No. & UI                                                   N   SD  "
                 End With
           Chartden2.ChartGroups(1).Data.IsBatched = False
           Chartden2.ChartGroups(2).Data.IsBatched = False
        Chartden2.Visible = True
End If
Endp:

End Sub


Private Sub Form_Resize()
Dim size As Integer
If (frmDenRank.Height > 360) And (frmDenRank.Height - 360 < frmDenRank.Width) Then
   size = frmDenRank.Height - 360
ElseIf (frmDenRank.Height > 360) And (frmDenRank.Height - 360 > frmDenRank.Width) Then
   size = frmDenRank.Width
Else
   size = 0
End If

Chartden2.Height = size
Chartden2.Width = size
End Sub




Private Sub Form_Activate()
'Set the focus to the command button to force the allowance of
' the <F1> key to bring up help
     Set CurGraphicOBJ = Chartden2
        CurWindowNum = 2
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
    'cmdReset.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
'End the program.
    Set CurGraphicOBJ = Nothing
    
    CurWindowNum = 2
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

    'End
End Sub



