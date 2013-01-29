VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmCum 
   Caption         =   "Cumulative Frequency"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   Icon            =   "frmCum.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   8940
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   2220
      TabIndex        =   4
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   3270
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1170
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdOpenCum 
      Caption         =   "Open File"
      Default         =   -1  'True
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D ChartCum 
      Height          =   7605
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   8745
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   15425
      _ExtentY        =   13414
      _StockProps     =   0
      ControlProperties=   "frmCum.frx":0442
   End
End
Attribute VB_Name = "frmCum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim startprint As Boolean
Dim leftx, topy, widthx, heighty As Integer


Private Sub cmdCancel_Click()
    Unload frmCum
End Sub

Private Sub cmdOpenCum_Click()
    frmOpenCum.Show 1
    'cmdOpenCum.Enabled = False
        If txt_file_cum = "" Then
            Exit Sub
        End If
        If Dir(CurDir + "\" + txt_file_cum) = "" Then
            Beep
            MsgBox "Input file does not exist in current directory, please try again"
            Exit Sub
        End If
        If filelen(CurDir + "\" + txt_file_cum) = 0 Then
            Beep
            MsgBox "Input file is empty, please try again"
            Exit Sub
        End If
       startup
       OpenFile
End Sub

Private Sub cmdPrint_Click()
'Dim h, w As Integer
'MsgBox "Select a print area using pointer"
   ChartCum.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
   ' ChartCum.PrintChart oc2dFormatEnhMetafile, oc2dScaleToFit, 0, 0, 0, 0
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub Command1_Click()

     Set CurChartSaveObj = ChartCum
     ChartSaveAs

'     ChartCum.DrawToFile Filename, oc2dFormatBitmap
'    Dim ImageName As String
'    ImageName = InputBox("Please give a file name without extension", "Save chart as image (.EMF)", 1)
'    If ImageName <> "" Then
'        ChartCum.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'        ChartCum.DrawToFile ImageName & ".EMF", oc2dFormatEnhMetafile
'        ChartCum.Save ImageName & ".OC2"
'    End If
End Sub

Private Sub Form_Activate()
     Set CurGraphicOBJ = ChartCum
    
    CurWindowNum = 1
    CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
     Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    cmdOpenCum.Enabled = True
    'Set CurGraphicOBJ = ChartCum
    'startup
    'openFile
    
    CurWindowNum = 1
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd
    
End Sub


Public Sub startup()
Dim I, J As Integer

ChartCum.ChartGroups(1).Styles(1).Symbol.size = 0

With ChartCum.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 1  '**** changed to 2 from 1
    .NumPoints(1) = 30
End With

ChartCum.ChartGroups(1).Data.IsBatched = False

End Sub


Public Sub OpenFile()
Dim FileNum As Integer
Dim Filename As String
Dim InputWell As String
Dim InputCum As String
Dim Temp As String
Dim numWells() As Integer
Dim Cum() As Integer
Dim MaxNumWell As Integer
Dim tempPost As Integer
Dim tempTitle As String
Dim I As Integer, J As Integer

FileNum = FreeFile
If txt_file_cum <> "" Then
    Open CurDir + "\" + txt_file_cum For Input As FileNum
Else
    MsgBox "Input file: " + CurDir + "\" + txt_file_cum + " does not exist, please try again"
Exit Sub
End If

InputWell = ""
InputCum = ""
MaxNumWell = 0

Do While EOF(FileNum) = False
    If EOF(FileNum) = False Then
        Line Input #FileNum, Temp
    End If
    If EOF(FileNum) = False Then
        Line Input #FileNum, Temp
    End If
    If EOF(FileNum) = False Then
        Line Input #FileNum, InputWell
    End If
    If EOF(FileNum) = False Then
        Line Input #FileNum, InputCum
    End If
    
    If (Trim(InputWell) = "" Or Trim(InputCum) = "") And MaxNumWell = 0 Then
           MsgBox "The input file " + CurDir + "\" + txt_file_cum + " is empty or the format is  not correct, " + Chr$(13) + "Please check it and try again." + Chr$(13)
           Close #FileNum
    '       OpenFileKey = 0
           Exit Sub
    End If
    If (Trim(InputWell) = "" Or Trim(InputCum) = "") And MaxNumWell > 0 Then
            Close #FileNum
    '       OpenFileKey = 0
           Exit Do
    End If
    If (Len(Trim(InputWell)) <> Len(Trim(InputCum))) Then
           MsgBox "The input file " + CurDir + "\" + txt_file_cum + " is empty or the format is  not correct, " + Chr$(13) + "Please check it and try again." + Chr$(13)
           Close #FileNum
    '       OpenFileKey = 0
           Exit Sub
    End If

    
    MaxNumWell = MaxNumWell + (Len(Trim(InputWell)) - 17) / 4
    'MsgBox Str(MaxNumWell)
Loop
Close #FileNum

ReDim numWells(1 To MaxNumWell)
ReDim Cum(1 To MaxNumWell)
'Read the data
Open CurDir + "\" + txt_file_cum For Input As FileNum
I = 0
Do While EOF(FileNum) = False
    If EOF(FileNum) = False Then
        Line Input #FileNum, Temp
    End If
    If EOF(FileNum) = False Then
        Line Input #FileNum, Temp
    End If
    If EOF(FileNum) = False Then
        Line Input #FileNum, InputWell
    End If
    If EOF(FileNum) = False Then
        Line Input #FileNum, InputCum
    End If
    
    If (Trim(InputWell) = "" Or Trim(InputCum) = "") And MaxNumWell = 0 Then
           MsgBox "The input file " + CurDir + "\" + txt_file_cum + " is empty or the format is  not correct, " + Chr$(13) + "Please check it and try again." + Chr$(13)
           Close #FileNum
    '       OpenFileKey = 0
           Exit Sub
    End If
    If (Trim(InputWell) = "" Or Trim(InputCum) = "") And MaxNumWell > 0 Then
            Close #FileNum
    '       OpenFileKey = 0
           Exit Do
    End If
    
    For J = 1 To (Len(Trim(InputWell)) - 17) Step 4
        I = I + 1
        numWells(I) = Mid(Trim(InputWell), J + 17, 4)
        Cum(I) = Mid(Trim(InputCum), J + 17, 4)
    Next J
    'MsgBox Str(MaxNumWell)
Loop
Close #FileNum

'For I = 1 To MaxNumWell
'    numWells(I) = Mid(InputWell, 18 + 4 * (I - 1), 4)
'    Cum(I) = Mid(InputCum, 18 + 4 * (I - 1), 4)
''MsgBox Str(numWells(I)) + "   " + Str(Cum(I))
'Next I
 
ChartCum.ChartGroups(1).Styles(1).Symbol.size = 0
ChartCum.ChartGroups(1).Styles(1).Line.Width = 2
tempPost = InStr(1, txt_file_cum, ".")
tempTitle = Mid(txt_file_cum, 1, tempPost - 1)

ChartCum.ChartLabels(1).Text = "File Name: " + tempTitle

With ChartCum.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 1
    .NumPoints(1) = MaxNumWell
    
      For I = 1 To MaxNumWell
        .X(1, I) = numWells(I)
        .Y(1, I) = Cum(I)
      Next I
End With

ChartCum.ChartGroups(1).Data.IsBatched = False
ChartCum.Visible = True
End Sub

Private Sub Form_Resize()
Dim size As Integer
If (frmCum.Height > 360) And (frmCum.Height - 360 < frmCum.Width) Then
   size = frmCum.Height - 360
ElseIf (frmCum.Height > 360) And (frmCum.Height - 360 >= frmCum.Width) Then
   size = frmCum.Width
Else
   size = 0
End If

ChartCum.Height = size
ChartCum.Width = size
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set CurGraphicOBJ = Nothing
   
   CurWindowNum = 1
   Call MDIWindowsMenuDelete(CurWindowNum)
   WindowsHwnd(CurWindowNum) = -1

End Sub
