VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmChartShow 
   Caption         =   "Display Graph"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   Icon            =   "frmChartShow.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdCopyToClipboard 
      Caption         =   "Copy to Clipboard"
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   90
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   90
      Width           =   852
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Top             =   90
      Width           =   852
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   90
      Width           =   852
   End
   Begin VB.CommandButton cmdOpenCum 
      Caption         =   "Open File"
      Default         =   -1  'True
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D ChartShow 
      Height          =   7605
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   11085
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   19553
      _ExtentY        =   13414
      _StockProps     =   0
      ControlProperties=   "frmChartShow.frx":0442
   End
End
Attribute VB_Name = "frmChartShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim startprint As Boolean
Dim leftx, topy, widthx, heighty As Integer


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCopytoClipboard_Click()
    If ChartShow.Visible Then
        ChartShow.CopyToClipboard oc2dFormatEnhMetafile
    End If
End Sub

Private Sub cmdOpenCum_Click()
    Call OpenSavedChart
End Sub

Private Sub cmdPrint_Click()
'Dim h, w As Integer
'MsgBox "Select a print area using pointer"
   ChartShow.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
   ' ChartShow.PrintChart oc2dFormatEnhMetafile, oc2dScaleToFit, 0, 0, 0, 0
   'Some printer program will change current path, here recover the workplace one time
   ChDir CurrentDir
   ChDrive CurrentDrive

End Sub

Private Sub Command1_Click()

     Set CurChartSaveObj = ChartShow
     ChartSaveAs

'     ChartShow.DrawToFile Filename, oc2dFormatBitmap
'    Dim ImageName As String
'    ImageName = InputBox("Please give a file name without extension", "Save chart as image (.EMF)", 1)
'    If ImageName <> "" Then
'        ChartShow.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'        ChartShow.DrawToFile ImageName & ".EMF", oc2dFormatEnhMetafile
'        ChartShow.Save ImageName & ".OC2"
'    End If
End Sub

Private Sub Form_Activate()
     Set CurGraphicOBJ = ChartShow
    
    CurWindowNum = 29
    CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
     Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    
    CurWindowNum = 29
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd
    
    Set CurChartShowObj = ChartShow
    
End Sub


Public Sub startup()
    Dim I, J As Integer
    
    ChartShow.ChartGroups(1).Styles(1).Symbol.size = 0
    
    With ChartShow.ChartGroups(1).Data
        .Layout = oc2dDataGeneral
        .IsBatched = True
        .NumSeries = 1  '**** changed to 2 from 1
        .NumPoints(1) = 30
    End With
    
    ChartShow.ChartGroups(1).Data.IsBatched = False
    
End Sub


Private Sub Form_Resize()
        Dim size As Integer
        If (frmChartShow.Height > 360) And (frmChartShow.Height - 360 < frmChartShow.Width) Then
           size = frmChartShow.Height - 360
        ElseIf (frmChartShow.Height > 360) And (frmChartShow.Height - 360 >= frmChartShow.Width) Then
           size = frmChartShow.Width
        Else
           size = 0
        End If
        
        ChartShow.Height = size
        ChartShow.Width = size
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set CurGraphicOBJ = Nothing
   
   CurWindowNum = 29
   Call MDIWindowsMenuDelete(CurWindowNum)
   WindowsHwnd(CurWindowNum) = -1
   
   Set CurChartShowObj = Nothing
   ChartShow.Visible = False

End Sub
