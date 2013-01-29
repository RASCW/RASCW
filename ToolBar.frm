VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ToolBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "ToolBar"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ToolBar.frx":0000
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCopyGraphToClipboard 
      Height          =   375
      Left            =   540
      Picture         =   "ToolBar.frx":29084
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Copy Graph to Clipboard"
      Top             =   2100
      Width           =   375
   End
   Begin VB.CommandButton CmdViewSavedGraph 
      Height          =   375
      Left            =   930
      Picture         =   "ToolBar.frx":2953C
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Display Saved Graphs"
      Top             =   2100
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   960
      Top             =   6390
   End
   Begin VB.CommandButton CmdAbout 
      BackColor       =   &H00FFC0C0&
      Caption         =   "A"
      Height          =   375
      Left            =   570
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "About"
      Top             =   7800
      Width           =   765
   End
   Begin VB.CommandButton CmdHelp 
      BackColor       =   &H00C0FFC0&
      Caption         =   "H"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Help"
      Top             =   7800
      Width           =   405
   End
   Begin VB.CommandButton CmdWindowArrange 
      BackColor       =   &H0080C0FF&
      Caption         =   "W4"
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "ToolBar.frx":29A30
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Arrange Icons"
      Top             =   7290
      Width           =   405
   End
   Begin VB.CommandButton CmdWindowArrange 
      BackColor       =   &H0080C0FF&
      Caption         =   "W3"
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "ToolBar.frx":2A334
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Windows Cascade"
      Top             =   6900
      Width           =   405
   End
   Begin VB.CommandButton CmdWindowArrange 
      BackColor       =   &H0080C0FF&
      Caption         =   "W2"
      Height          =   375
      Index           =   1
      Left            =   540
      Picture         =   "ToolBar.frx":2AC38
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Windows Tile Vertically"
      Top             =   6900
      Width           =   405
   End
   Begin VB.CommandButton CmdWindowArrange 
      BackColor       =   &H0080C0FF&
      Caption         =   "W1"
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ToolBar.frx":2B53C
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Windows Tile Horizontally"
      Top             =   6900
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T8"
      Height          =   375
      Index           =   7
      Left            =   540
      Picture         =   "ToolBar.frx":2BE40
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Dictionary"
      Top             =   6390
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T7"
      Height          =   375
      Index           =   6
      Left            =   120
      Picture         =   "ToolBar.frx":2C744
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cyclicity Table"
      Top             =   6390
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T6"
      Height          =   375
      Index           =   5
      Left            =   960
      Picture         =   "ToolBar.frx":2D048
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Normality Test Table"
      Top             =   6000
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T5"
      Height          =   375
      Index           =   4
      Left            =   540
      Picture         =   "ToolBar.frx":2D94C
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Penalty Points Table"
      Top             =   6000
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T4"
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "ToolBar.frx":2E250
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Occurrence Table"
      Top             =   6000
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T3"
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "ToolBar.frx":2EB54
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Well/Section Names Table"
      Top             =   5610
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T2"
      Height          =   375
      Index           =   1
      Left            =   540
      Picture         =   "ToolBar.frx":2F458
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Warnings Table"
      Top             =   5610
      Width           =   405
   End
   Begin VB.CommandButton CmdTables 
      BackColor       =   &H00C0FFC0&
      Caption         =   "T1"
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ToolBar.frx":2FD5C
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Summary Table"
      Top             =   5610
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CD5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   540
      Picture         =   "ToolBar.frx":30660
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Depth Scaling Dendrogram"
      Top             =   5130
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CD4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      Picture         =   "ToolBar.frx":30F64
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Transformed Depth Differences"
      Top             =   5130
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CD3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   960
      Picture         =   "ToolBar.frx":31868
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Depth Difference Q-Q Plot"
      Top             =   4740
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CD2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   540
      Picture         =   "ToolBar.frx":3216C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Depth Difference Histogram"
      Top             =   4740
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CD1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "ToolBar.frx":32A70
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Depth Difference Optimum Sequence"
      Top             =   4740
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C4"
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "ToolBar.frx":33374
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Correlation_Scaling"
      Top             =   4260
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C3"
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "ToolBar.frx":33C78
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Correlation_Ranking"
      Top             =   3870
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C2"
      Height          =   375
      Index           =   1
      Left            =   540
      Picture         =   "ToolBar.frx":3457C
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Scattergram_Scaling"
      Top             =   3870
      Width           =   405
   End
   Begin VB.CommandButton CmdCASCGraphics 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C1"
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ToolBar.frx":34E80
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Scattergram_Ranking"
      Top             =   3870
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R9"
      Height          =   375
      Index           =   8
      Left            =   960
      Picture         =   "ToolBar.frx":35784
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Event Ranges_Scaling"
      Top             =   3390
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R8"
      Height          =   375
      Index           =   7
      Left            =   540
      Picture         =   "ToolBar.frx":36088
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Event Ranges_Ranking"
      Top             =   3390
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R7"
      Height          =   375
      Index           =   6
      Left            =   120
      Picture         =   "ToolBar.frx":3698C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Variance Analysis_Scaling"
      Top             =   3390
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R6"
      Height          =   375
      Index           =   5
      Left            =   960
      Picture         =   "ToolBar.frx":37290
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Variance Analysis_Ranking"
      Top             =   3000
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R5"
      Height          =   375
      Index           =   4
      Left            =   540
      Picture         =   "ToolBar.frx":37B94
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Scattergram_Scaling"
      Top             =   3000
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R4"
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "ToolBar.frx":38498
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Scattergram_Ranking"
      Top             =   3000
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R3"
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "ToolBar.frx":38D9C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Scaled Optimum Sequence"
      Top             =   2610
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R2"
      Height          =   375
      Index           =   1
      Left            =   540
      Picture         =   "ToolBar.frx":396A0
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ranked Optimum Sequence"
      Top             =   2610
      Width           =   405
   End
   Begin VB.CommandButton CmdRASCGraphics 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R1"
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ToolBar.frx":39FA4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cumulative Frequency"
      Top             =   2610
      Width           =   405
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1830
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdRunCASC 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Run CASC"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Run CASC"
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton CmdRunRASC 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Run RASC"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Run RASC"
      Top             =   990
      Width           =   1245
   End
   Begin VB.CommandButton CmdFile 
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "ToolBar.frx":3A8A8
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Launch QSCreator"
      Top             =   480
      Width           =   405
   End
   Begin VB.CommandButton CmdFile 
      Height          =   375
      Index           =   1
      Left            =   540
      Picture         =   "ToolBar.frx":3B00C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Edit Text Files"
      Top             =   480
      Width           =   405
   End
   Begin VB.CommandButton CmdFile 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ToolBar.frx":3B770
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Set Project Space"
      Top             =   480
      Width           =   405
   End
   Begin VB.CommandButton GraphicControl 
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "ToolBar.frx":3BED4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Reset Graph"
      Top             =   2100
      Width           =   405
   End
   Begin VB.CommandButton GraphicControl 
      Height          =   375
      Index           =   2
      Left            =   930
      Picture         =   "ToolBar.frx":3C4AC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Graphic Panning"
      Top             =   1710
      Width           =   405
   End
   Begin VB.CommandButton GraphicControl 
      Height          =   375
      Index           =   1
      Left            =   540
      Picture         =   "ToolBar.frx":3CA84
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Graphic Scaling"
      Top             =   1710
      Width           =   375
   End
   Begin VB.CommandButton GraphicControl 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "ToolBar.frx":3D05C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Graphic Zooming"
      Top             =   1710
      Width           =   405
   End
   Begin VB.Frame frZoomType 
      Caption         =   "Zoom Type"
      Enabled         =   0   'False
      Height          =   645
      Left            =   1680
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   1800
      Begin VB.OptionButton optGraphicZoom 
         Caption         =   "Graphic"
         Enabled         =   0   'False
         Height          =   330
         Left            =   780
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAxisZoom 
         Caption         =   "Axis "
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   645
      End
   End
End
Attribute VB_Name = "ToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------------------
'*****************************************************************************
'*
'* Copyright (c) 1999, APEX Software Corporation.
'* Portions Copyright (c) 1999, KL GROUP INC.
'* http://www.apexsc.com
'*
'* This file is provided for demonstration and educational uses only.
'* Permission to use, copy, modify and distribute this file for
'* any purpose and without fee is hereby granted, provided that the
'* above copyright notice and this permission notice appear in all
'* copies, and that the name of APEX not be used in advertising
'* or publicity pertaining to this material without the specific,
'* prior written permission of an authorized representative of
'* APEX.
'*
'* APEX MAKES NO REPRESENTATIONS OR WARRANTIES ABOUT THE SUITABILITY
'* OF THE SOFTWARE, EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
'* TO THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'* PURPOSE, OR NON-INFRINGEMENT. APEX SHALL NOT BE LIABLE FOR ANY
'* DAMAGES SUFFERED BY USERS AS A RESULT OF USING, MODIFYING OR
'* DISTRIBUTING THIS SOFTWARE OR ITS DERIVATIVES.
'*
'*****************************************************************************
Option Explicit

'Action Constant
Dim CurrentAction As Integer
'---------------------------------------------------------------------------------------------------------------------------------------------

Dim ToolWindowTop As Single, ToolWindowLeft As Single
Dim deltax As Single, deltay As Single, CursorX As Single
Dim FileTypeIndex As Integer

'For Timer1 setting
Dim MDIWindowCurStatus As Integer


'Codes  for display control

Private Sub EnableZoomType(flag As Boolean)
'Enable/disable the zoom type frame, and all its contents.
    frZoomType.Enabled = flag
    optAxisZoom.Enabled = flag
    optGraphicZoom.Enabled = flag
End Sub


Private Sub ClearEvents()
'Remove all of the currently defined Action Maps from the ActionMap structure.
    If Not (CurGraphicOBJ Is Nothing) Then
        CurGraphicOBJ.ActionMaps.RemoveAll
        'Allow user to edit Chart properties
        CurGraphicOBJ.ActionMaps.Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
    End If
 
End Sub


Private Sub CmdAbout_Click()
'    MDIFrmCascRasc.ActiveTimer.Enabled = False
'    MDIFrmCascRasc.DeActiveTimer.Enabled = False
    FrmFrontPage.Show 1
End Sub

Private Sub CmdCASCGraphics_Click(Index As Integer)
    Select Case Index
        Case 0
            'Scattergram Ranking
            file_type = "DE1"
            Load frmscatterDE1
            frmscatterDE1.SetFocus
            
        Case 1
            'Scattergram Scaling
            file_type = "DE2"
            Load frmscatterDE2
            frmscatterDE2.SetFocus

        Case 2
            'Correlation Ranking
            file_type = "Ca1"
            Load frmwellPlot1
            frmwellPlot1.SetFocus

        Case 3
            'Correlation Scaling
            file_type = "Ca2"
            Load frmwellPlot2
            frmwellPlot2.SetFocus
            
        Case 4
            'Depth Difference Optimum Sequence
            frmFirstOrderDepthDiff.Show
            frmFirstOrderDepthDiff.SetFocus

        Case 5
            'Depth Difference Histogram
            frmDiffHistogram.Show
            frmDiffHistogram.SetFocus

        Case 6
            'Depth Difference Q-Q Plot
            frmDepthDiffQQplot.Show
            frmDepthDiffQQplot.SetFocus
            
        Case 7
            'Transformed Depth Differences
            frmTransDepthDiffQQplot.Show
            frmTransDepthDiffQQplot.SetFocus

        Case 8
            'Depth Scaling Dendrogram
            Load frmDepthDem
            frmDepthDem.SetFocus
            
    End Select

End Sub

Private Sub CmdCopyGraphToClipboard_Click()
     If Not (CurGraphicOBJ Is Nothing) Then
            If CurGraphicOBJ.Visible Then
                CurGraphicOBJ.CopyToClipboard oc2dFormatEnhMetafile
            End If
     End If

End Sub

Private Sub CmdFile_Click(Index As Integer)
    Select Case Index
        Case 0
            'Set Workplace
            dlgsetpath.Show vbModal
            
        Case 1
            'Edit Text Files
             Call EditTextFile

        Case 2
            'Launch QSCreator
              Dim X, StrOutput As String
              Me.MousePointer = vbHourglass
        
              'StrOutput = "notepad.exe " + Filename
              If Dir(App.Path + "\" + "QSCreator.jar") <> "" Then
                  StrOutput = App.Path + "\" + "QSCreator.jar"
              Else
                  MsgBox "In RASCW directory, QSCreator.jar program was not found. Please check it and run again. "
                  Exit Sub
              End If
              X = Shell("javaw -jar " + StrOutput, vbMaximizedFocus)
              
              'AppActivate StrOutput
              
              'Set the mouse cursor back to default
              Me.MousePointer = vbDefault

            
    End Select

  
End Sub

Private Sub cmdReset_Click()
'Reset the view, and the scrollbars.
    If Not (CurGraphicOBJ Is Nothing) Then
        CurGraphicOBJ.CallAction oc2dActionReset, 0, 0
    End If
    
    'ResetScrollbars HScroll.Enabled, VScroll.Enabled
End Sub

Private Sub CmdHelp_Click()
    CommonDialog1.DialogTitle = "RASCW Help File"
    'CommonDialog1.InitDir = App.Path

    CommonDialog1.HelpFile = App.Path + "\" + "RASCW.hlp"
    CommonDialog1.HelpCommand = cdlHelpContents
    CommonDialog1.ShowHelp    ' 显示系统帮助目录主题。
End Sub

Private Sub CmdRASCGraphics_Click(Index As Integer)
    Select Case Index
        Case 0
            'Cumuative Frequency
            Load frmCum
            frmCum.SetFocus
            
        Case 1
            'Ranked Optimum Sequence
            Load frmDenRank
            frmDenRank.SetFocus

        Case 2
            'Scaled Optimum Sequence
            Load frmDen
            frmDen.SetFocus

        Case 3
            'Scatttergram Ranking
            file_type = "Sc1"
            Load frmscatterSC1
            frmscatterSC1.SetFocus
            
        Case 4
            'Scatttergram Scaling
            file_type = "Sc2"
            Load frmscatterSC2
            frmscatterSC2.SetFocus

        Case 5
            'Variance Analysis Ranking
            file_type = "Va1"
            Load frmVar1
            frmVar1.SetFocus

        Case 6
            'Variance Analysis Scaling
            file_type = "Va2"
            Load frmVar2
            frmVar2.SetFocus
            
        Case 7
            'Event Ranges Ranking
            file_type = "Ra1"
           ' Load frmRan2
            Load frmRan1
            frmRan1.SetFocus

        Case 8
            'Event Ranges Scaling
            file_type = "Ra2"
           ' Load frmRan4
            Load frmRan3
            frmRan3.SetFocus
            
    End Select

End Sub

Private Sub CmdRunCASC_Click()
    frmCascinput.Show
     frmCascinput.SetFocus
End Sub

Private Sub CmdRunRASC_Click()
    frmRascW.Show
    frmRascW.SetFocus
End Sub

Private Sub CmdTables_Click(Index As Integer)
    Select Case Index
        Case 0
            'Summary Table
                TableCaption = "Summary Table"
                tableType = 1
                'Load frmSumTable
                If CheckExistWindows(19) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(19))
                Else
                    Opentab
                End If
            
        Case 1
            'Warnings Table
                TableCaption = "Warnings"
                tableType = 2
                'Load frmSumTable
                If CheckExistWindows(20) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(20))
                Else
                    Opentab
                End If

        Case 2
            'Well/Sections Name Table
                TableCaption = "Well/Section Names"
                tableType = 3
                'Load frmSumTable
                If CheckExistWindows(21) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(21))
                Else
                    Opentab
                End If

        Case 3
            'Occurrence Table
                TableCaption = "Occurrence Table"
                tableType = 4
                'Load frmSumTable
                If CheckExistWindows(22) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(22))
                Else
                    Opentab
                End If
            
        Case 4
            'Penalty Points Table
                TableCaption = "Penalty Points"
                tableType = 5
                'Load frmSumTable
                If CheckExistWindows(23) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(23))
                Else
                    Opentab
                End If

        Case 5
            'Normality Test Table
                TableCaption = "Normality Test"
                tableType = 6
                'Load frmSumTable
                If CheckExistWindows(24) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(24))
                Else
                    Opentab
                End If

        Case 6
            'Cyclicity Table
                TableCaption = "Cyclicity"
                tableType = 7
                'Load frmSumTable
                If CheckExistWindows(25) > 0 Then
                    Call CurWindowSetFocus(CheckExistWindows(25))
                Else
                    Opentab
                End If
            
        Case 7
            'Dictionary
                frmDicInput.Show
                frmDicInput.SetFocus

        Case 8
            '

   End Select


End Sub

Private Sub CmdViewSavedGraph_Click()
    Load frmChartShow
    frmChartShow.SetFocus
End Sub

Private Sub CmdWindowArrange_Click(Index As Integer)
    Select Case Index
        Case 0
            'Tile horizontally
                MDIFrmCascRasc.Arrange vbTileHorizontal
            
        Case 1
            'Tile vertically
                MDIFrmCascRasc.Arrange vbTileVertical

        Case 2
            'Cascade
                MDIFrmCascRasc.Arrange vbCascade

        Case 3
            'Arrange Icons
                MDIFrmCascRasc.Arrange vbArrangeIcons
            
        Case 4
            '
             
        Case 5
            '

        Case 6
            '
            
        Case 7
            '
             
        Case 8
            '

    End Select


End Sub

Private Sub Form_Activate()
    ToolBarVisible = 1

'  Dim ToolbarTop As Long, ToolbarLeft As Long
'     'Change the Twips to Pixels Unit
'     If MdiFrmFirstTimeLoad = 1 Then
'         Exit Sub
'    End If
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
''    If MDIfrmActive = 1 Then
''     If GetFocus() = MDIFrmCascRasc.hwnd Then
''         SetWindowPos ToolBar.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''         MDIFrmCascRasc.SetFocus
''    Else
'    If Screen.ActiveForm.hwnd <> MDIFrmCascRasc.hwnd And Screen.ActiveForm.hwnd <> ToolBar.hwnd Then
'          SetWindowPos ToolBar.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'    End If
    'MDIFrmCascRasc.Show
End Sub

Private Sub Form_DblClick()
    Me.Hide
    ToolBarVisible = 0
    MDIFrmCascRasc.mnuToolbar.Checked = False
End Sub

Private Sub Form_GotFocus()
'Dim ToolbarTop As Long, ToolbarLeft As Long
'     'Change the Twips to Pixels Unit
'     If MdiFrmFirstTimeLoad = 1 Then
'         Exit Sub
'    End If
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
'  '   If MDIfrmActive = 1 Then
'     If GetActiveWindow() = MDIFrmCascRasc.hwnd Then
'         SetWindowPos Me.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'         MDIFrmCascRasc.SetFocus
'    Else
'         SetWindowPos Me.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'    End If
'
'
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Watch for the letter r and reset the chart to where we started if we get it.

'    If (KeyAscii = 82 Or KeyAscii = 114) And Not (TypeOf Me.ActiveControl Is TextBox) Then
'        cmdReset_Click
'    End If
End Sub

Private Sub Form_Load()
   ToolWindowLeft = MDIFrmCascRasc.left + MDIFrmCascRasc.Width - 1600
   ToolWindowTop = MDIFrmCascRasc.top + 800
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 100, 560, Flags
   Me.top = ToolWindowTop
   Me.left = ToolWindowLeft
  
  hh = Me.left
  vv = Me.top
  
  MDIWindowCurStatus = 1

   'FileTypeIndex for file type filter index in CommonDialog1
   FileTypeIndex = 0
   DialogInitPath = CurDir

End Sub

Private Sub Form_LostFocus()
'  Dim ToolbarTop As Long, ToolbarLeft As Long
'     'Change the Twips to Pixels Unit
'     If MdiFrmFirstTimeLoad = 1 Then
'         Exit Sub
'    End If
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
''    If MDIfrmActive = 1 Then
''     If GetFocus() = MDIFrmCascRasc.hwnd Then
''         SetWindowPos ToolBar.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''         MDIFrmCascRasc.SetFocus
''    Else
'    If Screen.ActiveForm.hwnd <> MDIFrmCascRasc.hwnd And Screen.ActiveForm.hwnd <> ToolBar.hwnd Then
'          SetWindowPos ToolBar.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'    End If


End Sub

Private Sub Form_Paint()
'Dim ToolbarTop As Long, ToolbarLeft As Long
'     'Change the Twips to Pixels Unit
'     If MdiFrmFirstTimeLoad = 1 Then
'         Exit Sub
'    End If
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
''    If MDIfrmActive = 1 Then
''     If GetFocus() = MDIFrmCascRasc.hwnd Then
''    Else
''         SetWindowPos Me.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''    End If
'    If Screen.ActiveForm.hwnd <> MDIFrmCascRasc.hwnd And Screen.ActiveForm.hwnd <> ToolBar.hwnd Then
'          SetWindowPos ToolBar.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'    Else
'         SetWindowPos Me.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''         MDIFrmCascRasc.SetFocus
'   End If
   

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'End the program.
    MDIFrmCascRasc.mnuToolbar.Checked = False
    ToolBarVisible = 0
    MDIFrmCascRasc.ActiveTimer.Enabled = False
    MDIFrmCascRasc.DeActiveTimer.Enabled = False
    'End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Const LEFT_BUTTON = 1
  Const RIGHT_BUTTON = 2
  Dim leftdown
  leftdown = (Button And LEFT_BUTTON) > 0
  If leftdown Then
    deltax = X: deltay = Y
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Const LEFT_BUTTON = 1
  Const RIGHT_BUTTON = 2
  Dim leftdown, xx As Single, yy As Single
  leftdown = (Button And LEFT_BUTTON) > 0
  If leftdown Then
    xx = X - deltax
    yy = Y - deltay
    Me.left = hh + xx
    Me.top = vv + yy
    hh = Me.left
    vv = Me.top
  End If
End Sub


Private Sub GraphicControl_Click(Index As Integer)
    Select Case Index
        Case 0
            'Graphic Zoom
                If Not (CurGraphicOBJ Is Nothing) Then
                    ClearEvents
                    With CurGraphicOBJ.ActionMaps
                        .Add WM_LBUTTONDOWN, 0, 0, oc2dActionZoomStart
                        .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionZoomUpdate
                        .Add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
                        .Add WM_KEYDOWN, MK_LBUTTON, VK_ESCAPE, oc2dActionZoomCancel
            '            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
                    End With
                    CurrentAction = oc2dActionZoomStart
                End If
            
        Case 1
            'Graphic Scale
                If Not (CurGraphicOBJ Is Nothing) Then
                    ClearEvents
                    With CurGraphicOBJ.ActionMaps
                        .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
                        .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionScale
                        .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
            '            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
                    End With
                    CurrentAction = oc2dActionScale
                End If

        Case 2
            'Graphic Pan
                If Not (CurGraphicOBJ Is Nothing) Then
                    ClearEvents
                    With CurGraphicOBJ.ActionMaps
                        .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
                        .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionTranslate
                        .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
            '            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
                    End With
                    CurrentAction = oc2dActionTranslate
                End If

        Case 3
            'Graphic Reset
                If Not (CurGraphicOBJ Is Nothing) Then
                    ClearEvents
                    CurGraphicOBJ.CallAction oc2dActionReset, 0, 0
                End If
            
    End Select

End Sub


Private Sub optAxisZoom_Click()
'Select the zoom that keeps the axes visible.
   If Not (CurGraphicOBJ Is Nothing) Then
        CurGraphicOBJ.ActionMaps.Remove WM_LBUTTONUP, 0, 0 'oc2dActionZoomEnd
        CurGraphicOBJ.ActionMaps.Add WM_LBUTTONUP, 0, 0, oc2dActionZoomAxisEnd
    End If
End Sub

Private Sub optGraphicZoom_Click()
'Select the zoom that doesn't keep the axes visible.
    If Not (CurGraphicOBJ Is Nothing) Then
         CurGraphicOBJ.ActionMaps.Remove WM_LBUTTONUP, 0, 0  'oc2dActionAxisBound
         CurGraphicOBJ.ActionMaps.Add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
    End If
End Sub

'Private Sub optMove_Click()
''Clear any previous ActionMaps, and construct the new ones.
''Make the rotation constraints inaccessible, and reset the scrollbars.
'
'    If optMove And Not (CurGraphicOBJ Is Nothing) Then
'        ClearEvents
'
'        With CurGraphicOBJ.ActionMaps
'            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
'            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionTranslate
'            .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
''            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
'        End With
'
''        EnableDepthSetting False
''        EnableZoomType False
''        ResetScrollbars True, True
'        CurrentAction = oc2dActionTranslate
'    End If
'End Sub

'Private Sub optNone_Click()
''Clear any previous ActionMaps.
''Make the rotation constraints inaccessible, and reset the scrollbars.
'
'    If optNone And Not (CurGraphicOBJ Is Nothing) Then
'        ClearEvents
''        EnableDepthSetting False
''        EnableZoomType False
''        ResetScrollbars False, False
'        CurrentAction = oc2dActionNone
'        CurGraphicOBJ.ActionMaps.Reset
''        With CurGraphicOBJ.ActionMaps
''            .Add WM_LBUTTONDOWN, MK_CONTROL, 0, oc2dActionZoomStart
'''            .Add WM_LBUTTONUP, MK_CONTROL, 0, oc2dActionZoomEnd
''            .Add WM_LBUTTONDOWN, WM_RBUTTONDOWN, MK_CONTROL, oc2dActionScale
''            .Add WM_LBUTTONDOWN, WM_RBUTTONDOWN, MK_SHIFT, oc2dActionModifyStart  'move-pan
'''            .Add WM_LBUTTONUP, WM_RBUTTONUP, VK_SHIFT, oc2dActionModifyEnd  'move-pan
'''            .Add WM_LBUTTONDOWN, MK_ALT, 0, oc2dActionReset
''            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
''
''        End With
'
'    End If
'End Sub
'
'
'Private Sub optScale_Click()
''Clear any previous ActionMaps, and construct the new ones.
''Make the rotation constraints inaccessible, and reset the scrollbars.
'
'    If optScale And Not (CurGraphicOBJ Is Nothing) Then
'        ClearEvents
'
'        With CurGraphicOBJ.ActionMaps
'            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
'            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionScale
'            .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
''            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
'        End With
'
'        CurrentAction = oc2dActionScale
'    End If
'End Sub
'
'Private Sub optZoom_Click()
''Clear any previous ActionMaps, and construct the new ones.
''Make the rotation constraints inaccessible, and reset the scrollbars.
'
'    If optZoom And Not (CurGraphicOBJ Is Nothing) Then
'        ClearEvents
'
'        With CurGraphicOBJ.ActionMaps
'            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionZoomStart
'            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionZoomUpdate
'            .Add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
'            .Add WM_KEYDOWN, MK_LBUTTON, VK_ESCAPE, oc2dActionZoomCancel
''            .Add WM_RBUTTONDOWN, 0, 0, oc2dActionProperties
'        End With
'
'        CurrentAction = oc2dActionZoomStart
'    End If
'End Sub


'Private Sub EventPrint(msg As Integer, modf As Integer, key As Integer, action As Integer)
''Debugging output for the ActionMaps.  Uncomment in CurGraphicOBJ_Click to view these.
'
'    Dim Output As String
'
'    Select Case msg
'        Case WM_MOUSEMOVE
'            Output = "MouseMove"
'        Case WM_LBUTTONDOWN
'            Output = "LButtonDown"
'        Case WM_LBUTTONUP
'            Output = "LButtonUp"
'        Case WM_LBUTTONDBLCLK
'            Output = "LButtonDbl"
'        Case WM_RBUTTONDOWN
'            Output = "RButtonDown"
'        Case WM_RBUTTONUP
'            Output = "RButtonUp"
'        Case WM_RBUTTONDBLCLK
'            Output = "RButtonDbl"
'        Case WM_MBUTTONDOWN
'            Output = "MButtonDown"
'        Case WM_MBUTTONUP
'            Output = "MButtonUp"
'        Case WM_MBUTTONDBLCLK
'            Output = "MButtonDbl"
'        Case WM_KEYDOWN
'            Output = "KeyDown"
'        Case WM_KEYUP
'            Output = "KeyUp"
'    End Select
'
'    Output = Output & ", "
'    If (modf And MK_ALT) = MK_ALT Then Output = Output & "ALT+"
'    If (modf And MK_MBUTTON) = MK_MBUTTON Then Output = Output & "MBUTTON+"
'    If (modf And MK_CONTROL) = MK_CONTROL Then Output = Output & "CONTROL+"
'    If (modf And MK_SHIFT) = MK_SHIFT Then Output = Output & "SHIFT+"
'    If (modf And MK_RBUTTON) = MK_RBUTTON Then Output = Output & "RBUTTON+"
'    If (modf And MK_LBUTTON) = MK_LBUTTON Then Output = Output & "LBUTTON+"
'
'    Output = Output & ", "
'
'    If key <> 0 Then Output = Output & Chr(key)
'
'    Output = Output & " = "
'
'    Select Case action
'        Case oc2dActionNone
'            Output = Output & "None"
'        Case oc2dActionModifyStart
'            Output = Output & "ModifyStart"
'        Case oc2dActionModifyEnd
'            Output = Output & "ModifyEnd"
'        Case oc2dActionRotate
'            Output = Output & "Rotate"
'        Case oc2dActionScale
'            Output = Output & "Scale"
'        Case oc2dActionTranslate
'            Output = Output & "Translate"
'        Case oc2dActionZoomStart
'            Output = Output & "ZoomStart"
'        Case oc2dActionZoomUpdate
'            Output = Output & "ZoomUpdate"
'        Case oc2dActionZoomEnd
'            Output = Output & "ZoomEnd"
'        Case oc2dActionZoomCancel
'            Output = Output & "ZoomCancel"
'        Case oc2dActionProperties
'            Output = Output & "Properties Page"
'        Case oc2dActionReset
'            Output = Output & "Reset"
'    End Select
'
'   ' Debug.Print Output
'End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------

Private Sub EditTextFile()
Dim X
Dim StrOutput As String
Dim I As Integer
Dim Y As Integer
Dim Z As Integer

Dim FileNames$()          'the array to save selected filenames
' Dim ComeInTime As Integer

'Routine to use the Common Dialog so we can select
' a  file to use.
        
     'Set up all of the required items and then call the OpenFile Common Dialog
    CommonDialog1.DialogTitle = "Load Text Files"
    CommonDialog1.DefaultExt = ".*"
    CommonDialog1.InitDir = DialogInitPath
    CommonDialog1.Filter = "All Files (*.*)|*.*|ALL Files (*.ALL)|*.ALL|CUM Files (*.CUM)|*.CUM|CA Files (*.CA*)|*.CA*|Den/Dem Files (*.DE*)|*.DE*|DF Files (*.DF*)|*.DF*|DI Files (*.DI*)|*.DI*|FL Files (*.FL*)|*.FL*|Out Files (*.out)|*.out|Parameter Files (*.IN*)|*.IN*|PAR Files (*.par)|*.par|RA Files (*.RA*)|*.RA*|SC Files (*.SC*)|*.SC*|Data Files (*.DAT)|*.DAT"
    CommonDialog1.FilterIndex = FileTypeIndex
    CommonDialog1.Filename = ""
    CommonDialog1.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer   '它指定文件名列表框允许多重选择，返回文件'串中各文件名用空格隔开。并且对话框以浏览器的形式出现
    CommonDialog1.Action = 1   'showopen file
    
    CommonDialog1.Filename = CommonDialog1.Filename & Chr(32)
    '从返回的字符串中分离出文件名
     If Trim(CommonDialog1.Filename) = " " Then
'         'Some printer program will change current path, here recover the workplace one time
'         ChDir CurrentDir
'         ChDrive CurrentDrive
         Exit Sub
    End If
    
    Z = 1
    Y = 0
    For I = 1 To Len(CommonDialog1.Filename)
        'InStr函数，返回 Variant (Long)，指定一字符串在另一字符串中最先出现的位置。
        '语法 InStr(起点位置, string1, string2)
       'Debug.Print Asc(Mid(CommonDialog1.Filename, I, 1))
       'Next
        I = InStr(Z, CommonDialog1.Filename, Chr$(0))       '注意上面Flags参数决定返回字符串的形式，若CommonDialog1.Flags = cdlOFNAllowMultiselect，则返回字符串文件之间的分割符全为空格；若CommonDialog1.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer，文件之间的分割符全为chr$(0),末尾为一个空格。
        If I = 0 Then
           I = InStr(Z, CommonDialog1.Filename, Chr$(32))    'last char is chr$(32)
        End If
        ReDim Preserve FileNames(Y)
        'Mid函数，返回 Variant (String)，其中包含字符串中指定数量的字符。
        '语法 Mid(string, start[, length])
        FileNames(Y) = Mid(CommonDialog1.Filename, Z, I - Z)
        'Debug.Print I, Y, FileNames(Y)
        Z = I + 1
        Y = Y + 1
    Next
       
'   Filename = CommonDialog1.Filename
'Save last time FilterIndex
    FileTypeIndex = CommonDialog1.FilterIndex
   If CommonDialog1.Filename <> " " Then
          DialogInitPath = FileNames(0)
   End If
    
'Apply the selections made by the user.
   If Y = 1 Then
              If FileNames(0) <> "" Then
                  'Set the mouse cursor to an HourGlass
                  Me.MousePointer = vbHourglass
        
                  'StrOutput = "notepad.exe " + Filename
                  If Dir(App.Path + "\" + "editpadlite.exe") <> "" Then
                      StrOutput = App.Path + "\" + "editpadlite.exe " + FileNames(0)
                  Else
                      StrOutput = "NotePAD.exe " + FileNames(0)        'Using system editor
                  End If
                  X = Shell(StrOutput, vbMaximizedFocus)
                
                  'Set the mouse cursor back to default
                  Me.MousePointer = vbDefault
              Else
                  'MsgBox "No File Selected!", vbCritical, "Error!"
              End If
    End If
    If Y > 1 Then
           For I = 1 To Y - 1
              If FileNames(I) <> "" Then
                  'Set the mouse cursor to an HourGlass
                  Me.MousePointer = vbHourglass
        
                  'StrOutput = "notepad.exe " + Filename
                  'StrOutput = App.Path + "\" + "editpadlite.exe " + FileNames(0) + "\" + FileNames(I)
                  If Dir(App.Path + "\" + "editpadlite.exe") <> "" Then
                      StrOutput = App.Path + "\" + "editpadlite.exe " + FileNames(0) + "\" + FileNames(I)
                  Else
                      StrOutput = "NotePAD.exe " + FileNames(0) + "\" + FileNames(I)       'Using system editor
                  End If

                  X = Shell(StrOutput, vbMaximizedFocus)
                
                  'Set the mouse cursor back to default
                  Me.MousePointer = vbDefault
              Else
                  'MsgBox "No File Selected!", vbCritical, "Error!"
              End If
          Next
   End If
       
       'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub Timer1_Timer()
'  Dim ToolbarTop As Long, ToolbarLeft As Long
''     'Change the Twips to Pixels Unit
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
''     If MdiFrmFirstTimeLoad = 1 Then
''         Exit Sub
''    End If
'''    If MDIfrmActive = 1 Then
'     If GetActiveWindow() = MDIFrmCascRasc.hwnd And MDIWindowCurStatus = 0 Then
'          SetWindowPos ToolBar.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'          MDIFrmCascRasc.SetFocus
'          MDIWindowCurStatus = 1
'    End If
'    If GetActiveWindow() <> MDIFrmCascRasc.hwnd And MDIWindowCurStatus = 1 Then
'         SetWindowPos Me.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'         MDIWindowCurStatus = 0
'    End If
'
End Sub
