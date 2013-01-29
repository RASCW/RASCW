VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmActions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Olectra Chart 2D - Actions Demo"
   ClientHeight    =   6180
   ClientLeft      =   2190
   ClientTop       =   1740
   ClientWidth     =   6390
   HelpContextID   =   20
   Icon            =   "actions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6180
   ScaleWidth      =   6390
   Begin VB.Frame frDepth 
      Caption         =   "Depth"
      Enabled         =   0   'False
      Height          =   645
      Left            =   4095
      TabIndex        =   12
      Top             =   5460
      Width           =   1275
      Begin VB.CommandButton cmdDDown 
         Caption         =   "¯"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   870
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton cmdDUP 
         Caption         =   "­"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   630
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   210
         Width           =   255
      End
      Begin VB.TextBox txtDepth 
         Enabled         =   0   'False
         Height          =   330
         Left            =   210
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "10"
         Top             =   210
         Width           =   435
      End
   End
   Begin VB.Frame frZoomType 
      Caption         =   "Zoom Type"
      Enabled         =   0   'False
      Height          =   645
      Left            =   2100
      TabIndex        =   9
      Top             =   5460
      Width           =   1800
      Begin VB.OptionButton optGraphicZoom 
         Caption         =   "Graphic"
         Enabled         =   0   'False
         Height          =   330
         Left            =   840
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAxisZoom 
         Caption         =   "Axis "
         Enabled         =   0   'False
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Default         =   -1  'True
      Height          =   540
      Left            =   840
      TabIndex        =   0
      Top             =   5565
      Width           =   1065
   End
   Begin VB.Frame frEvents 
      Caption         =   "Events"
      Height          =   645
      Left            =   735
      TabIndex        =   3
      Top             =   4725
      Width           =   4740
      Begin VB.OptionButton optRotate 
         Caption         =   "Rotate"
         Height          =   330
         Left            =   3885
         TabIndex        =   8
         Top             =   210
         Width           =   790
      End
      Begin VB.OptionButton optZoom 
         Caption         =   "Zoom"
         Height          =   330
         Left            =   2940
         TabIndex        =   7
         Top             =   210
         Width           =   750
      End
      Begin VB.OptionButton optScale 
         Caption         =   "Scale"
         Height          =   330
         Left            =   1995
         TabIndex        =   6
         Top             =   210
         Width           =   750
      End
      Begin VB.OptionButton optMove 
         Caption         =   "Move"
         Height          =   330
         Left            =   1050
         TabIndex        =   5
         Top             =   210
         Width           =   750
      End
      Begin VB.OptionButton optNone 
         Caption         =   "None"
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   4425
      Left            =   6105
      TabIndex        =   2
      Top             =   0
      Width           =   225
   End
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   4425
      Width           =   6105
   End
   Begin OlectraChart2D.Chart2D Chart2D1 
      Height          =   4425
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6105
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   10769
      _ExtentY        =   7805
      _StockProps     =   0
      ControlProperties=   "actions.frx":0442
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   105
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutDemo 
         Caption         =   "About This &Demo"
      End
      Begin VB.Menu mnuSkip 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutOlectra 
         Caption         =   "&About Olectra Chart"
      End
   End
End
Attribute VB_Name = "frmActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'Scrollbar related
Dim HScrollStarted As Boolean
Dim VScrollStarted As Boolean
Dim HScrollValue As Long
Dim VScrollValue As Long

'Action Constant
Dim CurrentAction As Integer

Private Sub EnableDepthSetting(flag As Boolean)
'Enable/disable the depth value frame, and all its contents.

    frDepth.Enabled = flag
    txtDepth.Enabled = flag
    cmdDUP.Enabled = flag
    cmdDDown.Enabled = flag
End Sub

Private Sub EnableZoomType(flag As Boolean)
'Enable/disable the zoom type frame, and all its contents.

    frZoomType.Enabled = flag
    optAxisZoom.Enabled = flag
    optGraphicZoom.Enabled = flag
End Sub

Private Sub ResetScrollbars(hflag As Boolean, vflag As Boolean)
'Reset the scrollbars.  Move the value to the middle, and
' make them enabled/disabled.  Note how the global *value is
' set before the .value - this makes sure the scroll.changed
' event does nothing.

    With HScroll
        'The if...endif here is to prevent the scroll bar from flickering as it is updated
        If hflag = True Then
            HScrollValue = (.Max - .Min) / 2
            .Value = HScrollValue
            .Enabled = hflag
        Else
            .Enabled = hflag
        End If
    End With

    With VScroll
        'The if...endif here is to prevent the scroll bar from flickering as it is updated
        If vflag = True Then
            VScrollValue = (.Max - .Min) / 2
            .Value = VScrollValue
            .Enabled = vflag
        Else
            .Enabled = vflag
        End If
    End With
End Sub

Private Sub ClearEvents()
'Remove all of the currently defined Action Maps from the ActionMap structure.
    
    Chart2D1.ActionMaps.RemoveAll
End Sub

Private Sub EventPrint(msg As Integer, modf As Integer, key As Integer, action As Integer)
'Debugging output for the ActionMaps.  Uncomment in Chart2D1_Click to view these.

    Dim Output As String
    
    Select Case msg
        Case WM_MOUSEMOVE
            Output = "MouseMove"
        Case WM_LBUTTONDOWN
            Output = "LButtonDown"
        Case WM_LBUTTONUP
            Output = "LButtonUp"
        Case WM_LBUTTONDBLCLK
            Output = "LButtonDbl"
        Case WM_RBUTTONDOWN
            Output = "RButtonDown"
        Case WM_RBUTTONUP
            Output = "RButtonUp"
        Case WM_RBUTTONDBLCLK
            Output = "RButtonDbl"
        Case WM_MBUTTONDOWN
            Output = "MButtonDown"
        Case WM_MBUTTONUP
            Output = "MButtonUp"
        Case WM_MBUTTONDBLCLK
            Output = "MButtonDbl"
        Case WM_KEYDOWN
            Output = "KeyDown"
        Case WM_KEYUP
            Output = "KeyUp"
    End Select
    
    Output = Output & ", "
    If (modf And MK_ALT) = MK_ALT Then Output = Output & "ALT+"
    If (modf And MK_MBUTTON) = MK_MBUTTON Then Output = Output & "MBUTTON+"
    If (modf And MK_CONTROL) = MK_CONTROL Then Output = Output & "CONTROL+"
    If (modf And MK_SHIFT) = MK_SHIFT Then Output = Output & "SHIFT+"
    If (modf And MK_RBUTTON) = MK_RBUTTON Then Output = Output & "RBUTTON+"
    If (modf And MK_LBUTTON) = MK_LBUTTON Then Output = Output & "LBUTTON+"
        
    Output = Output & ", "
    
    If key <> 0 Then Output = Output & Chr(key)
    
    Output = Output & " = "
    
    Select Case action
        Case oc2dActionNone
            Output = Output & "None"
        Case oc2dActionModifyStart
            Output = Output & "ModifyStart"
        Case oc2dActionModifyEnd
            Output = Output & "ModifyEnd"
        Case oc2dActionRotate
            Output = Output & "Rotate"
        Case oc2dActionScale
            Output = Output & "Scale"
        Case oc2dActionTranslate
            Output = Output & "Translate"
        Case oc2dActionZoomStart
            Output = Output & "ZoomStart"
        Case oc2dActionZoomUpdate
            Output = Output & "ZoomUpdate"
        Case oc2dActionZoomEnd
            Output = Output & "ZoomEnd"
        Case oc2dActionZoomCancel
            Output = Output & "ZoomCancel"
        Case oc2dActionProperties
            Output = Output & "Properties Page"
        Case oc2dActionReset
            Output = Output & "Reset"
    End Select
    
    Debug.Print Output
End Sub

Private Sub Chart2D1_Click()
'Uncomment these lines to debug the action maps.

    'Dim Event As ActionMap
    
    'For Each Event In Chart2D1.ActionMaps
    'EventPrint Event.Message, Event.Modifier, Event.KeyCode, Event.action
    'Next
End Sub

Private Sub cmdDDown_Click()
'Act like a spinbox control.

    If Val(txtDepth) - 1 >= 0 Then
        txtDepth = Val(txtDepth) - 1
    End If
    
    Chart2D1.ChartArea.View3D.Depth = Val(txtDepth)
End Sub

Private Sub cmdDUP_Click()
'Act like a spinbox control.

    If Val(txtDepth) + 1 <= 100 Then
        txtDepth = Val(txtDepth) + 1
    End If
    
    Chart2D1.ChartArea.View3D.Depth = Val(txtDepth)
End Sub

Private Sub cmdReset_Click()
'Reset the view, and the scrollbars.

    Chart2D1.CallAction oc2dActionReset, 0, 0
    
    ResetScrollbars HScroll.Enabled, VScroll.Enabled
End Sub

Private Sub Form_Activate()
'Set the focus to the command button to force the allowance of
' the <F1> key to bring up help

    cmdReset.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Watch for the letter r and reset the chart to where we started if we get it.

    If (KeyAscii = 82 Or KeyAscii = 114) And Not (TypeOf Me.ActiveControl Is TextBox) Then
        cmdReset_Click
    End If
End Sub

Private Sub Form_Load()
'Start out with no ActionMaps.  Initialize the scrollbars to cover
' the ChartArea, and start out at the center position.
 
    ClearEvents

    'Prevent the user from bringing up the property pages at run-time
    Chart2D1.AllowUserChanges = False
    
    'Start with the form in the top-left corner
    Me.Top = 50
    Me.Left = 50
    
    'Setup the horizontal scroll bar
    With HScroll
        .Min = Chart2D1.ChartArea.Location.Left
        .Max = .Min + Chart2D1.ChartArea.Location.Width - 1
        .Value = (.Max - .Min) / 2
        .SmallChange = 5
        .LargeChange = 30
        HScrollValue = .Value
    End With
     
    'Setup the vertical scroll bar
    With VScroll
        .Min = Chart2D1.ChartArea.Location.Top
        .Max = .Min + Chart2D1.ChartArea.Location.Height - 1
        .Value = (.Max - .Min) / 2
        .SmallChange = 5
        .LargeChange = 30
        VScrollValue = .Value
    End With
    
    'Clear any currently selected action
    CurrentAction = oc2dActionNone
    
    'Batch the updates to the Chart so all the changes will occur at once
    Chart2D1.IsBatched = True
    
    With Chart2D1.ChartGroups(1)
        .ChartType = oc2dTypeBar
        'Reset how many series and points we have
        .Data.NumSeries = 4
        .Data.NumPoints(1) = 2
        .PointLabels.Add "Group 1"
        .PointLabels.Add "Group 2"
    End With
    
    Chart2D1.ChartArea.Axes("x").AnnotationMethod = oc2dAnnotatePointLabels
    
    'Set some defaults for the 3D settings of the bar chart
    Chart2D1.ChartArea.View3D.Depth = Val(txtDepth)
    Chart2D1.ChartArea.View3D.Elevation = Val(txtDepth)
    Chart2D1.ChartArea.View3D.Rotation = Val(txtDepth)
    
    'Change some colors in the Chart
    Chart2D1.Interior.BackgroundColor = RGB(&HFA, &HFA, &HD2)                       'LightGoldenRodYellow
    Chart2D1.Interior.ForegroundColor = RGB(&H0, &H0, &H0)                          'Black
    Chart2D1.ChartArea.PlotArea.Interior.BackgroundColor = RGB(&HD2, &HB4, &H8C)    'Tan
    
    'Resume normal updating of the Chart
    Chart2D1.IsBatched = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End the program.

    End
End Sub

Private Sub HScroll_Change()
'Called when the value of the scrollbar has changed.  This could
' be preceded by a series of HScroll_Scroll() calls, depending on
' how the user is changing the scrollbar.  The Scroll event always
' sets the HScrollStarted flag to True.

    With Chart2D1
        If HScroll <> HScrollValue Then
            If Not HScrollStarted Then
                'Start a scroll from the old value to the current value
                HScrollStarted = True
                .CallAction oc2dActionModifyStart, HScrollValue, VScroll
                .CallAction CurrentAction, HScrollValue, VScroll
                .CallAction CurrentAction, HScroll, VScroll
            End If
            .CallAction oc2dActionModifyEnd, HScroll, VScroll
            
            HScrollStarted = False
            HScrollValue = HScroll      'Save current value
        End If
    End With
End Sub

Private Sub HScroll_Scroll()
'Called when the user drags the scrollbar.  On the first drag,
' start the actions rolling.  HScrollValue is the previous value,
' before the drag started.  Its reset in HScroll_Change.

    With Chart2D1
        If Not HScrollStarted Then
            .CallAction oc2dActionModifyStart, HScrollValue, VScroll
            HScrollStarted = True
        End If
        
        .CallAction CurrentAction, HScroll, VScroll
    End With
End Sub

Private Sub mnuAboutDemo_Click()
'User wants to see what to do in this demo.

    With CommonDialog1
        .HelpCommand = cdlHelpContext
        .HelpContext = 20
        .HelpFile = App.HelpFile
        .ShowHelp
    End With
End Sub

Private Sub mnuAboutOlectra_Click()
'User wants to see what Olectra Chart 2D is all about.

    With CommonDialog1
        .HelpCommand = cdlHelpContext
        .HelpContext = 19
        .HelpFile = App.HelpFile
        .ShowHelp
    End With
End Sub

Private Sub mnuExit_Click()
'End the program.

    Unload Me
End Sub

Private Sub optAxisZoom_Click()
'Select the zoom that keeps the axes visible.
            
    Chart2D1.ActionMaps.Remove WM_LBUTTONUP, 0, 0 'oc2dActionZoomEnd
    Chart2D1.ActionMaps.Add WM_LBUTTONUP, 0, 0, oc2dActionZoomAxisEnd
End Sub

Private Sub optGraphicZoom_Click()
'Select the zoom that doesn't keep the axes visible.

    Chart2D1.ActionMaps.Remove WM_LBUTTONUP, 0, 0  'oc2dActionAxisBound
    Chart2D1.ActionMaps.Add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
End Sub

Private Sub optMove_Click()
'Clear any previous ActionMaps, and construct the new ones.
'Make the rotation constraints inaccessible, and reset the scrollbars.

    If optMove Then
        ClearEvents
        
        With Chart2D1.ActionMaps
            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionTranslate
            .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
        End With
    
        EnableDepthSetting False
        EnableZoomType False
        ResetScrollbars True, True
        CurrentAction = oc2dActionTranslate
    End If
End Sub

Private Sub optNone_Click()
'Clear any previous ActionMaps.
'Make the rotation constraints inaccessible, and reset the scrollbars.

    If optNone Then
        ClearEvents
        EnableDepthSetting False
        EnableZoomType False
        ResetScrollbars False, False
        CurrentAction = oc2dActionNone
    End If
End Sub

Private Sub optRotate_Click()
'Clear any previous ActionMaps, and construct the new ones.
'Make the rotation constraints accessible, and reset the scrollbars.

    If optRotate Then
        ClearEvents
        
        With Chart2D1.ActionMaps
            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionRotate
            .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
        End With
        
        EnableDepthSetting True
        EnableZoomType False
        ResetScrollbars True, True
        CurrentAction = oc2dActionRotate
    End If
End Sub

Private Sub optScale_Click()
'Clear any previous ActionMaps, and construct the new ones.
'Make the rotation constraints inaccessible, and reset the scrollbars.
    
    If optScale Then
        ClearEvents
        
        With Chart2D1.ActionMaps
            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionScale
            .Add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
        End With
        
        EnableDepthSetting False
        EnableZoomType False
        ResetScrollbars False, True         'Vertical scrollbar only
        CurrentAction = oc2dActionScale
    End If
End Sub

Private Sub optZoom_Click()
'Clear any previous ActionMaps, and construct the new ones.
'Make the rotation constraints inaccessible, and reset the scrollbars.

    If optZoom Then
        ClearEvents
        
        With Chart2D1.ActionMaps
            .Add WM_LBUTTONDOWN, 0, 0, oc2dActionZoomStart
            .Add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionZoomUpdate
            .Add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
            .Add WM_KEYDOWN, MK_LBUTTON, VK_ESCAPE, oc2dActionZoomCancel
        End With
        
        EnableDepthSetting False
        EnableZoomType True
        ResetScrollbars False, False        'Scrollbars can't be used
        CurrentAction = oc2dActionZoomStart
    End If
End Sub

Private Sub txtDepth_KeyPress(KeyAscii As Integer)
'Only allow numbers, Enter, Backspace and Tab to be entered.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 9 Then
        KeyAscii = 0
    End If
    
    'If we get an Enter, move away so the chart updates
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdReset.SetFocus
    End If
End Sub

Private Sub txtDepth_LostFocus()
'Allow the user to change the depth of the bars.

    Dim NewVal As Integer
    
    NewVal = Val(txtDepth)
    
    If NewVal > 100 Then
        NewVal = 100
    ElseIf NewVal < 0 Then
        NewVal = 0
    End If
    
    Chart2D1.ChartArea.View3D.Depth = NewVal
    txtDepth = NewVal
End Sub

Private Sub VScroll_Change()
'Called when the value of the scrollbar has changed.  This could
' be preceded by a series of VScroll_Scroll() calls, depending on
' how the user is changing the scrollbar.  The Scroll event always
' sets the VScrollStarted flag to True.

    With Chart2D1
        If VScrollValue <> VScroll Then
            If Not VScrollStarted Then
                'Start a scroll from the old value to the current value
                VScrollStarted = True
                .CallAction oc2dActionModifyStart, HScroll, VScrollValue
                .CallAction CurrentAction, HScroll, VScrollValue
                .CallAction CurrentAction, HScroll, VScroll
            End If
            .CallAction oc2dActionModifyEnd, HScroll, VScroll
            
            VScrollStarted = False
            VScrollValue = VScroll          'Save current value
        End If
    End With
End Sub

Private Sub VScroll_Scroll()
'Called when the user drags the scrollbar.  On the first drag,
' start the actions rolling.  VScrollValue is the previous value,
' before the drag started.  Its reset in VScroll_Change.

    With Chart2D1
        If Not VScrollStarted Then
            .CallAction oc2dActionModifyStart, HScroll, VScrollValue
            VScrollStarted = True
        End If
        
        .CallAction CurrentAction, HScroll, VScroll
    End With
End Sub
