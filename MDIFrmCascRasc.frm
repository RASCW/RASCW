VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIFrmCascRasc 
   BackColor       =   &H8000000C&
   Caption         =   "RASC and CASC"
   ClientHeight    =   6780
   ClientLeft      =   270
   ClientTop       =   885
   ClientWidth     =   10755
   Icon            =   "MDIFrmCascRasc.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer DeActiveTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   600
      Top             =   2460
   End
   Begin VB.Timer ActiveTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   600
      Top             =   1770
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSetProjectSpace 
         Caption         =   "&Set Project Space..."
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputFile 
         Caption         =   "&Edit Text Files..."
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu menuQSCreator 
         Caption         =   "&Run QSCreator..."
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu munExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu numRascw 
      Caption         =   "&Rasc"
      HelpContextID   =   200000001
   End
   Begin VB.Menu mnuCascw 
      Caption         =   "&Casc"
      HelpContextID   =   300000001
   End
   Begin VB.Menu munGraphs 
      Caption         =   "&Graphs"
      Begin VB.Menu mnuRasc 
         Caption         =   "&RASC Graphs"
         Begin VB.Menu Muncum_freq 
            Caption         =   "&Cumulative Frequency"
         End
         Begin VB.Menu mnuDenRank 
            Caption         =   "&Ranked Optimum Sequence"
         End
         Begin VB.Menu MunDendrogram 
            Caption         =   "&Scaled Optimum Sequence  "
         End
         Begin VB.Menu MunScatter_Rank 
            Caption         =   "&Scattergram_Ranking"
         End
         Begin VB.Menu MunScatter_Scaling 
            Caption         =   "&Scattergram_Scaling"
         End
         Begin VB.Menu MunVar_Ana_Ranking 
            Caption         =   "&Variance Analysis_Ranking"
         End
         Begin VB.Menu MunVar_Ana_Scaling 
            Caption         =   "&Variance Analysis_Scaling"
         End
         Begin VB.Menu munEvent_Ranges_Ranking 
            Caption         =   "&Event Ranges_Ranking"
         End
         Begin VB.Menu munEvent_Ranges_Scaling 
            Caption         =   "&Event Ranges_Scaling"
         End
      End
      Begin VB.Menu mnudash212 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCasc 
         Caption         =   "&CASC Graphs"
         Begin VB.Menu mnuSC_DE_Ranking 
            Caption         =   "&Scattergram_Ranking"
         End
         Begin VB.Menu mnuSC_DE_Scaling 
            Caption         =   "&Scattergram_Scaling"
         End
         Begin VB.Menu munCasc_Ranking 
            Caption         =   "&Correlation_Ranking"
         End
         Begin VB.Menu munCasc_Scaling 
            Caption         =   "&Correlation_Scaling"
         End
         Begin VB.Menu MenuDash213 
            Caption         =   "-"
         End
         Begin VB.Menu MenuDepthFrequencyDistribution 
            Caption         =   "&Depth Difference Optimum Sequence"
         End
         Begin VB.Menu MenuDiffHistogram 
            Caption         =   "&Depth Difference Histogram"
         End
         Begin VB.Menu MenuDepthDiffQQPlot 
            Caption         =   "&Depth Difference Q-Q Plot "
         End
         Begin VB.Menu MenuTransQQPlot 
            Caption         =   "&Transformed Depth Differences"
         End
         Begin VB.Menu MenuDepthDendrogram 
            Caption         =   "&Depth Scaling Dendrogram"
         End
      End
      Begin VB.Menu mnudash214 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewGraphs 
         Caption         =   "&Display Saved Graphs ..."
      End
   End
   Begin VB.Menu munTab 
      Caption         =   "&Tables"
      Begin VB.Menu mnuSumTable 
         Caption         =   "&Summary Table"
      End
      Begin VB.Menu mnuwarnings 
         Caption         =   "&Warnings"
      End
      Begin VB.Menu mnuWellNames 
         Caption         =   "&Well/Section Names"
      End
      Begin VB.Menu mnuOccuenceTable 
         Caption         =   "&Occurrence Table"
      End
      Begin VB.Menu mnuPenaltypoints 
         Caption         =   "&Penalty Points"
      End
      Begin VB.Menu mnuNormalityTest 
         Caption         =   "&Normality Test"
      End
      Begin VB.Menu mnuCyclicity 
         Caption         =   "&Cyclicity"
      End
   End
   Begin VB.Menu mnuDictionary 
      Caption         =   "&Dictionary"
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuTile 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuAggrange 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuDash71 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCurWindows 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu numHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContent 
         Caption         =   "&Content..."
      End
      Begin VB.Menu menuaboutrasc 
         Caption         =   "&About RASC..."
      End
   End
End
Attribute VB_Name = "MDIFrmCascRasc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim aWell_Name_Temp() As String
Dim Filename As String
Dim FileTypeIndex As Integer

'Private Sub ActiveTimer_Timer()
'    Dim k As Integer
'    Dim ToolbarTop As Long, ToolbarLeft As Long
'    Dim WinTop As Long, WinLeft As Long, WinWidth As Long, WinLong As Long
'
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
'
'         'if current active windows is of project itself, then continue to show toolbar
'         For k = 0 To Forms.count - 1
'            If GetActiveWindow() = Forms(k).hwnd Then
'                'SetActiveWindow (MDIFrmCascRasc.hwnd)
'                ActiveTimer.Enabled = False
'                DeActiveTimer.Enabled = True
'                If MDIFrmCascRasc.mnuToolbar.Checked And MdiFrmFirstTimeLoad <> 1 Then
'                    'ShowWindow ToolBar.hwnd, SW_SHOW
'                    SetWindowPos ToolBar.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
'                    If Forms(k).hwnd <> MDIFrmCascRasc.hwnd Then
'                                WinTop = MDIFrmCascRasc.top / Screen.TwipsPerPixelY
'                                WinLeft = MDIFrmCascRasc.left / Screen.TwipsPerPixelX
'                                WinWidth = MDIFrmCascRasc.Width / Screen.TwipsPerPixelX
'                                WinLong = MDIFrmCascRasc.Height / Screen.TwipsPerPixelY
'                                SetWindowPos MDIFrmCascRasc.hwnd, HWND_TOP, WinLeft, WinTop, WinWidth, WinLong, Flags
'                        'ShowWindow MDIFrmCascRasc.hwnd, SW_SHOWMAXIMIZED
'''                        SetActiveWindow (MDIFrmCascRasc.hwnd)
'                    End If
''                   If Forms(k).hwnd <> ToolBar.hwnd Then
'''                                WinTop = Forms(k).top / Screen.TwipsPerPixelY
'''                                WinLeft = Forms(k).left / Screen.TwipsPerPixelX
'''                                WinWidth = Forms(k).Width / Screen.TwipsPerPixelX
'''                                WinLong = Forms(k).Height / Screen.TwipsPerPixelY
'''                                SetWindowPos Forms(k).hwnd, HWND_TOP, WinLeft, WinTop, WinWidth, WinLong, SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOSIZE
''                        'ShowWindow Forms(k).hwnd, SW_SHOW
''                       SetActiveWindow (Forms(k).hwnd)
''                    End If
'                End If
''                SetActiveWindow (Forms(k).hwnd)
'                Exit Sub
'            End If
'        Next k
'
'End Sub
'
'Private Sub DeActiveTimer_Timer()
'    Dim k As Integer, ActiveWinHwnd As Long
'    Dim ToolbarTop As Long, ToolbarLeft As Long
'    Dim WinTop As Long, WinLeft As Long, WinWidth As Long, WinLong As Long
'
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
'     'Note: "ToolBar.hwnd" will let Toolbar form show
''     If GetActiveWindow() <> MDIFrmCascRasc.hwnd And GetActiveWindow() <> ToolBar.hwnd Then
'         If MdiFrmFirstTimeLoad <> 1 Then
'            ActiveWinHwnd = GetActiveWindow()
'             'if current active windows is of project itself, then do nothing and exit sub
'             For k = 0 To Forms.count - 1
'                If ActiveWinHwnd = Forms(k).hwnd Then
'                    Exit Sub
'                End If
'             Next k
'             'If ToolBarVisible = 1 Then
'                 'ToolBar.Hide
'
'
'             'ShowWindow ToolBar.hwnd, SW_HIDE
'             SetWindowPos ToolBar.hwnd, HWND_BOTTOM, ToolbarLeft, ToolbarTop, 100, 560, SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOSIZE
'             'ShowWindow GetActiveWindow(), SW_SHOWNA
'
'              DeActiveTimer.Enabled = False
'             ActiveTimer.Enabled = True
'            '    ToolBarVisible = 0
'             'End If
'         End If
''     End If
'End Sub

Private Sub MDIForm_Activate()
'  Dim ToolbarTop As Long, ToolbarLeft As Long
'     'Change the Twips to Pixels Unit
'     If MdiFrmFirstTimeLoad = 1 Then
'         Exit Sub
'    End If
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
''    If MDIfrmActive = 1 Then
''     If GetFocus() = MDIFrmCascRasc.hwnd Then
'     SetWindowPos ToolBar.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''         MDIFrmCascRasc.SetFocus
''    Else
''         SetWindowPos Me.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''    End If


End Sub


Private Sub MDIForm_Click()
'  Dim ToolbarTop As Long, ToolbarLeft As Long
'     'Change the Twips to Pixels Unit
'     If MdiFrmFirstTimeLoad = 1 Then
'         Exit Sub
'    End If
'     ToolbarTop = vv / Screen.TwipsPerPixelY
'     ToolbarLeft = hh / Screen.TwipsPerPixelX
''    If MDIfrmActive = 1 Then
''     If GetFocus() = MDIFrmCascRasc.hwnd Then
'     SetWindowPos ToolBar.hwnd, HWND_TOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''         MDIFrmCascRasc.SetFocus
''    Else
''         SetWindowPos Me.hwnd, HWND_NOTOPMOST, ToolbarLeft, ToolbarTop, 100, 560, Flags
''    End If
End Sub

Private Sub MDIForm_Deactivate()
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

Private Sub MDIForm_Initialize()
    'For the toolbar windows display control
    MdiFrmFirstTimeLoad = 1
    MDIfrmActive = 1
    ToolBarVisible = 0
End Sub

Private Sub MDIForm_Load()
    Dim drv As String, I As Integer
    
    ChDir App.Path
    drv = left(App.Path, 2)
    ChDrive (drv)
    CurrentDir = App.Path
    CurrentDrive = drv
    Me.Show
    FrmFrontPage.Show 1
    
    'RASCW Help File
    App.HelpFile = App.Path + "\RASCW.HLP"

    
    'FileTypeIndex for file type filter index in CommonDialog1
    FileTypeIndex = 0
    DialogInitPath = CurDir
    ChartFileTypeIndex = 1
    ChartSaveInitPath = CurDir
    
    ChartOpenInitPath = CurDir
    
    'For rascw.exe and rascwO.exe running selction control  0-- rascw.exe; 1-- rascwO.exe
    RASCWVersionControlKey = 0
    
    'For the toolbar windows display control
    'MdiFrmFirstTimeLoad = 1
    'MDIBackground.Height = Me.Height
    'MDIBackground.Width = Me.Width
    
    'Record all possible MDIWindows name
    WindowsName(1) = "Cumulative Frequency"
    WindowsName(2) = "Ranked Optimum Sequence"
    WindowsName(3) = "Scaled Optimum Sequence"
    WindowsName(4) = "RASC Scattergram - Ranking"
    WindowsName(5) = "RASC Scattergram - Scaling"
    WindowsName(6) = "Variance Analysis - Ranking"
    WindowsName(7) = "Variance Analysis - Scaling"
    WindowsName(8) = "Event Ranges - Ranking"
    WindowsName(9) = "Event Ranges - Scaling"
    WindowsName(10) = "CASC Scattergram - Ranking"
    WindowsName(11) = "CASC Scattergram - Scaling"
    WindowsName(12) = "Correlation - Ranking"
    WindowsName(13) = "Correlation - Scaling"
    WindowsName(14) = "Depth Difference Optimum Sequence"
    WindowsName(15) = "Depth Difference Histogram"
    WindowsName(16) = "Depth Difference Q-Q Plot"
    WindowsName(17) = "Transformed Depth Differences"
    WindowsName(18) = "Depth-Scaled Dendrogram"
    WindowsName(19) = "Summary Table"
    WindowsName(20) = "Warning  Table"
    WindowsName(21) = "Well/Section Name Table"
    WindowsName(22) = "Occurrence Table"
    WindowsName(23) = "Penalty Point Table"
    WindowsName(24) = "Normality Test Table"
    WindowsName(25) = "Cyclicity Table"
    WindowsName(26) = "Dictionary"
    WindowsName(27) = "Run RASC"
    WindowsName(28) = "Run CASC"
    WindowsName(29) = "Display Graph"
    WindowsName(30) = ""
    For I = 1 To 30
           WindowsHwnd(I) = -1
    Next I
    
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      MsgBox "I am here"
End Sub

Private Sub MDIForm_Resize()
    'MDIBackground.Height = Me.Height
    'MDIBackground.Width = Me.Width
   ' 如果父窗体被最小化...
   If Me.WindowState = vbMinimized Then
   ' ...隐藏 Form
      ToolBar.Visible = False
   ' 如果父窗体不再是最小化...
   Else
      'For the first time of Toolbar windows control
'      If MdiFrmFirstTimeLoad = 1 Then
'          MdiFrmFirstTimeLoad = 0
'          Exit Sub
'      End If
      ' ...恢复 Form
      If Me.mnuToolbar.Checked And MdiFrmFirstTimeLoad = 0 Then
          ToolBar.Visible = True
          'MDIFrmCascRasc.SetFocus
      End If
   End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload ToolBar
End Sub

Private Sub menuaboutrasc_Click()
    MDIFrmCascRasc.ActiveTimer.Enabled = False
    MDIFrmCascRasc.DeActiveTimer.Enabled = False
    FrmFrontPage.Show 1
End Sub

Private Sub menuCor_Click()
    frmCor.Show
End Sub

Private Sub MenuDepthDendrogram_Click()
    
    Load frmDepthDem
    frmDepthDem.SetFocus
End Sub

Private Sub MenuDepthDiffQQPlot_Click()
    frmDepthDiffQQplot.Show
    frmDepthDiffQQplot.SetFocus
End Sub

Private Sub MenuDepthFrequencyDistribution_Click()
     frmFirstOrderDepthDiff.Show
     frmFirstOrderDepthDiff.SetFocus
End Sub

Private Sub MenuDiffHistogram_Click()
     frmDiffHistogram.Show
      frmDiffHistogram.SetFocus
End Sub

Private Sub menuOutputFile_Click()
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
    CommonDialog1.DialogTitle = "Select Text File(s) to Edit ..."
    CommonDialog1.DefaultExt = ".*"
    CommonDialog1.InitDir = DialogInitPath
    CommonDialog1.Filter = "All Files (*.*)|*.*|ALL Files (*.ALL)|*.ALL|CUM Files (*.CUM)|*.CUM|CA Files (*.CA*)|*.CA*|Den/Dem Files (*.DE*)|*.DE*|DF Files (*.DF*)|*.DF*|DI Files (*.DI*)|*.DI*|FL Files (*.FL*)|*.FL*|Out Files (*.out)|*.out|Parameter Files (*.IN*)|*.IN*|PAR Files (*.par)|*.par|RA Files (*.RA*)|*.RA*|SC Files (*.SC*)|*.SC*|Data Files (*.DAT)|*.DAT"
    CommonDialog1.FilterIndex = FileTypeIndex
    CommonDialog1.Filename = ""
    CommonDialog1.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer   '它指定文件名列表框允许多重选择，返回文件'串中各文件名用空格隔开。并且对话框以浏览器的形式出现
    CommonDialog1.Action = 1   'showopen file
    
    CommonDialog1.Filename = CommonDialog1.Filename & Chr(32)
    '从返回的字符串中分离出文件名
     If Trim(CommonDialog1.Filename) = "" Then
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
   
'Dim X

'X = Shell("notepad.exe", vbNormalFocus)

'frmOpenOutputFiles.Show 1

    'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

Private Sub menuQSCreator_Click()
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
End Sub

Private Sub MenuTransQQPlot_Click()
     frmTransDepthDiffQQplot.Show
     frmTransDepthDiffQQplot.SetFocus
End Sub

Private Sub mnuCascw_Click()
    frmCascinput.Show
    frmCascinput.SetFocus
End Sub


Private Sub munCum_Click()

End Sub


Private Sub mnuClose_Click()
     myDBName = ""
End Sub

Private Sub mnuContent_Click()
    CommonDialog1.DialogTitle = "RASCW Help File"
    'CommonDialog1.InitDir = App.Path
    
    CommonDialog1.HelpFile = App.Path + "\" + "RASCW.hlp"
    CommonDialog1.HelpCommand = cdlHelpContents
    CommonDialog1.ShowHelp    ' 显示系统帮助目录主题。
  
End Sub

Private Sub mnuCyclicity_Click()
    TableCaption = "Cyclicity"
    tableType = 7
    'Load frmSumTable
    If CheckExistWindows(25) > 0 Then
        Call CurWindowSetFocus(CheckExistWindows(25))
    Else
        Opentab
    End If
End Sub

Private Sub mnuDenRank_Click()
    Load frmDenRank
    frmDenRank.SetFocus
End Sub

Private Sub mnuDictionary_Click()
    frmDicInput.Show
    frmDicInput.SetFocus
End Sub

Private Sub mnuHeader_Click()
Dim TabIndex As Integer
Dim TabCount As Integer
  
'CmDlg1.ShowOpen
'CmDlg1.Filter = "(*.mdb)|*.MDB"

'myDBName = CmDlg1.Filename

If myDBName <> "" Then
   frmMakedatheader.Show
Else
   MsgBox "Open a database from file menu", , ""
End If
End Sub

Private Sub mnuinput_Rasc_Click()
' input for RASC from Makedat
Dim newfrmWell() As New frmMakedatWell
Dim TabIndex As Integer
Dim TabCount As Integer
Dim I, J, k As Integer
Dim gFileNum As Integer
Dim dbsDic As Database
Dim rstEvent_Name As Recordset
Dim rstWellTable, rstwellHeader As Recordset
Dim NumFossilRecord As Integer
Dim TempFossilName As String
Dim tempSpace As String
Dim TempDepth, tempRTH As Double
Dim tempNumber As Integer
Dim InputFileDic As String
Dim TempLen, tempId, tempNameLen As Integer
Dim tempUnit As String
Dim TempIndex As Integer
Dim TempYN As Boolean
'Dim TempLength As Integer


 frmOpenTable_RC.Show 1
'CmDlg1.Filter = "(*.dat)|*.DAT"
'CmDlg1.ShowSave
'RascInputFile = CmDlg1.Filename
If RascInputFile = "" Then Exit Sub
     
    ' the new database.
   
If Len(Dir(RascInputFile)) <> 0 Then
  response = MsgBox("Over write the existing file?", vbYesNoCancel)
  If response = vbYes Then
  Kill RascInputFile
   Else
   Exit Sub
   End If
End If

If myDBName <> "" Then
frmSelect.Caption = "Select wells to prepare input file for runing RASC"
frmSelect.Show 1
Else
MsgBox "Open a database from file menu", , ""
Exit Sub
End If

If Num_Edit_Well < 1 Then Exit Sub

frmSelectevent.Show 1

'**********************
'write out dictionary file to the input for runing RASC

'Fill gCascInput with the currently displayed data
gFileNum = FreeFile
InputFileDic = RascInputFile
TempLen = Len(RascInputFile)
Mid(InputFileDic, TempLen - 2) = "dic"
'MsgBox InputFileDic
Open CurDir + "\" + InputFileDic For Output As gFileNum

'*********** get dictinary information from event_name
 Set dbsDic = OpenDatabase(myDBName)
    With dbsDic
        ' Open table-type Recordset and show RecordCount
        ' property.
     Set rstEvent_Name = .OpenRecordset("event_name")
    
       If rstEvent_Name.EOF = False Then
          rstEvent_Name.MoveLast
          NumFossilRecord = rstEvent_Name.RecordCount
        Else
         MsgBox "Dictionary file is Empty!"
         Num_Edit_Well = 0
       Exit Sub
       End If
    End With
  
    If NumFossilRecord > 0 Then
       rstEvent_Name.MoveFirst
       For I = 0 To NumFossilRecord - 1
        If IsNull(rstEvent_Name.Fields(0).Value) = True And IsNull(rstEvent_Name.Fields(2).Value) = True Then
           TempFossilName = "dummy"
         ElseIf IsNull(rstEvent_Name.Fields(0).Value) = True Then
           TempFossilName = rstEvent_Name.Fields(2).Value
         ElseIf IsNull(rstEvent_Name.Fields(2).Value) = True Then
            TempFossilName = rstEvent_Name.Fields(0).Value
         Else
            TempFossilName = rstEvent_Name.Fields(0).Value + " " + rstEvent_Name.Fields(2).Value
         End If
         
         'add number after each name
         tempSpace = ""
         tempSpace = Space(40 - Len(Trim(TempFossilName)))
         TempFossilName = TempFossilName + tempSpace + Str(I + 1)
         
         Print #gFileNum, TempFossilName
       If rstEvent_Name.EOF = False Then
          rstEvent_Name.MoveNext
          End If
       Next I
    End If
  dbsDic.Close
  Print #gFileNum, "LAST"
Close gFileNum

'******************************************************

ReDim newfrmWell(1 To Num_Edit_Well)
ReDim aWell_Name_Temp(1 To Num_Edit_Well)
gFileNum = FreeFile
Open CurDir + "\" + RascInputFile For Output As gFileNum

Set DBsSAGA = OpenDatabase(myDBName)
 TabCount = DBsSAGA.TableDefs.count
   For J = 1 To Num_Edit_Well
      For I = 0 To TabCount - 1
   
       If DBsSAGA.TableDefs(I).name = "well" + Trim(Str(Edit_Well_Num(J))) + "_inf" Then
        Well_Name_Temp = DBsSAGA.TableDefs(I).name
        aWell_Name_Temp(J) = Trim(Well_Name_Temp)
 '       MsgBox aWell_Name_Temp(J)
        'call function for writing out the values from table with table name well_Name_temp
         With DBsSAGA
           Set rstWellTable = .OpenRecordset(Well_Name_Temp)
           If rstWellTable.EOF = False Then
             rstWellTable.MoveLast
             NumFossilRecord = rstWellTable.RecordCount
             Else
              MsgBox "Well " + Edit_Well_Names(J) + " is Empty!"
              Num_Edit_Well = 0
             NumFossilRecord = 0
'             Exit Sub
            End If
         End With
  
    If NumFossilRecord > 0 Then
       ' write out the name of well and RT hight first
           Set rstwellHeader = DBsSAGA.OpenRecordset("header")
             tempNameLen = Len(Trim(Well_Name_Temp))
             If tempNameLen > 8 Then
             tempId = Int(Mid(Trim(Well_Name_Temp), 5, tempNameLen - 8))
             Else
             tempId = -1
             End If
               
           If tempId > 0 Then
             rstwellHeader.Move (tempId - 1)
               If IsNull(rstwellHeader.Fields(4).Value) = True Then
               tempUnit = " "
               Else
                tempUnit = Trim(rstwellHeader.Fields(4).Value)
               End If
               If IsNull(rstwellHeader.Fields(3).Value) = True Then
               tempRTH = -1
               Else
               tempRTH = rstwellHeader.Fields(3).Value
               End If
             Print #gFileNum, Trim(Edit_Well_Names(J))
             Print #gFileNum, Format(tempUnit, "#"); Format(tempRTH, "00.00")
           End If
           
       rstWellTable.MoveFirst
       For k = 0 To NumFossilRecord - 1
        If IsNull(rstWellTable.Fields(0).Value) = True Then
        TempDepth = "  "
        Else
        TempDepth = rstWellTable.Fields(0).Value
        End If
        If IsNull(rstWellTable.Fields(2).Value) = True Then
           tempNumber = " "
        Else
           tempNumber = rstWellTable.Fields(2).Value
        End If
    'add condition to delete some events from the list
        TempYN = False
        For TempIndex = 1 To Del_Event_Total
           If tempNumber <> Del_Event_Numbers(TempIndex) Then
           TempYN = False
           Else
           TempYN = True
           Exit For
           End If
        Next TempIndex
        If TempYN = False Then
        Print #gFileNum, Format(TempDepth, "00000.00"); Format(tempNumber, "0000")
        End If
            
          If rstWellTable.EOF = False Then
          rstWellTable.MoveNext
          End If
        Next k
     Print #gFileNum, "  "
     End If
            
        ' Well_Caption = Trim(Edit_Well_Names(J))
        ' newfrmWell(J).Caption = Edit_Well_Names(J)
         
        ' newfrmWell(J).Show
       End If
    Next I
   Next J
 DBsSAGA.Close
Close gFileNum
 MsgBox "Proceed with RASC menu"
End Sub

Private Sub mnuIPS_Click()
Dim IPSfile As String
Dim X
Dim max_Event As Integer
'old makedat file is ips file

frmOpenIPS.Show 1
If Old_Makedat_File = "" Or New_Makedat_File = "" Then
MsgBox "Select an Input File"
Exit Sub
End If

If Len(Dir(Old_Makedat_File)) = 0 Then
   MsgBox "The file name  " + Old_Makedat_File + " does not exist"
   Exit Sub
 End If
X = Count_Well_IPS(Old_Makedat_File)
myDBName = New_Makedat_File

CreateNewMakeDat (X)

Add_wellheader_IPS (Old_Makedat_File)
max_Event = Int(Add_welldata_IPS(Old_Makedat_File))
X = Add_Dic_IPS(Old_Makedat_File, max_Event)
MsgBox "Database " + myDBName + " has been created, please use Makedat", , "Makedat, Rasc, Casc and Cor"
End Sub


Private Sub mnuMakedatFile_Click()
Dim LSTfile As String
Dim X
frmOpenMake.Show 1
If Old_Makedat_File = "" Or New_Makedat_File = "" Then
MsgBox "Select an Input File"
Exit Sub
End If

If Len(Dir(Old_Makedat_File)) = 0 Then
   MsgBox "The file name  " + Old_Makedat_File + " does not exist"
   Exit Sub
 End If
X = Count_Well(Old_Makedat_File)
myDBName = New_Makedat_File

CreateNewMakeDat (X)
Add_Dic (Old_Makedat_File)
Add_wellheader (Old_Makedat_File)
Add_welldata (Old_Makedat_File)
MsgBox "Database " + myDBName + " has been created, please use Makedat", , "Makedat, Rasc, Casc and Cor"
End Sub

Private Sub mnuMakeDictionary_Click()
'Dim newfrmEvent As New frmMakedatEvent
Dim response
'CmDlg1.ShowOpen
'myDBName = CmDlg1.Filename
'Load frmKamedatEvent
'MsgBox "Frits, This is test#1 "
If myDBName <> "" Then
Load frmMakedatEvent
Else
MsgBox "Open a database from file menu", , " "
End If
End Sub

Private Sub mnuRasccasc_Click()

End Sub

Private Sub mnuNormalityTest_Click()
TableCaption = "Normality Test"
tableType = 6
'Load frmSumTable
    If CheckExistWindows(24) > 0 Then
        Call CurWindowSetFocus(CheckExistWindows(24))
    Else
        Opentab
    End If
End Sub

Private Sub mnuOccuenceTable_Click()
TableCaption = "Occurrence Table"
tableType = 4
'Load frmSumTable
    If CheckExistWindows(22) > 0 Then
        Call CurWindowSetFocus(CheckExistWindows(22))
    Else
        Opentab
    End If
End Sub

Private Sub mnuPenaltypoints_Click()
TableCaption = "Penalty Points"
tableType = 5
'Load frmSumTable
    If CheckExistWindows(23) > 0 Then
        Call CurWindowSetFocus(CheckExistWindows(23))
    Else
        Opentab
    End If
End Sub

Private Sub mnuSC_DE_Ranking_Click()
file_type = "DE1"
Load frmscatterDE1
frmscatterDE1.SetFocus
End Sub

Private Sub mnuSC_DE_Scaling_Click()
file_type = "DE2"
Load frmscatterDE2
frmscatterDE2.SetFocus
End Sub

Private Sub mnuSelFile_Click()
  frmDialog.Show vbModal
End Sub

Private Sub mnuSetProjectSpace_Click()
  dlgsetpath.Show vbModal
End Sub

Private Sub mnuSumTable_Click()
TableCaption = "Summary Table"
tableType = 1
'Load frmSumTable
If CheckExistWindows(19) > 0 Then
    Call CurWindowSetFocus(CheckExistWindows(19))
Else
    Opentab
End If
End Sub

Private Sub mnuSupport_Click()
'frmpage2.Show 1
End Sub


Private Sub mnuTile_Click()
     MDIFrmCascRasc.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertical_Click()
    MDIFrmCascRasc.Arrange vbTileVertical
End Sub

Private Sub mnuToolbar_Click()
   If mnuToolbar.Checked = False Then
       ToolBar.Show
       mnuToolbar.Checked = True
       MDIFrmCascRasc.ActiveTimer.Enabled = True
       MDIFrmCascRasc.DeActiveTimer.Enabled = True
       Me.SetFocus
   Else
      ToolBar.Hide
      mnuToolbar.Checked = False
      MDIFrmCascRasc.ActiveTimer.Enabled = False
       MDIFrmCascRasc.DeActiveTimer.Enabled = False
    End If
End Sub

Private Sub mnuViewGraphs_Click()
    Load frmChartShow
    frmChartShow.SetFocus
End Sub

Private Sub mnuwarnings_Click()
    TableCaption = "Warnings"
    tableType = 2
    'Load frmSumTable
    If CheckExistWindows(20) > 0 Then
        Call CurWindowSetFocus(CheckExistWindows(20))
    Else
        Opentab
    End If
End Sub

Private Sub mnuWellNames_Click()
    TableCaption = "Well/Section Names"
    tableType = 3
    'Load frmSumTable
    If CheckExistWindows(21) > 0 Then
        Call CurWindowSetFocus(CheckExistWindows(21))
    Else
        Opentab
    End If
End Sub

Private Sub mnuCurWindows_Click(Index As Integer)
     Call CurWindowSetFocus(Index)
End Sub

Private Sub munCasc_Ranking_Click()
    file_type = "Ca1"
    Load frmwellPlot1
    frmwellPlot1.SetFocus
End Sub

Private Sub munCasc_Scaling_Click()
    file_type = "Ca2"
    Load frmwellPlot2
    frmwellPlot2.SetFocus
End Sub

Private Sub Muncum_freq_Click()
    Load frmCum
    frmCum.SetFocus
End Sub

Private Sub MunDendrogram_Click()
    Load frmDen
    frmDen.SetFocus
End Sub

Private Sub munDictonary_Click()
'If txt_file_dic = "" Then
 
' Else
'    frmDic.Show
'End If
End Sub

Private Sub munEvent_Ranges_Ranking_Click()
      file_type = "Ra1"
'      frmRan2.Show
      frmRan1.Show
      frmRan1.SetFocus
End Sub

Private Sub munEvent_Ranges_Scaling_Click()
        file_type = "Ra2"
'        Load frmRan4
        Load frmRan3
        frmRan3.SetFocus
End Sub

Private Sub MunExit_Click()
    End
End Sub

Private Sub munGraphs_Click()
'RunGraph
End Sub

Private Sub munMake_Click()
Dim X, Y

X = Shell("newmake.exe", vbNormalFocus)

End Sub

Private Sub munOpen_Click()

End Sub

Private Sub munPrint_Click()

ActiveForm.PrintForm

End Sub

Private Sub MunScatter_Rank_Click()
file_type = "Sc1"
Load frmscatterSC1
frmscatterSC1.SetFocus
End Sub

Private Sub MunScatter_Scaling_Click()
file_type = "Sc2"
Load frmscatterSC2
frmscatterSC2.SetFocus
End Sub

Private Sub MunVar_Ana_Ranking_Click()
file_type = "Va1"
Load frmVar1
frmVar1.SetFocus
End Sub

Private Sub MunVar_Ana_Scaling_Click()
file_type = "Va2"
Load frmVar2
frmVar2.SetFocus
End Sub

Private Sub numHelp_Click()
'MsgBox "Please see the introduction"
End Sub

Private Sub numRascw_Click()
    frmRascW.Show
    frmRascW.SetFocus
End Sub

Public Sub RunGraph()
Dim X, Y
X = Shell(App.Path & "\rascwin.exe", vbNormalFocus)
'y = MsgBox(" Done! Result are .dum .den .sac . var. ran", 0, "")
End Sub


'******************************************************
'from makedat menu


Private Sub mnuAbout_Click()
  Load frmAboutBox
End Sub

Private Sub mnuAggrange_Click()
  MDIFrmCascRasc.Arrange vbArrangeIcons

End Sub

Private Sub mnuCascade_Click()
  MDIFrmCascRasc.Arrange vbCascade
  
End Sub

Private Sub mnuDictoinary_Click()

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuIamge_Click()

End Sub

'This example uses CreateDatabase to create a new, encrypted Database object.

Sub CreateDatabaseX()
    Dim wrkDefault As Workspace
    Dim dbsNew As Database
    Dim prpLoop As Property
    Dim count As Integer
    Dim response
    Dim NumOfWellNames
    Dim Number As Integer
    Dim newfrmWell As New frmMakedatWell
   

   Dim I As Integer
   Set wrkDefault = DBEngine.Workspaces(0)

'CmDlg1.Filter = "(*.mdb)|*.MDB"
'    CmDlg1.ShowSave
  
'    myDBName = CmDlg1.Filename
    frmSaveDBS.Show 1
       
   If myDBName = "" Then Exit Sub
    
    ' Make sure there isn't already a file with the name of
    ' the new database.
   
    If Len(Dir(myDBName)) <> 0 Then
         response = MsgBox("Over write the existing file?", vbYesNoCancel)
         If response = vbYes Then
          Kill myDBName
         Else
          Exit Sub
         End If
    End If
    ' Create a new encrypted database with the specified
    ' collating order.
Set dbsNew = wrkDefault.CreateDatabase(myDBName, _
        dbLangGeneral, dbEncrypt)
dbsNew.Close

'With dbsNew
'        Debug.Print "Properties of "
        ' Enumerate the Properties collection of the new
        ' Database object.
'        For Each prpLoop In .Properties
'            If prpLoop <> "" Then Debug.Print "    " & _
'                prpLoop.Name & " = " & prpLoop
'        Next prpLoop
 '   End With

  NumOfWellNames = InputBox("How many wells?")
  If Len(Trim(NumOfWellNames)) = 0 Or IsNumeric(NumOfWellNames) = False Or NumOfWellNames <= 0 Then
     MsgBox "Positive numeric value only!!", vbExclamation
     Exit Sub
  Else
       
   Number = NumOfWellNames
   'Create header table
   CreateWellHeader (Number)
        
     For I = 1 To NumOfWellNames
      Well_Name_Temp = Trim("well") & Trim(Str(I))
'      Well_Name_Header = Trim("awell") & Trim(Str(I)) + "_hdr"
      CreateWellTableDefX (Well_Name_Temp)
      Well_Name_Temp = Well_Name_Temp + "_inf"
      
'        newfrmWell.Show
'        Load frmMakedatWell
     Next I
  CreateDicTableDefX
 ' MsgBox "Database " + myDBName + " has been created, please use Makedat"
  End If
  
  
'dbsNew.Close

'Load frmMakedatEvent

End Sub

'This example creates a new TableDef object in the Northwind database.

Sub CreateDicTableDefX()
    Dim prpLoop As Property
    If myDBName = "" Then Exit Sub
    
    Set DBsSAGA = OpenDatabase(myDBName)

    ' Create a new TableDef object.
    Set SAGADfNewDic = DBsSAGA.CreateTableDef("Event_name")

    With SAGADfNewDic
        ' Create fields and append them to the new TableDef
        ' object. This must be done before appending the
        ' TableDef object to the TableDefs collection of the
        ' Northwind database.

        .Fields.Append .CreateField("Fossil_Name", dbText)
        .Fields.Append .CreateField("Event_Number", dbLong)
        .Fields.Append .CreateField("Event_Type", dbText, 3)
        .Fields.Append .CreateField("Class_of_Fossils", dbText, 3)
        .Fields.Append .CreateField("Authors", dbText)
        .Fields.Append .CreateField("Year", dbInteger)
        .Fields.Append .CreateField("Reference", dbText)
        .Fields.Append .CreateField("Type_Locality", dbText)
        .Fields.Append .CreateField("Biozone", dbText)
        .Fields.Append .CreateField("Age", dbText)
        .Fields.Append .CreateField("Synonyms", dbText, 1)
        .Fields.Append .CreateField("Image", dbText, 1)

        .Fields(0).AllowZeroLength = True
        .Fields(2).AllowZeroLength = True
        .Fields(3).AllowZeroLength = True
        .Fields(4).AllowZeroLength = True
        .Fields(6).AllowZeroLength = True
        .Fields(7).AllowZeroLength = True
        .Fields(8).AllowZeroLength = True
        .Fields(9).AllowZeroLength = True
        .Fields(10).AllowZeroLength = True
        .Fields(11).AllowZeroLength = True
     
                
        Debug.Print "Properties of new TableDef object " & _
            "before appending to collection:"

        ' Enumerate Properties collection of new TableDef
        ' object.
               
        For Each prpLoop In .Properties
            On Error Resume Next
            If prpLoop <> "" Then
            Debug.Print "    " & prpLoop.name & " = " & prpLoop
            On Error GoTo 0
            End If
            
        Next prpLoop

        ' Append the new TableDef object to the Northwind
        ' database.
        DBsSAGA.TableDefs.Append SAGADfNewDic

        Debug.Print "Properties of new TableDef object " & _
            "after appending to collection:"

        ' Enumerate Properties collection of new TableDef
        ' object.
        For Each prpLoop In .Properties
            On Error Resume Next
            If prpLoop <> "" Then
            Debug.Print "    " & prpLoop.name & " = " & prpLoop
            On Error GoTo 0
            End If
        Next prpLoop

    End With
 DBsSAGA.Close
 
End Sub
'***************************************
Sub CreateWellTableDefX(well As String)
'creating new sample table
    Dim prpLoop As Property
    If myDBName = "" Then Exit Sub
    
    Set DBsSAGA = OpenDatabase(myDBName)


  Set SAGADfNewWell = DBsSAGA.CreateTableDef(well + "_inf")

    With SAGADfNewWell
        ' Create fields and append them to the new TableDef
        ' object. This must be done before appending the
        ' TableDef object to the TableDefs collection of the
        ' Northwind database.

        .Fields.Append .CreateField("Sample_depth", dbLong)
        .Fields.Append .CreateField("Sample_type", dbText, 1)
        .Fields.Append .CreateField("Event_number", dbLong)
        .Fields.Append .CreateField("Fossil_Names", dbText)
        .Fields.Append .CreateField("Sample_Quality", dbText, 1)
        .Fields.Append .CreateField("Count", dbInteger)
        .Fields.Append .CreateField("Rel_Abundance", dbText, 1)
        .Fields.Append .CreateField("Event_Qualifier", dbText, 3)
        
        .Fields(1).AllowZeroLength = True
        .Fields(3).AllowZeroLength = True
        .Fields(4).AllowZeroLength = True
        .Fields(6).AllowZeroLength = True
        .Fields(7).AllowZeroLength = True
        
        Debug.Print "Properties of new TableDef object " & _
            "before appending to collection:"

        ' Enumerate Properties collection of new TableDef
        ' object.
               
        For Each prpLoop In .Properties
            On Error Resume Next
            If prpLoop <> "" Then
            Debug.Print "    " & prpLoop.name & " = " & prpLoop
            On Error GoTo 0
            End If
            
        Next prpLoop

        ' Append the new TableDef object to the Northwind
        ' database.
        DBsSAGA.TableDefs.Append SAGADfNewWell

        Debug.Print "Properties of new TableDef object " & _
            "after appending to collection:"

        ' Enumerate Properties collection of new TableDef
        ' object.
        For Each prpLoop In .Properties
            On Error Resume Next
            If prpLoop <> "" Then
            Debug.Print "    " & prpLoop.name & " = " & prpLoop
            On Error GoTo 0
            End If
        Next prpLoop
  End With
  DBsSAGA.Close
         
End Sub


'Private Sub mnuMakedatFiles_Click()
'Dim makeDatFileName As String
'On Error GoTo Errorhandler

'CmDlg1.Filter = "(*.lst)|*.LST|(*.dct)|*.DCT"
'CmDlg1.ShowOpen
'makeDatFileName = CmDlg1.Filename
'MsgBox makeDatFileName

'Errorhandler:
'   Exit Sub
'
'End Sub

Private Sub mnuNew_Click()
CreateDatabaseX
MsgBox "Database " + myDBName + " has been created, please use Makedat", , "Makedat, Rasc, Casc and Cor"
End Sub

Private Sub mnuOpen_Click()

'CmDlg1.Filter = "(*.mdb)|*.MDB"
'CmDlg1.ShowOpen
'myDBName = CmDlg1.Filename
frmOpenDBS.Show 1

If Len(myDBName) = 0 Then Exit Sub
If Len(Dir(myDBName)) = 0 Then
  MsgBox "Database does not exist, try again!"
  Exit Sub
End If
MsgBox "Database " + myDBName + " has been opened, please use Makedat", , "Makedat, Rasc, Casc and Cor"
End Sub

Private Sub mnuWells_Click()
Dim newfrmWell() As New frmMakedatWell
Dim TabIndex As Integer
Dim TabCount As Integer
Dim I, J As Integer

'CmDlg1.ShowOpen
'CmDlg1.Filter = "(*.mdb)|*.MDB"

'myDBName = CmDlg1.Filename
If myDBName <> "" Then
frmSelect.Show 1
Else
MsgBox "Open a database from file menu", , " "
Exit Sub
End If

If Num_Edit_Well < 1 Then Exit Sub
ReDim newfrmWell(1 To Num_Edit_Well)
Set DBsSAGA = OpenDatabase(myDBName)
 TabCount = DBsSAGA.TableDefs.count
   For J = 1 To Num_Edit_Well
   
    For I = 0 To TabCount - 1
   
       If DBsSAGA.TableDefs(I).name = "well" + Trim(Str(Edit_Well_Num(J))) + "_inf" Then
        Well_Name_Temp = DBsSAGA.TableDefs(I).name
         Well_Caption = Trim(Edit_Well_Names(J))
         newfrmWell(J).Caption = Edit_Well_Names(J)
         
         newfrmWell(J).Show
       End If
    Next I
   Next J
  
 DBsSAGA.Close
 
'frmEditWells.Show
'Dim newfrmWell As New frmMakedatWell
'CmDlg1.ShowOpen
'myDBName = CmDlg1.filename
'load frmMakedatWell
' select a well to display

End Sub


'**************************************************************
'End of makedat menu




Public Sub CreateWellHeader(Number As Integer)
'creating new sample table
'****************************************8
'creat well header
    Dim prpLoop As Property
    If myDBName = "" Then Exit Sub
    
    Set DBsSAGA = OpenDatabase(myDBName)

   Set SagaDfNewHeader = DBsSAGA.CreateTableDef("header")

    With SagaDfNewHeader
        ' Create fields and append them to the new TableDef
        ' object. This must be done before appending the
        ' TableDef object to the TableDefs collection of the
        ' Northwind database.
        .Fields.Append .CreateField("Well_ID", dbText)
        .Fields.Append .CreateField("Well_name", dbText)
        .Fields.Append .CreateField("Geo_Coordinate", dbText)
        .Fields.Append .CreateField("Rot_Table_Height", dbDouble)
        .Fields.Append .CreateField("Unit", dbText, 1)
        .Fields.Append .CreateField("Water_Depth", dbDouble)
        .Fields.Append .CreateField("Comments", dbText)
        
        .Fields(0).AllowZeroLength = True
        .Fields(1).AllowZeroLength = True
        .Fields(2).AllowZeroLength = True
        .Fields(4).AllowZeroLength = True
        .Fields(6).AllowZeroLength = True
        
        Debug.Print "Properties of new TableDef object " & _
            "before appending to collection:"

        ' Enumerate Properties collection of new TableDef
        ' object.
               
        For Each prpLoop In .Properties
            On Error Resume Next
            If prpLoop <> "" Then
            Debug.Print "    " & prpLoop.name & " = " & prpLoop
            On Error GoTo 0
            End If
            
        Next prpLoop

        ' Append the new TableDef object to the Northwind
        ' database.
        DBsSAGA.TableDefs.Append SagaDfNewHeader

        Debug.Print "Properties of new TableDef object " & _
            "after appending to collection:"

        ' Enumerate Properties collection of new TableDef
        ' object.
        For Each prpLoop In .Properties
            On Error Resume Next
            If prpLoop <> "" Then
            Debug.Print "    " & prpLoop.name & " = " & prpLoop
            On Error GoTo 0
            End If
        Next prpLoop
     End With
  DBsSAGA.Close
End Sub

'Public Sub OpenMakeDic(File As String)''

'End Sub

'Public Sub OpenMakeWells()
'
'End Sub


Public Function Count_Well(ByVal File_Name As String) As Integer
Dim sum As Integer
Dim Temp As String
Dim gFileNum As Integer
sum = 0
gFileNum = FreeFile

Open CurDir + "\" + File_Name For Input As gFileNum

While Not (EOF(gFileNum))
Input #gFileNum, Temp
   If Mid(Temp, 1, 5) = "Depth" Then
   sum = sum + 1
   End If
Wend
Count_Well = sum
Close gFileNum
End Function

Public Function Count_Well_IPS(ByVal File_Name As String) As Integer
Dim sum As Integer
Dim Temp As String
Dim gFileNum As Integer
sum = 0
gFileNum = FreeFile

Open CurDir + "\" + File_Name For Input As gFileNum

While Not (EOF(gFileNum))
Line Input #gFileNum, Temp
   If Len(Temp) > 3 And LCase(Mid(Temp, 1, 4)) = "well" Then
   sum = sum + 1
   End If
Wend
Count_Well_IPS = sum
Close gFileNum
End Function
Sub CreateNewMakeDat(ByVal Well_Num As Integer)
    Dim wrkDefault As Workspace
    Dim dbsNew As Database
    Dim prpLoop As Property
    Dim count As Integer
    Dim response
    Dim NumOfWellNames As Integer
    Dim Number As Integer
    Dim newfrmWell As New frmMakedatWell
   

   Dim I As Integer
   Set wrkDefault = DBEngine.Workspaces(0)
 
    If myDBName = "" Then Exit Sub
    
    ' Make sure there isn't already a file with the name of
    ' the new database.
   
    If Len(Dir(myDBName)) <> 0 Then
         response = MsgBox("Over write the existing file?", vbYesNoCancel)
         If response = vbYes Then
          Kill myDBName
         Else
          Exit Sub
         End If
    End If
    ' Create a new encrypted database with the specified
    ' collating order.
Set dbsNew = wrkDefault.CreateDatabase(myDBName, _
        dbLangGeneral, dbEncrypt)
dbsNew.Close

 

  NumOfWellNames = Well_Num
  If Len(Trim(NumOfWellNames)) = 0 Or IsNumeric(NumOfWellNames) = False Or NumOfWellNames <= 0 Then
     MsgBox "Positive numeric value only!!", vbExclamation
     Exit Sub
  Else
       
   Number = NumOfWellNames
   'Create header table
     CreateWellHeader (Number)
        
     For I = 1 To NumOfWellNames
      Well_Name_Temp = Trim("well") & Trim(Str(I))
'      Well_Name_Header = Trim("awell") & Trim(Str(I)) + "_hdr"
      CreateWellTableDefX (Well_Name_Temp)
      Well_Name_Temp = Well_Name_Temp + "_inf"
      
'        newfrmWell.Show
'        Load frmMakedatWell
     Next I
  CreateDicTableDefX
 ' MsgBox "Database " + myDBName + " has been created, please use Makedat"
  End If
  
  
'dbsNew.Close

'Load frmMakedatEvent

End Sub


Public Function Add_Dic(ByVal File_Name As String)
Dim Temp, name As String
Dim TempLen, Index As Integer
Dim gFileNum As Integer
Dim Currentpos As Integer
Dim rstEvent_Name As Recordset
Dim DBsSAGA As Database
Dim MyRecordset As Recordset
Dim tempName As String

  If File_Name = "" Then
    Exit Function
  ElseIf LCase(right(Trim(File_Name), 4)) <> ".lst" Then
       MsgBox "Wrong file name has been selected"
       Exit Function
  Else
    TempLen = Len(Trim(File_Name))
    If TempLen < 5 Then
    Exit Function
    Else
    tempName = Trim(File_Name)
    
    Mid(tempName, TempLen - 2) = "dic"
    End If
  End If
  
 If Len(Dir(tempName)) = 0 Then
    MsgBox "The file  " + tempName + "does not exist"
    Exit Function
 End If
Index = 0

   Set DBsSAGA = OpenDatabase(myDBName)
   Set MyRecordset = DBsSAGA.OpenRecordset("Event_name")
      
      
gFileNum = FreeFile

Open CurDir + "\" + tempName For Input As gFileNum


While Not (EOF(gFileNum))
Line Input #gFileNum, Temp
    If Trim(Temp) <> "LAST" Then
      name = Mid(Temp, 1, 41)
      Index = Index + 1
      'add data into dic event_name table
   
   With MyRecordset
       Currentpos = .RecordCount
       
       .AddNew
             
        .Fields(0).Value = name
        .Fields(1).Value = Index
        .Fields(2).Value = ""
        .Fields(3).Value = ""
        .Fields(4).Value = ""
        .Fields(5).Value = 0
        .Fields(6).Value = ""
        .Fields(7).Value = ""
        .Fields(8).Value = ""
        .Fields(9).Value = ""
        .Fields(10).Value = ""
        .Fields(11).Value = ""
                 
       .Update
      .Bookmark = .LastModified
    
    End With
   End If
Wend
Close gFileNum
MyRecordset.Close
DBsSAGA.Close

End Function
Public Function Add_Dic_IPS(ByVal File_Name As String, ByVal Max_Event_IPS As Integer)
Dim Temp, name As String
Dim TempLen, Index As Integer
Dim gFileNum As Integer
Dim Currentpos As Integer
Dim rstEvent_Name As Recordset
Dim DBsSAGA As Database
Dim MyRecordset As Recordset
Dim tempName As String
Dim I, J As Integer
Dim Temp1, temp3 As String
Dim slashPS1, slashPS2 As Integer
  If File_Name = "" Then
    Exit Function
  ElseIf LCase(right(Trim(File_Name), 4)) <> ".ips" Then
       MsgBox "Wrong file name has been selected"
       Exit Function
  End If
 If Len(Dir(File_Name)) = 0 Then
    MsgBox "The file  " + File_Name + "does not exist"
    Exit Function
 End If
Index = 0

   Set DBsSAGA = OpenDatabase(myDBName)
   Set MyRecordset = DBsSAGA.OpenRecordset("Event_name")
     

 For I = 1 To Max_Event_IPS
 FossilNM = "dummy"
 gFileNum = FreeFile
  Open CurDir + "\" + File_Name For Input As gFileNum
    While Not (EOF(gFileNum))
    Line Input #gFileNum, Temp1
    If Len(Trim(Temp1)) > 3 And LCase(Mid(Trim(Temp1), 1, 4)) <> "well" Then
            slashPS1 = InStr(1, Trim(Temp1), ",", 1)
             If slashPS1 > 0 Then
                 temp3 = Mid(Trim(Temp1), slashPS1 + 1)
                 slashPS2 = InStr(1, Trim(temp3), ",", 1)
                   
                 'Depth = Left(Trim(Temp2), slashPS1 - 1)
                 'SampleTP = Mid(Trim(Temp2), slashPS + 1, 1)
                 EventNum = Int(left(Trim(temp3), slashPS2 - 1))
                 If EventNum = I Then
                 'Freq = Mid(Trim(Temp2), slashPS + 14, 1)
                 FossilNM = Trim(Mid(Trim(temp3), slashPS2 + 1))
                 End If
             End If
     End If
     Wend
           
   With MyRecordset
       Currentpos = .RecordCount
       
       .AddNew
             
        .Fields(0).Value = FossilNM
        .Fields(1).Value = I
        .Fields(2).Value = ""
        .Fields(3).Value = ""
        .Fields(4).Value = ""
        .Fields(5).Value = 0
        .Fields(6).Value = ""
        .Fields(7).Value = ""
        .Fields(8).Value = ""
        .Fields(9).Value = ""
        .Fields(10).Value = ""
        .Fields(11).Value = ""
                 
       .Update
      .Bookmark = .LastModified
    
    End With
    Close gFileNum
   Next I

MyRecordset.Close
DBsSAGA.Close

End Function

Public Function Add_wellheader(ByVal File_Name As String)
Dim Temp1, Temp2, temp3, Temp4, name As String
Dim TempLen, Index As Integer
Dim gFileNum As Integer
Dim Currentpos As Integer
Dim DBsSAGA As Database
Dim MyRecordsetHeader As Recordset
Dim HeightUnit As String
Dim Height As Double
Dim WaterD As Double
Dim WaterDUnit As String
Dim Comments As String
  If File_Name = "" Then
    Exit Function
  End If
  
 If Len(Dir(File_Name)) = 0 Then
    MsgBox "The file  " + File_Name + "does not exist"
    Exit Function
 End If
Index = 0

   Set DBsSAGA = OpenDatabase(myDBName)
   Set MyRecordsetHeader = DBsSAGA.OpenRecordset("header")
With MyRecordsetHeader
   gFileNum = FreeFile

  Open CurDir + "\" + File_Name For Input As gFileNum
  Line Input #gFileNum, Temp1
  Line Input #gFileNum, Temp2

While Not (EOF(gFileNum))
Line Input #gFileNum, Temp1
   If Len(Trim(Temp1)) > 6 And Mid(Trim(Temp1), 1, 6) = "Rotary" Then
      Index = Index + 1
      name = Temp2
      Height = Mid(Trim(Temp1), 21, 3)
      HeightUnit = right(Trim(Temp1), 1)
      Line Input #gFileNum, temp3
      WaterD = Mid(Trim(temp3), 13, 7)
      waterUnit = right(Trim(temp3), 1)
      Line Input #gFileNum, Temp4
      Comments = Mid(Trim(Temp4), 9)
        
         Currentpos = .RecordCount
       
       .AddNew
             
       .Fields(0).Value = Index
        .Fields(1).Value = name
        .Fields(2).Value = ""
        .Fields(3).Value = Height
        .Fields(4).Value = HeightUnit
        .Fields(5).Value = WaterD
        .Fields(6).Value = Comments
               
       .Update
      .Bookmark = .LastModified
        
      
   ElseIf Not (EOF(gFileNum)) Then
      Line Input #gFileNum, Temp2
      If Len(Trim(Temp2)) > 6 And Mid(Trim(Temp2), 1, 6) = "Rotary" Then
        name = Temp1
        Index = Index + 1
      
      Height = Mid(Trim(Temp2), 21, 3)
      HeightUnit = right(Trim(Temp2), 1)
      Line Input #gFileNum, temp3
      WaterD = Mid(Trim(temp3), 13, 7)
      waterUnit = right(Trim(temp3), 1)
      Line Input #gFileNum, Temp4
      Comments = Mid(Trim(Temp4), 9)
        
         Currentpos = .RecordCount
       
       .AddNew
             
       .Fields(0).Value = Index
        .Fields(1).Value = name
        .Fields(2).Value = ""
        .Fields(3).Value = Height
        .Fields(4).Value = HeightUnit
        .Fields(5).Value = WaterD
        .Fields(6).Value = Comments
               
       .Update
      .Bookmark = .LastModified
           
      End If
   End If
 
Wend
End With
Close gFileNum
MyRecordsetHeader.Close
DBsSAGA.Close

End Function


Public Function Add_wellheader_IPS(ByVal File_Name As String)
Dim Temp1, Temp2 As String
Dim TempLen, Index As Integer
Dim gFileNum As Integer
Dim Currentpos As Integer
Dim DBsSAGA As Database
Dim MyRecordsetHeader As Recordset
Dim HeightUnit As String
'Dim Height As Double
'Dim WaterD As Double
'Dim WaterDUnit As String
'Dim Comments As String
  If File_Name = "" Then
    Exit Function
  End If
  
 If Len(Dir(File_Name)) = 0 Then
    MsgBox "The file  " + File_Name + "does not exist"
    Exit Function
 End If
Index = 0

   Set DBsSAGA = OpenDatabase(myDBName)
   Set MyRecordsetHeader = DBsSAGA.OpenRecordset("header")
With MyRecordsetHeader
   gFileNum = FreeFile

  Open CurDir + "\" + File_Name For Input As gFileNum
'  Line Input #gFileNum, Temp1
'  Line Input #gFileNum, Temp2

While Not (EOF(gFileNum))
Line Input #gFileNum, Temp1
   If Len(Trim(Temp1)) > 3 And LCase(Mid(Trim(Temp1), 1, 4)) = "well" Then
      Index = Index + 1
      Temp2 = Mid(Temp1, 5)
        Currentpos = .RecordCount
       .AddNew
             
       .Fields(0).Value = Index
        .Fields(1).Value = Temp2
        .Fields(2).Value = ""
        .Fields(3).Value = 0
        .Fields(4).Value = ""
        .Fields(5).Value = 0
        .Fields(6).Value = ""
               
       .Update
      .Bookmark = .LastModified
   End If
Wend
End With
Close gFileNum
MyRecordsetHeader.Close
DBsSAGA.Close

End Function

Public Function Add_welldata(ByVal File_Name As String)
Dim Temp1, Temp2, temp3, Temp, name As String
Dim TempLen, Index As Integer
Dim gFileNum As Integer
Dim Currentpos As Integer
Dim DBsSAGA As Database
Dim MyRecordsetWell As Recordset
Dim Freq As String
Dim Depth As Double
Dim SampleTP As String
Dim EventNum As Integer
Dim Blank_lines As Integer
Dim slashPS As Integer
Dim FossilNM As String

  If File_Name = "" Then
    Exit Function
  End If
  
 If Len(Dir(File_Name)) = 0 Then
    MsgBox "The file  " + File_Name + "does not exist"
    Exit Function
 End If
Index = 0

   Set DBsSAGA = OpenDatabase(myDBName)
   
   
'   Set MyRecordsetHeader = DBsSAGA.OpenRecordset("header")
'With MyRecordsetHeader
   
   gFileNum = FreeFile

  Open CurDir + "\" + File_Name For Input As gFileNum
Index = 0
While Not (EOF(gFileNum))
Line Input #gFileNum, Temp1
 If Len(Trim(Temp1)) > 5 And Mid(Trim(Temp1), 1, 5) = "Depth" Then
         Index = Index + 1
         Line Input #gFileNum, Temp 'skip one line
   Set MyRecordsetWell = DBsSAGA.OpenRecordset(Trim("well") & Trim(Str(Index)) + "_inf")
    With MyRecordsetWell
      Blank_lines = 0
         While Not (EOF(gFileNum)) And Blank_lines < 2
            Line Input #gFileNum, Temp2
            
            If Len(Temp2) = 0 Then
                 Blank_lines = Blank_lines + 1
            ElseIf Blank_lines = 1 Then
                slashPS = InStr(1, Trim(Temp2), "/", 1)
                 
                 If slashPS > 0 Then
                 Depth = left(Trim(Temp2), slashPS - 1)
                 SampleTP = Mid(Trim(Temp2), slashPS + 1, 1)
                 EventNum = Mid(Trim(Temp2), slashPS + 2, 11)
                 Freq = Mid(Trim(Temp2), slashPS + 14, 1)
                 FossilNM = Mid(Trim(Temp2), slashPS + 15)
                 Blank_lines = 0
                 
            'add data to well table
    
        
                  Currentpos = .RecordCount
       
                    .AddNew
               
                    .Fields(0).Value = Depth
                    .Fields(1).Value = SampleTP
                    .Fields(2).Value = EventNum
                    .Fields(3).Value = FossilNM
                    .Fields(4).Value = ""
                    .Fields(5).Value = 0
                    .Fields(6).Value = Freq
                    .Fields(7).Value = ""
                      
                    .Update
                    .Bookmark = .LastModified
         
                  End If
                             
             ElseIf Blank_lines = 0 Then
             
                slashPS = InStr(1, Trim(Temp2), "/", 1)
                 
                 If slashPS > 0 Then
                 EventNum = left(Trim(Temp2), slashPS - 1)
                 Freq = Mid(Trim(Temp2), slashPS + 1, 1)
                 FossilNM = Mid(Trim(Temp2), slashPS + 2)
                 Blank_lines = 0
                 
             'add data to well table
   
                Currentpos = .RecordCount
       
                    .AddNew
               
                    .Fields(0).Value = Depth
                    .Fields(1).Value = SampleTP
                    .Fields(2).Value = EventNum
                    .Fields(3).Value = FossilNM
                    .Fields(4).Value = ""
                    .Fields(5).Value = 0
                    .Fields(6).Value = Freq
                    .Fields(7).Value = ""
                      
                    .Update
                    .Bookmark = .LastModified
                  End If
                   
              End If
              
            Wend
              End With
                MyRecordsetWell.Close
   End If
Wend
 
Close gFileNum
 
DBsSAGA.Close

End Function

Public Function Add_welldata_IPS(ByVal File_Name As String) As Integer
Dim Temp1, Temp2, temp3 As String
Dim TempLen, Index As Integer
Dim gFileNum As Integer
Dim Currentpos As Integer
Dim DBsSAGA As Database
Dim MyRecordsetWell As Recordset
'Dim Freq As String
Dim Depth As Double
'Dim SampleTP As String
Dim EventNum As Integer
'Dim Blank_lines As Integer
Dim slashPS1, slashPS2 As Integer
Dim FossilNM As String
Dim Temp_Number As Integer
   Temp_Number = 0
  If File_Name = "" Then
    Exit Function
  End If
  
 If Len(Dir(File_Name)) = 0 Then
    MsgBox "The file  " + File_Name + "does not exist"
    Exit Function
 End If
Index = 0

   Set DBsSAGA = OpenDatabase(myDBName)
   
   
'   Set MyRecordsetHeader = DBsSAGA.OpenRecordset("header")
'With MyRecordsetHeader
   
   gFileNum = FreeFile

  Open CurDir + "\" + File_Name For Input As gFileNum
Index = 0
Temp1 = "tttttt"
While Not (EOF(gFileNum))
  If LCase(Mid(Trim(Temp1), 1, 4)) <> "well" Then
   Line Input #gFileNum, Temp1
  End If
 If Len(Trim(Temp1)) > 3 And LCase(Mid(Trim(Temp1), 1, 4)) = "well" Then
         Index = Index + 1
   Set MyRecordsetWell = DBsSAGA.OpenRecordset(Trim("well") & Trim(Str(Index)) + "_inf")
    With MyRecordsetWell
        Temp2 = "ssssss"
         While Not (EOF(gFileNum)) And LCase(Mid(Temp2, 1, 4)) <> "well"
            Line Input #gFileNum, Temp2
 
            If Len(Temp2) > 3 And LCase(Mid(Temp2, 1, 4)) = "well" Then
            Temp1 = Temp2
        
            Else
                slashPS1 = InStr(1, Trim(Temp2), ",", 1)
                 
                 If slashPS1 > 0 Then
                 temp3 = Mid(Trim(Temp2), slashPS1 + 1)
                 slashPS2 = InStr(1, Trim(temp3), ",", 1)
                   
                 Depth = left(Trim(Temp2), slashPS1 - 1)
                 'SampleTP = Mid(Trim(Temp2), slashPS + 1, 1)
                 EventNum = left(Trim(temp3), slashPS2 - 1)
                 If EventNum > Temp_Number Then Temp_Number = EventNum
                 'Freq = Mid(Trim(Temp2), slashPS + 14, 1)
                 FossilNM = Mid(Trim(temp3), slashPS2 + 1)
                               
            'add data to well table
    
        
                  Currentpos = .RecordCount
       
                    .AddNew
               
                    .Fields(0).Value = Depth
                    .Fields(1).Value = ""
                    .Fields(2).Value = EventNum
                    .Fields(3).Value = FossilNM
                    .Fields(4).Value = ""
                    .Fields(5).Value = 0
                    .Fields(6).Value = ""
                    .Fields(7).Value = ""
                      
                    .Update
                    .Bookmark = .LastModified
         
                  End If
               End If
           Wend
         End With
       MyRecordsetWell.Close
   End If
Wend
Close gFileNum
 
DBsSAGA.Close
Add_welldata_IPS = Temp_Number
End Function


