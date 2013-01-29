Attribute VB_Name = "Module3"
Option Explicit
Global txt_file_inp As String
Global txt_file_dic As String
Global txt_file_RASCout As String
Global txt_file_den As String
Global txt_file_cum As String
Global txt_file_ran1 As String
Global txt_file_hout As String
Global txt_file_par As String
Global txt_file_dem As String
Global txt_file_df1 As String
Global txt_file_df2 As String
Global txt_file_di1 As String
Global txt_file_inc As String
Global txt_file_table As String


Global OpenFileKey As Integer
Global CancelKeyPress As Integer
Global DepthOrderNum As Integer
''For the RASC Parameter file
Global CurRASCParaFile(1 To 4) As String
'Global RASCParaFileChangeKey As Integer
''For the CASC Parameter file
Global CurCASCParaFile As String
'For new and old RASCW.exe Running control on probablic cycling function
Global RASCWVersionControlKey As Integer

'For graphics display control
Global CurGraphicOBJ As Object

Global file_type As String
Global well As Integer
Global ChartTableLoader As Integer
Global well_numbers(1 To 20) As Integer
Global event_numbers(1 To 20) As Integer
Global event_names(1 To 20) As String
Global max_event_num, max_well_num As Integer
Global prob_opt, obs_opt, norm_opt
Global orders() As Integer
Global modifyflag As Integer
Global testing_flag As Integer
Global Check1 As String 'frmran1 and frmran2 EventRanges-Ranking
Global txt_file_range1 ' 'frmran1 and frmran2 EventRanges-Ranking
'Global checking As Boolean
Global well_label_used() As String
Global TableMinObs(1 To 20, 1 To 20) As Double 'data table for casc result with plot
Global TableMaxObs(1 To 20, 1 To 20) As Double
Global TableObs(1 To 20, 1 To 20) As Double
Global TableMinProb(1 To 20, 1 To 20) As Double
Global TableMaxProb(1 To 20, 1 To 20) As Double
Global TableProb(1 To 20, 1 To 20) As Double
Global tableType As Integer 'value indicates different tables
Global TableEvent() As String  ' names of events to be putted into data table
Global TableRatio(1 To 20) As Double  'ratio values of Casc result to be included in data table
Global Table_col As Integer
Global Table_row As Integer
Global TableColName() As String
Global DataTableChk As Integer
Global Table_Title As String

'For frmwells2, frmwellplot2, frmDataChart2
Global max_event_num2, max_well_num2 As Integer
Global prob_opt2, obs_opt2, norm_opt2
Global modifyflag2 As Integer
Global well_numbers2(1 To 20) As Integer
Global event_numbers2(1 To 20) As Integer
Global event_names2(1 To 20) As String
Global orders2() As Integer
Global testing_flag2 As Integer
Global Check2 As String 'frmran3 and frmran4 EventRanges-Scaling
Global txt_file_range2 ' 'frmran3 and frmran42 EventRanges-Scaling
Global well_label_used2() As String
Global TableMinObs2(1 To 20, 1 To 20) As Double 'data table for casc result with plot
Global TableMaxObs2(1 To 20, 1 To 20) As Double
Global TableObs2(1 To 20, 1 To 20) As Double
Global TableMinProb2(1 To 20, 1 To 20) As Double
Global TableMaxProb2(1 To 20, 1 To 20) As Double
Global TableProb2(1 To 20, 1 To 20) As Double
'Global tableType2 As Integer 'value indicates different tables
'Global TableEvent2() As String  ' names of events to be putted into data table
Global TableRatio2(1 To 20) As Double  'ratio values of Casc result to be included in data table
'Global Table_col2 As Integer
'Global Table_row2 As Integer
'Global TableColName2() As String
'Global DataTableChk2 As Integer
Global Table_Title2 As String


Global MakeDat_File As String
Global DialogInitPath As String
Global CurrentDir As String
Global CurrentDrive As String


Global well_fossil As String
Global sort_field As String
Global sort_method As String
Global myDBName As String
Global DBsSAGA As Database
Global SAGADfNewDic As TableDef
Global SAGADfNewWell As TableDef
Global SagaDfNewHeader As TableDef
Global Well_Name_Temp As String
Global Well_Name_Header As String
Global Event_Name_Temp As String
Global Edit_Well_Names() As String
Global Edit_Well_Num() As Integer
Global Num_Edit_Well As Integer
Global Well_Caption As String
Global RascInputFile As String
Global Old_Makedat_File As String
Global New_Makedat_File As String
Global Del_Event_Names() As String
Global Del_Event_Numbers() As Integer
Global Del_Event_Total As Integer
Global TableCaption As String

'For MDI windows processing
Global WindowsName(1 To 30) As String
Global WindowsHwnd(1 To 30) As Long
Global CurWindowNum As Integer
'For Toolbar windows control
Global MdiFrmFirstTimeLoad As Integer
Global MDIfrmActive As Integer
Global hh As Single, vv As Single
Global ToolBarVisible As Integer
'For Chart Save As dialog
Global CurChartSaveObj As Object
Global ChartSaveInitPath As String
Global ChartFileTypeIndex As Integer
'For ChartShow file open dialog
Global ChartOpenInitPath As String
Global CurChartShowObj As Object



Sub Opentab()

    Dim NewTab As frmSumTable
    Set NewTab = New frmSumTable
    With NewTab
        .Show
    End With

End Sub

Sub RefleshDir()
    ChDir CurrentDir
    ChDrive CurrentDrive
    DialogInitPath = CurDir  'For load text files dialog
End Sub

Function CheckExistWindows(num As Integer) As Integer
    Dim I As Integer, J As Integer

    CheckExistWindows = 0
    If MDIFrmCascRasc.mnuCurWindows.count > 1 Then
        For I = 0 To MDIFrmCascRasc.mnuCurWindows.count - 1
            If MDIFrmCascRasc.mnuCurWindows(I).Caption = WindowsName(num) Then
               CheckExistWindows = I
            End If
        Next I
    End If
End Function

Sub MDIWindowsMenuAdd(num As Integer)
    Dim I As Integer, J As Integer

    Load MDIFrmCascRasc.mnuCurWindows(MDIFrmCascRasc.mnuCurWindows.count)
    MDIFrmCascRasc.mnuCurWindows(MDIFrmCascRasc.mnuCurWindows.count - 1).Caption = WindowsName(num)
    If MDIFrmCascRasc.mnuCurWindows.count > 1 Then
        For I = 0 To MDIFrmCascRasc.mnuCurWindows.count - 1
            MDIFrmCascRasc.mnuCurWindows(I).Visible = True
            If MDIFrmCascRasc.mnuCurWindows(I).Caption = WindowsName(num) Then
               MDIFrmCascRasc.mnuCurWindows(I).Checked = True
            Else
               MDIFrmCascRasc.mnuCurWindows(I).Checked = False
            End If
        Next I
    End If

End Sub

Sub MDIWindowsMenuDelete(num As Integer)
    Dim I As Integer, J As Integer

    For I = 0 To MDIFrmCascRasc.mnuCurWindows.count - 1
         If MDIFrmCascRasc.mnuCurWindows(I).Caption = WindowsName(num) Then
           If I < MDIFrmCascRasc.mnuCurWindows.count - 1 Then
              For J = I + 1 To MDIFrmCascRasc.mnuCurWindows.count - 1
                  MDIFrmCascRasc.mnuCurWindows(J - 1).Caption = MDIFrmCascRasc.mnuCurWindows(J).Caption
              Next J
           End If
           'Arrange to the last item and exit FOR Loop
           I = MDIFrmCascRasc.mnuCurWindows.count - 1
           Unload MDIFrmCascRasc.mnuCurWindows(I)
        End If
    Next I
    
    If MDIFrmCascRasc.mnuCurWindows.count > 1 Then
        For I = 0 To MDIFrmCascRasc.mnuCurWindows.count - 1
            MDIFrmCascRasc.mnuCurWindows(I).Visible = True
            If I = MDIFrmCascRasc.mnuCurWindows.count - 1 Then
               MDIFrmCascRasc.mnuCurWindows(I).Checked = True
               CurWindowSetFocus (I)
            Else
               MDIFrmCascRasc.mnuCurWindows(I).Checked = False
            End If
        Next I
    Else
        For I = 0 To MDIFrmCascRasc.mnuCurWindows.count - 1
            MDIFrmCascRasc.mnuCurWindows(I).Visible = False
        Next I
    End If

End Sub

Sub CurWindowSetFocus(Index As Integer)
    Dim I As Integer, J As Integer, k As Integer
    Dim WinTop As Long, WinLeft As Long, WinWidth As Long, WinLong As Long


        For I = 0 To MDIFrmCascRasc.mnuCurWindows.count - 1
            If I = Index Then
                MDIFrmCascRasc.mnuCurWindows(I).Checked = True
                For J = 1 To 30
                   If WindowsName(J) = MDIFrmCascRasc.mnuCurWindows(I).Caption Then
                        For k = 0 To Forms.count - 1
                            If WindowsHwnd(J) = Forms(k).hwnd Then
                                WinTop = Forms(k).top / Screen.TwipsPerPixelY
                                WinLeft = Forms(k).left / Screen.TwipsPerPixelX
                                WinWidth = Forms(k).Width / Screen.TwipsPerPixelX
                                WinLong = Forms(k).Height / Screen.TwipsPerPixelY
                                SetWindowPos Forms(k).hwnd, HWND_TOP, WinLeft, WinTop, WinWidth, WinLong, Flags
                             
                                ShowWindow Forms(k).hwnd, SW_SHOW
                                'SetFocus Forms(k).hwnd
                                'Forms(k).SetFocus
                                'Exit Sub
                            End If
                         Next k
                   End If
                Next J
            Else
                MDIFrmCascRasc.mnuCurWindows(I).Checked = False
            End If
        Next I
End Sub

Sub ChartSaveAs()
Dim X

'    ChDir CurrentDir
'    ChDrive CurrentDrive

'Routine to use the Common Dialog so we can select
' a  file to use.
        
     'Set up all of the required items and then call the OpenFile Common Dialog
    MDIFrmCascRasc.CommonDialog1.DialogTitle = "Save Chart As ..."
    'MDIFrmCascRasc.CommonDialog1.DefaultExt = "*."
    MDIFrmCascRasc.CommonDialog1.InitDir = ChartSaveInitPath
    MDIFrmCascRasc.CommonDialog1.Filter = "OC2 File (*.oc2)|*.oc2|EMF File (*.emf)|*.emf|WMF File (*.wmf)|*.wmf|JPEG File (*.jpg)|*.jpg|PNG File (*.png)|*.png|BMP File (*.bmp)|*.bmp"
    MDIFrmCascRasc.CommonDialog1.FilterIndex = ChartFileTypeIndex
    MDIFrmCascRasc.CommonDialog1.Filename = ""
    MDIFrmCascRasc.CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNExplorer '它指定对话框以浏览器的形式出现
    MDIFrmCascRasc.CommonDialog1.Action = 2   'showsave file
    
    '文件名空则退出
     If Trim(MDIFrmCascRasc.CommonDialog1.Filename) = "" Then
'         'Some printer program will change current path, here recover the workplace one time
'         ChDir CurrentDir
'         ChDrive CurrentDrive
         Exit Sub
    End If

'Save last time FilterIndex
    ChartFileTypeIndex = MDIFrmCascRasc.CommonDialog1.FilterIndex

    Select Case MDIFrmCascRasc.CommonDialog1.FilterIndex
    Case 1
           CurChartSaveObj.Save MDIFrmCascRasc.CommonDialog1.Filename
    
    Case 2
           CurChartSaveObj.DrawToFile MDIFrmCascRasc.CommonDialog1.Filename, oc2dFormatEnhMetafile
   
    Case 3
            CurChartSaveObj.DrawToFile MDIFrmCascRasc.CommonDialog1.Filename, oc2dFormatMetafile
   
    Case 4
            CurChartSaveObj.SaveImageAsJpeg MDIFrmCascRasc.CommonDialog1.Filename, 100, False, False, False
   
    Case 5
            CurChartSaveObj.SaveImageAsPng MDIFrmCascRasc.CommonDialog1.Filename, False
   
    Case 6
           CurChartSaveObj.DrawToFile MDIFrmCascRasc.CommonDialog1.Filename, oc2dFormatBitmap
    
    End Select
    
    'save current directory
    ChartSaveInitPath = CurDir
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

    
End Sub

Sub OpenSavedChart()
Dim X

'    ChDir CurrentDir
'    ChDrive CurrentDrive

'Routine to use the Common Dialog so we can select
' a  file to use.
        
     'Set up all of the required items and then call the OpenFile Common Dialog
    MDIFrmCascRasc.CommonDialog1.DialogTitle = "Open Graphic File ..."
    'MDIFrmCascRasc.CommonDialog1.DefaultExt = "*."
    MDIFrmCascRasc.CommonDialog1.InitDir = ChartOpenInitPath
    MDIFrmCascRasc.CommonDialog1.Filter = "OC2 File (*.oc2)|*.oc2"
    MDIFrmCascRasc.CommonDialog1.FilterIndex = 1
    MDIFrmCascRasc.CommonDialog1.Filename = ""
    MDIFrmCascRasc.CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNExplorer '它指定对话框以浏览器的形式出现+文件名和路径必须是存在和合法的
    MDIFrmCascRasc.CommonDialog1.Action = 1   'showopen file
    
    '文件名空则退出
     If Trim(MDIFrmCascRasc.CommonDialog1.Filename) = "" Then
'         'Some printer program will change current path, here recover the workplace one time
'         ChDir CurrentDir
'         ChDrive CurrentDrive
         Exit Sub
    End If

'Save last time FilterIndex
'    ChartFileTypeIndex = MDIFrmCascRasc.CommonDialog1.FilterIndex

'Load OC2 saved chart file to display
     CurChartShowObj.Load MDIFrmCascRasc.CommonDialog1.Filename
     
    CurChartShowObj.Visible = True
    
    'save current directory
    ChartOpenInitPath = CurDir
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

    
End Sub
