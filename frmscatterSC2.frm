VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#6.0#0"; "olch2x32.ocx"
Begin VB.Form frmscatterSC2 
   Caption         =   "Scattergram - Scaling"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmscatterSC2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.TextBox TipText 
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   450
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   252
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   612
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Data"
      Height          =   252
      Left            =   3810
      TabIndex        =   6
      Top             =   150
      Width           =   1092
   End
   Begin VB.ListBox List_Well 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   3645
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   252
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin OlectraChart2D.Chart2D ChartScatter 
      Height          =   7680
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Visible         =   0   'False
      Width           =   11295
      _Version        =   393216
      _Revision       =   1
      _ExtentX        =   19923
      _ExtentY        =   13547
      _StockProps     =   0
      ControlProperties=   "frmscatterSC2.frx":0442
   End
   Begin VB.Label LabelWell 
      Caption         =   "Select a well"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "frmscatterSC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim E_Value() As Double
Dim x_Value() As Double
Dim y_value() As Double
Dim Dev_Value() As Double
Dim NameOfEvent() As String
Dim MaxNumWell As Integer
Dim NumOfEvent() As Integer
Dim well As Integer
Dim MaxArray As Integer
Dim nameOfWell() As String

Dim CurWell As Integer
Dim TempWellEvent As Integer
Dim EmaxTemp As Double


'For diaplay event name
'Some constants
'Const NumPoints As Integer = 200
Const HugeVal As Double = 1E+308
Const Closeness As Integer = 3

'Storage for tracking the mouse
Dim px As Long
Dim py As Long

'Storage for user interaction values
Dim Series As Long
Dim Pnt As Long
Dim Distance As Long
Dim Region As Long
Dim XVal As Double
Dim YVal As Double



 

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        ChartTableLoader = 8
        frmChartTable.Show
    Else
       frmChartTable.Hide
    End If
  
End Sub

Private Sub cmdCancel_Click()
    Unload frmscatterSC2
End Sub

Private Sub cmdOpen_Click()
    file_type = "Sc2"
    frmOpenRan1.Show 1
    'cmdOpen.Enabled = False
    If CancelKeyPress = 1 Then
        CancelKeyPress = 0
        Exit Sub
    End If
    
    If txt_file_ran1 = "" Then
      Beep
      Exit Sub
    End If

    startup
    OpenFile
End Sub

Private Sub cmdPrint_Click()
    ChartScatter.PrintChart oc2dFormatStandardMetafile, oc2dScaleToFit, 0, 0, 0, 0
        'Some printer program will change current path, here recover the workplace one time
    ChDir CurrentDir
    ChDrive CurrentDrive

End Sub

 

Private Sub Command1_Click()

     Set CurChartSaveObj = ChartScatter
     ChartSaveAs


'    Dim ImageName As String
'    ImageName = InputBox("Please give a file name without extension", "Save chart as image (.JPG)", 1)
'    If ImageName <> "" Then
'        ChartScatter.SaveImageAsJpeg ImageName & ".jpg", 100, False, False, False
'    End If
End Sub

Private Sub Form_Activate()
    Set CurGraphicOBJ = ChartScatter
        CurWindowNum = 5
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_Deactivate()
    Set CurGraphicOBJ = Nothing
End Sub

Private Sub Form_Load()
    txt_file_ran1 = ""
    CancelKeyPress = 0
    cmdOpen.Enabled = True
    
    TipText.Visible = False
    
    CurWindowNum = 5
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

End Sub

'For display event number and name
Private Sub ChartScatter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The user is moving the mouse over the chart control so either just update
' the header to reflect the position, and if the user is on a dot, then
' change the mouse cursor to a cross-hair to reflect this.

    Dim HoldY As Double
    Dim I As Integer
    Dim J As Integer
    
    px = X / Screen.TwipsPerPixelX
    py = Y / Screen.TwipsPerPixelY
    
    'Check to see if the mouse is over a point, and if so, display the values in the header
    With ChartScatter.ChartGroups(2)
        Region = .CoordToDataCoord(px, py, XVal, YVal)
        If Region = oc2dRegionInChartArea Then
            .CoordToDataIndex px, py, oc2dFocusXY, Series, Pnt, Distance
            If (Series <> -1) And (Pnt <> -1) Then
                HoldY = .Data.Y(Series, Pnt)
                If Distance <= Closeness And HoldY <> HugeVal Then
                    'if the point is the event point, then show result
                    'If Series = 2 Then
                      XVal = .Data.X(Series, Pnt)
                      
                      'Find out corresponding event name
                      TempWellEvent = 0
                      For J = 1 To CurWell - 1
                         TempWellEvent = TempWellEvent + NumOfEvent(J)
                      Next J
                      For I = 1 To NumOfEvent(CurWell)
                         If x_Value(TempWellEvent + I) = XVal And (EmaxTemp - y_value(TempWellEvent + I)) = HoldY Then
                            TipText.Text = LTrim(NameOfEvent(TempWellEvent + I))
                            TipText.left = px * Screen.TwipsPerPixelX
                            TipText.top = py * Screen.TwipsPerPixelY
                            TipText.Width = Len(Trim(NameOfEvent(TempWellEvent + I))) * 85
                            TipText.Visible = True
                         End If
                      Next I
       
                      'TipText.Text = "X = " + Format$(XVal, "0") + ", Y = " + Format$(HoldY, "0")
                      
                      Me.MousePointer = vbCrosshair
                    'End If
                Else
                    TipText.Visible = False
                    Me.MousePointer = vbNormal
                End If
            Else
                TipText.Visible = False
                Me.MousePointer = vbNormal
            End If
        Else
            TipText.Visible = False
            Me.MousePointer = vbNormal
        End If
    End With
End Sub

Public Sub startup()
Dim I, J As Integer

ChartScatter.Visible = False
List_Well.Clear

ChartScatter.ChartGroups(1).Styles(1).Symbol.size = 0
ChartScatter.ChartGroups(1).ChartType = oc2dTypePlot
With ChartScatter.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 1  '**** changed to 2 from 1
 '   .NumPoints(1) = 150
End With

ChartScatter.ChartGroups(1).Data.IsBatched = False

End Sub


Public Sub OpenFile()
Dim FileNum As Integer
Dim Filename As String

Dim Temp As String
Dim tempfrom As Integer
Dim tempto As Integer
Dim tempstr As String
Dim tempstr1 As String
Dim minx, maxx As Double
Dim I, J As Integer
Dim title_str As String

Dim ArrayNum As Integer
Dim tnum

FileNum = FreeFile
If Dir(CurDir + "\" + txt_file_ran1) <> "" Then
    Open CurDir + "\" + txt_file_ran1 For Input As FileNum
Else
    MsgBox "Input file does not exist, please try again"
Exit Sub
End If
    
Line Input #FileNum, Temp
Line Input #FileNum, Temp
Line Input #FileNum, Temp
'    MsgBox temp
MaxNumWell = Val(Mid(Temp, 25, 4))
' MsgBox MaxNumWell
Line Input #FileNum, Temp
ReDim NumOfEvent(1 To MaxNumWell)
ReDim nameOfWell(1 To MaxNumWell)

'read number of events per well
  For I = 1 To MaxNumWell
      Line Input #FileNum, Temp
      NumOfEvent(I) = Mid(Temp, 16, 4)
' MsgBox "Number of Events   : " + Str(NumOfEvent(I))

  Next I
' read in all data

  For I = 1 To 3
      Line Input #FileNum, Temp
  Next I
  
 Input #FileNum, Temp
 nameOfWell(1) = Mid(Temp, 1, 40)

' For I = 1 To 3
Input #FileNum, Temp
Line Input #FileNum, Temp
Input #FileNum, Temp
'Next I
' define array x-Value, E_Valie, Dev_Value
MaxArray = 0

For I = 1 To MaxNumWell
   MaxArray = MaxArray + NumOfEvent(I)
Next I
'   MsgBox "MAx array is   " + Str(MaxArray)
   
ReDim x_Value(1 To MaxArray)
ReDim E_Value(1 To MaxArray)
ReDim y_value(1 To MaxArray)
ReDim NameOfEvent(1 To MaxArray)
'ReDim RowContent(1 To MaxArray)

'ReDim Dev_Value(1 To MaxArray)

' start to read in data
ArrayNum = 0
    For I = 1 To MaxNumWell
             If I > 1 Then
                For J = 1 To 6
                    Input #FileNum, Temp
                Next J
                
                'read in the name of well
                Input #FileNum, Temp
                nameOfWell(I) = Mid(Temp, 1, 40)
                
                For J = 1 To 3
                    Input #FileNum, Temp
                Next J
             End If
             
        For J = 1 To NumOfEvent(I)
            ArrayNum = ArrayNum + 1
            Line Input #FileNum, Temp
'            If J < 10 Then
'                tnum = 2
'            Else
'                tnum = 1
'            End If
                                 
            x_Value(ArrayNum) = Val(Mid(Temp, 4, 10))
            y_value(ArrayNum) = Val(Mid(Temp, 14, 10))
            E_Value(ArrayNum) = Val(Mid(Temp, 24, 10))
            NameOfEvent(ArrayNum) = Mid(Temp, 45, Len(Temp) - 44)
         '   RowContent(ArrayNum) = Temp
         '   Dev_Value(ArrayNum) = Mid(temp, 34 - tnum, 10)
            
'            If I = MaxNumWell Then
 '           MsgBox Str(x_Value(ArrayNum)) + "   " + Str(E_Value(ArrayNum)) + "  " + Str(Dev_Value(ArrayNum))
'            End If
            Next J
     Next I
'     MsgBox "The actural num of MaxArray is   " + Str(ArrayNum)

Close FileNum

'select well to plot
List_Well.Enabled = True
LabelWell.Enabled = True

  For I = 1 To MaxNumWell
    List_Well.AddItem "Well No. " + Str(I) + ": " + nameOfWell(I)
  Next I
End Sub


Private Sub Form_Resize()
Dim size As Integer
If (frmscatterSC2.Height > 360) And (frmscatterSC2.Height - 360 < frmscatterSC2.Width / 1.5) Then
   size = frmscatterSC2.Height - 360
ElseIf (frmscatterSC2.Height > 360) And (frmscatterSC2.Height - 360 > frmscatterSC2.Width / 1.5) Then
   size = frmscatterSC2.Width / 1.5
Else
   size = 0
End If

ChartScatter.Height = size
ChartScatter.Width = 1.5 * size
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CurGraphicOBJ = Nothing
    
    CurWindowNum = 5
    Call MDIWindowsMenuDelete(CurWindowNum)
    WindowsHwnd(CurWindowNum) = -1

End Sub

Private Sub List_Well_Click()
Dim gridspace As Double
Dim I, J As Integer
Dim well As Integer
Dim Emin As Double, Emax As Double, E_temp As Double
Dim Name_temp As String
Dim title_str As String
Dim tempfrom As Integer
Dim TempY As Double
Dim Ymin As Double, Ymax As Double

ChartScatter.Visible = False

well = List_Well.ListIndex + 1
'MsgBox "Well No. == " + Str(well)
'LabelWell.Visible = False
'List_Well.Visible = False

'Save Current Well Number
CurWell = well


ChartScatter.ChartGroups(1).Styles(1).Symbol.size = 0
'ChartScatter.ChartGroups(1).Styles(1).Line.Width = 2
'ChartScatter.ChartGroups(1).Styles(1).Line.Color = 0
ChartScatter.ChartGroups(1).ChartType = oc2dTypePlot

If UCase(file_type) = "SC1" Then
    title_str = "Ranked"
ElseIf UCase(file_type) = "SC2" Then
    title_str = "Scaled"
End If
    
'ChartScatter.Header.Text.Text = "Scattergram   " + txt_file_ran1 + "  (Well No. " + Str(well) + ")"
ChartScatter.ChartArea.Axes("x").Title.Text = title_str + " Optimum Sequence of Events"


With ChartScatter.ChartGroups(1).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 2
    .NumPoints(1) = NumOfEvent(well)
     tempfrom = 0
         For J = 1 To well - 1
              tempfrom = tempfrom + NumOfEvent(J)
         Next J
    Emax = 0
    Emin = 0
 'sort data according to x values
       For I = tempfrom + 1 To tempfrom + NumOfEvent(well) - 1
        For J = I + 1 To tempfrom + NumOfEvent(well)
            If x_Value(I) > x_Value(J) Then
               E_temp = x_Value(I)
               x_Value(I) = x_Value(J)
               x_Value(J) = E_temp
               E_temp = E_Value(I)
               E_Value(I) = E_Value(J)
               E_Value(J) = E_temp
               E_temp = y_value(I)
               y_value(I) = y_value(J)
               y_value(J) = E_temp
               Name_temp = NameOfEvent(I)
               NameOfEvent(I) = NameOfEvent(J)
               NameOfEvent(J) = Name_temp
'               Name_temp = RowContent(I)
'               RowContent(I) = RowContent(J)
'               RowContent(J) = Name_temp
             End If
        Next J
        Next I
        
        For I = 1 To NumOfEvent(well)
            If y_value(tempfrom + I) < Emin Then
               Emin = y_value(tempfrom + I)
               Ymin = y_value(tempfrom + I)
            End If
            If y_value(tempfrom + I) > Emax Then
               Emax = y_value(tempfrom + I)
               Ymax = y_value(tempfrom + I)
            End If
        Next I
        For I = 1 To NumOfEvent(well)
            If Emax - E_Value(tempfrom + I) < Ymin Then
               Ymin = Emax - E_Value(tempfrom + I)
            End If
            If Emax - E_Value(tempfrom + I) > Ymax Then
               Ymax = Emax - E_Value(tempfrom + I)
            End If
        Next I
        
        frmChartTable.listOfData.Clear
        frmChartTable.listOfData.AddItem "   SN " + "        X" + "      Y" + "     No" + "  Event Name"
    
       For I = 1 To NumOfEvent(well)
        .X(1, I) = x_Value(tempfrom + I)
        .Y(1, I) = Emax - E_Value(tempfrom + I)
       ' .Y(1, I) = Emax - E_Value(tempfrom + I) + 1
         TempY = Emax - y_value(tempfrom + I)
       frmChartTable.listOfData.AddItem Space(5 - Len(Str(I))) + Str(I) _
                    + Space(10 - Len(Str(x_Value(tempfrom + I)))) _
                    + Str(x_Value(tempfrom + I)) _
                    + Space(7 - Len(Str(TempY))) _
                    + Str(TempY) _
                    + Space(2) + NameOfEvent(tempfrom + I)
        Next I
        
End With


ChartScatter.ChartArea.Axes("y").Min.Value = Int(Ymin) - 1
ChartScatter.ChartArea.Axes("y").Max.Value = Int(Ymax) + 1
ChartScatter.ChartArea.Axes("y2").Min.Value = Int(Ymin) - 1
ChartScatter.ChartArea.Axes("y2").Max.Value = Int(Ymax) + 1
ChartScatter.ChartArea.Axes("y2").Multiplier = 1
  
ChartScatter.ChartArea.Axes("x").Origin.Value = 0
ChartScatter.ChartArea.Axes("y").Origin.Value = Int(Ymax) + 1

ChartScatter.ChartArea.Axes("x").Min.Value = 0
ChartScatter.ChartArea.Axes("x").Max.Value = x_Value(tempfrom + NumOfEvent(well)) + 1

ChartScatter.ChartLabels(1).Text.Text = nameOfWell(well)
ChartScatter.ChartLabels(1).Anchor = oc2dAnchorSouthWest
ChartScatter.ChartLabels(1).AttachDataCoord.X = ChartScatter.ChartArea.Axes("x").Max.Value
ChartScatter.ChartLabels(1).AttachDataCoord.Y = ChartScatter.ChartArea.Axes("y").Min.Value

If UCase(file_type) = "SC2" Then
 gridspace = x_Value(tempfrom + NumOfEvent(well)) / 10
 ChartScatter.ChartArea.Axes("x").MajorGrid.Spacing = gridspace
End If

'Define the label spacing and precision of  X and Y axe
'ChartScatter.ChartArea.Axes("x").NumSpacing.Value = 5
ChartScatter.ChartArea.Axes("x").Precision = 0
'ChartScatter.ChartArea.Axes("y").NumSpacing.Value = 5
ChartScatter.ChartArea.Axes("y").Precision = 0


'Setting chart label number spacing for Y axes
Dim ChartYAxeNuberSpace As Double
Select Case (ChartScatter.ChartArea.Axes("y").Max.Value - ChartScatter.ChartArea.Axes("y").Min.Value)
Case Is < 2
   ChartYAxeNuberSpace = 0.5
Case 2 To 5
   ChartYAxeNuberSpace = 1
Case 5 To 20
   ChartYAxeNuberSpace = 2
Case 20 To 50
   ChartYAxeNuberSpace = 5
Case 50 To 100
   ChartYAxeNuberSpace = 10
Case 100 To 300
   ChartYAxeNuberSpace = 20
Case 300 To 500
   ChartYAxeNuberSpace = 25
Case 500 To 1000
   ChartYAxeNuberSpace = 50
Case 1000 To 2000
   ChartYAxeNuberSpace = 100
Case 2000 To 3500
   ChartYAxeNuberSpace = 200
Case 3500 To 5000
   ChartYAxeNuberSpace = 250
Case Else
   ChartYAxeNuberSpace = 500
End Select
ChartScatter.ChartArea.Axes("y").NumSpacing = ChartYAxeNuberSpace
ChartScatter.ChartArea.Axes("y2").NumSpacing = ChartYAxeNuberSpace
'Setting chart label number spacing for Y axes
Dim ChartXAxeNuberSpace As Double
Select Case (ChartScatter.ChartArea.Axes("x").Max.Value - ChartScatter.ChartArea.Axes("x").Min.Value)
Case Is < 2
   ChartXAxeNuberSpace = 0.5
Case 2 To 5
   ChartXAxeNuberSpace = 1
Case 5 To 20
   ChartXAxeNuberSpace = 2
Case 20 To 50
   ChartXAxeNuberSpace = 5
Case 50 To 100
   ChartXAxeNuberSpace = 10
Case 100 To 300
   ChartXAxeNuberSpace = 20
Case 300 To 500
   ChartXAxeNuberSpace = 25
Case 500 To 1000
   ChartXAxeNuberSpace = 50
Case 1000 To 2000
   ChartXAxeNuberSpace = 100
Case 2000 To 3500
   ChartXAxeNuberSpace = 200
Case 3500 To 5000
   ChartXAxeNuberSpace = 250
Case Else
   ChartXAxeNuberSpace = 500
End Select
ChartScatter.ChartArea.Axes("x").NumSpacing = ChartXAxeNuberSpace


'ChartScatter.ChartArea.Axes("y2").Origin.Value = Int(Emin) - 1

ChartScatter.ChartGroups(1).Data.IsBatched = False

  
''ChartScatter.ChartGroups(2).Styles(1).Symbol.size = 5
''ChartScatter.ChartGroups(2).ChartType = oc2dTypePlot


With ChartScatter.ChartGroups(2).Data
    .Layout = oc2dDataGeneral
    .IsBatched = True
    .NumSeries = 2
    .NumPoints(2) = NumOfEvent(well)
       For I = 1 To NumOfEvent(well)
        .X(2, I) = x_Value(tempfrom + I)
        .Y(2, I) = Emax - y_value(tempfrom + I)
'        .Y(2, I) = Emax - y_value(tempfrom + I) + 1
      Next I
End With

'Save Emax to EmaxTemp for event name display
EmaxTemp = Emax

ChartScatter.ChartGroups(2).Data.IsBatched = False
'ChartScatter.Visible = True
ChartScatter.Visible = True

End Sub

