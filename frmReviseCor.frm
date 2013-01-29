VERSION 5.00
Begin VB.Form frmReviseCor 
   Caption         =   "Revision of Parameters to run Cor"
   ClientHeight    =   3375
   ClientLeft      =   1215
   ClientTop       =   2610
   ClientWidth     =   8295
   Icon            =   "frmReviseCor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8295
   Begin VB.CheckBox chkWellOrder 
      Caption         =   "Maintain Well Order?"
      Height          =   252
      Left            =   600
      TabIndex        =   18
      Top             =   1320
      Width           =   2172
   End
   Begin VB.CheckBox ChkRankingOrder 
      Caption         =   "Maintain Ranking Order?"
      Height          =   492
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   2072
   End
   Begin VB.Frame Frame1 
      Height          =   972
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   2532
   End
   Begin VB.Frame FrameLimit 
      Caption         =   "Interval within Ranked or Scaled Optimum Sequence"
      Height          =   852
      Left            =   480
      TabIndex        =   12
      Top             =   2400
      Width           =   4032
      Begin VB.TextBox txt_base 
         Height          =   288
         Left            =   2880
         TabIndex        =   14
         Text            =   " "
         Top             =   360
         Width           =   612
      End
      Begin VB.TextBox Txt_Top 
         Height          =   288
         Left            =   960
         TabIndex        =   13
         Text            =   " "
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label_base 
         Caption         =   "Base"
         Height          =   252
         Left            =   2400
         TabIndex        =   16
         Top             =   396
         Width           =   372
      End
      Begin VB.Label Label_top 
         Caption         =   "Top"
         Height          =   252
         Left            =   600
         TabIndex        =   15
         Top             =   396
         Width           =   372
      End
   End
   Begin VB.CommandButton cmdOpt_seq 
      Caption         =   "Optimum Sequence"
      Height          =   252
      Left            =   5760
      TabIndex        =   11
      Top             =   840
      Width           =   1572
   End
   Begin VB.ListBox List_events 
      Height          =   1620
      Left            =   4800
      MultiSelect     =   1  'Simple
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   3272
   End
   Begin VB.ComboBox cboOut 
      Height          =   288
      Left            =   3240
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   240
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.FileListBox FileOut 
      Height          =   675
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "New Name"
      Height          =   252
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   932
   End
   Begin VB.TextBox txtOut 
      Height          =   288
      Left            =   3240
      TabIndex        =   6
      Top             =   480
      Width           =   1212
   End
   Begin VB.CheckBox ChkScaling 
      Caption         =   "Use Scaling Results?"
      Height          =   492
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   1932
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   6960
      TabIndex        =   3
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton cmdApplyRevise 
      Caption         =   "Save"
      Height          =   252
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox txtMinWell 
      Height          =   288
      Left            =   3240
      TabIndex        =   0
      Top             =   1860
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "Select Two Events."
      Height          =   252
      Left            =   4800
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label lblMinWell 
      Caption         =   "Minimum Number of Wells in which each Event (per interval) occurs"
      Height          =   372
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   2628
   End
End
Attribute VB_Name = "frmReviseCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim Counter
Dim Totalwell As Integer
Dim MinEvents As Integer
Dim MinWells As Integer     'minmum number of wells per event in interval zone
Dim Scaling As Integer      'Binary variable for Scaling Result yes = 1
Dim RankOrder As Integer ' binary variable for changing ranking order
Dim Opt_top As Integer  'Optimum sequence limits Top
Dim Opt_Base As Integer 'Optimum sequence limits Base
Dim UseCasc As Integer
Dim Top_space, Base_space As String
Dim Temp As String
Dim Event_Num() As Integer
Dim Event_name() As String
Dim NumberOfEvents As Integer

Private Sub cboOut_Click()
   FileOut.Pattern = "*.par"
   
End Sub


Private Sub ChkRankingOrder_Click()
If ChkRankingOrder.Value = 1 Then
   chkWellOrder.Value = 0
End If

End Sub

Private Sub chkWellOrder_Click()
If chkWellOrder.Value = 1 Then
  ChkRankingOrder.Value = 0
  End If
  
End Sub

Private Sub cmdApplyRevise_Click()
'Save the current record.
If Trim(txtOut.Text) = "" Then
    Beep
    MsgBox "Choose a output file"
    Exit Sub
Else
   Dim pos As Integer
   pos = InStr(txtOut.Text, ".")
   If pos > 0 Then
      txtOut.Text = Left(txtOut.Text, pos - 1)
      If Trim(txtOut.Text) = "" Then
        Beep
        MsgBox "Wrong output file"
        Exit Sub
      End If
   End If
End If

SaveCurrentRecord
'Close the file.
'Close gFileNum
Unload frmReviseCor
End Sub

Private Sub cmdCancel_Click()
Unload frmReviseCor
End Sub

Private Sub cmdOpt_seq_Click()
Dim I, J, k As Integer

If cmdOpt_seq.Caption = "Optimum Sequence" Then

'Add the updated list of events

If Trim(txt_file_hout) = "" Then
   MsgBox "Choose a h.Out file"
   Exit Sub
End If

OpenFile    'open *h.out file and get the list of events
List_events.Clear
Label1.Visible = True

For J = 1 To NumberOfEvents
      List_events.AddItem Space(2 * (4 - Len(Str(Event_Num(J))))) + Str(Event_Num(J)) + Space(5) + Event_name(J)
    Next J

List_events.Visible = True

cmdOpt_seq.Caption = "Add Zone Limits"
Else

k = 0
For I = 0 To List_events.ListCount - 1
     If List_events.Selected(I) Then
       If k = 0 Then
       Txt_Top.Text = Event_Num(I + 1)
       ElseIf k = 1 Then
       txt_base.Text = Event_Num(I + 1)
       Else
       Beep
       MsgBox "You should select only two limits"
       Exit Sub
       End If
      k = k + 1
    End If
 Next I

 cmdOpt_seq.Caption = "Optimum Sequence"
  Label1.Visible = False
 ' update the events list so that only event occur on wells are displayed
 List_events.Clear
 List_events.Visible = False
 End If


End Sub

Private Sub cmdOut_Click()
cboOut.AddItem "Par files (*.Par)"
cboOut.ListIndex = 0
FileOut.Visible = True
End Sub

Private Sub FileOut_Click()
txtOut.Text = FileOut.Filename
FileOut.Visible = False
End Sub

Private Sub Form_Load()
Txt_Top.Text = ""       'set default to text top
txt_base.Text = ""      'set default to text base
txtMinWell.Text = ""    'minimum number of wells

'Display the current record.
ShowCurrentRecord

End Sub

Public Sub ShowCurrentRecord()

Dim I, J As Integer
Dim gFileNum As Integer
Dim Temp As String

gFileNum = FreeFile

If Dir(txt_file_inp) = "" Then
    If InStr(txt_file_inp, ".") <= 0 Then
       txt_file_inp = txt_file_inp & ".par"
    End If
    Open CurDir & "\" & txt_file_inp For Output As gFileNum
    Print #gFileNum, " " & "1 1   1    1"
    Close gFileNum
End If

Open CurDir + "\" + txt_file_inp For Input As gFileNum

'Get the five parameters from input file .par
If EOF(gFileNum) = False Then
  Input #gFileNum, Temp
 
  Scaling = Val(Mid(Temp, 1, 1))
  RankOrder = Val(Mid(Temp, 2, 2))
  Opt_top = Val(Mid(Temp, 4, 4))
  Opt_Base = Val(Mid(Temp, 8, 4))
  MinWells = Val(Mid(Temp, 12, 2))
Else
    MsgBox "File " + txt_file_inp + " is empty"
End If
  
  Close gFileNum

Txt_Top.Text = Str(Opt_top)
txt_base.Text = Str(Opt_Base)
txtMinWell.Text = Str(MinWells)
ChkScaling.Value = Scaling
ChkRankingOrder.Value = RankOrder

End Sub



Public Sub SaveCurrentRecord()
'Fill gCascInput with the currently displayed data
Dim I As Integer
Dim gFileNum As Integer

gFileNum = FreeFile

If Trim(txtOut.Text) <> "" Then
   Open CurDir + "\" + txtOut.Text + ".par" For Output As gFileNum
Else
Beep
  MsgBox "Choose a output file"
Exit Sub
End If

Scaling = ChkScaling.Value
   If chkWellOrder.Value = 1 Then
   RankOrder = 2
   Else
    RankOrder = ChkRankingOrder.Value
   End If
MinWells = Trim(txtMinWell.Text)
Opt_top = Trim(Txt_Top.Text)
Opt_Base = Trim(txt_base.Text)

'Save the parameters into a new par file

' To format the output of Top and Base limits
If Opt_top < 10 Then
    Top_space = "   "
 ElseIf Opt_top < 100 Then
    Top_space = "  "
 Else
    Top_space = " "
 End If
 
 If Opt_Base < 10 Then
    Base_space = "   "
 ElseIf Opt_Base < 100 Then
    Base_space = "  "
 Else
    Base_space = " "
 End If
 
Print #gFileNum, " "; Trim(Scaling); " "; Trim(RankOrder); _
      Top_space; Trim(Opt_top); Base_space; Trim(Opt_Base); _
      Trim(MinWells)
Close gFileNum

End Sub


' function to open h.out file and get the list of events

Public Sub OpenFile()

Dim gFileNum As Integer

Dim I As Integer
Dim Temp_Num As Double
Dim Num_temp As Integer
Dim Temp As String

gFileNum = FreeFile
If Dir(txt_file_hout) = "" Then
Exit Sub
End If

 Open CurDir + "\" + txt_file_hout For Input As #gFileNum
'skip four lines
If EOF(gFileNum) Then
Beep
  MsgBox "The h.out file is Empty"
Exit Sub
End If

Input #gFileNum, Temp
Input #gFileNum, Temp
Input #gFileNum, Temp
Input #gFileNum, Temp

Input #gFileNum, NumberOfEvents
'    MsgBox Str(NumberOfEvents)
ReDim Event_Num(1 To NumberOfEvents)
ReDim Event_name(1 To NumberOfEvents)

    
If ChkScaling.Value = 1 Then    ' skip lines of Scaling = 1
    For I = 1 To 2 * NumberOfEvents + 6
    Line Input #gFileNum, Temp
    Next I
End If
    
' read in the list
 If ChkScaling.Value = 1 Then
   For I = 1 To NumberOfEvents
    Input #gFileNum, Event_Num(I), Temp_Num, Event_name(I), Temp
   Next I
 Else
     For I = 1 To NumberOfEvents
    Input #gFileNum, Num_temp, Event_Num(I), Num_temp, Num_temp, Event_name(I)
   Next I
 End If
 
Close gFileNum

End Sub
