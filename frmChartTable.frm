VERSION 5.00
Begin VB.Form frmChartTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data"
   ClientHeight    =   3960
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmChartTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   1890
      TabIndex        =   2
      Top             =   30
      Width           =   1065
   End
   Begin VB.CommandButton CmdCopytoClipboard 
      Caption         =   "Copy to clipborad"
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   1755
   End
   Begin VB.ListBox listOfData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   330
      Width           =   7605
   End
End
Attribute VB_Name = "frmChartTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCopytoClipboard_Click()
   Dim I As Integer
   Dim Temp As String
   
   Temp = ""
   Clipboard.Clear
   For I = 1 To frmChartTable.listOfData.ListCount
       Temp = Temp + frmChartTable.listOfData.List(I - 1) + Chr$(13) + Chr$(10)
   Next I
   Clipboard.SetText Temp
End Sub

Private Sub Form_Activate()
    If ChartTableLoader = 1 Then
        Me.Caption = "Data - CASC Scattergram - Ranking"
    End If
    If ChartTableLoader = 2 Then
        Me.Caption = "Data - Depth Differences Frequency Histogram"
    End If
    If ChartTableLoader = 3 Then
        Me.Caption = "Data - Scattergram - Ranking"
    End If
    If ChartTableLoader = 4 Then
        Me.Caption = "Data - Depth Differences Q-Q Plot"
    End If
    If ChartTableLoader = 5 Then
        Me.Caption = "Data - Cumulative Depth Differences"
    End If
    If ChartTableLoader = 6 Then
        Me.Caption = "Data - Transformed Depth Difference Plot"
    End If
    If ChartTableLoader = 7 Then
        Me.Caption = "Data - CASC Scattergram - Scaling"
    End If
    If ChartTableLoader = 8 Then
        Me.Caption = "Data - Scattergram - Scaling"
    End If


End Sub

Private Sub Form_Resize()
'    listOfData.left = 0
'    listOfData.top = 480
'    If Me.Width - 20 > 0 Then
'        listOfData.Width = Me.Width - 100      'old: 20
'    End If
'    If Me.Height - 500 > 0 Then
'        listOfData.Height = Me.Height - 100    'old: 500
'    End If
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = 1
  If ChartTableLoader = 1 Then
      frmscatterDE1.Check1.Value = 0
  End If
  If ChartTableLoader = 2 Then
      frmDiffHistogram.Check1.Value = 0
  End If
  If ChartTableLoader = 3 Then
      frmscatterSC1.Check1.Value = 0
  End If
  If ChartTableLoader = 4 Then
      frmDepthDiffQQplot.Check1.Value = 0
  End If
  If ChartTableLoader = 5 Then
      frmFirstOrderDepthDiff.Check1.Value = 0
  End If
  If ChartTableLoader = 6 Then
      frmTransDepthDiffQQplot.Check1.Value = 0
  End If
  If ChartTableLoader = 7 Then
      frmscatterDE2.Check1.Value = 0
  End If
  If ChartTableLoader = 8 Then
      frmscatterSC2.Check1.Value = 0
  End If
  
  Me.Hide
End Sub

