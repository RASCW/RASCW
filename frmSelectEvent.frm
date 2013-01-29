VERSION 5.00
Begin VB.Form frmSelectevent 
   Caption         =   "Delete Events from this Run "
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   Icon            =   "frmSelectEvent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Sort Dictionary "
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   7440
      Width           =   3492
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   480
      Width           =   5148
   End
   Begin VB.CommandButton cmdbackwd 
      Caption         =   "<<"
      Height          =   252
      Left            =   5520
      TabIndex        =   5
      Top             =   4320
      Width           =   372
   End
   Begin VB.CommandButton CmdForward 
      Caption         =   ">>"
      Height          =   252
      Left            =   5520
      TabIndex        =   4
      Top             =   3480
      Width           =   372
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      ItemData        =   "frmSelectEvent.frx":0442
      Left            =   6120
      List            =   "frmSelectEvent.frx":0444
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Width           =   5150
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   9480
      TabIndex        =   2
      Top             =   7680
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   252
      Left            =   7440
      TabIndex        =   1
      Top             =   7680
      Width           =   1212
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   5148
   End
   Begin VB.ListBox List3 
      Height          =   1620
      Left            =   5400
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Deleted events"
      Height          =   372
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "Dictionary "
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "frmSelectevent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num_Select() As Integer

 

Private Sub Check1_Click()
 
Dim I As Integer
     
If Check1.Value = 1 Then
     List1.Visible = False
     List4.Refresh
Else
     List1.Visible = True
        
     List1.Refresh
End If

End Sub

Private Sub cmdbackwd_Click()

Dim I As Integer
    For I = 0 To List2.ListCount - 1
    If List2.Selected(I) = False Then
    List3.AddItem List2.List(I)
   ' Else
   ' List1.AddItem List2.List(I)
    End If
    Next I
  '  List1.Refresh
    List2.Clear
    
    For I = 0 To List3.ListCount - 1
    List2.AddItem List3.List(I)
    Next I
    List3.Clear
 
End Sub

Private Sub cmdCancel_Click()
Del_Event_Total = 0
Unload frmSelectevent
End Sub

 

Private Sub CmdForward_Click()
Dim I As Integer
   For I = 0 To List1.ListCount - 1
    If Check1.Value = 0 Then
      If List1.Selected(I) = True Then
      List2.AddItem List1.List(I)
      End If
    ElseIf Check1.Value = 1 Then
      If List4.Selected(I) = True Then
      List2.AddItem List4.List(I)
      End If
    End If
   Next I
    
End Sub

Private Sub cmdOK_Click()
Dim I As Integer
  Del_Event_Total = List2.ListCount
If Del_Event_Total > 0 Then
  ReDim Del_Event_Names(1 To Del_Event_Total)
  ReDim Del_Event_Numbers(1 To Del_Event_Total)
  
  For I = 0 To Del_Event_Total - 1
     List2.ListIndex = I
     Del_Event_Names(I + 1) = Mid(Trim(List2.Text), 1, 40)
     Del_Event_Numbers(I + 1) = Int(Right(Trim(List2.Text), 4))
 
  Next
End If
  
Unload frmSelectevent
  
End Sub
 
 

Private Sub Form_Load()
Dim rstMy As Recordset
Dim dbsMy As Database
Dim NumOfrecord As Integer
Dim I As Integer
Dim tempName As String
  Del_Event_Total = 0
  
  List1.Clear
  List4.Clear
  
  Set dbsMy = OpenDatabase(myDBName)
    With dbsMy
        ' Open table-type Recordset and show RecordCount
        ' property.
     Set rstMy = .OpenRecordset("Event_name")
    
       If rstMy.EOF = False Then
          rstMy.MoveLast
          NumOfrecord = rstMy.RecordCount
        Else
         MsgBox "Dictonary File is Empty!"
         Num_Edit_Well = 0
       Exit Sub
       End If
   End With
   
   rstMy.MoveFirst
   I = 0
   Do While rstMy.EOF = False
  
   tempName = Trim(rstMy.Fields(0).Value) + Space(42 - Len(Trim(rstMy.Fields(0).Value))) + "    " _
                 + Str(rstMy.Fields(1).Value)
   List1.AddItem tempName, I
  List4.AddItem tempName ', I
   I = I + 1
   rstMy.MoveNext
   Loop
   List1.Refresh
   List4.Refresh
 
   dbsMy.Close
     
End Sub


