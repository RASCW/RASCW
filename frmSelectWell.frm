VERSION 5.00
Begin VB.Form frmSelect 
   Caption         =   "Select Well(s) for Editing"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "frmSelectWell.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6495
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   480
      Width           =   2412
   End
   Begin VB.CommandButton cmdbackwd 
      Caption         =   "<<"
      Height          =   252
      Left            =   2760
      TabIndex        =   5
      Top             =   4080
      Width           =   372
   End
   Begin VB.CommandButton CmdForward 
      Caption         =   ">>"
      Height          =   252
      Left            =   2760
      TabIndex        =   4
      Top             =   3480
      Width           =   372
   End
   Begin VB.ListBox List2 
      Height          =   6105
      ItemData        =   "frmSelectWell.frx":0442
      Left            =   3360
      List            =   "frmSelectWell.frx":0444
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Width           =   2292
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   4680
      TabIndex        =   2
      Top             =   6840
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   252
      Left            =   3360
      TabIndex        =   1
      Top             =   6840
      Width           =   1092
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   3360
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "List of wells"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "Selected wells"
      Height          =   252
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1812
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num_Select() As Integer



Private Sub Command1_Click()

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

'Dim NR As Integer
  
'  Num_Edit_Well = List2.ListCount
'   If Num_Edit_Well < 1 Then Exit Sub
' For NR = 0 To Num_Edit_Well - 1
'    If List2.Selected(NR) = True Then
'    List2.RemoveItem NR
'    End If
' Next NR
End Sub

Private Sub cmdCancel_Click()
Num_Edit_Well = 0
Unload frmSelect
End Sub

Private Sub CmdForward_Click()
Dim I As Integer
    For I = 0 To List1.ListCount - 1
    If List1.Selected(I) Then
    List2.AddItem List1.List(I)
    End If
    Next I
    List1.Refresh
    
' remover from list1

 '   For I = 0 To List1.ListCount - 1
 '   If List1.Selected(I) = False Then
 '   List4.AddItem List1.List(I)
 '   End If
 '   Next I
 '   List1.Clear
    
 '   For I = 0 To List4.ListCount - 1
 '   List1.AddItem List4.List(I)
 '   Next I
 '   List4.Clear
End Sub

Private Sub cmdOK_Click()
Dim I As Integer
  Num_Edit_Well = List2.ListCount
  If Num_Edit_Well < 1 Then Exit Sub
  ReDim Edit_Well_Names(1 To Num_Edit_Well)
  ReDim Edit_Well_Num(1 To Num_Edit_Well)
  
  For I = 0 To Num_Edit_Well - 1
     List2.ListIndex = I
     Edit_Well_Names(I + 1) = Mid(Trim(List2.Text), 5)
     Edit_Well_Num(I + 1) = Int(Trim(Mid(Trim(List2.Text), 1, 5)))
 
  Next
Unload frmSelect
  
End Sub

Private Sub Form_Load()
Dim rstMy As Recordset
Dim dbsMy As Database
Dim NumOfrecord As Integer
Dim I As Integer
Dim tempName As String

  Set dbsMy = OpenDatabase(myDBName)
    With dbsMy
        ' Open table-type Recordset and show RecordCount
        ' property.
     Set rstMy = .OpenRecordset("header")
    
       If rstMy.EOF = False Then
          rstMy.MoveLast
          NumOfrecord = rstMy.RecordCount
        Else
         MsgBox "Well Header File is Empty!"
         Num_Edit_Well = 0
       Exit Sub
       End If
   End With
   
   rstMy.MoveFirst
   I = 0
   Do While rstMy.EOF = False
  
   tempName = rstMy.Fields(0).Value + "    " + rstMy.Fields(1).Value
   List1.AddItem tempName, I
   I = I + 1
   rstMy.MoveNext
   Loop
 
   dbsMy.Close
    
End Sub


