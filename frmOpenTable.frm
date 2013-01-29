VERSION 5.00
Begin VB.Form frmOpenTable 
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmOpenTable.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File_ran1 
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   252
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   732
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   252
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtRan1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.ComboBox cboRan1 
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "frmOpenTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRan1_Click()
    'File_ran1.Pattern = "*.sum"
    If tableType = 1 Then
    File_ran1.Pattern = "*.sum"
    ElseIf tableType = 2 Then
    File_ran1.Pattern = "*.fl0"
    ElseIf tableType = 3 Then
    File_ran1.Pattern = "*.wel"
    ElseIf tableType = 4 Then
    File_ran1.Pattern = "*.occ"
    ElseIf tableType = 5 Then
    File_ran1.Pattern = "*.pen"
    ElseIf tableType = 6 Then
    File_ran1.Pattern = "*.nor"
    ElseIf tableType = 7 Then
    File_ran1.Pattern = "*.cyc"
    End If
End Sub

Private Sub cmdBrowse_Click()
        If tableType = 1 Then
        cboRan1.AddItem "Ran files (*.sum)"
        ElseIf tableType = 2 Then
        cboRan1.AddItem "Err files (*.fl0)"
        ElseIf tableType = 3 Then
        cboRan1.AddItem "Wel files (*.wel)"
        ElseIf tableType = 4 Then
        cboRan1.AddItem "Occ files (*.occ)"
        ElseIf tableType = 5 Then
        cboRan1.AddItem "Pen files (*.pen)"
        ElseIf tableType = 6 Then
        cboRan1.AddItem "Nor files (*.nor)"
        ElseIf tableType = 7 Then
        cboRan1.AddItem "Cyc files (*.cyc)"
        End If
        cboRan1.ListIndex = 0
        File_ran1.Visible = True
End Sub

Private Sub cmdCancel_Click()
   Unload frmOpenTable
End Sub

Private Sub cmdOpen_Click()
If txtRan1.Text <> "" Then
    txt_file_table = txtRan1.Text
 '   Check = txt_file_ran1
    
    Unload frmOpenTable
Else
    Beep
End If
End Sub

Private Sub File_ran1_Click()
    txtRan1.Text = File_ran1.Filename
    File_ran1.Visible = False
End Sub

