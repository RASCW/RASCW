VERSION 5.00
Begin VB.Form frmDicInput 
   Caption         =   "Open a Dictionary File"
   ClientHeight    =   1440
   ClientLeft      =   1410
   ClientTop       =   2055
   ClientWidth     =   3150
   Icon            =   "frmDicInput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   3150
   Begin VB.FileListBox File_Dic 
      Height          =   870
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cboDic 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   360
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   732
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   252
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   732
   End
   Begin VB.CommandButton cmdDicInput 
      Caption         =   "Browse"
      Height          =   252
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtDicInput 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmDicInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDic_Click()
    File_Dic.Pattern = "*.Dic"
End Sub

Private Sub cmdCancel_Click()
    Unload frmDicInput

End Sub

Private Sub cmdDicInput_Click()
    cboDic.AddItem "Dic files (*.Dic)"
    cboDic.ListIndex = 0
    File_Dic.Visible = True

End Sub

Private Sub cmdOpen_Click()

If txtDicInput.Text <> "" Then
    txt_file_dic = txtDicInput.Text
    txtDicInput.Text = ""
    Unload frmDicInput
    Unload frmDic  'press the content refresh
    frmDic.Show
Else
    Beep
End If
End Sub

Private Sub File_Dic_Click()
    txtDicInput.Text = File_Dic.Filename
    File_Dic.Visible = False
End Sub

Private Sub Form_Load()
   txtDicInput.Text = ""
End Sub
