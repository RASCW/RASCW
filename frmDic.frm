VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dictionary"
   ClientHeight    =   5250
   ClientLeft      =   5310
   ClientTop       =   1155
   ClientWidth     =   6315
   Icon            =   "frmDic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6315
   Begin VB.TextBox Filename 
      BackColor       =   &H8000000F&
      Height          =   255
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1725
   End
   Begin RichTextLib.RichTextBox RichTextDic 
      Height          =   4695
      Left            =   30
      TabIndex        =   3
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8281
      _Version        =   393217
      BackColor       =   14737632
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   3
      TextRTF         =   $"frmDic.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   288
      Left            =   240
      TabIndex        =   2
      Top             =   84
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   252
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   252
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "frmDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim textToSearch As String * 40
Dim MyPos As String * 40

Private Sub cmdCancel_Click()
    'RichTextDic.Text = ""
    RichTextDic.Refresh
    Unload frmDic
End Sub

Private Sub CmdHelp_Click()
   MsgBox "Select a Dic file"
End Sub

Private Sub cmdSearch_Click()
 Dim FoundPos As Integer
 Dim FoundLine As Integer
    ' Find the text specified in the TextBox control.
   FoundPos = RichTextDic.Find(txtSearch.Text, 1, , (rtfWholeWord = 1 Or rtfWholeWord = 0) _
           And (rtfNoHighlight = 6))

    ' Show message based on whether the text was found or not.

    'If FoundPos <> -1 Then
        ' Returns number of line containing found text.
     '   FoundLine = RichTextDic.GetLineFromChar(FoundPos)
     '   MsgBox "Word found on line " & CStr(FoundLine)
    'Else
     '   MsgBox "Word not found."

    'End If
  txtSearch.Text = ""
End Sub



Private Sub Form_Activate()
        CurWindowNum = 26
        CurWindowSetFocus (CheckExistWindows(CurWindowNum))
End Sub

Private Sub Form_GotFocus()
      ' ShowWindow MDIFrmCascRasc.hwnd, SW_SHOW
End Sub

Private Sub Form_Load()
    Dim FileOut As String
    On Error GoTo errhandler
    Dim FileNum As Integer
    
    'FileOut =Trim(frmRascW.txt_Out.Text)
    '****************
       'Unload frmRascW
       
      FileOut = txt_file_dic
      FileNum = FreeFile
        
        'If Dir(FileOut + ".all") = "" Then
        If Dir(CurDir & "\" + FileOut) = "" Then
            If Trim(FileOut) = "" Then
                 MsgBox "Dictionary filename is empty. " + CurDir & "\" + FileOut + "Dictionary is not available." + Chr$(13) + "Hint: Select a name for datafile in Run RASC Dialog Window. "
            Else
                  MsgBox "Dictionary file " + CurDir & "\" + FileOut + " is not available." + Chr$(13) + "Please check it and try again. "
            End If
                 Exit Sub
        End If
        
          
        Open CurDir & "\" + FileOut For Input As FileNum
           RichTextDic.Text = Input(LOF(FileNum), FileNum)
        Close FileNum
        Filename.Text = "File: " + FileOut
        
        CurWindowNum = 26
        Call MDIWindowsMenuAdd(CurWindowNum)
        WindowsHwnd(CurWindowNum) = Me.hwnd

                
      Exit Sub
errhandler:
      MsgBox "Data file isn't in right format(ASC)", vbCritical, "File Error !"
End Sub


Private Sub Form_Unload(Cancel As Integer)
        CurWindowNum = 26
        Call MDIWindowsMenuDelete(CurWindowNum)
       WindowsHwnd(CurWindowNum) = -1
End Sub
