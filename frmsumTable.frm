VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSumTable 
   Caption         =   "Summary Table"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "frmsumTable.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   10485
   Visible         =   0   'False
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2130
      TabIndex        =   3
      Top             =   120
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4980
      Top             =   30
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Print 
      Caption         =   "Print"
      Height          =   252
      Left            =   1170
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
   Begin RichTextLib.RichTextBox RichtxtTable 
      Height          =   4005
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7064
      _Version        =   393217
      BackColor       =   12648447
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmsumTable.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frmSumTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Dev_Value() As Double
Dim Well_Num() As Integer
Dim frequency(1 To 10) As Integer
Dim classfrom(1 To 10), classto(1 To 10) As Double

Dim MaxNumWell As Integer
Dim MaxNumEvent As Integer
Dim NumOfEvent() As Integer
Dim EventName() As String * 50
Dim NumOfWell() As Integer
Dim EventNum As Integer
Dim MaxArray As Integer

Dim Temp As String
Dim FileNum As Integer
Dim FilenameTemp As String


Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    frmOpenTable.Show 1
    'cmdOpen.Enabled = False
    
    OpenFile

End Sub

Private Sub Form_Load()

Me.Caption = TableCaption

Select Case Me.Caption
Case Is = "Summary Table"
    CurWindowNum = 19
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

Case Is = "Warnings"
    CurWindowNum = 20
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

Case Is = "Well/Section Names"
    CurWindowNum = 21
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

Case Is = "Occurrence Table"
    CurWindowNum = 22
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

Case Is = "Penalty Points"
    CurWindowNum = 23
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

Case Is = "Normality Test"
    CurWindowNum = 24
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

Case Is = "Cyclicity"
    CurWindowNum = 25
    Call MDIWindowsMenuAdd(CurWindowNum)
    WindowsHwnd(CurWindowNum) = Me.hwnd

End Select


txt_file_table = ""
cmdOpen.Enabled = True
End Sub




Public Sub OpenFile()
Dim tempfrom As Integer
Dim tempto As Integer
Dim I, J As Integer
Dim title_str As String

Dim ArrayNum As Integer
Dim tnum
Dim FileNum As Integer

FileNum = FreeFile
If txt_file_table <> "" Then
    Open CurDir + "\" + txt_file_table For Input As FileNum
Else
    MsgBox "Input file is not exist, please try again"
   Exit Sub
End If

RichtxtTable.Text = Input(LOF(FileNum), FileNum)
Close FileNum

End Sub


Private Sub Form_Resize()
    RichtxtTable.left = 0
    RichtxtTable.top = 480
    If Me.Width - 20 > 0 Then
        RichtxtTable.Width = Me.Width - 100      'old: 20
    End If
    If Me.Height - 500 > 0 Then
        RichtxtTable.Height = Me.Height - 1000    'old: 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case Me.Caption
    Case Is = "Summary Table"
        CurWindowNum = 19
        Call MDIWindowsMenuDelete(CurWindowNum)
        WindowsHwnd(CurWindowNum) = -1
   
    Case Is = "Warnings"
        CurWindowNum = 20
        Call MDIWindowsMenuDelete(CurWindowNum)
         WindowsHwnd(CurWindowNum) = -1
   
    Case Is = "Well/Section Names"
        CurWindowNum = 21
        Call MDIWindowsMenuDelete(CurWindowNum)
         WindowsHwnd(CurWindowNum) = -1
    
    Case Is = "Occurrence Table"
        CurWindowNum = 22
        Call MDIWindowsMenuDelete(CurWindowNum)
          WindowsHwnd(CurWindowNum) = -1
   
    Case Is = "Penalty Points"
        CurWindowNum = 23
        Call MDIWindowsMenuDelete(CurWindowNum)
          WindowsHwnd(CurWindowNum) = -1
    
    Case Is = "Normality Test"
        CurWindowNum = 24
        Call MDIWindowsMenuDelete(CurWindowNum)
          WindowsHwnd(CurWindowNum) = -1
    
    Case Is = "Cyclicity"
        CurWindowNum = 25
        Call MDIWindowsMenuDelete(CurWindowNum)
          WindowsHwnd(CurWindowNum) = -1
    
    End Select

End Sub

Private Sub Print_Click()
 Dim fontname As String
 Dim fontsize As Double
  CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
   If RichtxtTable.SelLength = 0 Then
      CommonDialog1.Flags = CommonDialog1.Flags + cdlPDAllPages
   Else
      CommonDialog1.Flags = CommonDialog1.Flags + cdlPDSelection
   End If
     On Error GoTo errorhandle
     With CommonDialog1
         .CancelError = True
         .ShowPrinter
         
     End With
     fontname = RichtxtTable.Font.name
     Printer.Font = RichtxtTable.Font
   
    Printer.fontsize = RichtxtTable.Font.size
    RichtxtTable.SelPrint CommonDialog1.hDC  'Printer.hDC
'End Sub
errorhandle:

End Sub
