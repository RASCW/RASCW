VERSION 5.00
Begin VB.Form frmChartSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Chart"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextFilename 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   1290
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   1680
      TabIndex        =   2
      Top             =   180
      Width           =   3615
      Begin VB.OptionButton Option3 
         Caption         =   "OC2"
         Height          =   345
         Left            =   2550
         TabIndex        =   5
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "JPG"
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.OptionButton Option1 
         Caption         =   "EMF"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      Top             =   2700
      Width           =   1065
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   3330
      TabIndex        =   0
      Top             =   2700
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "File Name:"
      Height          =   285
      Left            =   570
      TabIndex        =   7
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "File Type:"
      Height          =   315
      Left            =   570
      TabIndex        =   6
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "frmChartSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
