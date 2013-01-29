VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDataTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASC Data Table - Ranking"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frmDataTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDG1 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8493
      _Version        =   393216
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim I, J, k As Integer
   
   With MSFlexGrid1
      .Cols = max_well_num + 1
      .Rows = max_event_num * 7 + 1
    
      For I = 0 To max_well_num 'add column headers
         .Col = I
         .row = 0
         .CellAlignment = 1
         .RowHeight(0) = 500
         .CellBackColor = RGB(255, 255, 100)
         If I = 0 Then
         .Text = Trim(Table_Title)
         Else
         .Text = Trim(well_label_used(I))
         End If
      Next I
      
        
    For J = 1 To max_event_num
         k = (J - 1) * 7
       For I = 0 To max_well_num
         .Col = I
         .row = k + 1
        .CellAlignment = 1
          .RowHeight(k + 1) = 400
         .CellBackColor = RGB(255, 255, 150)
           If I = 0 Then
            .Text = Str(Trim(event_numbers(J))) + "   " + Trim(event_names(J))
           ElseIf (I = 1) And (max_well_num > 0) Then
             .ColWidth(1) = 1500
            .Text = "SD/Ave SD = " + Str(Trim(TableRatio(J)))
           End If
           
          .row = k + 2
         .CellAlignment = 1
           If I = 0 Then
             If J = 1 Then
             .Text = "Observed Depth (all depths in m)"
             Else
            .Text = "Observed Depth"
             End If
           Else
              If TableObs(I, J) > -1 Then
            .Text = Str(Trim(TableObs(I, J)))
              End If
              
           End If
          
          .row = k + 3
         .CellAlignment = 1
           If I = 0 Then
            .Text = "Min Observed Depth"
           Else
                If TableMinObs(I, J) > -1 Then
            .Text = Str(Trim(TableMinObs(I, J)))
                End If
           End If
          
          .row = k + 4
         .CellAlignment = 1
           If I = 0 Then
            .Text = "Max Observed Depth"
           Else
            If TableMaxObs(I, J) > -1 Then
            .Text = Str(Trim(TableMaxObs(I, J)))
            End If
           End If
          
          .row = k + 5
         .CellAlignment = 1
           If I = 0 Then
            .ColWidth(0) = 2500
            .Text = "Probable Depth"
           Else
            If TableProb(I, J) > -1 Then
            .Text = Str(Trim(TableProb(I, J)))
            End If
           End If
          
          .row = k + 6
         .CellAlignment = 1
           If I = 0 Then
            .Text = "Min Probable Depth"
           Else
              If TableMinProb(I, J) > -1 Then
               .Text = Str(Trim(TableMinProb(I, J)))
              End If
           End If
          
          .row = k + 7
         .CellAlignment = 1
           If I = 0 Then
            .Text = "Max Probable Depth"
           Else
             If TableMaxProb(I, J) > -1 Then
            .Text = Str(Trim(TableMaxProb(I, J)))
              End If
           End If
               
        Next I
      Next J
    End With
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

       frmwellPlot1.chkTable.Value = 0

End Sub

Private Sub MSFlexGrid1_DblClick()
Dim response

response = MsgBox("Do you want to print the table?", vbYesNo)
If response = vbYes Then    ' User chose Yes.

CDG1.ShowPrinter

frmDataTable.PrintForm
End If

End Sub
