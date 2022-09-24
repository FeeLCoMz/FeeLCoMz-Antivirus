VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FEELCOMZ ANTIVIRUS"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameVDB 
      Caption         =   "Virus Database"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8655
      Begin MSComctlLib.ListView LVirus 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   8655
      Begin VB.CommandButton CmdHide 
         Caption         =   "HIDE"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdProses 
         Caption         =   "PROSES"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdScan 
         Caption         =   "SCAN"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdKeluar 
         Appearance      =   0  'Flat
         Caption         =   "KELUAR"
         Height          =   375
         Left            =   7560
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   8655
      Begin MSComctlLib.ListView LLog 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdHide_Click()
    Me.Visible = False
End Sub

Private Sub CmdProses_Click()
    FProses.Show
End Sub

Private Sub CmdScan_Click()
    Dim I As Integer
    
    LLog.ListItems.Clear
    
    With LVirus.ListItems
        For I = 1 To .Count
            DoEvents
            LLog.ListItems.Add.Text = .Item(I).Text
            If BunuhVirus(.Item(I).SubItems(1)) > 0 Then
                LLog.ListItems.Item(I).SubItems(1) = "Virus telah dibasmi!"
            Else
                LLog.ListItems.Item(I).SubItems(1) = "Bersih"
            End If
        Next I
    End With
End Sub

Private Sub Form_Load()
    With LVirus
        .ColumnHeaders.Add 1, , "Nama Virus", 2000
        .ColumnHeaders.Add 2, , "Teks Window Virus", 3000
        .ColumnHeaders.Add 3, , "File Induk Virus", .Width - 5050
    End With
    
    With LLog
        .ColumnHeaders.Add 1, , "Nama Virus", 2000
        .ColumnHeaders.Add 2, , "Status", .Width - 2050
    End With
    
    TampilDaftarVirus
    
    frameVDB.Caption = "Virus Database (Total: " & LVirus.ListItems.Count & " Viruses)"
    
End Sub

Private Sub TampilDaftarVirus()
    Dim FileNum As Long
    Dim I, j As Integer
    Dim VirContent As String
    Dim VirusData As Virus
    Dim LFile
    
    I = 1
    j = 0
    FileNum = FreeFile

    Open App.Path & "\VirDef.txt" For Input As #FileNum
    Do Until EOF(FileNum)
        DoEvents
        Input #FileNum, VirContent
        With VirusData
            .Nama = Split(VirContent, "|")(0)
            .TeksWindow = Split(VirContent, "|")(1)
            .FileMaster = ""
            If UBound(Split(VirContent, "|")) > 1 Then
                .FileMaster = Split(VirContent, "|")(2)
            End If
        End With
        With LVirus.ListItems
            .Add.Text = VirusData.Nama
            .Item(I).SubItems(1) = VirusData.TeksWindow
            If VirusData.FileMaster <> "" Then
                For j = 0 To UBound(Split(VirusData.FileMaster, "?"))
                    .Item(I).SubItems(2) = Split(VirusData.FileMaster, "?")(j)
                Next j
            End If
        End With
        I = I + 1
    Loop
    Close #FileNum
    
End Sub
Private Sub CmdKeluar_Click()
    End
End Sub
