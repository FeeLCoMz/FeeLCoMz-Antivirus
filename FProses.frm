VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FProses 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Window List"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FProses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRefresh 
      Appearance      =   0  'Flat
      Caption         =   "REFRESH"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ListView LProses 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
   Begin VB.CommandButton CmdTutup 
      Appearance      =   0  'Flat
      Caption         =   "TUTUP"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "FProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim KlikStat As Boolean
Private Sub CmdRefresh_Click()
    ShowDaftarWindow FProses.LProses
End Sub
Private Sub CmdTutup_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    KlikStat = True
    ShowDaftarWindow FProses.LProses
End Sub
'Private Sub Form_Resize()
'    LProses.Width = Me.ScaleWidth - 230
'    LProses.Height = Me.ScaleHeight - 750
'    CmdTutup.Top = LProses.Height + 250
'    CmdTutup.Left = LProses.Width - CmdTutup.Width + 80
'End Sub
Private Sub LProses_Click()
    If KlikStat = True Then
        'ShowWindow LProses.SelectedItem.SubItems(1), SW_SHOW
        KlikStat = False
    Else
        'ShowWindow LProses.SelectedItem.SubItems(1), SW_HIDE
        KlikStat = True
    End If
End Sub
