Attribute VB_Name = "ModulUtama"
Option Explicit

Type Virus
    Nama As String
    TeksWindow As String
    FileMaster As String
End Type

'*** Text Format ***
Public Function FormatReport(Field As String, Data As String)
    FormatReport = Field & vbTab & ": " & Data & vbCrLf
End Function

Public Function BunuhVirus(NamaProsesVirus As String) As Integer
    
    Do While FindWindow(vbNullString, NamaProsesVirus) <> 0
        DoEvents
        WindowHandle FindWindow(vbNullString, NamaProsesVirus), 0
        BunuhVirus = BunuhVirus + 1
    Loop
    
End Function



