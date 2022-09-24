Attribute VB_Name = "RoNzModule"
Option Explicit

'*****************************
'Deklarasi Fungsi Icon Systray
'*****************************
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
  cbSize As Long           ' size of the structure
  hWnd As Long             ' the handle of the window
  uID As Long              ' an unique ID for the icon
  uFlags As Long           ' flags(see below)
  uCallbackMessage As Long ' the Msg that call back when a user do something to the icon
  hIcon As Long            ' the memory location of the icon
  szTip As String * 64     ' tooltip max 64 characters
End Type

Private Const NIM_ADD = &H0      ' add an icon to the system tray
Private Const NIM_MODIFY = &H1   ' modify an icon in the system tray
Private Const NIM_DELETE = &H2   ' delete an icon in the system tray
Private Const NIF_MESSAGE = &H1  ' whether a message is sent to the window procedure for events
Private Const NIF_ICON = &H2     ' whether an icon is displayed
Private Const NIF_TIP = &H4      ' tooltip availibility

'*************************
'Deklarasi Fungsi Registry
'*************************
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As _
        String, ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
    As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_WRITE = &H20006
Private Const REG_SZ = 1

'**********************
'Deklarasi Fungsi Sound
'**********************
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'***********************
'Deklarasi Window Region
'***********************
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'*********************************
'Deklarasi Posisi Window dan Mouse
'*********************************
Public Type POINTAPI
    X               As Long                ' Position X
    Y               As Long                ' Position Y
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'***********
'Sub Program
'***********
Public Function Nama_Aplikasi() As String

    Nama_Aplikasi = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
End Function

Public Function Versi_Aplikasi() As String

    Versi_Aplikasi = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
End Function

'************
'Registry Run
'************
Public Sub Jalankan_Saat_Startup()

    Dim hregkey, retval As Long
    Dim subkey, stringbuffer As String

    subkey = "Software\Microsoft\Windows\CurrentVersion\Run"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_WRITE, hregkey)
    stringbuffer = App.Path & "\" & App.EXEName & ".exe" & vbNullChar
    retval = RegSetValueEx(hregkey, App.Title, 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer))
    
    RegCloseKey hregkey
    
End Sub

'*********************
'Bikin icon di systray
'*********************
Public Sub Buat_Icon_Systray(ByVal FormGue As Form)
    
    Dim nid As NOTIFYICONDATA
    
    With nid
        .cbSize = Len(nid) 'Ukuran struktur
        .hWnd = FormGue.hWnd 'memory location(handle) for the processor of its message and icon
        .uID = 0 'ID yg unik untuk icon. Harus berbeda dari icon systray yang lain
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'notify of message, display icon, display tooltip
        .uCallbackMessage = 1400 'message used to be notified when there's event. any number greater than 1300 will do
        .hIcon = FormGue.Icon 'assign the icon to the form's icon
        .szTip = App.Title & " Systray Icon" & vbNullChar 'terminate the string with vbNullChar or Chr(0)
    End With
  
    Shell_NotifyIconA NIM_ADD, nid
    Oldproc = SetWindowLongA(FormGue.hWnd, -4, AddressOf proc)
    
End Sub

'*********************
'Hapus icon di systray
'*********************
Public Sub Hapus_Icon_Systray(ByVal FormGue As Form)

    Dim nid As NOTIFYICONDATA
    With nid
        .cbSize = Len(nid)
        .hWnd = FormGue.hWnd
        .uID = 0
    End With
  
    Shell_NotifyIconA NIM_DELETE, nid 'hapus icon dari systray
    SetWindowLongA FormGue.hWnd, -4, Oldproc 'kembalikan prosedur window ke semula
    
End Sub

'************
'Bikin Region
'************
Public Function BikinRegion(ByVal FormGue As Form, UkuranKontrol As Control, RoundX, RoundY As Integer) As Long
    
    FormGue.ScaleMode = 3 'Pixel
    BikinRegion = SetWindowRgn(FormGue.hWnd, CreateRoundRectRgn(1, 1, UkuranKontrol.Width, UkuranKontrol.Height, RoundX, RoundY), True)
    
End Function

'************
'PlaySound
'************

Public Sub PlaySound(ByVal NamaWave As String) 'File Wave harus diletakkan di folder \wave

    sndPlaySound App.Path & "\Sounds\" & NamaWave, 1

End Sub

Public Function proc&(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)

    ' 513 -- left button down
    ' 514 -- left button up
    ' 515 -- left button double click
    ' 516 -- right button down
    ' 517 -- right button up
    ' 518 -- right button double click
    ' 519 -- middle button down ( mouse yg memiliki tombol tengah )
    ' 520 -- middle button up
    ' 521 -- middle button double click
    
    If Msg = 1400 And lParam = 517 Then
        MainForm.PopupMenu MainForm.mnu
    ElseIf Msg = 1400 And lParam = 514 Then
        If ClientFormHide = True Then
            ClientForm.Show
        Else
            Unload ClientForm
        End If
    End If

    proc = CallWindowProcA(Oldproc, hWnd, Msg, wParam, lParam)

End Function

