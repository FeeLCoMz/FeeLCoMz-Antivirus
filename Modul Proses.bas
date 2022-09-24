Attribute VB_Name = "ModulProses"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_Maximize = 3
Public Const SW_Minimize = 6
Public Const SW_Normal = 1
Public Const SW_SHOW = 5
Public Const WM_CLOSE = &H10

Sub WindowHandle(hWindow, mCase As Long)
Dim X As Long
Select Case mCase
    Case 0
        X = SendMessage(hWindow, WM_CLOSE, 0, 0)
    Case 1
        X = ShowWindow(hWindow, SW_SHOW)
    Case 2
        X = ShowWindow(hWindow, SW_HIDE)
    Case 3
        X = ShowWindow(hWindow, SW_Maximize)
    Case 4
        X = ShowWindow(hWindow, SW_Minimize)
    Case 5
        X = ShowWindow(hWindow, SW_Normal)
End Select

End Sub

Public Function GetWindowTitle(ByVal hWnd As Long) As String

    Dim L As Long
    Dim S As String
    
    On Error Resume Next
    
    L = GetWindowTextLength(hWnd)
    S = Space(L + 1)

    GetWindowText hWnd, S, L + 1
    GetWindowTitle = Left$(S, L)
    
End Function

Public Sub ShowDaftarWindow(vListView As ListView)

    Dim RefreshD As Boolean
    Dim I, z, AppCap As Integer
    Dim WinTitle As String
    Dim hW As Long
    
    With vListView
        .Checkboxes = False
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ListItems.Clear
        
        With .ColumnHeaders
            .Clear
            .Add 1, , "No.", 500
            .Add 2, , "hWnd", 700
            .Add 3, , "Status", 900
            .Add 4, , "Window Title", vListView.Width - (.Item(1).Width + .Item(2).Width + .Item(3).Width)
        End With

    End With
    
    For I = 1 To 10000
        DoEvents
        WinTitle$ = GetWindowTitle(I)
        z = FindWindow(vbNullString, WinTitle$)
        If z <> 0 Then
            If WinTitle$ <> vbNullString And LCase(WinTitle$) <> LCase(AppCap) Then
            '***
                If (IsWindowEnabled(z) = 0) And (IsWindowVisible(z) = 0) Then
                        With vListView.ListItems
                            .Add.Text = Format(.Count, "00")
                            .Item(.Count).SubItems(1) = I
                            .Item(.Count).SubItems(2) = ""
                            .Item(.Count).SubItems(3) = WinTitle$
                        End With
                ElseIf (IsWindowEnabled(z) = 1) And (IsWindowVisible(z) = 0) Then
                        With vListView.ListItems
                            .Add.Text = Format(.Count, "00")
                            .Item(.Count).SubItems(1) = I
                            .Item(.Count).SubItems(2) = "Active"
                            .Item(.Count).SubItems(3) = WinTitle$
                        End With
                ElseIf (IsWindowEnabled(z) = 1) And (IsWindowVisible(z) = 1) Then
                        With vListView.ListItems
                            .Add.Text = Format(.Count, "00")
                            .Item(.Count).SubItems(1) = I
                            .Item(.Count).SubItems(2) = "Visible"
                            .Item(.Count).SubItems(3) = WinTitle$
                        End With
                End If
            '***
            End If
        End If
    Next I
    
End Sub

