
' ------------------------------------------------------------
' 1. **Open the VBA editor** in Excel by pressing `Alt + F11`.
' 2. **Insert a new module** by right-clicking on any of the existing modules or the VBA project, then selecting `Insert > Module`.
' 3. **Copy and paste the following code** into the new module:

' ------------------------------------------------------------
' Microsoft Copilot:
' 1: Please give me an example of a Windows Excel vba macro. I want to maximize another window.
' 2: I want to maximize that window even when it is already maximized.

' --------------------
' Here's an example of a VBA macro that maximizes another window in Excel:
'
' Replace `"Title of the Window"` with the exact title of the window you want to maximize.
' This macro uses Windows API functions to find and maximize the specified window.

' This code will minimize the window first and then maximize it,
' ensuring it is maximized regardless of its initial state.

' ------------------------------------------------------------
' Declare the necessary Windows API functions
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' Constants for the ShowWindow function
Private Const SW_MINIMIZE As Long = 6
Private Const SW_MAXIMIZE As Long = 3


Sub MaximizeWindow_ww_O()
' CTRL + O(rders) - work orders (radni nalozi)
' date: 2024-12M-13 16:10:05
' EDIT:
'
' I actually use this macro every day as part of
' https://github.com/mlabrkic/my_office_excel_app

    Dim hwnd As Long
    Dim windowTitle As String
    Dim result As Long

    ' Set the title of the window you want to maximize
    ' windowTitle = "Title of the Window"
    ' windowTitle = "Untitled - Notepad" ' Change this to the title of your application window
    windowTitle = "Traženje svih kartica"

    ' Find the window handle based on the title
    hwnd = FindWindow(vbNullString, windowTitle)

    If hwnd <> 0 Then
        ' Minimize the window first
        result = ShowWindow(hwnd, SW_MINIMIZE)

        ' Then maximize the window
        result = ShowWindow(hwnd, SW_MAXIMIZE)

        If result = 0 Then
            MsgBox "Failed to maximize the window.", vbExclamation
        Else
            ' MsgBox "Window maximized successfully!", vbInformation
        End If
    Else
        MsgBox "Window not found.", vbExclamation
    End If
End Sub

