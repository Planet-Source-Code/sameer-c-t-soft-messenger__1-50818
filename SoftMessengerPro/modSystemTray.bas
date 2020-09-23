Attribute VB_Name = "modSystemTray"
'-------------------------------------------------------------------------------------------
'Project Name   :   Soft Messenger (SoftMessenger.vbp)
'Started On     :   2001 November 28
'Last Modified  :   2001 December 06
'Version 2003   :   2003 January 21
'Description    :   Message sending using windows's Net Send
'                    and chat using winsock among the networked machines

'Module Name    :   System Tray Module(modSystemTray/modSystemTray.bas)
'Description    :   Procedures for showing the application in system tray
'Developed By   :   Sameer C T
'Last Modified  :   2003 March 28
'-------------------------------------------------------------------------------------------
Option Explicit


Public Declare Function SetForegroundWindow _
Lib "user32" (ByVal hwnd As Long) As Long

'-------------------------------------------------------------------------------------------
'Declaration which referes to image shown in TaskBar
'-------------------------------------------------------------------------------------------
   'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
      Public Type NOTIFYICONDATA
         cbSize As Long
         hwnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type
      
      
      'Declare the constants for the API function. These constants can be
      'found in the header file Shellapi.h.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Public Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.

      'Left-click constants.
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up

      'Declare the API function call.
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
      Public nid As NOTIFYICONDATA
      
    Public Sub gprDeleteSysTrayIcon()
        Dim intRet As Integer
        nid.cbSize = Len(nid)
        nid.hwnd = frmMain.hwnd
        nid.uId = 1&
        intRet = Shell_NotifyIcon(NIM_DELETE, nid)
    End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------


