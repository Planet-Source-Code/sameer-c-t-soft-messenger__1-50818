VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soft Messenger 2003"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "SoftMessenger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "SoftMessenger.frx":0442
   ScaleHeight     =   2595
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1260
      Left            =   5490
      TabIndex        =   19
      Top             =   1260
      Width           =   3720
      Begin VB.Timer tmrSM 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2280
         Top             =   720
      End
      Begin VB.PictureBox picAbout 
         Height          =   555
         Left            =   3045
         ScaleHeight     =   495
         ScaleWidth      =   510
         TabIndex        =   25
         Top             =   405
         Width           =   570
      End
      Begin MSComctlLib.ImageList imgSM 
         Left            =   1260
         Top             =   420
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":074C
               Key             =   "ArrowDN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":0B9E
               Key             =   "ArrowLT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":0FF0
               Key             =   "ArrowRT"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":1442
               Key             =   "ArrowUP"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":1894
               Key             =   "LightON"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":1CE6
               Key             =   "LightOFF"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SoftMessenger.frx":2138
               Key             =   "CheckMark"
            EndProperty
         EndProperty
      End
      Begin MSWinsockLib.Winsock sokChat 
         Left            =   2280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         Caption         =   "This is a free ware"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   23
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         Caption         =   "All rigts reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   705
         Width           =   1245
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Top             =   465
         Width           =   570
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   210
         Width           =   435
      End
   End
   Begin VB.Frame fraControls 
      Height          =   1215
      Left            =   5490
      TabIndex        =   13
      Top             =   15
      Width           =   3720
      Begin VB.CheckBox chkUseTextFile 
         Caption         =   "To list from text file"
         Height          =   225
         Left            =   1965
         TabIndex        =   18
         ToolTipText     =   "Check this to populate the To User List from the text file (SoftMessenger.txt)"
         Top             =   930
         Width           =   1725
      End
      Begin VB.CheckBox chkLoginName 
         Caption         =   "Show login name"
         Height          =   225
         Left            =   1965
         TabIndex        =   17
         ToolTipText     =   "Check this to show the login name as your identity"
         Top             =   562
         Width           =   1725
      End
      Begin VB.CheckBox chkSysTray 
         Caption         =   "&Sys tray minimising"
         Height          =   225
         Left            =   1965
         TabIndex        =   16
         ToolTipText     =   "Check this if minimising to Sytem Tray is required"
         Top             =   195
         Width           =   1725
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh To List"
         Height          =   300
         Left            =   90
         TabIndex        =   15
         ToolTipText     =   "Refreshes the To list with computers in the network or from the text file"
         Top             =   795
         Width           =   1740
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "Include User Names"
         Height          =   225
         Left            =   90
         TabIndex        =   14
         ToolTipText     =   "Check this to include user names in the To list"
         Top             =   195
         Width           =   1785
      End
      Begin MSComctlLib.ProgressBar prbUsers 
         Height          =   210
         Left            =   90
         TabIndex        =   26
         Top             =   495
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   1905
         X2              =   1905
         Y1              =   105
         Y2              =   1185
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1890
         X2              =   1890
         Y1              =   105
         Y2              =   1185
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   1260
      Left            =   75
      TabIndex        =   10
      Top             =   1260
      Width           =   5310
      Begin VB.CommandButton cmdClear 
         Height          =   450
         Left            =   4665
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Clear sent messages(F3)"
         Top             =   435
         Width           =   525
      End
      Begin RichTextLib.RichTextBox txtDetails 
         Height          =   915
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "Sent messages"
         Top             =   210
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   1614
         _Version        =   393217
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"SoftMessenger.frx":258A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraPopup 
      Height          =   1215
      Left            =   75
      TabIndex        =   24
      Top             =   15
      Width           =   5310
      Begin VB.TextBox txtToUser 
         Height          =   315
         Left            =   2130
         TabIndex        =   2
         ToolTipText     =   "Type a valid computer name here"
         Top             =   195
         Width           =   1860
      End
      Begin VB.CommandButton cmdBuzz 
         Caption         =   "&Buzz"
         Height          =   450
         Left            =   4035
         TabIndex        =   5
         ToolTipText     =   "Click this to send a Buzz alarm during chat"
         Top             =   615
         Width           =   525
      End
      Begin VB.CheckBox chkChat 
         Caption         =   "&Chat"
         Height          =   195
         Left            =   4005
         TabIndex        =   4
         ToolTipText     =   "Check this to turn on chat; Otherwise message will be sent as popup"
         Top             =   255
         Width           =   630
      End
      Begin VB.CommandButton cmdDetails 
         Height          =   450
         Left            =   4665
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   645
         Width           =   525
      End
      Begin VB.CommandButton cmdControl 
         Height          =   450
         Left            =   4665
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   525
      End
      Begin VB.TextBox txtMessage 
         Height          =   495
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Type the message here and press Enter(F9)  to send"
         Top             =   600
         Width           =   3885
      End
      Begin VB.ComboBox cboToUser 
         Height          =   315
         Left            =   2130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select the Recipient"
         Top             =   195
         Width           =   1860
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   600
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Your Identity"
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "&From"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   255
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "&To"
         Height          =   195
         Left            =   1905
         TabIndex        =   0
         ToolTipText     =   "Double click here to toggle editing of To machine list and enter only machine name "
         Top             =   255
         Width           =   195
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project Name   :   Soft Messenger (SoftMessenger.vbp)
'Started On     :   2001 November 28
'Last Modified  :   2001 December 06
'Version 3      :   2003 January 21
'Last Modified  :   2003 March 29
'Version 4      :   2003 May 01

'Description    :   Message sending using windows's Net Send
'                    and chat using winsock among the networked machines

'Module Name    :   Main Form(frmMain/SoftMessenger.frm)
'Developed By   :   Sameer C T
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
'Registry Entries
'-------------------------------------------------------------------------------------------
'Key:HKEY_CURRENT_USER\Software\VB and VBA Program Settings\SoftMessenger\
'1.\Settings
'   SysTrayMin      : 1 if minimise to system tray is on; else 0
'   ShowLoginName   : 1 if login name should be displayed as User Name, else 0
'   UseTextFile     : 1 if To List should be populated from a text file, else 0
'2.\UserProfile
'   Name            : Name of the user
'-------------------------------------------------------------------------------------------
'Hidden features
'-------------------------------------------------------------------------------------------
'1.Easter Egg: Click the Show settings and Show detail buttons to reveal the about screen
' Click on the Light picture to toggle it on and off and to reveal author's name as tool tip
'2.Shortcut key: Hit Escape to minimise the application
'3.Message codes: Send 'buzz' to invoke Buzz
'4.Message codes: Send 'clrscr' to clear the recipient's message details
'5.Tip: Click on the To label to make the To Address editable
'-------------------------------------------------------------------------------------------
'History
'-------------------------------------------------------------------------------------------
'2001 November (Version 1)
'Started From Soft Sytems Cochin, to send messages to colleagues sitting in different machines
'and different labs. The conventional method was using net send through command prompt, but
'it was time consuming and tedious. There were some utilities available developed by other
'colleagues, but they were storing the Users list in registry, which made it difficult during
'shifting of computers. So developed this utility with Net Send and chat options. The users
'list was stored in a text file and was possible to edit. The chat section was using UDP and
'provided a safe option for chatting, without being tracked by anyone.
'2003 March (Version 3)
'After more than one year since its development, this tool was already in regular use by many
'of my colleagues in Soft Systems, mainly for sending messages and some for regular chatting.
'But I was not happy with its looks and features and was plaaning to make more compact and
'nice and feature rich. At last I got time for that, when I was working in Kenya, for Sasini
'Tea and Coffe Limited, and changed the whole look into a tiny application. The User list was
'populated from network domain. Chat was done using TCP. Option for Buzz during chat was another
'attraction. System Tray minimising and showing login name were also given in the settings
'section.
'2003 May (Version 4)
'Some machines were not populating from the network domain, and some user names were also not
'displaying properly. So added the option of populating the users from a text file and released
'the final version named as Soft Messenger Pro (Professional Edition), under the banner of
'Sameeriya Soft (the latest name I decided, after Samtech and Sams World).
'2003 July (Version 4.2)
'Again some more modifications from Soft Systems, Cochin, as per the feedback of friends.
'Gave some short cut keys for the operations. Systray icon will change now when a message
'is received during chat. The recently logged in user name will be displayed when a machine
'name is selected from the To user list. Added strict encryption for chat.
'-------------------------------------------------------------------------------------------
Option Explicit
'-------------------------------------------------------------------------------------------
'Win API Declarations
'-------------------------------------------------------------------------------------------
'To get the User Name
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
'To get the Computer Name
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
'For flashing the window if it is minimised
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'-------------------------------------------------------------------------------------------
'Module level constants
'-------------------------------------------------------------------------------------------
Private Const mintMINWIDTH = 5565
Private Const mintMAXWIDTH = 9375
Private Const mintMINHEIGHT = 1680
Private Const mintMAXHEIGHT = 2970
'-------------------------------------------------------------------------------------------
'Module level variables
'-------------------------------------------------------------------------------------------
Private mintSecretCode As Integer
Private mblnControlsShown As Boolean
Private mblnDetailsShown As Boolean
Private mblnLightON As Boolean
Private mblnChatON As Boolean
Private mstrChatMsg As String
Private mintBuzzCounter As Integer
Private mblnBuzzON As Boolean
Private mintPort As Integer
Private mblnToIsEditable As Boolean

'-------------------------------------------------------------------------------------------
'Form: Event handling procedures
'-------------------------------------------------------------------------------------------
Private Sub Form_Load()
    Dim lBuffLen As Long
    Dim sBuffer As String
    Dim lRet As Long
    
    On Error GoTo ErrorTrap

    'Assigning the Title and About details
    gstrMsgTitle = "Soft Messenger"
    lblAbout(0).Caption = "Soft Messenger™ Professional Edition"
    lblAbout(1).Caption = "Version 4.00.02 (110720031245)"
    lblAbout(2).Caption = "Copyright © 2001 - 2003 Sameeriya Soft"
    'lblAbout(3).Caption = "Licensed to Soft Systems Ltd., Cochin"
    lblAbout(3).Caption = "Licensed to You!"
    
    'Exit, If an instance of this application is already running
    If App.PrevInstance = True Then
        MsgBox "Soft Messenger is already running", vbExclamation, gstrMsgTitle
        Unload Me
        Exit Sub
    End If
    
    'Getting the Computer Name
    lBuffLen = 128
    sBuffer = String$(lBuffLen, vbNullChar)
    lRet = GetComputerName(sBuffer, lBuffLen)
    gstrComputerName = Left$(sBuffer, lBuffLen)
  
    'Getting the Login Name
    lBuffLen = 128
    lRet = GetUserName(sBuffer, lBuffLen)
    gstrLoginName = StrConv(Left$(sBuffer, lBuffLen - 1), vbProperCase)
    
    'Displaying the Computer Name with caption
    Me.Caption = gstrMsgTitle & "  [" & gstrComputerName & "]"
    
    'Getting the ShowLoginName setting from registry
    chkLoginName.Value = Val(GetSetting(App.Title, "Settings", "ShowLoginName", 1))
    gblnShowLoginName = chkLoginName.Value
    
    'Getting the UseTextFile setting from registry
    chkUseTextFile.Value = Val(GetSetting(App.Title, "Settings", "UseTextFile", 0))
    gblnUseTextFile = chkUseTextFile.Value
    
    'Getting User Name from registry
    Call mprSetUserName
    
    'Getting the System Tray minimising setting from registry
    chkSysTray.Value = Val(GetSetting(App.Title, "Settings", "SysTrayMin", 1))
    gblnSysTrayMin = chkSysTray.Value
    
    'Assigning the picture and tool tips for buttons
    cmdControl.Picture = imgSM.ListImages("ArrowRT").Picture
    cmdControl.ToolTipText = "Show Settings"
    cmdDetails.Picture = imgSM.ListImages("ArrowDN").Picture
    cmdDetails.ToolTipText = "Show Details"
    cmdClear.Picture = imgSM.ListImages("CheckMark").Picture
    picAbout.Picture = imgSM.ListImages("LightON").Picture
    
    'Initial values for the flags
    mblnControlsShown = False: mblnDetailsShown = False
    mblnLightON = True
    
    mblnToIsEditable = False
    txtToUser.Visible = mblnToIsEditable
    cboToUser.Visible = Not mblnToIsEditable
    
    'Initial size of the form
    Me.Width = mintMINWIDTH
    Me.Height = mintMINHEIGHT
    Call cmdRefresh_Click
    Call gfnAutosizeCombo(cboToUser, 100)
    
    'Chat Protocol setting
    mblnChatON = False
    sokChat.Protocol = sckUDPProtocol
    mintPort = 15
    
    mintSecretCode = Int(Rnd() * 1000)
    
    Me.Icon = imgSM.ListImages("LightOFF").Picture
    Exit Sub
ErrorTrap:
    Call gprShowErrorMessage("Error...Application cannot be loaded!")
    End
End Sub

Private Sub Form_Activate()
    Me.Icon = imgSM.ListImages("LightOFF").Picture
End Sub

Private Sub Form_Resize()
    'If the System Tray setting is on,
    'When clicked on minimize button, show the application in the System Tray
    On Error Resume Next
    If Me.WindowState = vbMinimized Then
        If gblnSysTrayMin = True Then
            Call prShowInSysTray
            Me.Hide
        End If
    End If
    If Me.WindowState = vbMaximized Then
            Call gprDeleteSysTrayIcon
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Handling the events, when the Application is in system tray
    
    Dim result As Long
    
    On Error Resume Next
    
    'result = SetForegroundWindow(Me.hwnd)
    'the following is a code context example ,assuming that you already have your
    'application icon on the system tray
    
    If Button = 2 Then
        DoEvents
        result = SetForegroundWindow(Me.hwnd)
    End If
    
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
            Me.Icon = imgSM.ListImages("LightOFF").Picture
            nid.hIcon = Me.Icon
            Shell_NotifyIcon NIM_MODIFY, nid
            Me.WindowState = 0
            Me.Show
            Call gprDeleteSysTrayIcon
       Case WM_RBUTTONDOWN
          Dim ToolTipString As String
'                ToolTipString = InputBox("Enter the new ToolTip:", _
'                                  "Change ToolTip")
'                If ToolTipString <> "" Then
'                   nid.szTip = ToolTipString & vbNullChar
'                   Shell_NotifyIcon NIM_MODIFY, nid
'                End If
       Case WM_RBUTTONUP
          PopupMenu Me.mnuFile
       Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If Escape is pressed, minimise the form
    On Error Resume Next
    If KeyAscii = 27 Then Me.WindowState = vbMinimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'When closed, remove the icon in system tray
    On Error Resume Next
    Call gprDeleteSysTrayIcon
End Sub

'-------------------------------------------------------------------------------------------
'Menu: Event handling procedures
'The hidden menu which pops up, when rt mouse button is clicked on the icon in system tray
'-------------------------------------------------------------------------------------------
Private Sub mnuFileRestore_Click()
    'Restores the Application into normal position
    On Error Resume Next
    Me.WindowState = 0
    Me.Show
    nid.hIcon = Me.Icon
    Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'-------------------------------------------------------------------------------------------
'Controls: Event handling procedures
'-------------------------------------------------------------------------------------------
Private Sub txtName_LostFocus()
    On Error Resume Next
    gstrUserName = Trim(txtName.Text)
    
    If gstrUserName <> gstrLoginName Then
        SaveSetting App.Title, "UserProfile", "Name", gstrUserName
    End If
End Sub

Private Sub lblTo_DblClick()
    mblnToIsEditable = Not mblnToIsEditable
    txtToUser.Visible = mblnToIsEditable
    cboToUser.Visible = Not mblnToIsEditable
    SendKeys ("{TAB}")
End Sub

Private Sub cboToUser_Click()
    chkChat.Value = 0
End Sub

Private Sub cboToUser_LostFocus()
    Dim strToMachine As String
    Dim strToUser As String
    
    On Error Resume Next
    
    Me.Caption = "Resolving the User Name... Please wait..."
    
    If chkUser.Value = 0 And chkUseTextFile.Value = 0 And Trim(cboToUser.Text) <> "" Then
        Me.MousePointer = vbHourglass
        strToMachine = mfnGetRemoteMachineName
        strToUser = mfnGetUserName(strToMachine)
        If Trim(strToUser) <> "" And InStr(strToUser, "$") = 0 Then
            MsgBox "Recently logged in User Name on the machine " & vbCrLf & vbCrLf & strToMachine & " is " & strToUser
        Else
            MsgBox "Recently logged in User Name on the machine " & vbCrLf & vbCrLf & strToMachine & " is not available!"
        End If
        Me.MousePointer = vbDefault
    End If
    
    Me.Caption = gstrMsgTitle & "  [" & gstrComputerName & "]"
End Sub

Private Sub chkChat_Click()
    On Error GoTo ErrorTrap
    
    mblnChatON = chkChat.Value
    
    If mblnChatON = True Then
        'If Chat option is selected establishing connection to the remote machine
        If mfnValidateData = False Then Exit Sub
        With sokChat
            .RemoteHost = mfnGetRemoteMachineName
            .RemotePort = mintPort
            If Not sokChat.State = 1 Then
                .Bind mintPort
            End If
        End With
        Me.Height = mintMAXHEIGHT
    Else
        Me.Height = mintMINHEIGHT
    End If
    
    Exit Sub
ErrorTrap:
    Call gprShowErrorMessage("Error... Cannot establish connection to the remote machine!" & vbCrLf & "Try again to use another port")
    mintPort = mintPort + 1
End Sub

Private Sub txtMessage_GotFocus()
    Me.Icon = imgSM.ListImages("LightOFF").Picture
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then Call txtMessage_KeyPress(13)
    If KeyCode = vbKeyF3 Then txtDetails.Text = ""
    If KeyCode = vbKeyF4 Then Unload Me
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    'Sends the message using Windows Net Send or Win Sock
    
    Dim strMsg As String
    Dim strToMachine As String
    Dim strEncryptedMsg As String
    Dim varRetValue As Variant
    
    On Error GoTo ErrorTrap
    
    'If the Key pressed is not Enter Key, exit
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    If mfnValidateData = False Then Exit Sub
    strToMachine = mfnGetRemoteMachineName
    
    strMsg = Trim(txtName.Text) & " Says " & Chr(187) & " " & Trim(txtMessage.Text)
    
    If mblnChatON = True Then
        'If Chat option is selected, send the message through WinSock
        strEncryptedMsg = gfnEncodeMsg(strMsg, Str(mintSecretCode))
        sokChat.SendData strEncryptedMsg
    Else
        'Otherwise, send the message as popup using windows net send
        strMsg = "Net Send " & strToMachine & " " & strMsg
        varRetValue = Shell(strMsg, vbHide) 'Sending the message
        If varRetValue = 0 Then MsgBox "Error...Problem in sending message!" & vbCrLf & _
            "(May be Network Problems)", vbInformation + vbOKOnly, gstrMsgTitle
    End If
    
    strMsg = "To " & Trim(strToMachine) & " " & Chr(187) & " " & Trim(txtMessage.Text)
    txtMessage.Text = ""
    
    'Coloring the sent text and scrolling the text box
    With txtDetails
        .Text = .Text & vbCrLf & strMsg
        .SelStart = Len(.Text) - Len(strMsg)
        .SelLength = Len(strMsg)
        .SelColor = vbBlue
        .SelLength = Len(.Text)
    End With
    
    txtMessage.SetFocus
    SendKeys "^{Home}"
    
    Exit Sub
ErrorTrap:
    Call gprShowErrorMessage("Error...Problem encountered while sending message!" & _
    vbCrLf & "Please make sure that the remote machine name is correct if you have typed it.")
End Sub

Private Sub cmdBuzz_Click()
    On Error Resume Next
    If mblnChatON = True Then
        txtMessage.Text = "BUZZ"
        Call txtMessage_KeyPress(13)
    End If
End Sub

Private Sub cmdControl_Click()
    'Showing or hiding the settings screen
    On Error Resume Next
    If mblnControlsShown = False Then
        Me.Width = mintMAXWIDTH
        cmdControl.Picture = imgSM.ListImages("ArrowLT").Picture
        cmdControl.ToolTipText = "Hide Settings"
    Else
        Me.Width = mintMINWIDTH
        cmdControl.Picture = imgSM.ListImages("ArrowRT").Picture
        cmdControl.ToolTipText = "Show Settings"
    End If
    mblnControlsShown = Not mblnControlsShown
End Sub

Private Sub cmdDetails_Click()
    'Showing or hiding the details screen
    On Error Resume Next
    If mblnDetailsShown = False Then
        Me.Height = mintMAXHEIGHT
        cmdDetails.Picture = imgSM.ListImages("ArrowUP").Picture
        cmdDetails.ToolTipText = "Hide Details"
    Else
        Me.Height = mintMINHEIGHT
        cmdDetails.Picture = imgSM.ListImages("ArrowDN").Picture
        cmdDetails.ToolTipText = "Show Details"
    End If
    mblnDetailsShown = Not mblnDetailsShown
End Sub

Private Sub cmdClear_Click()
    'Clears the details text box
    txtDetails.Text = ""
End Sub

Private Sub cmdRefresh_Click()
    'Populates the network machines
    On Error GoTo ErrorTrap
    Me.MousePointer = vbHourglass
    
    If gblnUseTextFile = True Then
      Call gprPopulateFromTextFile(cboToUser, "SoftMessenger.txt")
    Else
      prbUsers.Value = 0
      If chkUser.Value = 1 Then
          Call gprPopulateNetworkUsers(cboToUser, prbUsers, True)
      ElseIf chkUser.Value = 0 Then
          Call gprPopulateNetworkUsers(cboToUser, prbUsers, False)
      End If
    End If
    
    Me.MousePointer = vbDefault
    Exit Sub
ErrorTrap:
    Call gprShowErrorMessage("Error...To list cannot be populated!")
    Me.MousePointer = vbDefault
End Sub

Private Sub chkSysTray_Click()
    'Save the System Tray minimising setting to registry
    On Error Resume Next
    gblnSysTrayMin = chkSysTray.Value
    SaveSetting App.Title, "Settings", "SysTrayMin", chkSysTray.Value
End Sub

Private Sub chkLoginName_Click()
    'Save the ShowLoginName setting to registry
    On Error Resume Next
    gblnShowLoginName = chkLoginName.Value
    SaveSetting App.Title, "Settings", "ShowLoginName", chkLoginName.Value
    Call mprSetUserName
End Sub

Private Sub chkUseTextFile_Click()
    'Save the UseTextFile setting to registry
    On Error Resume Next
    gblnUseTextFile = chkUseTextFile.Value
    SaveSetting App.Title, "Settings", "UseTextFile", chkUseTextFile.Value
    Call cmdRefresh_Click
End Sub

Private Sub picAbout_Click()
    'Easter Egg; Blinks the light in the about window when clicked on it
    On Error Resume Next
    If mblnLightON = True Then
        picAbout.Picture = imgSM.ListImages("LightOFF").Picture
    Else
        picAbout.Picture = imgSM.ListImages("LightON").Picture
    End If
    mblnLightON = Not mblnLightON
    picAbout.ToolTipText = "Developed By Sameer C T"
End Sub

Private Sub sokChat_DataArrival(ByVal bytesTotal As Long)

    Dim intMsgLength As Integer
    Dim intCount As Integer
    Dim lngCount As Long
    Dim strReceivedMessage() As String
    
    On Error GoTo ErrorTrap
    
    'Getting the arrived message from port
    sokChat.GetData mstrChatMsg
    mstrChatMsg = gfnDecodeMsg(mstrChatMsg, Str(mintSecretCode))
    intMsgLength = Len(mstrChatMsg)
    intCount = 1
    
    'If the application is minimised, flash it
    If Me.WindowState = 1 Then FlashWindow Me.hwnd, 1
    
    'If the Apln in Systray, it should show a sign that a message has arrived.
    If Me.WindowState = vbMinimized Then
        If gblnSysTrayMin = True Then
            'Me.WindowState = 0
            'Me.Show
            Me.Icon = imgSM.ListImages("LightON").Picture
            nid.hIcon = Me.Icon
            Shell_NotifyIcon NIM_MODIFY, nid
            'Call prShowInSysTray
            'Me.Hide
        End If
    End If
    
    'If the arrived message is a code...
    strReceivedMessage = Split(mstrChatMsg, Chr(187))
    '...for Buzzing; shake the form calling timer codes
    If StrComp(UCase(Trim(strReceivedMessage(1))), "BUZZ", vbTextCompare) = 0 Then
        mstrChatMsg = mstrChatMsg & "...!"
        tmrSM.Enabled = True
    '...for clearing; clear the message details
    ElseIf StrComp(UCase(Trim(strReceivedMessage(1))), "CLRSCR", vbTextCompare) = 0 Then
        txtDetails.Text = ""
        mstrChatMsg = Replace(strReceivedMessage(0), "Says", " HAS CLEARED ALL MESSAGE DETAILS!")
    End If
    
    'Changing the color of the arrived text and scrolling the text box to new msg
    With txtDetails
        .Text = .Text & vbCrLf & mstrChatMsg
        .SelStart = Len(txtDetails.Text) - Len(mstrChatMsg)
        .SelLength = Len(mstrChatMsg)
        .SelColor = vbRed
        .SelLength = Len(txtDetails.Text)
    End With
    
    Exit Sub
ErrorTrap:
    Call gprShowErrorMessage("Error while sending or receiving messages!" & vbCrLf & _
                                "Make sure that the other party has turned on the Chat option")
End Sub

Private Sub tmrSM_Timer()
    'Shakes the form by moving it; for Buzz during chat
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    mintBuzzCounter = mintBuzzCounter + 1
    If mblnBuzzON = True Then
        Me.Move Me.Left + 100, Me.Top + 100
        mblnBuzzON = False
        If mintBuzzCounter >= 30 Then tmrSM.Enabled = False: mintBuzzCounter = 0
        Exit Sub
    End If
    If Not mblnBuzzON Then
        Me.Move Me.Left - 100, Me.Top - 100
        mblnBuzzON = True
        If mintBuzzCounter >= 4 Then tmrSM.Enabled = False: mintBuzzCounter = 0
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------------------
'Module level procedures
'-------------------------------------------------------------------------------------------
Private Sub prShowInSysTray()
    ' This fn Keeps the icon in taskbar

     'Set the individual values of the NOTIFYICONDATA data type.
     nid.cbSize = Len(nid)
     nid.hwnd = Me.hwnd
     nid.uId = vbNull
     nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     nid.uCallBackMessage = WM_MOUSEMOVE
     nid.hIcon = Me.Icon
     nid.szTip = gstrMsgTitle & vbNullChar

     'Call the Shell_NotifyIcon function to add the icon to the taskbar
     'status area.
     Shell_NotifyIcon NIM_ADD, nid
     'Max = 0
End Sub

Private Sub mprSetUserName()
    'Displaying the Login Name as User Name if it is set; other wise the stored name
    gstrUserName = GetSetting(App.Title, "UserProfile", "Name")
    If gblnShowLoginName = True Then
        txtName.Text = gstrLoginName
    Else
        txtName.Text = gstrUserName
    End If
End Sub

Private Function mfnGetRemoteMachineName() As String
    'Extracts the remote machine's name from the combo
    'If the combo contains user name and computer name, it will be separated by a colon
    'So based on that colon, only computer name is extracted
    
    Dim strToAddress() As String

    If mblnToIsEditable = True Then
        mfnGetRemoteMachineName = Trim(txtToUser.Text)
    Else
        strToAddress = Split(Trim(cboToUser.Text), ":")
        
        If UBound(strToAddress) > 0 Then
            mfnGetRemoteMachineName = Trim(strToAddress(1))
        Else
            mfnGetRemoteMachineName = Trim(cboToUser.Text)
        End If
    End If
End Function

Private Function mfnValidateData() As Boolean
    mfnValidateData = False
    If mblnToIsEditable = True Then
        If Trim(txtToUser.Text) = "" Then
            MsgBox "Type a valid computer name", vbInformation, gstrMsgTitle
            txtToUser.SetFocus
            Exit Function
        End If
    Else
        If cboToUser.ListIndex = -1 Then
            MsgBox "To address must be selected", vbInformation, gstrMsgTitle
            cboToUser.SetFocus
            Exit Function
        End If
    End If
    mfnValidateData = True
End Function
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------


