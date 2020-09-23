Attribute VB_Name = "modSMMain"
'-------------------------------------------------------------------------------------------
'Project Name   :   Soft Messenger (SoftMessenger.vbp)
'Started On     :   2001 November 28
'Last Modified  :   2001 December 06
'Version 2003   :   2003 January 21
'Description    :   Message sending using windows's Net Send
'                    and chat using winsock among the networked machines

'Module Name    :   Main Module(modSMMain/SoftMessenger.bas)
'Developed By   :   Sameer C T
'Last Modified  :   2003 March 28
'-------------------------------------------------------------------------------------------
Option Explicit
'-------------------------------------------------------------------------------------------
'Global Variables
'-------------------------------------------------------------------------------------------
Public gstrMsgTitle As String
Public gstrLoginName As String
Public gstrUserName As String
Public gstrComputerName As String
Public gstrDomainName As String
Public gstrFullDomainName As String
Public gblnSysTrayMin As Boolean
Public gblnShowLoginName As Boolean
Public gblnUseTextFile As Boolean
'-------------------------------------------------------------------------------------------
'General Declarations
'-------------------------------------------------------------------------------------------
Public TheDomain As IADsDomain
Public Computer As IADsComputer
Public luserName As IADsUser

Public iMousePointer As Integer
Public username1 As String
Public frm As Form
Public lWindowExists As Boolean

'-------------------------------------------------------------------------------------------
'Declaration for getting username for each mechine
'-------------------------------------------------------------------------------------------
Private Const NERR_SUCCESS As Long = 0&
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const ERROR_MORE_DATA As Long = 234&
Private Const LB_SETTABSTOPS As Long = &H192

Private Const PLATFORM_ID_DOS As Long = 300
Private Const PLATFORM_ID_OS2 As Long = 400
Private Const PLATFORM_ID_NT  As Long = 500
Private Const PLATFORM_ID_OSF As Long = 600
Private Const PLATFORM_ID_VMS As Long = 700
   
'for use on Win NT/2000 only
Private Type WKSTA_INFO_102
  wki102_platform_id As Long
  wki102_computername As Long
  wki102_langroup As Long
  wki102_ver_major As Long
  wki102_ver_minor As Long
  wki102_lanroot As Long
  wki102_logged_on_users As Long
End Type

Private Declare Function NetWkstaGetInfo Lib "netapi32" _
  (ByVal servername As Long, _
   ByVal level As Long, _
   bufptr As Long) As Long
   
Private Type WKSTA_USER_INFO_0 'not used; provided for completeness
  wkui0_username  As Long
End Type

Private Type WKSTA_USER_INFO_1
  wkui1_username As Long
  wkui1_logon_domain As Long
  wkui1_oth_domains As Long
  wkui1_logon_server As Long
End Type

Private Declare Function NetWkstaUserEnum Lib "netapi32" _
  (ByVal servername As Long, _
   ByVal level As Long, _
   bufptr As Long, _
   ByVal prefmaxlen As Long, _
   entriesread As Long, _
   totalentries As Long, _
   resume_handle As Long) As Long
   
Private Declare Function NetApiBufferFree Lib "netapi32" _
   (ByVal Buffer As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, _
   uFrom As Any, _
   ByVal lSize As Long)
   
Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lparam As Any) As Long
'-------------------------------------------------------------------------------------------
'Used for autosize combo
'-------------------------------------------------------------------------------------------
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const DT_CALCRECT = &H400

Private Type rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lparam As Long) As Long

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As rect, ByVal wFormat _
    As Long) As Long
'-------------------------------------------------------------------------------------------
'Global Procedures
'-------------------------------------------------------------------------------------------
Public Sub gprPopulateNetworkUsers(cboCombo As ComboBox, prgBar As ProgressBar, blnUsersRequired As Boolean)
    'Populates network computers and the users
       
    'Declaration for getting user names
    Dim bufptr          As Long
    Dim dwServer        As Long
    Dim nStructSize     As Long
    Dim strServer       As String
    Dim ws102           As WKSTA_INFO_102
    Dim wui1            As WKSTA_USER_INFO_1
    Dim User            As String
    Dim intCount        As Integer
    Dim strName As String
    
    On Error Resume Next

    strServer = gstrComputerName
    dwServer = StrPtr(strServer)
    wui1 = mfnGetWorkstationUserInfo(dwServer)
    gstrDomainName = mfnGetPointerToByteStringW(wui1.wkui1_logon_domain)
    gstrFullDomainName = "WinNT://" & gstrDomainName
    
    'Take only the Computers in the Domain, filtering out other objects
    Set TheDomain = GetObject(gstrFullDomainName)
    TheDomain.Filter = Array("Computer")
    
    'Finding out the number of computer by looping, to set max for progress bar
    cboCombo.Clear
    intCount = 0
    For Each Computer In TheDomain
        intCount = intCount + 1
    Next Computer
    prgBar.Min = 0
    prgBar.Max = intCount
    prgBar.Value = 0
    intCount = 0
    
    For Each Computer In TheDomain
        On Error Resume Next
        If blnUsersRequired = True Then
            strName = mfnGetUserName(Computer.Name)
            If Trim(strName) <> "" Then
                cboCombo.AddItem (strName & " : " & Trim(Computer.Name))
            Else
                cboCombo.AddItem (Trim(Computer.Name))
            End If
        Else
            cboCombo.AddItem (Trim(Computer.Name))
        End If
        If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Next Computer
      
    'Clean up
    Set Computer = Nothing
    Set TheDomain = Nothing
End Sub

Public Sub gprPopulateFromTextFile(cboCombo As ComboBox, TextFileName As String)
    'Populates the combo with lines from a text file
    Dim objFSO As Object
    Dim filData As Variant
    Dim datData As Variant
    Dim strLine As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Set filData = objFSO.getfile(TextFileName)
    Set datData = filData.OpenAsTextStream(1)
    
    cboCombo.Clear
    Do While datData.AtEndOfLine <> True
        strLine = datData.ReadLine
        If Left(strLine, 1) <> "-" Then cboCombo.AddItem strLine
    Loop
    datData.Close
End Sub

Public Sub gprShowErrorMessage(strMessage As String)
        MsgBox strMessage & vbCrLf & vbCrLf & _
            "Technical Details:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrMsgTitle
End Sub

Public Function gfnAutosizeCombo(Combo As ComboBox, length As Integer) As Boolean
    Dim lngRet As Long
    Dim lngCurrentWidth As Single
    Dim rectCboText As rect
    Dim lngParentHDC As Long
    Dim lngListCount As Long
    Dim lngCounter As Long
    Dim lngTempWidth As Long
    Dim lngWidth As Long
    Dim strSavedFont As String
    Dim sngSavedSize As Single
    Dim blnSavedBold As Boolean
    Dim blnSavedItalic As Boolean
    Dim blnSavedUnderline As Boolean
    Dim blnFontSaved As Boolean
On Error GoTo ErrorHandler
    'Grab the combo handle and list count
      lngParentHDC = Combo.Parent.hdc
      lngListCount = Combo.ListCount
      If lngParentHDC = 0 Or lngListCount = 0 Then Exit Function
          'Save combo box fonts, etc. to the parent
          'object (form), for testing lengths with the API
            With Combo.Parent
                  strSavedFont = .FontName
                  sngSavedSize = .FontSize
                  blnSavedBold = .FontBold
                  blnSavedItalic = .FontItalic
                  blnSavedUnderline = .FontUnderline
                  .FontName = Combo.FontName
                  .FontSize = Combo.FontSize
                  .FontBold = Combo.FontBold
                  .FontItalic = Combo.FontItalic
                  .FontUnderline = Combo.FontItalic
            End With
      blnFontSaved = True
      'Get the width of the widest item
      For lngCounter = 0 To lngListCount
            DrawText lngParentHDC, Combo.List(lngCounter), -1, rectCboText, _
            DT_CALCRECT
            lngTempWidth = rectCboText.Right - rectCboText.Left + length
            If (lngTempWidth > lngWidth) Then
                  lngWidth = lngTempWidth
            End If
      Next
     'Get current width of combo
     lngCurrentWidth = SendMessageLong(Combo.hwnd, CB_GETDROPPEDWIDTH, 0, 0)

    'If big enough then that's all A-OK
      If lngCurrentWidth > lngWidth Then
            gfnAutosizeCombo = True
            GoTo ErrorHandler
            Exit Function
      End If
      '... but if not big enough, first calculate the screen width to ensure we don't exceed it!
      If lngWidth > Screen.Width \ Screen.TwipsPerPixelX - length Then _
        lngWidth = Screen.Width \ Screen.TwipsPerPixelX - length
      'Set the width of our combo
      lngRet = SendMessageLong(Combo.hwnd, CB_SETDROPPEDWIDTH, lngWidth, 0)
      'Set the function to True/False depending on API success
      gfnAutosizeCombo = lngRet > 0
ErrorHandler:
      'If anything goes wrong, revert back!
      On Error Resume Next
      If blnFontSaved Then
      With Combo.Parent
            .FontName = strSavedFont
            .FontSize = sngSavedSize
            .FontUnderline = blnSavedUnderline
            .FontBold = blnSavedBold
            .FontItalic = blnSavedItalic
      End With
      End If
End Function

'-------------------------------------------------------------------------------------------
'Module level Functions
'-------------------------------------------------------------------------------------------
Public Function mfnGetUserName(bServer As String) As String
    Dim bufptr          As Long
    Dim dwServer        As Long
    Dim wui1            As WKSTA_USER_INFO_1

    dwServer = StrPtr(bServer)
    wui1 = mfnGetWorkstationUserInfo(dwServer)
    If wui1.wkui1_username <> 0 Then
        mfnGetUserName = mfnGetPointerToByteStringW(wui1.wkui1_username)
    Else
        mfnGetUserName = ""
    End If
End Function

Private Function mfnGetPointerToByteStringW(ByVal dwData As Long) As String
   Dim tmp() As Byte
   Dim tmplen As Long
   
   If dwData <> 0 Then
      tmplen = lstrlenW(dwData) * 2
      If tmplen <> 0 Then
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         mfnGetPointerToByteStringW = tmp
     End If
   End If
End Function

Private Function mfnGetWorkstationUserInfo(ByVal dwWorkstation As Long) _
                                        As WKSTA_USER_INFO_1

   Dim bufptr          As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim success         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long
   Dim wui1            As WKSTA_USER_INFO_1
   
   success = NetWkstaUserEnum(dwWorkstation, _
                              1, _
                              bufptr, _
                              MAX_PREFERRED_LENGTH, _
                              dwEntriesread, _
                              dwTotalentries, _
                              dwResumehandle)

   If success = NERR_SUCCESS And _
      success <> ERROR_MORE_DATA Then
      
      nStructSize = LenB(wui1)
      
      If dwEntriesread > 0 Then
         
        'cast data into WKSTA_USER_INFO_1
        'and return the type. Although this
        'API enumerates and returns information
        'about all users currently logged on to
        'the workstation, including interactive,
        'service and batch logons, chances are
        'that the first user enumerated was the
        'user who logged on the session so we
        'exit after the first user info is returned.
        '
        'If this presumption is incorrect, please
        'let me know via the VBnet Comments link.
         CopyMemory mfnGetWorkstationUserInfo, _
                    ByVal bufptr, _
                    nStructSize
         
        'clean up before exiting
         Call NetApiBufferFree(bufptr)
         Exit Function

      End If
      
   End If
   
  'clean up
   Call NetApiBufferFree(bufptr)
End Function
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
