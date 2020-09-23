Attribute VB_Name = "ModAPI"
' Window API calls
Public Declare Function GetTickCount& Lib "Kernel32" ()
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function FlushFileBuffers Lib "Kernel32" (ByVal hFile As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CreateThread Lib "Kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeThread Lib "Kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Sub ExitThread Lib "Kernel32" (ByVal dwExitCode As Long)
Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
'Windows API calls
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function WinExec Lib "Kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetEnvironmentVariable Lib "Kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "Kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function IsPwrShutdownAllowed Lib "Powrprof.dll" () As Long
Public Declare Function IsPwrSuspendAllowed Lib "Powrprof.dll" () As Long
Public Declare Function IsPwrHibernateAllowed Lib "Powrprof.dll" () As Long
' Drive and file API calls
Public Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function MoveFile Lib "Kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function DeleteFile Lib "Kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' Internet / Network API calls
Public Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function NetUserChangePassword Lib "NETAPI32.DLL" (ByVal domainname As String, ByVal userName As String, ByVal oldpassword As String, ByVal newpassword As String) As Long
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
'Graphic API Calls
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' Window stayle consts
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
' Browse folder consts
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
' windows version consts
Private Const VER_PLATFORM_WIN32_NT = 2
' TextBox Consts
Public Const EM_SETMARGINS = &HD3
Public Const EC_LEFTMARGIN = &H1
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETSEL = &HB0
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETLINE = &HC4

Public Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    Y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum TSpecialFolders
    DM_DESKTOP = &H0
    DM_PROGRAMS = &H2
    DM_Controls = &H3
    DM_PRINTERS = &H4
    DM_PERSONAL = &H5
    DM_FAVORITES = &H6
    DM_STARTUP = &H7
    DM_RECENT = &H8
    DM_SENDTO = &H9
    DM_BITBUCKET = &HA
    DM_STARTMENU = &HB
    DM_DESKTOPDIRECTORY = &H10
    DM_DRIVES = &H11
    DM_NETWORK = &H12
    DM_NETHOOD = &H13
    DM_FONTS = &H14
    DM_TEMPLATES = &H15
End Enum

Enum RegOp
    Register = 1
    UnRegister
End Enum

Public Osver As OSVERSIONINFO
Public nPos As POINTAPI

Public Function GetShortPath(lzPathName As String) As String
    Dim iRet As Long, StrA As String
    StrA = String$(512, vbNullChar)
    iRet = GetShortPathName(lzPathName, StrA, 164)
    GetShortPath = Left$(StrA, iRet)
    StrA = ""
    iRet = 0
End Function

Public Function isWinNT() As Boolean
    Osver.dwOSVersionInfoSize = Len(Osver)
    GetVersionEx Osver
    isWinNT = (Osver.dwPlatformId) = VER_PLATFORM_WIN32_NT
End Function

Function GetFolder(ByVal hwndOwner As Long, mTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim OffSet As Integer
    If Len(mTitle) = 0 Then mTitle = "Look in:"
    bInf.hOwner = hwndOwner
    bInf.lpszTitle = mTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        OffSet = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, OffSet - 1)
    End If
End Function

Public Function tFlashWindow(mWnd As Long, mIntVal As Long)
   FlashWindow mWnd, mIntVal
End Function

Public Function GetUser() As String
Dim S As String, iRet As Long
    S = Space(165)
    iRet = GetUserName(S, 165)
    If iRet <> 1 Then
        GetUser = ""
        S = ""
        Exit Function
    Else
        GetUser = Left(S, InStr(S, Chr(0)) - 1)
        S = ""
    End If
End Function

Public Function SysComputerName() As String
Dim S As String, iRet As Long
    S = Space(160)
    iRet = GetComputerName(S, 160)
    
    If iRet <> 1 Then
        SysComputerName = ""
        S = ""
        Exit Function
    Else
        SysComputerName = Left(S, InStr(S, Chr(0)) - 1)
        S = ""
    End If
    
End Function

Public Function DMGetSystemPath() As String
Dim StrBuff As String
    StrBuff = String(255, Chr(0))
    GetSystemDirectory StrBuff, 255
    DMGetSystemPath = Left(StrBuff, InStr(StrBuff, Chr(0)) - 1)
    StrBuff = ""
End Function

Public Function DMGetTempPath() As String
Dim StrBuff As String
    StrBuff = String(255, Chr(0))
    GetTempPath 255, StrBuff
    DMGetTempPath = Left(StrBuff, InStr(1, StrBuff, Chr(0)) - 1)
    StrBuff = ""
End Function

Public Function DMGetWindowsPath() As String
Dim StrBuff As String
    StrBuff = String(255, Chr(0))
    GetWindowsDirectory StrBuff, 255
    DMGetWindowsPath = Left(StrBuff, InStr(1, StrBuff, Chr(0)) - 1)
    StrBuff = ""
End Function

Public Sub FlatBorder(ByVal hwnd As Long)
  Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Function RegisterActiveX(lzAxDll As String, mRegOption As RegOp) As Boolean
Dim mLib As Long, DllProcAddress As Long
Dim mThread
Dim sWait As Long
Dim mExitCode As Long
Dim lpThreadID As Long

Dim slib As String

    slib = lzAxDll
    mLib = LoadLibrary(slib)
    
    If mLib <= 0 Then
        RegisterActiveX = False
        Exit Function
    End If
    
    If mRegOption = Register Then
        DllProcAddress = GetProcAddress(mLib, "DllRegisterServer")
    Else
        DllProcAddress = GetProcAddress(mLib, "DllUnregisterServer")
    End If
    
    If DllProcAddress = 0 Then
        RegisterActiveX = True
        Exit Function
    Else
        mThread = CreateThread(ByVal 0, 0, ByVal DllProcAddress, ByVal 0, 0, lpThreadID)
        
        If mThread = 0 Then
            FreeLibrary mLib
            RegisterActiveX = False
            Exit Function
        Else
            sWait = WaitForSingleObject(mThread, 10000)
            If sWait <> 0 Then
                FreeLibrary lLib
                mExitCode = GetExitCodeThread(mThread, mExitCode)
                ExitThread mExitCode
                Exit Function
            Else
                FreeLibrary mLib
                CloseHandle mThread
            End If
        End If
    End If
    slib = ""
    RegisterActiveX = True
    
End Function

