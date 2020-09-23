This file contains VBScript Functions Along with built in functions I have added.

ActivateWindow~(ByVal lzWindowName As String) As Long
AlphaBlend~(Hangle As Long. [AlphaBlendValue As Integer = 165])
AppPath~()
AppendFile~(Filename As String. Text As String)
Asc~(String) As Variant
Abs~(Number)
Array~(ParamArray ArgList() As Variant)
Atn~(Number As Double) as Variant
BeepA~()
BrowseForFolder~(Hangle As Long. [Title As String])
CenterDialog~()
CBool~(Expression) As Variant
CByte~(expression) As Variant
CCur~(expression) As Variant
CDate~(date) As Variant
CDbl~(expression) As Variant
Chr~(charcode)
CheckIni~() As Long
CheckKey~(Selection As String. KeyName As String) As String
CInt~(expression) As Variant
CLng~(expression) As Variant
CountChr~(ByVal Range As String. ByVal Criteria As String) As Long
CloseWindowA~(WndHangle As Long)
Cos~(number)
CreateObject~(Class As String. [ServerName As String])
CSng~(expression) As Variant
CStr~(expression) As Variant
CopyFile~(File1 As String. File2 As String) as Variant
CreateFolder~(lzPath As String)
DrawLine~(mObject As Object. X1. Y1. X2. Y2. mColor)
DriveType~(string) As Long
Date~()
DateAdd~(interval. number. date)
DateDiff~(interval. date1. date2 [.firstdayofweek[. firstweekofyear]])
DatePart~(interval. date[. firstdayofweek[. firstweekofyear]])
DateSerial~(year. month. day)
DateValue~(date As Variant)
Day~(date)
DeleteFile~(Filename As String) As Variant
DeleteFolder~(lzPath As String) As Variant
DisconnectNetworkDriveDlg~(ByVal Hangle As Long)
DownloadFile~(ByVal URL As String. ByVal LocalFilename As String) As Boolean
Exp~(number)
FindWindowA~(WndClsName As String. WndName As String)
FileDateTimeA~(Filename As String) As Date
FindFile~(lzFilename As String) As Integer
FixPath~(lzPath As String) As Variant
Filter~(InputStrings. Value[. Include[. Compare]])
Fix~(number)
FlashWindow~(hwnd As Long. mInterval As Long)
Int~(number)
FolderExists~(lzPath As String) As Integer
FormatCurrency~(Expression[.NumDigitsAfterDecimal [.IncludeLeadingDigit_ [.UseParensForNegativeNumbers [.GroupDigits]]]])
FormatDateTime~(Date[. NamedFormat])
FormatNumber~(Expression [.NumDigitsAfterDecimal [.IncludeLeadingDigit_ [.UseParensForNegativeNumbers [.GroupDigits]]]])
FormatPercent~(Expression[.NumDigitsAfterDecimal [.IncludeLeadingDigit_ [.UseParensForNegativeNumbers [.GroupDigits]]]])
GetFileAttributes~(Filename As String) As Long
GetFileSize~(Filename As String) As Long
GetObject~([pathname] [. class])
GetHDC~(ByVal Hangle As Long) As Long
GetPixelA~(ByVal tHDC As Long. ByVal X As Long. ByVal Y As Long)
GetClip~([ByVal zFormatType As Integer = 1]) As String
GetEnvVar~(sName As String) As String
GetSettingA~(ByVal tAppName As String. ByVal tSelection As String. _ByVal tKey As String. ByVal tDefault As String) As String
GetActiveWindowA~() As Long
GetComputerNameA~() As String
GetMousePos~()
GetOSVerType~() as integer
GetSpecialFolderLocation~(ByVal bsSpecialFolder As String) As String
GetTickCountA~() As Long
GetUserNameA~() As String
GetWindowPosition~(ByVal Hangle As Long) As Variant()
GetForegroundWindowA~()
Hex~(number)
Hour~(time)
INIDeleteKey~(ByVal Selection As String. ByVal KeyName As String) As Long
INIDeleteKeyValue~(ByVal Selection As String. ByVal KeyName As String) As Long
INIDeleteSelection~(ByVal Selection As String) As Long
INIReadKeyValue~(ByVal Selection As String. ByVal KeyName As String. _[ByVal DefaultKey As String]) As String
INIWriteKeyValue~(ByVal Selection As String. ByVal KeyName As String. _sKeyValue As String) As Long
InputBox~(prompt[. title][. default][. xpos][. ypos][. helpfile. context])
InStr~([start. ]string1. string2[. compare])
InStrRev~(string1. string2[. start[. compare]])
isHibernateAllowed~() As Long
isShutdownAllowed~() As Long
isSuspendAllowed~() As Long
isAdmin~()
IsArray~(VarName) As Boolean
IsDate~(Expression) As Boolean
sEmpty~(Expression) As Boolean
IsNull~(Expression) As Boolean
IsNumeric~(Expression) As Boolean
IsObject~(Expression) As Boolean
Join~(list[. delimiter]) as String
LBound~(arrayname[. dimension])
LCase~(string)
Left~(string. length)
Len~(Expression)
LoadPicture~(picturename)
Log~(number)
LTrim~(string)
RTrim~(string)
Trim~(string)
MciSendStringA~(ByVal lpstrCommand As String. ByVal lpstrReturnString As String. _ByVal uReturnLength As Long. ByVal hwndCallback As Long)
MapNetworkDriveDlg~(ByVal Hangle As Long) As Long
Mid~(string. start[. length])
Minute~(time)
Month~(date)
MonthName~(Month As Long. [Abbreviate As Boolean = False]) As String
MoveFileA~(ByVal File1 As String. ByVal File2 As String) As Long
MessageBox~(ByVal Prompt As String. _[Buttons As VbMsgBoxStyle]. [Title = "Message Box"]) As Long
MsgBox~(prompt[. buttons][. title][. helpfile. context])
Now~()
Oct~(number)
OpenFile~(Filename As String) As String
Replace~(expression. find. replacewith[. start[. count[. compare]]])
RGB~(red. green. blue)
Right~(string. length)
Rnd~[(number)]
Run~(Filename As String) As Long
RunControlPanelApp~(ByVal ProgName As String) As Long
RunDialog~(Hangle As Long. [ByVal Title As String = "Run"]. _[ByVal Prompt As String = "Enter the name of the program to run"])
Round~(expression[. numdecimalplaces])
Sgn~(number)
Sin~(number)
SaveSettingA~(ByVal tAppName As String. ByVal tSelection As String. _ByVal tKey As String. ByVal tSetting As String)
Space~(number)
SetMousePos~(ByVal X As Long. Y As Long)
SetWindowFocus~(ByVal Hangle As Long)
SendKeysA~(TKeys As String. [ByVal Wait As Integer])
SetClip~(ByVal StrBuff As String. [ByVal zFormatType As Integer = 1]) As Integer
SetWindowPosition~(ByVal Hangle As Long. ByVal X. ByVal Y. _ByVal mHeight As Long. ByVal mWidth As Long) As Long
SetWindowTextA~(ByVal Hangle As Long. ByVal lText As String)
SetEnvVar~(ByVal sName As String. ByVal sValue As String) As Long
StrFormatDateTime~(lExpression As String. [ByVal lFormat])
Swap~(a. b)
Split~(expression[. delimiter[. count[. compare]]])
Sqr~(number)
Second~(time)
StrComp~(string1. string2[. compare])
StrReverse~(string1)
String~(number. character)
Tan~(number)
TColor~(ByVal lColor As Integer) As Long
TextOutA~(ByVal tHDC As Long. ByVal X As Long. ByVal Y As Long. ByVal Text As String) As Long
Time~()
TimeSerial~(hour. minute. second) As Variant
TimeValue~(time) As Variant
TypeName~(varname) As Variant
UBound~(arrayname[. dimension])
UnloadDialog~()
UCase~(string)
VarType~(varname)
Weekday~(date. [firstdayofweek])
WeekdayName~(weekday. abbreviate. firstdayofweek)
WriteFile~(Filename As String. DataBuff As String)
Year~(date)
Pause~(ByVal Millisecond As Long)
Printf~(S As String)
Power~(ByVal iNum. ByVal iCount)
Plot~(mObject As Object. X. Y. bColor)