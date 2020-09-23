Attribute VB_Name = "ModMain"
Private Type TConfig
    ShowGrid As Boolean
    GridX As Integer
    GridY As Integer
    GridColor As Long
    MovableAtRunTime As Boolean
    ShowControlHints As Boolean
    FormBackColor As Long
    FormCaption As String
    EditFont As String
    EditFontSize As Integer
    EditFontColor As Long
    EditFontBackColor As Long
    EditTabSize As Integer
    EditCodeinsight As Boolean
    EditMarginBar As Boolean
End Type

Private Type inSightHelperT
    tName As String
    tKey As String
End Type

Private Type TCodeHelpers
    CodehelperCount As Integer
    CodehelperName() As String
    CodehelperFilePath() As String
End Type

Private Type TPlugins
    PlugInterFace() As String
    PlugFileName() As String
End Type

Public LanComment As String
Public AppConfig As TConfig
Public App_ConfigFile As String
Public inSightHelperList As String
Public ApplicationPath As String, TemplatePath As String, Function_List As String
Public DataPath As String ' Path to all the program data files
Public DevXML As New dmXML
Public TPlugin As TPlugins
Public inSightHelper() As inSightHelperT
Public CodeHelperTemplate As TCodeHelpers

Public Sub WriteXMLIni()
Dim StrBuff As String
Dim nFile As Long
    nFile = FreeFile
    Const Char34 = """"
    StrBuff = "<?xml version=" & Char34 & "1.0" & Char34 & " encoding=" & Char34 & "utf-8" & Char34 & " ?>" & vbCrLf
    StrBuff = StrBuff & "<config>" & vbCrLf
    StrBuff = StrBuff & vbTab & "<FormDesigner>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<ShowGrid value=" & Char34 & AppConfig.ShowGrid & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<GridX value=" & Char34 & AppConfig.GridX & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<GridY value=" & Char34 & AppConfig.GridY & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<GridColor value=" & Char34 & AppConfig.GridColor & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<MovableAtRunTime value=" & Char34 & AppConfig.MovableAtRunTime & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<ShowControlHints value=" & Char34 & AppConfig.ShowControlHints & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & "</FormDesigner>" & vbCrLf
    StrBuff = StrBuff & vbTab & "<Editor>" & vbCrLf ' Editor settings
    
    StrBuff = StrBuff & vbTab & vbTab & "<Font value=" & Char34 & AppConfig.EditFont & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<FontSize value=" & Char34 & AppConfig.EditFontSize & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<FontColour value=" & Char34 & AppConfig.EditFontColor & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<EditFontBackColor value=" & Char34 & AppConfig.EditFontBackColor & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<TabSize value=" & Char34 & AppConfig.EditTabSize & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<Codeinsight value=" & Char34 & AppConfig.EditCodeinsight & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & vbTab & "<EditMarginBar value=" & Char34 & AppConfig.EditMarginBar & Char34 & "/>" & vbCrLf
    StrBuff = StrBuff & vbTab & "</Editor>" & vbCrLf
    StrBuff = StrBuff & "</config>"
    
 
    Open App_ConfigFile For Output As #nFile
        Print #1, StrBuff
    Close #nFile
    StrBuff = ""
End Sub

Public Sub SaveDefaultSettings()
    AppConfig.EditCodeinsight = True
    AppConfig.EditFont = "Courier New"
    AppConfig.EditFontBackColor = vbWhite
    AppConfig.EditFontColor = vbBlack
    AppConfig.EditFontSize = 10
    AppConfig.EditMarginBar = True
    AppConfig.EditTabSize = 4
    AppConfig.GridColor = vbApplicationWorkspace
    AppConfig.GridX = 120
    AppConfig.GridY = 120
    AppConfig.MovableAtRunTime = False
    AppConfig.ShowControlHints = True
    AppConfig.ShowGrid = True
End Sub

Public Sub PhaseInsightList(lzFileList As String)
Dim StrLn As String
Dim iCnt As Long, iPos As Long, FuncName As String, FuncKey As String

    Erase inSightHelper
    iCnt = -1
    If Not IsFileHere(inSightHelperList & lzFileList) Then Exit Sub
    
    Open inSightHelperList & lzFileList For Input As #1
        Do While Not EOF(1)
            Input #1, StrLn
            
            iCnt = iCnt + 1
            ReDim Preserve inSightHelper(iCnt)
            
            iPos = InStr(1, StrLn, "~", vbBinaryCompare)
            If iPos > 0 Then
                FuncName = Mid(StrLn, 1, iPos - 1)
                FuncKey = ReplaceChr(FuncName & Mid(StrLn, iPos + 1, Len(StrLn)), ".", ",")
                inSightHelper(iCnt).tName = FuncName
                inSightHelper(iCnt).tKey = FuncKey
            End If
            DoEvents
        Loop
    Close #1
    
    iCnt = 0
    iPos = 0
    FuncName = ""
    FuncKey = ""
    StrLn = ""
End Sub

Function GetItemKey(ItemName As String) As String
Dim Cnt As Integer
    For Cnt = LBound(inSightHelper) To UBound(inSightHelper)
        If LCase(ItemName) = LCase(inSightHelper(Cnt).tName) Then
            GetItemKey = inSightHelper(Cnt).tKey
            Exit For
        End If
    Next
    Cnt = 0
End Function

Public Function ItemExists(ItemName As String) As Boolean
Dim Cnt As Integer
    ItemExists = False
    For Cnt = LBound(inSightHelper) To UBound(inSightHelper)
        If LCase(ItemName) = LCase(inSightHelper(Cnt).tName) Then
            ItemExists = True
            Exit For
        End If
    Next
    Cnt = 0
End Function

Private Function GetHelperData(lStr As String) As Variant()
Dim nInfo(2) As Variant, iPos As Long
    
    iPos = InStr(1, lStr, "::", vbTextCompare)
    If iPos = 0 Then
        nInfo(0) = 0
        GetHelperData = nInfo
        Exit Function
    End If
        
    nInfo(0) = 1
    nInfo(1) = Mid(lStr, 1, iPos - 1)
    nInfo(2) = Mid(lStr, iPos + 2, Len(lStr))
    GetHelperData = nInfo
    Erase nInfo
End Function

Public Sub PhaseHelperList(nLanguage As String)
Dim HelperPath As String, HelperSelection As String, HelperCout As Integer, Counter As Integer
Dim HelperString As String
Dim HelperStrVar As Variant
Dim GotError As Boolean

    CodeHelperTemplate.CodehelperCount = -1
    Erase CodeHelperTemplate.CodehelperFilePath()
    Erase CodeHelperTemplate.CodehelperName()
    
    GotError = False
    DevXML.XMLReset
    DevXML.XMLLoadFormFile DataPath & "Code Helpers.xml"
    HelperPath = DevXML.GetSelectionValue(DevXML.XMLData, "Path", "value", "")
 
    If Len(HelperPath) = 0 Then
        MsgBox "There as an error while loading:" & vbCrLf & DataPath & "Code Helpers.xml)", vbCritical
        DevXML.XMLReset
        Exit Sub
    Else
        HelperPath = Replace(HelperPath, "{AppDataPath}", DataPath)
        If LenB(Dir(HelperPath, vbDirectory)) = 0 Then
            MsgBox "Path not Found:" & vbCrLf & HelperPath, vbCritical, "Path Not Found"
            DevXML.XMLReset
            Exit Sub
        Else
            HelperCout = CInt(DevXML.GetSelectionValue(DevXML.XMLData, "Count", "Value", "0"))
        End If
        
        HelperSelection = DevXML.GetSelection(nLanguage)
        
        If LenB(HelperSelection) = 0 Then
            MsgBox "Unable to load main Helpers Data Selection", vbInformation, "Data Selection Not Found."
            HelperPath = ""
            HelperCout = 0
            Exit Sub
        End If
        
        For Counter = 1 To HelperCout
            HelperString = Trim(DevXML.GetSelectionValue(HelperSelection, "HelperName" & CStr(Counter), "Value", ""))
            If Len(HelperString) = 0 Then GotError = True: Exit For
            HelperStrVar = GetHelperData(HelperString)
            If Not CBool(HelperStrVar(0)) Then GotError = True: Exit For
            
            If IsFileHere(HelperPath & FixPath(nLanguage) & HelperStrVar(2)) Then
                CodeHelperTemplate.CodehelperCount = CodeHelperTemplate.CodehelperCount + 1
                ReDim Preserve CodeHelperTemplate.CodehelperName(CodeHelperTemplate.CodehelperCount)
                ReDim Preserve CodeHelperTemplate.CodehelperFilePath(CodeHelperTemplate.CodehelperCount)
                CodeHelperTemplate.CodehelperName(CodeHelperTemplate.CodehelperCount) = HelperStrVar(1)
                CodeHelperTemplate.CodehelperFilePath(CodeHelperTemplate.CodehelperCount) = HelperPath & FixPath(nLanguage) & HelperStrVar(2)
            End If
           Next
    End If
    If GotError Then
        MsgBox "There as an error while loading the contents of " _
        & vbCrLf & vbCrLf & DataPath & "Code Helpers.xml", vbCritical, "Phase Error 84"
    End If
    
    DevXML.XMLReset
    HelperPath = ""
    HelperSelection = ""
    HelperCout = 0
    Counter = 0
    HelperString = ""
    Erase HelperStrVar

End Sub



