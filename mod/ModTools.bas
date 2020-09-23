Attribute VB_Name = "ModTools"
Enum EditOp
    nCut = 1
    nCopy
    nPaste
    nSelectAll
    nDelete
End Enum

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Function GetFileExt(lzFile As String) As String
Dim I As Integer, iPos As Integer
    For I = 1 To Len(lzFile)
        If Mid(lzFile, I, 1) = "." Then iPos = I
    Next
    I = 0
    If iPos = 0 Then GetFileExt = "": lzFile = "": Exit Function
    GetFileExt = LCase(Trim(Mid(lzFile, iPos + 1, Len(lzFile))))
    iPos = 0: lzFile = ""
End Function

Public Function RemoveControlButton(Frm As Form, dMnuPosition As Integer, En As Boolean)
Dim hMenu As Long, iRet As Long
    hMenu = GetSystemMenu(Frm.hwnd, En)
    iRet = DeleteMenu(hMenu, dMnuPosition, MF_BYPOSITION)
End Function

Public Function IsFileHere(lzFilename As String) As Boolean
    If Dir(lzFilename) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FixPath(lzPath As String) As String
    If Right(FixPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function FindDir(lzPath As String) As Boolean
    If Not Dir(FixPath(lzPath), vbDirectory) = "." Then
        FindDir = False
        Exit Function
    Else
        FindDir = True
    End If
End Function

Public Function EnablePaste() As Boolean
   If Len(Clipboard.GetText(vbCFText)) > 0 Then
        EnablePaste = True
    Else
        EnablePaste = False
    End If
End Function

Public Function EditMenu(mOption As EditOp, txtBox As TextBox)
    Select Case mOption
        Case nCut
            Clipboard.SetText txtBox.SelText
            txtBox.SelText = ""
        Case nCopy
            Clipboard.SetText txtBox.SelText
        Case nPaste
            txtBox.SelText = Clipboard.GetText(vbCFText)
        Case nSelectAll
            txtBox.SelStart = 0
            txtBox.SelLength = Len(txtBox.Text)
            txtBox.SetFocus
        Case nDelete
            txtBox.SelText = ""
    End Select
End Function

Public Sub DrawGrid(Frm As Form, Optional GridColor As Long = vbBlack, Optional ShowGrid As Boolean = True)
    If Not ShowGrid Then Set frmWorkArea.Picture = Nothing: Exit Sub
    Frm.AutoRedraw = True
    
    For x = 0 To Frm.ScaleWidth Step AppConfig.GridX
        For y = 0 To Frm.ScaleHeight Step AppConfig.GridY
            Frm.PSet (x, y), AppConfig.GridColor
        Next
    Next
    Frm.Refresh
    Frm.AutoRedraw = False
    x = 0: y = 0
End Sub

Public Function OpenFile(lzFilename As String) As String
Dim StrA As String, iFile As Long
    iFile = FreeFile
        
    Open lzFilename For Binary As #iFile
        StrA = Space(LOF(iFile))
        Get #iFile, , StrA
    Close #iFile
    
    OpenFile = StrA
    StrA = ""
End Function

Public Sub lstVBFunctions(lzCode As String, cboLst As ComboBox)
Dim iCnt As Long, Ipart, lPart As Long, x As Long, y As Long, ch As Long
Dim LnStr, StrBuff As String, FuncName As String, SubName As String, StrLn As String
On Error Resume Next
' List VB Function names Added by Ben jones

    cboLst.Clear
    cboLst.AddItem "[General]"
    StrBuff = lzCode & vbCrLf
   
    For iCnt = 1 To Len(StrBuff)
        ch = Asc(Mid$(StrBuff, iCnt, 1))
        If ch <> 13 Then
            StrLn = StrLn & Chr(ch)
        Else
            Ipart = InStr(1, StrLn, "Function ", vbTextCompare)
            lPart = InStr(1, StrLn, "(")
            If Ipart > 0 And lPart > 0 Then
                FuncName = Trim$(Mid$(StrLn, Ipart + Len("Function"), lPart - Ipart - Len("Function")))
                cboLst.AddItem " " & FuncName
            End If
            
            x = InStr(1, StrLn, "Sub ", vbTextCompare)
            y = InStr(1, StrLn, "(")
            
            If x > 0 And y > 0 Then
                SubName = Trim(Mid(StrLn, x + Len("Sub"), y - x - Len("Sub")))
                cboLst.AddItem " " & SubName
            End If
            StrLn = ""
            iCnt = iCnt + 1
        End If
    Next iCnt
    cboLst.ListIndex = 0
    iCnt = 0: Ipart = 0: lPart = 0: x = 0: y = 0
    FuncName = ""
    StrBuff = ""
    LnStr = ""
    ch = ""
    
End Sub

Function GetAbsPath(lzPath As String) As String
Dim iPos As Long, I As Long
    For I = 1 To Len(lzPath)
        If InStr(I, lzPath, "\", vbBinaryCompare) Then
            iPos = I
        End If
    Next
    
    If iPos = 0 Then
        GetAbsPath = lzPath
    Else
        GetAbsPath = Mid(lzPath, 1, iPos)
    End If
    
    iPos = 0
End Function

Public Function WriteToFile(lzFile As String, lzData As String)
Dim iFile As Long
    iFile = FreeFile
    Open lzFile For Binary As #iFile
        Put #iFile, , lzData
    Close #iFile
End Function

Sub AppendErrorLog(StrError As String)
    Open FixPath(App.Path) & "error.log" For Append As #1
        Print #1, StrError
    Close #1
End Sub

Public Function SetMargin(nMarSize As Long, mTextBox As TextBox)
    SendMessage mTextBox.hwnd, EM_SETMARGINS, EC_LEFTMARGIN, nMarSize
End Function

Public Function GetLineCount(mTextBox As TextBox) As Long
Dim vCount As Variant
    vCount = Split(mTextBox.Text, vbCrLf)
    GetLineCount = UBound(vCount)
    Erase vCount
End Function

Public Function HighLightLine(LineNumber As Long, mTextBox As TextBox)
Dim iLength As Long, lnIdx As Long

    GotoLine LineNumber, mTextBox
    
    lnIdx = SendMessage(mTextBox.hwnd, EM_LINEINDEX, LineNumber, 0)
    iLength = SendMessage(mTextBox.hwnd, EM_LINELENGTH, lnIdx, 0)
    mTextBox.SelLength = iLength
    mTextBox.SetFocus
    
End Function

Public Sub GotoLine(vNewLineNum As Long, mTextBox As TextBox)
Dim lnNum As Long
    lnNum = SendMessage(mTextBox.hwnd, EM_LINEINDEX, ByVal vNewLineNum, ByVal 0&)
    If lnNum = -1 Then lnNum = 0
    mTextBox.SelStart = lnNum
    mTextBox.SetFocus
End Sub

Function GetCurrentLineNumber(mTextBox As TextBox) As Long
    GetCurrentLineNumber = SendMessage(mTextBox.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) '+ 1
End Function

Function GetCurrentLineLength(mTextBox As TextBox) As Long
    GetCurrentLineLength = SendMessage(mTextBox.hwnd, EM_LINELENGTH, _
    SendMessage(mTextBox.hwnd, EM_LINEINDEX, GetCurrentLineNumber(mTextBox), 0), 0)
End Function

Function GetLineText(mTextBox As TextBox) As String
Dim LineNo As Long
Dim sLineText As String
Dim OldSelStart As Long
On Error Resume Next
    OldSelStart = (mTextBox.SelStart)
    LineNo = GetCurrentLineNumber(mTextBox)
    GotoLine LineNo, mTextBox
    GetLineText = Mid(mTextBox, mTextBox.SelStart + 1, GetCurrentLineLength(mTextBox) - 1)
    mTextBox.SelStart = OldSelStart
End Function

Public Function GetColumn(mTextBox As TextBox) As Long
    GetColumn = mTextBox.SelStart - SendMessage(mTextBox.hwnd, EM_LINEINDEX, GetCurrentLineNumber(mTextBox), ByVal 0&)
End Function
Public Function ReplaceChr(lzStr As String, sOldChr As String, sNewChr As String)
Dim sByte() As Byte
   
    sByte() = StrConv(lzStr, vbFromUnicode)

    For I = LBound(sByte) To UBound(sByte)
       If sByte(I) = Asc(sOldChr) Then sByte(I) = Asc(sNewChr)
    Next
    ReplaceChr = StrConv(sByte(), vbUnicode)
    
    Erase sByte()
    I = 0

End Function

Public Function Invert(StrB As String) As String
Dim I As Long
Dim C As String * 1
' little function to invert string eg This will become tHIS
    For I = 1 To Len(StrB)
        C = Mid(StrB, I, 1)
        If C = UCase(C) Then Mid(StrB, I, 1) = LCase(C)
        If C = LCase(C) Then Mid(StrB, I, 1) = UCase(C)
    Next
    
    I = 0
    C = ""
    Invert = StrB
    StrB = ""
    
End Function

