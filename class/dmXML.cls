VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_XMLData As String
Private m_XMLVersion As String
Private m_Encoding As String
Private m_Loaded As Boolean
Private m_Error As Integer

Private Type TSelections
    Count As Integer
    XMLSelectionNameA() As String
End Type
' MY XML Reader Note only reads simple XML and for this project that all I require
Private XMLSelections As TSelections

Private Sub LastError()
    If Not XMLLoaded Then
        Err.Raise 1001, , "XML Data Not Found, or Not Loaded"
        Exit Sub
    End If
End Sub

Private Sub PhaseSelections()
    Dim vSelection As Variant, I As Long
    
    XMLSelections.Count = -1
    Erase XMLSelections.XMLSelectionNameA()
       
    vSelection = Split(m_XMLData, vbCrLf)
    For I = LBound(vSelection) To UBound(vSelection)
        sLine = Trim(vSelection(I))
        sLine = Replace(sLine, vbTab, "")
            
        If Left(sLine, 2) = "</" And Right(sLine, 1) = ">" Then
            XMLSelections.Count = XMLSelections.Count + 1
            ReDim Preserve XMLSelections.XMLSelectionNameA(XMLSelections.Count)
            XMLSelections.XMLSelectionNameA(XMLSelections.Count) = Mid(sLine, 3, Len(sLine) - 3)
        End If
    Next
    I = 0
    Erase vSelection
    
End Sub

Private Sub PhaseTop()
Dim vLine As Variant, tLine As Variant, sLine As String
Dim nFound As Boolean, iPos As Long

    nFound = False
    vLine = Split(m_XMLData, vbCrLf)
    
    For I = LBound(vLine) To UBound(vLine)
        sLine = Trim(vLine(I))
        If Left(sLine, 5) = "<?xml" And Right(sLine, 2) = "?>" Then
            nFound = True
            Exit For
        End If
    Next
    Erase vLine
    I = 0
    If Not nFound Then
        XMLEncoding = vbNullChar
        XMLVersion = vbNullChar
        Exit Sub
    Else
        tLine = Split(sLine, " ")
        
        For I = LBound(tLine) To UBound(tLine)
            iPos = InStr(1, tLine(I), "=", vbTextCompare)
            
            If iPos > 0 Then
                sLine = Trim(LCase(Mid(tLine(I), 1, iPos - 1)))
                Select Case sLine
                    Case "version"
                        sLine = Mid(tLine(I), iPos + 1, Len(tLine(I)))
                        If (Not Left(sLine, 1) = Chr(34)) Or (Not Right(sLine, 1) = Chr(34)) Then
                            XMLVersion = vbNullChar
                            sLine = ""
                            Erase tLine
                            Exit Sub
                        Else
                            XMLVersion = Mid(sLine, 2, Len(sLine) - 2)
                        End If
                    Case "encoding"
                        sLine = Mid(tLine(I), iPos + 1, Len(tLine(I)))
                        If (Not Left(sLine, 1) = Chr(34)) Or (Not Right(sLine, 1) = Chr(34)) Then
                            XMLVersion = vbNullChar
                            sLine = ""
                            Erase tLine
                            Exit Sub
                        Else
                            XMLEncoding = Mid(sLine, 2, Len(sLine) - 2)
                        End If
                End Select
            End If
        Next
    End If
    
    Erase tLine
    iPos = 0
    sLine = ""
    I = 0
End Sub
Function GetSelectionValue(XMLDOC As String, ValueName As String, ValueKey As String, Optional DefaultValue As String = " ") As String
Dim sTag1 As String, sTag2 As String, sTagEnd As String, sLine As String
Dim iPos As Long, nPos As Long, I As Long
Dim vLine As Variant
Dim sFound As Boolean
On Error Resume Next

    LastError
    sTag1 = "<" & ValueName
    sTagEnd = "/>" ' end of the the selection line
    sTag2 = ValueKey & "="
    
    vLine = Split(XMLDOC, vbCrLf)
    
    sFound = False
    
    For I = LBound(vLine) To UBound(vLine)
        iPos = InStr(1, CStr(vLine(I)), sTag1, vbTextCompare)
        If iPos > 0 Then
            sFound = True
            sLine = Trim(vLine(I)) ' store and remove any spaces in the string
            Exit For
        End If
    Next
    
    Erase vLine
    iPos = 0
    I = 0
    
    If Not sFound Then
        HasError = HasError + 1
        GetSelectionValue = DefaultValue
        Exit Function
    ElseIf Not Right(sLine, 2) = sTagEnd Then
        HasError = HasError + 1
        GetSelectionValue = DefaultValue
        sTag1 = "": sTag2 = "": sTagEnd = "": sLine = ""
        iPos = 0:  EndTag = 0
        Exit Function
    Else
        sLine = Left(sLine, Len(sLine) - Len(sTagEnd))
        nPos = InStr(1, sLine, ValueKey, vbTextCompare)
        If Not (nPos > 0) Then
            GetSelectionValue = DefaultValue
            HasError = HasError + 1
            sTag1 = "": sTag2 = "": sTagEnd = "": sLine = ""
            iPos = 0: nPos = 0: EndTag = 0
            Exit Function
        Else
            sLine = Trim(Mid(sLine, nPos + Len(sTag2), Len(sLine)))
            If (Not Left(sLine, 1) = Chr(34)) Or (Not Right(sLine, 1) = Chr(34)) Then
                GetSelectionValue = DefaultValue
                HasError = HasError + 1
                Exit Function
            End If
            
            If Not (Left(sLine, 1) = Chr(34)) Or Not (Right(sLine, 1) = Chr(34)) Then
                GetSelectionValue = DefaultValue
                HasError = HasError + 1
                sTag1 = "": sTag2 = "": sTagEnd = "": sLine = ""
                iPos = 0: nPos = 0: EndTag = 0
                Exit Function
            ElseIf Len(Mid(sLine, 2, Len(sLine) - 2)) = 0 Then
                GetSelectionValue = DefaultValue
                HasError = HasError + 1
            Else
                GetSelectionValue = Mid(sLine, 2, Len(sLine) - 2)
            End If
        End If
    End If
    
End Function

Function SelectionExists(SelectionName As String) As Boolean
Dim iPos As Long, lpos As Long
    
    LastError
    SelectionExists = False
    
    sTag1 = "<" & SelectionName & ">"
    sTag2 = "</" & SelectionName & ">"
    iPos = InStr(1, XMLData, sTag1, vbTextCompare)
    lpos = InStr(iPos + 1, XMLData, sTag2, vbTextCompare)
    
    If (iPos > 0) And (lpos > 0) Then SelectionExists = True: Exit Function
    iPos = 0: lpos = 0
End Function

Function GetSelection(SelectionName As String) As String
Dim iPos As Long, lpos As Long
Dim sTag1 As String, sTag2 As String, sBuff As String

    LastError
    
    sTag1 = "<" & SelectionName & ">"
    sTag2 = "</" & SelectionName & ">"
    iPos = InStr(1, XMLData, sTag1, vbTextCompare)
    lpos = InStr(iPos + 1, XMLData, sTag2, vbTextCompare)
    
    If (iPos > 0) And (lpos > 0) Then
        sBuff = Mid(XMLData, iPos + Len(sTag1), lpos - Len(sTag2) - iPos)
        If Left(sBuff, 2) = vbCrLf Then sBuff = Right(sBuff, Len(sBuff) - 2)
        GetSelection = sBuff
    Else
        HasError = HasError + 1
        GetSelection = ""
    End If
    
    sTag1 = "": sTag2 = "": sBuff = ""
    iPos = 0: lpos = 0
    
End Function


Public Sub XMLLoadFormFile(Filename As String)
Dim nFile As Long, sBuff As String
    
    If CBool(Dir(Filename) <> "") = False Then
        Err.Raise "53", , "File Not Found"
        Exit Sub
    End If
    
    nFile = FreeFile
    Open Filename For Binary As #nFile
        sBuff = Space(LOF(nFile))
        Get #nFile, , sBuff
    Close #nFile
    XMLData = sBuff
    m_Loaded = True

End Sub

Public Property Get XMLData() As String
    XMLData = m_XMLData
End Property

Public Property Let XMLData(ByVal vNewValue As String)
    m_XMLData = vNewValue
    m_Loaded = True
    PhaseTop
    PhaseSelections
End Property

Public Sub XMLReset()
    HasError = 0
    m_XMLData = ""
    m_Loaded = False
End Sub

Public Property Get XMLVersion() As String
    LastError
    XMLVersion = m_XMLVersion
End Property

Public Property Let XMLVersion(ByVal vNewValue As String)
    m_XMLVersion = vNewValue
End Property

Public Property Get XMLEncoding() As String
    LastError
    XMLEncoding = m_Encoding
End Property

Public Property Let XMLEncoding(ByVal vNewValue As String)
    m_Encoding = vNewValue
End Property

Public Property Get XMLLoaded() As Boolean
    XMLLoaded = m_Loaded
End Property

Public Function XMLSelectionCount() As Integer
    LastError
    XMLSelectionCount = XMLSelections.Count
End Function

Public Function XMLSelectionName(Index As Integer) As String
    LastError
    XMLSelectionName = XMLSelections.XMLSelectionNameA(Index)
End Function

Public Property Get HasError() As Integer
    HasError = m_Error
End Property

Public Property Let HasError(ByVal vNewValue As Integer)
    m_Error = vNewValue
End Property
