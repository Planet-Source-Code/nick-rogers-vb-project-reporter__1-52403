VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Enum VBMemberFlags
    HideThisMember = &H40
    UserInterfaceDefault = &H200
    DontShowInPropertyBrowser = &H400
    DataBound = &H4
    BindsToDataField = &H20
    ShowInDataBindingsCollection = &H10
    PropertyCallsCanPropertyChange = &H8
    UpdateImmediate = &H1000
End Enum

Private mastrDeclarations() As String
Private mastrAPI() As String
Private mastrEvents() As String
Private mastrTypes() As String
Private mastrEnums() As String
Private mastrSubroutines() As String
Private mastrFunctions() As String
Private mastrProperties() As String
Private mastrSubDescr() As String
Private mastrFunctionDescr() As String
Private mastrPropDescr() As String
Private mastrSubAttrib() As String
Private mastrFunctionAttrib() As String
Private mastrPropAttrib() As String
Private mintNumSubroutines As Integer
Private mintNumFunctions As Integer
Private mintNumProperties As Integer
Private mintNumDeclarations As Integer
Private mintNumAPI As Integer
Private mintNumTypes As Integer
Private mintNumEvents As Integer
Private mintNumEnums As Integer
Private mintPropCount As Integer
Private mstrFile As String
Private mstrPath As String
Private mstrType As String
Private mstrMTSMode As String
Private mstrPersist As String
Private mstrTitle As String
Private mstrMajor As String
Private mstrMinor As String
Private mstrRevision As String
Private mblnIncludeNAVBar As Boolean
Private mstrNAVBar As String
Private mstrFileType As String
Private mstrStyleSheetFile As String
Private mlngNumCodeLines As Long
Private mlngNumCommentLines As Long
Private mstrHeaderComments As String
Private mblnIncludeDeclarations As Boolean
Private mblnIncludeAPI As Boolean
Private mblnIncludeTypes As Boolean
Private mblnIncludeSubs As Boolean
Private mblnIncludeEvents As Boolean
Private mblnIncludeReferences As Boolean
Private mblnIncludeCounts As Boolean
Private mblnIncludeAttributes As Boolean

Public Property Get CodeLineCount() As Long
CodeLineCount = mlngNumCodeLines
End Property

Public Property Get CommentLineCount() As Long
CommentLineCount = mlngNumCommentLines
End Property

Public Property Let Filename(ByVal pstrFile As String)
mstrFile = ExtractFile(pstrFile, mstrPath)
End Property
Public Property Get Filename() As String
Filename = mstrFile
End Property

Public Property Let FilePath(ByVal pstrPath As String)
mstrPath = pstrPath
End Property

Public Property Let FileType(ByVal pstrData As String)
mstrFileType = pstrData
End Property

Public Property Let IncludeAPI(ByVal pblnData As Boolean)
mblnIncludeAPI = pblnData
End Property

Public Property Let IncludeAttributes(ByVal pblnData As Boolean)
mblnIncludeAttributes = pblnData
End Property

Public Property Let IncludeCounts(ByVal pblnData As Boolean)
mblnIncludeCounts = pblnData
End Property

Public Property Let IncludeDeclarations(ByVal pblnData As Boolean)
mblnIncludeDeclarations = pblnData
End Property

Public Property Let IncludeEvents(ByVal pblnData As Boolean)
mblnIncludeEvents = pblnData
End Property

Public Property Let IncludeNAVBar(ByVal pblnData As Boolean)
mblnIncludeNAVBar = pblnData
End Property

Public Property Let IncludeReferences(ByVal pblnData As Boolean)
mblnIncludeReferences = pblnData
End Property

Public Property Let IncludeSubs(ByVal pblnData As Boolean)
mblnIncludeSubs = pblnData
End Property

Public Property Let IncludeTypes(ByVal pblnData As Boolean)
mblnIncludeTypes = pblnData
End Property

Public Property Let NAVBar(ByVal pstrData As String)
mstrNAVBar = pstrData
End Property

Public Property Let ProjectTitle(ByVal pstrData As String)
mstrTitle = pstrData
End Property

Public Property Let ProjectVersionMajor(ByVal pstrMajor As String)
mstrMajor = pstrMajor
End Property

Public Property Let ProjectVersionMinor(ByVal pstrMinor As String)
mstrMinor = pstrMinor
End Property

Public Property Let ProjectVersionRevision(ByVal pstrRevision As String)
mstrRevision = pstrRevision
End Property

Public Property Let StyleSheetFile(ByVal pstrData As String)
mstrStyleSheetFile = pstrData
End Property

Private Function AddHTMLHeader() As String

Dim strOutput As String

strOutput = ""
strOutput = strOutput & "<html>" & vbCrLf
strOutput = strOutput & "<head>" & vbCrLf
strOutput = strOutput & "<title>" & FileOnly(mstrFile) & "</title>" & vbCrLf
strOutput = strOutput & "</head>" & vbCrLf
strOutput = strOutput & "<link rel=""stylesheet"" type=""text/css"" href=""" & IIf(mstrStyleSheetFile <> "", FileOnly(mstrStyleSheetFile), "general.css") & """>" & vbCrLf
strOutput = strOutput & "<body>" & vbCrLf
strOutput = strOutput & "<b>Project: " & mstrTitle & "</b><br>Version: " & mstrMajor & "." & mstrMinor & "." & mstrRevision & "<br>" & vbCrLf
strOutput = strOutput & "Report Date: " & Format$(Now, "dd mmm yyyy") & "<P>" & vbCrLf
strOutput = strOutput & "<b><h2>" & mstrFileType & ExtractName(mstrFile) & "</h2></b>" & vbCrLf

If InStr(mstrFile, ".cls") Then
    If mstrType <> "" Then
        strOutput = strOutput & "<b>Instancing: " & mstrType & "</b><br>" & vbCrLf
    End If
    If mstrMTSMode <> "" Then
        strOutput = strOutput & "<b>MTS Mode: " & mstrMTSMode & "</b><br>" & vbCrLf
    End If
    If mstrPersist <> "" Then
        strOutput = strOutput & "<b>Persistence: " & mstrPersist & "</b><br>" & vbCrLf
    End If
End If

strOutput = strOutput & "<p>" & vbCrLf

strOutput = strOutput & "<table class=""INTROPAGE"">" & vbCrLf
strOutput = strOutput & "<tr><td class=""INTROHEADER"">Properties</td>" & _
    "<td class=""INTROHEADER"">Subroutines</td>" & _
    "<td class=""INTROHEADER"">Functions</td></tr>" & vbCrLf
strOutput = strOutput & "<tr><td class=""INTROCELL"">" & mintPropCount & "</td>" & vbCrLf
strOutput = strOutput & "<td class=""INTROCELL"">" & mintNumSubroutines & "</td>" & vbCrLf
strOutput = strOutput & "<td class=""INTROCELL"">" & mintNumFunctions & "</td></tr></table><p>" & vbCrLf

If mblnIncludeCounts = True Then
    strOutput = strOutput & "Number of code lines: " & mlngNumCodeLines & "<br>" & vbCrLf
    strOutput = strOutput & "Number of comment lines: " & mlngNumCommentLines & "<p>" & vbCrLf
End If

If mstrHeaderComments <> "" Then
    strOutput = strOutput & mstrHeaderComments & "<p>" & vbCrLf
End If

AddHTMLHeader = strOutput

End Function

Private Function AddHTMLBody() As String

Dim strOutput As String, i As Long, j As Long
Dim strExtra() As String

strOutput = ""

If mblnIncludeNAVBar = True Then
    strOutput = strOutput & "<table class=""LAYOUT""><tr>" & vbCrLf
    strOutput = strOutput & "<td class=""LAYOUTNAV"">" & vbCrLf
    strOutput = strOutput & mstrNAVBar
    strOutput = strOutput & "</td>" & vbCrLf
    strOutput = strOutput & "<td class=""LAYOUTCELL"">" & vbCrLf
End If

If mblnIncludeDeclarations = True Then
    If mintNumDeclarations > 0 Then
        strOutput = strOutput & "<h3>Declarations</h3><a href=''></a><p>" & vbCrLf
        strOutput = strOutput & "<ul><table class=""GENERAL""><tr><td class=""HEADERBAND"">Name</td>" & _
            "<td class=""HEADERBAND"">Scope</td><td class=""HEADERBAND"">Type</td>" & _
            "<td class=""HEADERBAND"">Description</td></tr>" & vbCrLf
        For i = 0 To mintNumDeclarations - 1
            strOutput = strOutput & "<tr><td class=""CELL"">" & Split(mastrDeclarations(i), "|")(0) & "</td>"
            strOutput = strOutput & "<td class=""CELL"">" & Split(mastrDeclarations(i), "|")(1) & "</td>"
            strOutput = strOutput & "<td class=""CELL"">" & Split(mastrDeclarations(i), "|")(2) & "</td>"
            strOutput = strOutput & "<td class=""CELL"">" & Split(mastrDeclarations(i), "|")(3) & "</td></tr>"
        Next
        strOutput = strOutput & "</table><p>" & vbCrLf
    End If
    
    If mintNumTypes > 0 And mblnIncludeTypes = True Then
        If mintNumDeclarations = 0 Then
            strOutput = strOutput & "<h3>Declarations</h3><a href=''></a><p>" & vbCrLf
        End If
        For i = 0 To mintNumTypes - 1
            strOutput = strOutput & "<ul><b>TYPE: " & Split(mastrTypes(i), "|")(0) & " - " & Split(mastrTypes(i), "|")(1) & "</b></br>" & vbCrLf
            strExtra = Split(Split(mastrTypes(i), "|")(2), "~")
            strOutput = strOutput & "<ul><table class=""GENERAL""><tr><td class=""HEADERBAND"">Item Name</td>" & _
                "<td class=""HEADERBAND"">Type</td></tr>" & vbCrLf
            For j = 0 To UBound(strExtra)
                If strExtra(j) <> "" Then
                    strOutput = strOutput & "<tr><td class=""CELL"">" & Trim$(Split(strExtra(j), "^^^")(0)) & "</td>"
                    strOutput = strOutput & "<td class=""CELL"">" & Trim$(Split(strExtra(j), "^^^")(1)) & "</td></tr>"
                End If
            Next
            strOutput = strOutput & "</table></ul>" & vbCrLf
            strOutput = strOutput & "</ul><p>" & vbCrLf
        Next i
    End If
    
    If mintNumEnums > 0 And mblnIncludeTypes = True Then
        If mintNumTypes = 0 And mintNumDeclarations = 0 Then
            strOutput = strOutput & "<h3>Declarations</h3><a href=''></a><p>" & vbCrLf
        End If
        For i = 0 To mintNumEnums - 1
            strOutput = strOutput & "<ul><b>ENUM: " & Split(mastrEnums(i), "|")(0) & " - " & Split(mastrEnums(i), "|")(1) & "</b></br>" & vbCrLf
            strExtra = Split(Split(mastrEnums(i), "|")(2), "~")
            strOutput = strOutput & "<ul><table class=""GENERAL""><tr><td class=""HEADERBAND"">Item Name</td>" & _
                "<td class=""HEADERBAND"">Value</td></tr>" & vbCrLf
            For j = 0 To UBound(strExtra)
                If strExtra(j) <> "" Then
                    strOutput = strOutput & "<tr><td class=""CELL"">" & Trim$(Split(strExtra(j), "^^^")(0)) & "</td>"
                    strOutput = strOutput & "<td class=""CELL"">" & Trim$(Split(strExtra(j), "^^^")(1)) & "</td></tr>"
                End If
            Next
            strOutput = strOutput & "</table></ul>" & vbCrLf
            strOutput = strOutput & "</ul><p>" & vbCrLf
        Next i
    End If
    
    strOutput = strOutput & "</ul>" & vbCrLf
End If

If mintNumAPI > 0 And mblnIncludeAPI = True Then
    strOutput = strOutput & "<h3>API Declares</h3><a href=''></a><p>" & vbCrLf
    For i = 0 To mintNumAPI - 1
        strOutput = strOutput & ParseParamsHTML(mastrAPI(i), "")
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumEvents > 0 And mblnIncludeEvents = True Then
    strOutput = strOutput & "<h3>Event Declares</h3><a href=''></a><p>" & vbCrLf
    For i = 0 To mintNumEvents - 1
        strOutput = strOutput & ParseParamsHTML(mastrEvents(i), "")
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumProperties > 0 And mblnIncludeSubs = True Then
    strOutput = strOutput & "<h3>Properties</h3><a href=''></a><p>" & vbCrLf
    For i = 0 To mintNumProperties - 1
        strOutput = strOutput & ParseParamsHTML(mastrProperties(i), mastrPropDescr(i), mastrPropAttrib(i))
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumSubroutines > 0 And mblnIncludeSubs = True Then
    strOutput = strOutput & "<h3>Subroutines</h3><a href=''></a><p>" & vbCrLf
    For i = 0 To mintNumSubroutines - 1
        strOutput = strOutput & ParseParamsHTML(mastrSubroutines(i), mastrSubDescr(i), mastrSubAttrib(i))
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumFunctions > 0 And mblnIncludeSubs = True Then
    strOutput = strOutput & "<h3>Functions</h3><a href=''></a><p>" & vbCrLf
    For i = 0 To mintNumFunctions - 1
        strOutput = strOutput & ParseParamsHTML(mastrFunctions(i), mastrFunctionDescr(i), mastrFunctionAttrib(i))
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mblnIncludeNAVBar = True Then
    strOutput = strOutput & "</td></tr>" & vbCrLf
End If

AddHTMLBody = strOutput

End Function

Private Function AddHTMLFooter() As String

AddHTMLFooter = "</body></html>" & vbCrLf

End Function

Public Sub GetHeaderComments()

Dim strComments As String, strData As String, intInFile As Integer
Dim strHold As String, blnAbort As Boolean
Dim blnFound As Boolean

If mstrFile = "" Then Exit Sub

intInFile = FreeFile
Open mstrFile For Input As #intInFile

' get to the start of the code section in the file
' loop through the file looking for "Attribute"
' once we come to the end of the attributes, we are at the code section
Do While Not EOF(intInFile)
    Line Input #intInFile, strData
    If Left$(strData, 10) = "attribute " Then
        blnFound = True
    End If
    If Left$(strData, 10) <> "attribute " And blnFound = True Then Exit Do
Loop

Do While Not EOF(intInFile)
    'Line Input #intInFile, strData
    strHold = ""
    Do While Right$(Trim(strData), 1) = "_"
        strData = strHold & strData
        strHold = strData
        Line Input #intInFile, strData
    Loop
    strData = strHold & strData
    strData = Replace(strData, " _", " ")
    strData = Replace(strData, "  ", " ")

    ' count the code/comment lines
    If Left$(Trim$(strData), 1) = "'" Then
        mlngNumCommentLines = mlngNumCommentLines + 1
    ElseIf Trim$(strData) <> "" Then
        If Left$(strData, 10) <> "attribute " Then
            mlngNumCodeLines = mlngNumCodeLines + 1
        End If
    End If

    If Left$(Trim$(strData), 1) = "'" And blnAbort = False Then
        strComments = strComments & Trim$(Mid$(strData, 2)) & "<br>" & vbCrLf
    Else
        blnAbort = True
    End If
    
    If strData = "" And strComments = "" And mlngNumCodeLines = 0 Then blnAbort = False
    
    Line Input #intInFile, strData
Loop

Close #intInFile

mstrHeaderComments = strComments

'' remove the sub/function/property declaration code lines from the count
'mlngNumCodeLines = mlngNumCodeLines - (mintNumFunctions * 2 + mintNumSubroutines * 2 + mintNumProperties * 2)

End Sub

Public Sub ParseForm()

Dim astrSort() As String
Dim intInFile As Integer
Dim strData As String, strConstVal As String, strInLine As String, strValue As String
Dim strSubData As String, lngPos As Long, blnConfirm As Boolean
Dim strHold As String
Dim blnSubDescrOn As Boolean, blnFunctionDescrOn As Boolean, blnPropDescrOn As Boolean
Dim i As Integer
Dim blnCreatable As Boolean, blnExposed As Boolean, blnGlobal As Boolean
Dim intMTSMode As Integer, intPersist As Integer
Dim intPropCount As Integer, strTemp As String, blnFirstItem As Boolean
Dim blnDoDec As Boolean, blnDoAPI As Boolean, blnDoType As Boolean, blnDoEnum As Boolean
Dim blnDoEvent As Boolean, lngEnumCount As Long
Dim strOriginal As String, strItemName As String, strItemBreakdown As String
Dim strSplitDec() As String
Dim strCommentHold As String

If mstrFile = "" Then Exit Sub

On Error GoTo Handler

ReDim mastrSubroutines(50)
ReDim mastrFunctions(50)
ReDim mastrProperties(50)
ReDim mastrSubDescr(50)
ReDim mastrFunctionDescr(50)
ReDim mastrPropDescr(50)
ReDim mastrSubAttrib(50)
ReDim mastrFunctionAttrib(50)
ReDim mastrPropAttrib(50)
ReDim mastrDeclarations(50)
ReDim mastrAPI(50)
ReDim mastrTypes(50)
ReDim mastrEnums(50)
ReDim mastrEvents(50)

intMTSMode = -1
intPersist = -1

If InStr(mstrFile, ".") = 0 Then mstrFile = mstrFile & ".frm"

intInFile = FreeFile
Open mstrFile For Input As #intInFile

Do While Not EOF(intInFile)
    Line Input #intInFile, strData
    strHold = ""
    
    Do While Right$(Trim(strData), 1) = "_"
        strData = strHold & strData
        strHold = strData
        Line Input #intInFile, strData
    Loop
    strData = strHold & strData
    strData = Replace(strData, " _", " ")
    strData = Replace(strData, "  ", " ")
    
    ' do sub/function/property parsing
    If Left$(Trim$(strData), 1) <> "'" Then
        If Left$(Trim$(strData), 10) = "attribute " And (blnSubDescrOn = True Or blnFunctionDescrOn = True Or blnPropDescrOn = True) Then
            If blnSubDescrOn = True Then
                mastrSubAttrib(mintNumSubroutines - 1) = mastrSubAttrib(mintNumSubroutines - 1) & _
                    IIf(mastrSubAttrib(mintNumSubroutines - 1) = "", "", "~") & _
                    ProcessAttributes(strData)
            End If
            If blnFunctionDescrOn = True Then
                mastrFunctionAttrib(mintNumFunctions - 1) = mastrFunctionAttrib(mintNumFunctions - 1) & _
                    IIf(mastrFunctionAttrib(mintNumFunctions - 1) = "", "", "~") & _
                    ProcessAttributes(strData)
            End If
            If blnPropDescrOn = True Then
                mastrPropAttrib(mintNumProperties - 1) = mastrPropAttrib(mintNumProperties - 1) & _
                    IIf(mastrPropAttrib(mintNumProperties - 1) = "", "", "~") & _
                    ProcessAttributes(strData)
            End If
            strData = ""
        Else
            blnSubDescrOn = False
            blnFunctionDescrOn = False
            blnPropDescrOn = False
        End If
        
        ' check for declaration at start of file
        ' ie. we haven't come to a sub/function/prop yet
        If blnFirstItem = False Then
            blnDoDec = False
            blnDoAPI = False
            blnDoType = False
            blnDoEnum = False
            blnDoEvent = False
            strOriginal = strData
            
            ' look for a scope part
            strHold = "(Public)"
            If Left$(strData, 7) = "global " Then
                ' remove the dim/public/private item
                strData = Mid$(strData, InStr(strData, " ") + 1)
                strHold = "Public (Global)"
                blnDoDec = True
            End If
            If Left$(Trim$(strData), 4) = "dim " Then
                ' remove the dim/public/private item
                strData = Mid$(strData, InStr(strData, " ") + 1)
                strHold = "(Public)"
                blnDoDec = True
            End If
            If Left$(Trim$(strData), 8) = "private " Then
                If InStr(strData, " property ") = 0 And _
                    (InStr(strData, " sub ") = 0 Or _
                    (InStr(strData, " sub ") > 0 And _
                    InStr(strData, " declare ") > 0)) And _
                    (InStr(strData, " function ") = 0 Or _
                    (InStr(strData, " function ") > 0 And _
                    InStr(strData, " declare ") > 0)) Then
                    
                    ' remove the dim/public/private item
                    strData = Mid$(strData, InStr(strData, " ") + 1)
                    strHold = "Private"
                    blnDoDec = True
                End If
            End If
            If Left$(Trim$(strData), 7) = "public " Then
                If InStr(strData, " property ") = 0 And _
                    (InStr(strData, " sub ") = 0 Or _
                    (InStr(strData, " sub ") > 0 And _
                    InStr(strData, " declare ") > 0)) And _
                    (InStr(strData, " function ") = 0 Or _
                    (InStr(strData, " function ") > 0 And _
                    InStr(strData, " declare ") > 0)) Then
                
                    ' remove the dim/public/private item
                    strData = Mid$(strData, InStr(strData, " ") + 1)
                    strHold = "Public"
                    blnDoDec = True
                End If
            End If
            
            ' decide which type of declaration we are doing
            If Left$(strData, 5) = "type " Then
                blnDoType = True
                blnDoDec = False
            End If
            If Left$(strData, 5) = "enum " Then
                blnDoEnum = True
                blnDoDec = False
            End If
            If Left$(strData, 8) = "declare " Then
                blnDoAPI = True
                blnDoDec = False
            End If
            If Left$(strData, 6) = "event " Then
                blnDoEvent = True
                blnDoDec = False
            End If
            If Left$(Trim$(strData), 6) = "const " Then
                blnDoDec = True
            End If
            
            strSubData = ""
            
            ' process API
            If blnDoAPI = True Then
                If UBound(mastrAPI) < mintNumAPI Then
                    ReDim Preserve mastrAPI(UBound(mastrAPI) + 1)
                End If
                mastrAPI(mintNumAPI) = Replace(strOriginal, "declare ", "", , , vbTextCompare)
                mintNumAPI = mintNumAPI + 1
                strData = ""
                strCommentHold = ""
            End If
            
            ' process events
            If blnDoEvent = True Then
                If UBound(mastrEvents) < mintNumEvents Then
                    ReDim Preserve mastrEvents(UBound(mastrEvents) + 1)
                End If
                mastrEvents(mintNumEvents) = Replace(strOriginal, "event ", "sub ", , , vbTextCompare)
                mintNumEvents = mintNumEvents + 1
                strData = ""
                strCommentHold = ""
            End If
            
            ' process types
            If blnDoType = True Then
                If UBound(mastrTypes) < mintNumTypes Then
                    ReDim Preserve mastrTypes(UBound(mastrTypes) + 1)
                End If
                ' remove inline comments
                If InStr(strData, "'") > 0 Then
                    strData = Left$(strData, InStr(strData, "'") - 1)
                End If
                strItemName = Mid$(strData, InStr(strData, " ") + 1)
                strItemBreakdown = ""
                Do While Not EOF(intInFile)
                    Line Input #intInFile, strData
                    If Left$(Trim$(strData), 1) <> "'" Then
                        ' remove inline comments
                        If InStr(strData, "'") > 0 Then
                            strData = Left$(strData, InStr(strData, "'") - 1)
                        End If
                        If Trim$(strData) = "end type" Then Exit Do
                        strItemBreakdown = strItemBreakdown & IIf(strItemBreakdown = "", "", "~") & Trim$(Replace(strData, " As ", "^^^"))
                    End If
                Loop
                mastrTypes(mintNumTypes) = strItemName & "|" & strHold & "|" & strItemBreakdown
                mintNumTypes = mintNumTypes + 1
                strData = ""
                strCommentHold = ""
            End If
            
            ' process enum
            If blnDoEnum = True Then
                If UBound(mastrEnums) < mintNumEnums Then
                    ReDim Preserve mastrEnums(UBound(mastrEnums) + 1)
                End If
                ' remove inline comments
                If InStr(strData, "'") > 0 Then
                    strData = Left$(strData, InStr(strData, "'") - 1)
                End If
                strItemName = Mid$(strData, InStr(strData, " ") + 1)
                strItemBreakdown = ""
                lngEnumCount = 0
                Do While Not EOF(intInFile)
                    Line Input #intInFile, strData
                    If Left$(Trim$(strData), 1) <> "'" Then
                        ' remove inline comments
                        If InStr(strData, "'") > 0 Then
                            strData = Left$(strData, InStr(strData, "'") - 1)
                        End If
                        If Trim$(strData) = "end enum" Then Exit Do
                        If InStr(strData, " = ") = 0 Then
                            ' looks after case where enum values are implied
                            strItemBreakdown = strItemBreakdown & IIf(strItemBreakdown = "", "", "~") & Trim$(strData) & "^^^" & CStr(lngEnumCount)
                            lngEnumCount = lngEnumCount + 1
                        Else
                            ' standard explictly defined enum values
                            ' set the enum count in case the next enum item does not have a value (ie implied)
                            strValue = Mid$(strData, InStr(strData, " = ") + 3)
                            ' remove any brackets
                            strValue = Replace(strValue, ")", "")
                            strValue = Replace(strValue, "(", "")
                            ' remove trailing & symbol
                            If Right$(strValue, 1) = "&" Then
                                strValue = Left$(strValue, Len(strValue) - 1)
                            End If
                            ' if the value is not a number then ignore
                            If IsNumber(strValue) = True Then
                                lngEnumCount = CLng(strValue) + 1
                            End If
                            strItemBreakdown = strItemBreakdown & IIf(strItemBreakdown = "", "", "~") & Trim$(Replace(strData, " = ", "^^^"))
                        End If
                    End If
                Loop
                mastrEnums(mintNumEnums) = strItemName & "|" & strHold & "|" & strItemBreakdown
                mintNumEnums = mintNumEnums + 1
                strData = ""
                strCommentHold = ""
            End If
            
            ' process standard declaration
            If blnDoDec = True Then
                ' make sure we're only dealing with 1 declaration at a time
                strSplitDec = SpecialSplit(strData, ",")
                For i = 0 To UBound(strSplitDec)
                    If UBound(mastrDeclarations) < mintNumDeclarations Then
                        ReDim Preserve mastrDeclarations(UBound(mastrDeclarations) + 1)
                    End If
                    ' save the inline comments
                    strInLine = ""
                    If InStr(strSplitDec(i), "'") > 0 Then
                        strInLine = Trim$(Mid$(strSplitDec(i), InStr(strSplitDec(i), "'") + 1))
                        strSplitDec(i) = Left$(strSplitDec(i), InStr(strSplitDec(i), "'") - 1)
                    End If
                    ' remove any "As New" bits
                    strSplitDec(i) = Replace(strSplitDec(i), " As New ", " ")
                    ' remove any "WithEvents" bits
                    strSplitDec(i) = Replace(strSplitDec(i), "WithEvents ", "")
                    ' make sure we have a space char at the end
                    If Right$(strSplitDec(i), 1) <> " " Then
                        strSplitDec(i) = strSplitDec(i) & " "
                    End If
                    ' look for a const value
                    strConstVal = ""
                    If InStr(strSplitDec(i), " = ") > 0 Then
                        strConstVal = Mid$(strSplitDec(i), InStr(strSplitDec(i), " = "), InStr(InStr(strSplitDec(i), " = ") + 4, strSplitDec(i), " ") _
                            - InStr(strSplitDec(i), " = "))
                    End If
                    ' look for const
                    strSubData = ""
                    If Left$(strSplitDec(i), 6) = "const " Then
                        strSubData = "Constant "
                        strSplitDec(i) = Mid$(strSplitDec(i), 7)
                    End If
                    ' get the variable type
                    If InStr(strSplitDec(i), " as ") > 0 Then
                        strSubData = strSubData & Mid$(strSplitDec(i), InStr(strSplitDec(i), " as ") + 4, _
                            InStr(InStr(strSplitDec(i), " as ") + 4, strSplitDec(i), " ") _
                            - (InStr(strSplitDec(i), " as ") + 4))
                    
                        mastrDeclarations(mintNumDeclarations) = Left$(strSplitDec(i), InStr(strSplitDec(i), " as ") - 1) & _
                            IIf(strConstVal <> "", " " & strConstVal, "") & "|" & strHold & "|" & _
                            strSubData & "|" & strInLine
                    Else
                        If strSubData <> "Constant " Then
                            strSubData = strSubData & "Variant"
                        End If
                        
                        If InStr(strSplitDec(i), " = ") > 0 Then
                            mastrDeclarations(mintNumDeclarations) = Left$(strSplitDec(i), InStr(strSplitDec(i), " = ") - 1) & _
                                IIf(strConstVal <> "", " " & strConstVal, "") & "|" & strHold & "|" & _
                                strSubData & "|" & strInLine
                        Else
                            mastrDeclarations(mintNumDeclarations) = strSplitDec(i) & _
                                IIf(strConstVal <> "", " " & strConstVal, "") & "|" & strHold & "|" & _
                                strSubData & "|" & strInLine
                        End If
                    End If
                    
                    mintNumDeclarations = mintNumDeclarations + 1
                    
                Next
                strData = ""
                strCommentHold = ""
            End If
        End If
        
        If strData <> "" Then
            ' check for a sub
            If (InStr(strData, " sub ") > 0 Or InStr(strData, "sub ") = 1) And InStr(LTrim$(strData), "declare sub ") = 0 Then
                blnConfirm = True
                If InStr(strData, "sub ") > 1 Then
                    strSubData = Left$(strData, InStr(strData, "sub ") - 1)
                    blnConfirm = CheckForValidInfo(strSubData)
                End If
                If blnConfirm = True Then
                    If UBound(mastrSubroutines) < mintNumSubroutines Then
                        ReDim Preserve mastrSubroutines(UBound(mastrSubroutines) + 1)
                        ReDim Preserve mastrSubDescr(UBound(mastrSubroutines))
                        ReDim Preserve mastrSubAttrib(UBound(mastrSubroutines))
                    End If
                    blnSubDescrOn = True
                    mastrSubroutines(mintNumSubroutines) = strData
                    mintNumSubroutines = mintNumSubroutines + 1
                    If strCommentHold <> "" Then
                        mastrSubDescr(mintNumSubroutines - 1) = strCommentHold
                        strCommentHold = ""
                    End If
                    blnFirstItem = True
                End If
            End If
            
            ' check for a function
            If (InStr(strData, " function ") > 0 Or InStr(strData, "function ") = 1) And InStr(LTrim$(strData), "declare function ") = 0 Then
                blnConfirm = True
                If InStr(strData, "function ") > 1 Then
                    strSubData = Left$(strData, InStr(strData, "function ") - 1)
                    blnConfirm = CheckForValidInfo(strSubData)
                End If
                If blnConfirm = True Then
                    If UBound(mastrFunctions) < mintNumFunctions Then
                        ReDim Preserve mastrFunctions(UBound(mastrFunctions) + 1)
                        ReDim Preserve mastrFunctionDescr(UBound(mastrFunctions))
                        ReDim Preserve mastrFunctionAttrib(UBound(mastrFunctions))
                    End If
                    blnFunctionDescrOn = True
                    mastrFunctions(mintNumFunctions) = strData
                    mintNumFunctions = mintNumFunctions + 1
                    If strCommentHold <> "" Then
                        mastrFunctionDescr(mintNumFunctions - 1) = strCommentHold
                        strCommentHold = ""
                    End If
                    blnFirstItem = True
                End If
            End If
            
            ' check for a property
            If (InStr(strData, " property ") > 0 Or InStr(strData, "property ") = 1) Then
                blnConfirm = True
                If InStr(strData, "property ") > 1 Then
                    strSubData = Left$(strData, InStr(strData, "property ") - 1)
                    blnConfirm = CheckForValidInfo(strSubData)
                End If
                If blnConfirm = True Then
                    If UBound(mastrProperties) < mintNumProperties Then
                        ReDim Preserve mastrProperties(UBound(mastrProperties) + 1)
                        ReDim Preserve mastrPropDescr(UBound(mastrProperties))
                        ReDim Preserve mastrPropAttrib(UBound(mastrProperties))
                    End If
                    blnPropDescrOn = True
                    mastrProperties(mintNumProperties) = strData
                    mintNumProperties = mintNumProperties + 1
                    If strCommentHold <> "" Then
                        mastrPropDescr(mintNumProperties - 1) = strCommentHold
                        strCommentHold = ""
                    End If
                    blnFirstItem = True
                End If
            End If
            
            If InStr(strData, "attribute vb_creatable") = 1 Then
                blnCreatable = IIf(Trim$(Mid$(strData, InStr(strData, "=") + 1)) = "false", False, True)
            End If
            If InStr(strData, "attribute vb_exposed") = 1 Then
                blnExposed = IIf(Trim$(Mid$(strData, InStr(strData, "=") + 1)) = "false", False, True)
            End If
            If InStr(strData, "attribute vb_globalnamespace") = 1 Then
                blnGlobal = IIf(Trim$(Mid$(strData, InStr(strData, "=") + 1)) = "false", False, True)
            End If
            If InStr(strData, " mtstransactionmode ") = 1 Then
                intMTSMode = Val(Mid$(strData, InStr(strData, "=") + 1))
            End If
            If InStr(strData, " persistable ") = 1 Then
                intPersist = Val(Mid$(strData, InStr(strData, "=") + 1))
            End If
        End If
        strCommentHold = ""
    Else
        If blnSubDescrOn = True Then
            mastrSubDescr(mintNumSubroutines - 1) = mastrSubDescr(mintNumSubroutines - 1) & "<br>" & Trim$(Mid$(Trim$(strData), 2))
        End If
        If blnFunctionDescrOn = True Then
            mastrFunctionDescr(mintNumFunctions - 1) = mastrFunctionDescr(mintNumFunctions - 1) & "<br>" & Trim$(Mid$(Trim$(strData), 2))
        End If
        If blnPropDescrOn = True Then
            mastrPropDescr(mintNumProperties - 1) = mastrPropDescr(mintNumProperties - 1) & "<br>" & Trim$(Mid$(Trim$(strData), 2))
        End If
        If blnSubDescrOn = False And blnFunctionDescrOn = False And blnPropDescrOn = False Then
            strCommentHold = strCommentHold & "<br>" & Trim$(Mid$(Trim$(strData), 2))
        End If
    End If
Loop

Close #intInFile

If blnCreatable = False And blnExposed = True Then mstrType = "Public Not Creatable"
If blnCreatable = False And blnExposed = False Then mstrType = "Private"
If blnCreatable = True And blnExposed = True Then
    If blnGlobal = True Then
        mstrType = "Global MultiUse"
    Else
        mstrType = "MultiUse"
    End If
End If

Select Case intMTSMode
    Case 0: mstrMTSMode = "Not an MTS object"
    Case 1: mstrMTSMode = "Does not support transactions"
    Case 2: mstrMTSMode = "Requires transactions"
    Case 3: mstrMTSMode = "Uses transactions"
    Case 4: mstrMTSMode = "Requires a new transaction"
End Select

Select Case intPersist
    Case 0: mstrPersist = "Not Persistable"
    Case 1: mstrPersist = "Persistable"
End Select

' sort the lists
ReDim astrSort(UBound(mastrSubroutines))
For i = 0 To UBound(mastrSubroutines)
    astrSort(i) = GetItemName(mastrSubroutines(i))
Next i
Call SortList(astrSort(), mastrSubroutines(), mastrSubDescr(), mastrSubAttrib())

ReDim astrSort(UBound(mastrFunctions))
For i = 0 To UBound(mastrFunctions)
    astrSort(i) = GetItemName(mastrFunctions(i))
Next i
Call SortList(astrSort(), mastrFunctions(), mastrFunctionDescr(), mastrFunctionAttrib())

ReDim astrSort(UBound(mastrProperties))
For i = 0 To UBound(mastrProperties)
    astrSort(i) = GetItemName(mastrProperties(i))
Next i
Call SortList(astrSort(), mastrProperties(), mastrPropDescr(), mastrPropAttrib())

ReDim astrSort(UBound(mastrDeclarations))
For i = 0 To UBound(mastrDeclarations)
    astrSort(i) = mastrDeclarations(i)
Next i
Call SortList(astrSort(), mastrDeclarations())

ReDim astrSort(UBound(mastrTypes))
For i = 0 To UBound(mastrTypes)
    astrSort(i) = mastrTypes(i)
Next i
Call SortList(astrSort(), mastrTypes())

ReDim astrSort(UBound(mastrEnums))
For i = 0 To UBound(mastrEnums)
    astrSort(i) = mastrEnums(i)
Next i
Call SortList(astrSort(), mastrEnums())

ReDim astrSort(UBound(mastrAPI))
For i = 0 To UBound(mastrAPI)
    astrSort(i) = GetAPIItemName(mastrAPI(i))
Next i
Call SortList(astrSort(), mastrAPI())

ReDim astrSort(UBound(mastrEvents))
For i = 0 To UBound(mastrEvents)
    astrSort(i) = GetItemName(mastrEvents(i))
Next i
Call SortList(astrSort(), mastrEvents())

' count the properties
For i = 0 To UBound(mastrProperties)
    If strTemp <> GetItemName(mastrProperties(i)) And mastrProperties(i) <> "" Then
        mintPropCount = mintPropCount + 1
        strTemp = GetItemName(mastrProperties(i))
    End If
Next i

Call GetHeaderComments

Exit Sub

Handler:
Err.Raise Err.Number, IIf(Left$(Err.Source, 3) <> "cls", "clsFile.ParseForm", Err.Source), Err.Description

End Sub

Private Function ParseParamsHTML(ByVal pstrData As String, Optional ByVal pstrDescr As String, Optional ByVal pstrAttrib As String) As String

Dim astrWhole() As String
Dim astrParts() As String
Dim intPos1 As Integer
Dim intPos2 As Integer
Dim i As Integer, intStart As Integer
Dim intCount As Integer
Dim strExtra As String
Dim strType As String
Dim strOutput As String
Dim strLib As String
Dim strAlias As String
Dim strAttrib() As String

strOutput = ""
If pstrData = "" Then Exit Function

If Left$(Trim$(pstrData), 1) <> "'" Then
    
    ' remove inline comments
    If InStr(pstrData, "'") > 0 Then
        pstrData = Trim$(Left$(pstrData, InStr(pstrData, "'") - 1))
    End If
    
    If InStr(pstrData, " Lib ") > 0 Then
        ' get the library
        strLib = Mid$(pstrData, InStr(InStr(pstrData, " Lib "), pstrData, Chr$(34)) + 1, _
            InStr(InStr(InStr(pstrData, " Lib "), pstrData, Chr$(34)) + 1, pstrData, Chr$(34)) - _
            InStr(InStr(pstrData, " Lib "), pstrData, Chr$(34)) - 1)
        ' remove the library from the working string
        pstrData = Left$(pstrData, InStr(pstrData, " Lib ") - 1) & _
            Mid$(pstrData, InStr(InStr(InStr(pstrData, " Lib "), pstrData, Chr$(34)) + 1, pstrData, Chr$(34)) + 1)
    End If
    
    If InStr(pstrData, " Alias ") > 0 Then
        ' get the alias
        strAlias = Mid$(pstrData, InStr(InStr(pstrData, " Alias "), pstrData, Chr$(34)) + 1, _
            InStr(InStr(InStr(pstrData, " Alias "), pstrData, Chr$(34)) + 1, pstrData, Chr$(34)) - _
            InStr(InStr(pstrData, " Alias "), pstrData, Chr$(34)) - 1)
        ' remove the alias from the working string
        pstrData = Left$(pstrData, InStr(pstrData, " Alias ") - 1) & _
            Mid$(pstrData, InStr(InStr(InStr(pstrData, " Alias "), pstrData, Chr$(34)) + 1, pstrData, Chr$(34)) + 1)
    End If
    
    intPos1 = InStr(pstrData, "(") + 1
    intPos2 = InStrRev(pstrData, ")")
    ' look for array as return value and adjust position
    If Right$(pstrData, 2) = "()" Then
        If InStrRev(pstrData, ")", Len(pstrData) - 2) >= intPos1 Then
            intPos2 = InStrRev(pstrData, ")", Len(pstrData) - 2)
        End If
    End If
    ' if we didn't find an opening bracket, this isn't a sub/function/property, so exit.
    If intPos1 = 1 Then Exit Function
    ' if we didn't find an closing bracket, this isn't a sub/function/property, so exit.
    If intPos2 = 0 Then Exit Function
    
    astrWhole = Split(Mid$(pstrData, intPos1, intPos2 - intPos1), ",")
    For i = 0 To UBound(astrWhole)
        astrWhole(i) = Trim$(astrWhole(i))
    Next i
    If InStr(pstrData, "property ") > 0 Then
        If InStr(pstrData, " set ") > 0 Then strExtra = "[Set]"
        If InStr(pstrData, " let ") > 0 Then strExtra = "[Let]"
        If InStr(pstrData, " get ") > 0 Then strExtra = "[Get]"
    End If
    
    ' work out the scope
    If InStr(pstrData, "private") > 0 Then
        strExtra = strExtra & " - Private"
    ElseIf InStr(pstrData, "friend") > 0 Then
        strExtra = strExtra & " - Friend"
    Else
        strExtra = strExtra & " - Public"
    End If
    
    ' work out the return data type
    If InStr(pstrData, "Function ") > 0 Or InStr(pstrData, "property get ") > 0 Then
        intStart = -1
        If Right$(pstrData, 2) = "()" Then intStart = Len(pstrData) - 2
        If InStrRev(pstrData, ")", intStart) = Len(Trim$(pstrData)) Then
            strExtra = LTrim$(strExtra & " ") & "(Return Datatype: Variant)"
        Else
            strExtra = LTrim$(strExtra & " ") & "(Return Datatype: " & Mid$(pstrData, InStrRev(pstrData, " ") + 1) & ")"
        End If
    End If
    
    ' add in the API items, if found
    If Trim$(strLib) <> "" Or Trim$(strAlias) <> "" Then
        pstrData = Replace(pstrData, " (", "(")
        strExtra = strExtra & IIf(Trim$(strLib) <> "", " Lib: " & strLib, "") & _
            IIf(Trim$(strAlias) <> "", " Alias: " & strAlias, "")
    End If
    
    strOutput = strOutput & "<ul><b>" & GetItemName(pstrData) & IIf(strExtra = "", "", "  " & strExtra) & "</b><br>" & vbCrLf
    
    If Trim$(pstrDescr) <> "" Then
        strOutput = strOutput & "<ul>" & pstrDescr & "</ul><p>" & vbCrLf
    End If
    
    If UBound(astrWhole) = -1 Then
        ' no parameters for function
        strOutput = strOutput & "<ul><em>No Parameters</em>" & vbCrLf
    
    Else
        ' function has paramters
        strOutput = strOutput & "<ul><table class=""GENERAL""><tr><td class=""HEADERBAND"">Optional</td>" & _
            "<td class=""HEADERBAND"">ByVal/ByRef</td><td class=""HEADERBAND"">Variable</td>" & _
            "<td class=""HEADERBAND"">Datatype</td></tr>" & vbCrLf
        
        For i = 0 To UBound(astrWhole)
            
            ' process any default value
            ' put brackets around the value
            If InStr(astrWhole(i), " = ") > 0 Then
                astrWhole(i) = Replace(astrWhole(i), " = ", "(") & ")"
            End If
            
            astrParts = SpecialSplit(Trim(astrWhole(i)), " ")
            If UBound(astrParts) >= 0 Then
                intCount = 0
                strType = ""
                strOutput = strOutput & "<tr>" & vbCrLf
                If astrParts(intCount) = "Optional" Then
                    strOutput = strOutput & "<td class=""CELL"">Optional</td>" & vbCrLf
                    intCount = intCount + 1
                Else
                    strOutput = strOutput & "<td></td>" & vbCrLf
                End If
                If astrParts(intCount) = "ByVal" Or astrParts(intCount) = "ByRef" Then
                    strOutput = strOutput & "<td class=""CELL"">" & astrParts(intCount) & "</td>" & vbCrLf
                    If UBound(astrParts) > intCount Then intCount = intCount + 1
                Else
                    strOutput = strOutput & "<td></td>" & vbCrLf
                End If
                If astrParts(intCount) = "As" Then
                    intCount = intCount + 1
                Else
                    Select Case Right$(Trim$(astrParts(intCount)), 1)
                        Case "!": strType = "(Single)"
                        Case "#": strType = "(Double)"
                        Case "%": strType = "(Integer)"
                        Case "&": strType = "(Long)"
                        Case "$": strType = "(String)"
                        Case "@": strType = "(Currency)"
                    End Select
                End If
                strOutput = strOutput & "<td class=""CELL"">" & astrParts(intCount) & "</td>" & vbCrLf
                If UBound(astrParts) > intCount Then
                    intCount = intCount + 1
                    If astrParts(intCount) = "As" Then intCount = intCount + 1
                    strOutput = strOutput & "<td class=""CELL"">" & astrParts(intCount) & "</td>" & vbCrLf
                End If
                If strType <> "" Then
                    strOutput = strOutput & "<td class=""CELL"">" & strType & "</td>" & vbCrLf
                End If
                strOutput = strOutput & "</tr>" & vbCrLf
            End If
        Next i
    
        strOutput = strOutput & "</table>" & vbCrLf
    End If

    ' process any attributes for this procedure
    If Trim$(pstrAttrib) <> "" And mblnIncludeAttributes = True Then
        strOutput = strOutput & "<br><br><b>Procedure Attributes</b><br>" & vbCrLf
        strOutput = strOutput & "<table class=""GENERAL""><tr><td class=""HEADERBAND"">Attribute</td>" & _
            "<td class=""HEADERBAND"">Value</td></tr>" & vbCrLf
        
        strAttrib = Split(pstrAttrib, "~")
        For i = 0 To UBound(strAttrib)
            strOutput = strOutput & "<tr><td class=""CELL"">" & Split(strAttrib(i), "|")(0) & "</td>" & _
                "<td class=""CELL"">" & Split(strAttrib(i), "|")(1) & "</td></tr>" & vbCrLf
        Next i
        strOutput = strOutput & "</table>" & vbCrLf
    End If
    
    strOutput = strOutput & "</ul></ul>" & vbCrLf
End If

ParseParamsHTML = strOutput

End Function

Private Function ProcessAttributes(ByVal pstrData As String) As String

Dim strReturn As String
Dim strValue As String
Dim strSplit() As String

strValue = Trim$(AfterEqual(pstrData))

If InStr(pstrData, ".VB_Description") > 0 Then
    strReturn = "Description|" & strValue
End If

If InStr(pstrData, ".VB_HelpID") > 0 Then
    strReturn = "Help Context ID|" & strValue
End If

If InStr(pstrData, ".VB_MemberFlags") > 0 Then
    strReturn = ProcessMemberFlags(CLng("&H" & strValue))
End If

If InStr(pstrData, ".VB_UserMemID") > 0 Then
    strReturn = ProcessProcedureID(CLng(strValue))
End If

If InStr(pstrData, ".VB_ProcData.VB_Invoke_Property") > 0 Then
    strSplit = Split(strValue, ";")
    If UBound(strSplit) > 0 Then
        strReturn = "Property Browser Page|" & strSplit(0) & "," & _
            "Property Category|" & strSplit(1)
    Else
        strReturn = "Property Browser Page|" & strSplit(0)
    End If
End If

ProcessAttributes = strReturn

End Function

Public Function ProcessMemberFlags(ByVal plngData As Long) As String

Dim strReturn As String
Dim blnMemberFlags(7) As Boolean

If plngData >= VBMemberFlags.UpdateImmediate Then
    blnMemberFlags(7) = True
    plngData = plngData - VBMemberFlags.UpdateImmediate
End If
If plngData >= VBMemberFlags.DontShowInPropertyBrowser Then
    blnMemberFlags(6) = True
    plngData = plngData - VBMemberFlags.DontShowInPropertyBrowser
End If
If plngData >= VBMemberFlags.UserInterfaceDefault Then
    blnMemberFlags(5) = True
    plngData = plngData - VBMemberFlags.UserInterfaceDefault
End If
If plngData >= VBMemberFlags.HideThisMember Then
    blnMemberFlags(4) = True
    plngData = plngData - VBMemberFlags.HideThisMember
End If
If plngData >= VBMemberFlags.BindsToDataField Then
    blnMemberFlags(3) = True
    plngData = plngData - VBMemberFlags.BindsToDataField
End If
If plngData >= VBMemberFlags.ShowInDataBindingsCollection Then
    blnMemberFlags(2) = True
    plngData = plngData - VBMemberFlags.ShowInDataBindingsCollection
End If
If plngData >= VBMemberFlags.PropertyCallsCanPropertyChange Then
    blnMemberFlags(1) = True
    plngData = plngData - VBMemberFlags.PropertyCallsCanPropertyChange
End If
If plngData >= VBMemberFlags.DataBound Then
    blnMemberFlags(0) = True
    plngData = plngData - VBMemberFlags.DataBound
End If

If blnMemberFlags(4) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Attributes|Hide this member"
End If
If blnMemberFlags(5) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Attributes|User Interface Default"
End If
If blnMemberFlags(6) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Attributes|Don't show in property browser"
End If
If blnMemberFlags(0) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Data binding|Data bound"
End If
If blnMemberFlags(3) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Data binding|Property binds to data field"
End If
If blnMemberFlags(2) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Data binding|Show in DataBindings collection at run time"
End If
If blnMemberFlags(1) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Data binding|Property will call CanPropertyChange before changing"
End If
If blnMemberFlags(7) = True Then
    strReturn = strReturn & IIf(strReturn = "", "", ",") & "Data binding|Update Immediate"
End If
    
ProcessMemberFlags = strReturn

End Function

Private Function ProcessProcedureID(ByVal plngData As Long) As String

Dim strReturn As String

Select Case plngData
    Case 0: strReturn = "Default"
    Case -520: strReturn = "Appearance"
    Case -500: strReturn = "Autosize"
    Case -501: strReturn = "BackColor"
    Case -502: strReturn = "BackStyle"
    Case -503: strReturn = "BorderColor"
    Case -504: strReturn = "BorderStyle"
    Case -505: strReturn = "BorderWidth"
    Case -518: strReturn = "Caption"
    Case -507: strReturn = "DrawMode"
    Case -508: strReturn = "DrawStyle"
    Case -509: strReturn = "DrawWidth"
    Case -514: strReturn = "Enabled"
    Case -510: strReturn = "FillColor"
    Case -511: strReturn = "FillStyle"
    Case -512: strReturn = "Font"
    Case -513: strReturn = "ForeColor"
    Case -515: strReturn = "hWnd"
    Case -516: strReturn = "TabStop"
    Case -517: strReturn = "Text"
    Case -611: strReturn = "RightToLeft"
End Select

If strReturn <> "" Then ProcessProcedureID = "Procedure ID|" & strReturn

End Function

Public Sub SaveHTML(ByVal pstrPath As String)

Dim intFileNum As Integer

intFileNum = FreeFile

Open pstrPath & IIf(Right$(pstrPath, 1) = "\", "", "\") & FileOnly(ExtractFile(mstrFile, "")) & ".html" For Output As #intFileNum

Print #intFileNum, AddHTMLHeader
Print #intFileNum, AddHTMLBody
Print #intFileNum, AddHTMLFooter

Close #intFileNum

End Sub

