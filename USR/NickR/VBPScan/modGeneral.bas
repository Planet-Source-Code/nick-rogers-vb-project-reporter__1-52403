Attribute VB_Name = "modGeneral"
Option Explicit

Public Enum OutputType
    HTML = 0
    HTMLHelp = 1
End Enum

Public Type Struct_BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    mstrTitle As String
    Flags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As Struct_BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal lngItem As Long, ByVal sDir As String) As Long

Public Function AfterEqual(ByVal pstrData As String) As String
' Grabs everything after the equals

AfterEqual = RemoveQuotes(Mid$(pstrData, InStr(pstrData, "=") + 1))

End Function

Public Function CheckForValidInfo(ByVal pstrData As String) As Boolean

Dim blnCheck As Boolean

blnCheck = False

If InStr(LCase$(pstrData), "public ") > 0 Then blnCheck = True
If InStr(LCase$(pstrData), "private ") > 0 Then blnCheck = True
If InStr(LCase$(pstrData), "friend ") > 0 Then blnCheck = True
If Trim$(pstrData) = "" Then blnCheck = True

CheckForValidInfo = blnCheck

End Function

Public Function SpecialSplit(ByVal pstrData As String, ByVal pstrSep As String) As String()

Dim i As Long, blnOpenItem As Boolean, lngPos As Long
Dim lngCount As Long
Dim strReturn() As String

ReDim strReturn(0)

blnOpenItem = False
lngPos = 1
lngCount = 0

If InStr(pstrData, pstrSep) = 0 Then
    strReturn(0) = pstrData
Else
    pstrData = pstrData & pstrSep
    For i = 1 To Len(pstrData)
        Select Case Mid$(pstrData, i, 1)
        Case "(", Chr$(34)
            blnOpenItem = True
        Case ")", Chr$(34)
            blnOpenItem = False
        Case pstrSep
            If blnOpenItem = False Then
                If UBound(strReturn) < lngCount Then
                    ReDim Preserve strReturn(lngCount)
                End If
                strReturn(lngCount) = Trim$(Mid$(pstrData, lngPos, i - lngPos))
                lngPos = i + 1
                lngCount = lngCount + 1
            End If
        End Select
    Next i
    If lngCount = 0 Then
        strReturn(0) = Left$(pstrData, Len(pstrData) - 1)
    End If
End If

SpecialSplit = strReturn

End Function

Public Sub CompileHTMLHelp(ByVal pstrHelpProject As String, ByVal pstrCompiler As String)
' Runs the HTML Help compiler file to create the .CHM file

Dim strEXE As String

If FileExists(pstrCompiler) = False Then
    MsgBox "Help Compiler not found!", vbExclamation, "File Not Found"
    Exit Sub
End If

strEXE = pstrCompiler & " " & pstrHelpProject

Shell strEXE, vbMaximizedFocus

End Sub

Public Function ExtractFile(ByVal pstrData As String, ByVal pstrVBPPath As String) As String

Dim intCount As Integer, intPos As Integer
Dim strSubPath As String, i As Integer

If InStr(pstrData, ";") > 0 Then
    ExtractFile = Trim(Right$(pstrData, Len(pstrData) - InStr(pstrData, ";")))
Else
    ExtractFile = Trim$(pstrData)
End If

If InStr(ExtractFile, "\") > 0 Then
    If InStr(ExtractFile, ":") > 0 Then
        ' do nothing - the full path has been specified
    ElseIf InStr(ExtractFile, "..") > 0 Then
        intCount = 0
        intPos = 1
        Do
            If InStr(intPos, ExtractFile, "..") > 0 Then
                intCount = intCount + 1
                intPos = InStr(intPos, ExtractFile, "..") + 2
            End If
        Loop Until InStr(intPos, ExtractFile, "..") = 0
        strSubPath = Left$(pstrVBPPath, Len(pstrVBPPath) - 1)
        For i = 1 To intCount
            strSubPath = Left$(strSubPath, InStrRev(strSubPath, "\") - 1)
        Next i
        ExtractFile = strSubPath & Mid$(ExtractFile, InStrRev(ExtractFile, "..") + 2)
    Else
        ExtractFile = pstrVBPPath & IIf(Right$(pstrVBPPath, 1) = "\", "", "\") & ExtractFile
    End If
Else
    ExtractFile = pstrVBPPath & IIf(Right$(pstrVBPPath, 1) = "\", "", "\") & ExtractFile
End If

End Function

Public Function ExtractName(ByVal pstrData As String) As String

If InStr(pstrData, ";") = 0 Then
    ExtractName = pstrData
Else
    ExtractName = Trim$(Left$(pstrData, InStr(pstrData, ";") - 1))
End If

If InStr(ExtractName, "\") > 0 Then
    ExtractName = Mid$(ExtractName, InStrRev(ExtractName, "\") + 1)
End If
If InStrRev(ExtractName, ".") > 0 Then
    ExtractName = Left$(ExtractName, InStr(ExtractName, ".") - 1)
End If

End Function

Public Function FileExists(ByVal sFilename As String) As Boolean

' this function checks that a file exists

Dim i As Integer

On Error Resume Next

i = Len(Dir$(sFilename))
If Err Or i = 0 Then
    FileExists = False
Else
    FileExists = True
End If

On Error GoTo 0

End Function

Public Function FileOnly(ByVal pstrFullFile As String) As String
' Returns the filename part from a fully defined file (ie file with path info)

If InStr(pstrFullFile, "\") = 0 Then
    FileOnly = pstrFullFile
Else
    FileOnly = Mid$(pstrFullFile, InStrRev(pstrFullFile, "\") + 1)
End If

End Function

Public Function GetFolder(pstrTitle As String, pfrmOwnerForm As Form) As String

Dim browse_info As Struct_BrowseInfo
Dim lngItem As Long
Dim strDirName As String

browse_info.hwndOwner = pfrmOwnerForm.hWnd
browse_info.pidlRoot = 0
browse_info.sDisplayName = Space$(260)
browse_info.mstrTitle = pstrTitle
browse_info.Flags = 1
browse_info.lpfn = 0
browse_info.lParam = 0
browse_info.iImage = 0

lngItem = SHBrowseForFolder(browse_info)
If lngItem Then
    strDirName = Space$(260)
    If SHGetPathFromIDList(lngItem, strDirName) Then
        GetFolder = Left$(strDirName, InStr(strDirName, Chr$(0)) - 1)
    Else
        GetFolder = ""
    End If
End If

End Function

Public Function GetItemName(ByVal pstrData As String) As String

If InStr(pstrData, "(") = 0 Then Exit Function
GetItemName = Mid$(pstrData, InStrRev(pstrData, " ", InStr(pstrData, "(")) + 1, _
    InStr(pstrData, "(") - InStrRev(pstrData, " ", InStr(pstrData, "(")) - 1)

End Function

Public Function GetAPIItemName(ByVal pstrData As String) As String

If InStr(LCase$(pstrData), "(") = 0 Then Exit Function
GetAPIItemName = Mid$(pstrData, InStrRev(pstrData, " ", InStr(LCase$(pstrData), " lib ") - 1) + 1, _
    InStr(pstrData, " Lib ") - InStrRev(pstrData, " ", InStr(LCase$(pstrData), " lib ") - 1) - 1)

End Function

Public Sub OutputStyleSheet(ByVal pstrPath As String, Optional pstrFile As String)
' Outputs a standard style sheet, or copies the specified file

Dim intFileNum As Integer
Dim strOutput As String

If pstrFile <> "" Then
    FileCopy pstrFile, pstrPath & IIf(Right$(pstrPath, 1) = "\", "", "\") & FileOnly(pstrFile)
    Exit Sub
End If
    
strOutput = "<STYLE TYPE=""text/css"">" & vbCrLf
strOutput = strOutput & "A" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #003399;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "A:link" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #b03d91;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "A:visited" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #b03d91;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "A:active" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #b03d91;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "A:hover" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    COLOR: #ff6600;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: underline" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "B" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 12px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "P" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & ".copy" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & ".smlcopy" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 9px;" & vbCrLf
strOutput = strOutput & "    COLOR: #c6c3c6;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "EM" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-STYLE: italic;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "BODY" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    BACKGROUND-COLOR: white;" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TABLE.GENERAL" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    BORDER-RIGHT: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-TOP: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-LEFT: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-BOTTOM: 0px;" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TD.CELL" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    WIDTH: 260px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TD.HEADERBAND" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    WIDTH: 260px;" & vbCrLf
strOutput = strOutput & "    COLOR: white;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    BACKGROUND-COLOR: teal;" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TABLE.INTROPAGE" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    BORDER-RIGHT: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-TOP: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-LEFT: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-BOTTOM: 0px;" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TD.INTROCELL" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    COLOR: #336699;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TD.INTROHEADER" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    WIDTH: 15%;" & vbCrLf
strOutput = strOutput & "    COLOR: white;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    BACKGROUND-COLOR: teal;" & vbCrLf
strOutput = strOutput & "    TEXT-DECORATION: none" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TABLE.LAYOUT" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    BORDER-RIGHT: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-TOP: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-LEFT: 0px;" & vbCrLf
strOutput = strOutput & "    BORDER-BOTTOM: 0px;" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TD.LAYOUTNAV" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    WIDTH: 20%;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    BACKGROUND-COLOR: #eeeeee;" & vbCrLf
strOutput = strOutput & "    vertical-align: top;" & vbCrLf
strOutput = strOutput & "    padding: 10;" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "TD.LAYOUTCELL" & vbCrLf
strOutput = strOutput & "{" & vbCrLf
strOutput = strOutput & "    FONT-SIZE: 11px;" & vbCrLf
strOutput = strOutput & "    WIDTH: 100%;" & vbCrLf
strOutput = strOutput & "    FONT-FAMILY: ""Verdana"";" & vbCrLf
strOutput = strOutput & "    vertical-align: top;" & vbCrLf
strOutput = strOutput & "    padding: 10;" & vbCrLf
strOutput = strOutput & "}" & vbCrLf
strOutput = strOutput & "</STYLE>" & vbCrLf

intFileNum = FreeFile
Open pstrPath & IIf(Right$(pstrPath, 1) = "\", "", "\") & "general.css" For Output As #intFileNum
Print #intFileNum, strOutput
Close #intFileNum

End Sub

Public Function RemoveEndComment(ByVal pstrData As String) As String

RemoveEndComment = Left$(pstrData, Len(pstrData) - InStr(pstrData, "'"))

End Function

Public Function RemoveQuotes(ByVal pstrData As String) As String
' Removes quotes from a string

RemoveQuotes = Replace(pstrData, Chr(34), "")

End Function

Public Sub SortList(ByRef pstrList() As String, ByRef pstrRef() As String, Optional ByRef pstrRef2 As Variant, Optional ByRef pstrRef3 As Variant)

Dim blnCheck As Boolean
Dim strTemp As String, i As Long

Do
    blnCheck = False
    For i = 0 To UBound(pstrList) - 1
        If StrComp(pstrList(i), pstrList(i + 1), vbTextCompare) > 0 And pstrList(i + 1) <> "" Then
            blnCheck = True
            strTemp = pstrList(i)
            pstrList(i) = pstrList(i + 1)
            pstrList(i + 1) = strTemp
            strTemp = pstrRef(i)
            pstrRef(i) = pstrRef(i + 1)
            pstrRef(i + 1) = strTemp
            If IsArray(pstrRef2) = True Then
                strTemp = pstrRef2(i)
                pstrRef2(i) = pstrRef2(i + 1)
                pstrRef2(i + 1) = strTemp
            End If
            If IsArray(pstrRef3) = True Then
                strTemp = pstrRef3(i)
                pstrRef3(i) = pstrRef3(i + 1)
                pstrRef3(i + 1) = strTemp
            End If
        End If
    Next i
Loop Until blnCheck = False

End Sub

Public Function IsNumber(ByVal pstrText As String) As Boolean

Dim i As Long, blnCheck As Boolean

For i = 1 To Len(pstrText)
    If InStr("0123456789.-", Mid$(pstrText, i, 1)) = 0 Then blnCheck = True
Next i

If blnCheck = False Then IsNumber = True

End Function
