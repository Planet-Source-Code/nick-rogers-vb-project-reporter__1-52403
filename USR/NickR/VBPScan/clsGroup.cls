VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mstrFile As String
Private mstrPath As String
Private mstrOutputPath As String
Private mintNumProjects As Integer
Private mcProjects As Collection
Private mblnIncludeNAVBar As Boolean
Private mintOutputType As OutputType
Private mblnOutputStyleSheet As Boolean
Private mstrHelpTitle As String
Private mstrStyleSheetFile As String
Private mlngNumCodeLines As Long
Private mlngNumCommentLines As Long
Private mblnIncludeDeclarations As Boolean
Private mblnIncludeAPI As Boolean
Private mblnIncludeTypes As Boolean
Private mblnIncludeSubs As Boolean
Private mblnIncludeEvents As Boolean
Private mblnIncludeReferences As Boolean
Private mblnIncludeCounts As Boolean
Private mblnIncludeForms As Boolean
Private mblnIncludeModules As Boolean
Private mblnIncludeClasses As Boolean
Private mblnIncludeUserControls As Boolean
Private mblnIncludeUserDocuments As Boolean
Private mblnIncludeDesigners As Boolean
Private mblnIncludeRelatedDocs As Boolean
Private mblnIncludePropertyPages As Boolean
Private mblnIncludeAttributes As Boolean
Private mblnIncludeVersionInfo As Boolean

Public Property Let FileOutputType(ByVal pintType As OutputType)
mintOutputType = pintType
End Property

Public Property Let HelpTitle(ByVal pstrData As String)
mstrHelpTitle = pstrData
End Property

Public Property Let IncludeAPI(ByVal pblnData As Boolean)
mblnIncludeAPI = pblnData
End Property

Public Property Let IncludeAttributes(ByVal pblnData As Boolean)
mblnIncludeAttributes = pblnData
End Property

Public Property Let IncludeClasses(ByVal pblnData As Boolean)
mblnIncludeClasses = pblnData
End Property

Public Property Let IncludeCounts(ByVal pblnData As Boolean)
mblnIncludeCounts = pblnData
End Property

Public Property Let IncludeDeclarations(ByVal pblnData As Boolean)
mblnIncludeDeclarations = pblnData
End Property

Public Property Let IncludeDesigners(ByVal pblnData As Boolean)
mblnIncludeDesigners = pblnData
End Property

Public Property Let IncludeEvents(ByVal pblnData As Boolean)
mblnIncludeEvents = pblnData
End Property

Public Property Let IncludeForms(ByVal pblnData As Boolean)
mblnIncludeForms = pblnData
End Property

Public Property Let IncludeModules(ByVal pblnData As Boolean)
mblnIncludeModules = pblnData
End Property

Public Property Let IncludeNAVBar(ByVal pblnData As Boolean)
mblnIncludeNAVBar = pblnData
End Property

Public Property Let IncludePropertyPages(ByVal pblnData As Boolean)
mblnIncludePropertyPages = pblnData
End Property

Public Property Let IncludeReferences(ByVal pblnData As Boolean)
mblnIncludeReferences = pblnData
End Property

Public Property Let IncludeRelatedDocs(ByVal pblnData As Boolean)
mblnIncludeAPI = pblnData
End Property

Public Property Let IncludeSubs(ByVal pblnData As Boolean)
mblnIncludeSubs = pblnData
End Property

Public Property Let IncludeTypes(ByVal pblnData As Boolean)
mblnIncludeTypes = pblnData
End Property

Public Property Let IncludeVersionInfo(ByVal pblnData As Boolean)
mblnIncludeVersionInfo = pblnData
End Property

Public Property Let IncludeUserControls(ByVal pblnData As Boolean)
mblnIncludeUserControls = pblnData
End Property

Public Property Let IncludeUserDocuments(ByVal pblnData As Boolean)
mblnIncludeUserDocuments = pblnData
End Property

Public Property Let OutputPath(ByVal pstrData As String)
mstrOutputPath = pstrData
End Property

Public Property Let OutputStyleSheetFile(ByVal pblnData As Boolean)
mblnOutputStyleSheet = pblnData
End Property

Public Property Let StyleSheetFile(ByVal pstrData As String)
mstrStyleSheetFile = pstrData
End Property

Private Function AddHTMLHeader() As String

Dim strOutput As String

strOutput = "<html>" & vbCrLf
strOutput = strOutput & "<head>" & vbCrLf
strOutput = strOutput & "<title>Project Group: " & ExtractName(mstrFile) & "</title>" & vbCrLf
strOutput = strOutput & "</head>" & vbCrLf
strOutput = strOutput & "<link rel=""stylesheet"" type=""text/css"" href=""" & IIf(mstrStyleSheetFile <> "", FileOnly(mstrStyleSheetFile), "general.css") & """>" & vbCrLf
strOutput = strOutput & "<body>" & vbCrLf
strOutput = strOutput & "<h1>Project Group: " & ExtractName(mstrFile) & "</h1><br>" & vbCrLf

AddHTMLHeader = strOutput

End Function

Private Function AddHTMLBody() As String

Dim cProject As clsProject
Dim strOutput As String

If mintNumProjects > 0 Then
    strOutput = strOutput & "<h3>Projects</h3><a href=''></a><p>" & vbCrLf
    For Each cProject In mcProjects
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cProject.Filename, mstrPath)) & ".html'>" & ExtractName(cProject.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mblnIncludeCounts = True Then
    strOutput = strOutput & "Number of code lines: " & mlngNumCodeLines & "<br>" & vbCrLf
    strOutput = strOutput & "Number of comment lines: " & mlngNumCommentLines & "<p>" & vbCrLf
End If

AddHTMLBody = strOutput

End Function

Private Function AddHTMLFooter() As String

AddHTMLFooter = "</body></html>" & vbCrLf

End Function

Private Sub Class_Initialize()

Set mcProjects = New Collection

End Sub

Private Sub OutputHTMLHelpHHC()

Dim intFileNum As Integer
Dim strOutput As String
Dim cProject As clsProject
Dim i As Long

strOutput = strOutput & "<html>" & vbCrLf
strOutput = strOutput & "<!-- Sitemap 1.0 -->" & vbCrLf
strOutput = strOutput & "<object type=""text/site properties"">" & vbCrLf
strOutput = strOutput & "  <param name=""SiteType"" value=""toc"">" & vbCrLf
strOutput = strOutput & "  <param name=""Image Width"" value=""16"">" & vbCrLf
strOutput = strOutput & "  <param name=""Window Styles"" value=""0x800002"">" & vbCrLf
strOutput = strOutput & "  <param name=""ExWindow Styles"" value=""0x100"">" & vbCrLf
strOutput = strOutput & "</object>" & vbCrLf
strOutput = strOutput & "<ul>" & vbCrLf
strOutput = strOutput & "  <li><object type=""text/sitemap"">" & vbCrLf
strOutput = strOutput & "    <param name=""Name"" value=""" & ExtractName(mstrFile) & """>" & vbCrLf
strOutput = strOutput & "    <param name=""Local"" value=""" & FileOnly(mstrFile) & ".html"">" & vbCrLf
strOutput = strOutput & "  </object>" & vbCrLf
strOutput = strOutput & "  <ul>" & vbCrLf

For Each cProject In mcProjects
    strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "      <param name=""Name"" value=""" & ExtractName(cProject.Filename) & """>" & vbCrLf
    strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.Filename, "")) & ".html"">" & vbCrLf
    strOutput = strOutput & "    </object>" & vbCrLf
    
    If cProject.FormsCount > 0 And mblnIncludeForms = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""Forms"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.FormsCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.Forms(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.Forms(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
        
    If cProject.ModulesCount > 0 And mblnIncludeModules = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""Modules"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.ModulesCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.Modules(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.Modules(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
        
    If cProject.ClassesCount > 0 And mblnIncludeClasses = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""Classes"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.ClassesCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.Classes(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.Classes(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
    
    If cProject.UserControlsCount > 0 And mblnIncludeUserControls = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""User Controls"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.UserControlsCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.UserControls(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.UserControls(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
        
    If cProject.PropertyPagesCount > 0 And mblnIncludePropertyPages = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""Property Pages"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.PropertyPagesCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.PropertyPages(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.PropertyPages(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
        
    If cProject.DesignersCount > 0 And mblnIncludeDesigners = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""Designers"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.DesignersCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.Designers(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.Designers(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
        
    If cProject.UserDocumentsCount > 0 And mblnIncludeUserDocuments = True Then
        strOutput = strOutput & "    <ul>" & vbCrLf
        strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "        <param name=""Name"" value=""User Documents"">" & vbCrLf
        strOutput = strOutput & "      </object>" & vbCrLf
        strOutput = strOutput & "      <ul>" & vbCrLf
        For i = 0 To cProject.UserDocumentsCount - 1
            strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
            strOutput = strOutput & "        <param name=""Name"" value=""" & FileOnly(ExtractFile(cProject.UserDocuments(i).Filename, "")) & """>" & vbCrLf
            strOutput = strOutput & "        <param name=""Local"" value=""" & FileOnly(ExtractFile(cProject.UserDocuments(i).Filename, "")) & ".html"">" & vbCrLf
            strOutput = strOutput & "      </object>" & vbCrLf
        Next i
        strOutput = strOutput & "      </ul>" & vbCrLf
        strOutput = strOutput & "    </ul>" & vbCrLf
    End If
        
Next
strOutput = strOutput & "    </ul>" & vbCrLf
strOutput = strOutput & "</ul>" & vbCrLf
strOutput = strOutput & "</html>" & vbCrLf

intFileNum = FreeFile

Open mstrOutputPath & IIf(Right$(mstrOutputPath, 1) = "\", "", "\") & ExtractName(mstrFile) & ".HHC" For Output As #intFileNum
Print #intFileNum, strOutput
Close #intFileNum

End Sub

Private Sub OutputHTMLHelpHHP()

Dim strOutput As String
Dim intFileNum As Integer
Dim cProject As clsProject
Dim i As Long

strOutput = strOutput & "[OPTIONS]" & vbCrLf
strOutput = strOutput & "Compiled File=" & ExtractName(mstrFile) & ".chm" & vbCrLf
strOutput = strOutput & "Title=" & mstrHelpTitle & vbCrLf
strOutput = strOutput & "Contents File=" & ExtractName(mstrFile) & ".hhc" & vbCrLf
strOutput = strOutput & "Index File=" & vbCrLf
strOutput = strOutput & "Default topic=" & FileOnly(ExtractFile(mstrFile, "")) & ".html" & vbCrLf
strOutput = strOutput & "Default Window=NewWindow" & vbCrLf
strOutput = strOutput & "Error log file=" & vbCrLf
strOutput = strOutput & "Display compile progress=No" & vbCrLf
strOutput = strOutput & "Display compile notes=No" & vbCrLf
strOutput = strOutput & "Full-text search=Yes" & vbCrLf
strOutput = strOutput & "Binary Index=No" & vbCrLf
strOutput = strOutput & "Auto Index=No" & vbCrLf
strOutput = strOutput & "Enhanced decompilation=No" & vbCrLf
strOutput = strOutput & "Binary TOC=No" & vbCrLf
strOutput = strOutput & "Flat=No" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[WINDOWS]" & vbCrLf
strOutput = strOutput & "NewWindow=""" & mstrHelpTitle & """,""" & ExtractName(mstrFile) & ".hhc"","""","""","""",,,,,0x420,0,0x200e,[0,0,500,400],0x0,0x0,,0,0,0" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[FILES]" & vbCrLf
strOutput = strOutput & FileOnly(ExtractFile(mstrFile, "")) & ".html" & vbCrLf

For Each cProject In mcProjects
    strOutput = strOutput & FileOnly(ExtractFile(cProject.Filename, "")) & ".html" & vbCrLf
    
    If cProject.FormsCount > 0 And mblnIncludeForms = True Then
        For i = 0 To cProject.FormsCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.Forms(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
    If cProject.ModulesCount > 0 And mblnIncludeModules = True Then
        For i = 0 To cProject.ModulesCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.Modules(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
    If cProject.ClassesCount > 0 And mblnIncludeClasses = True Then
        For i = 0 To cProject.ClassesCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.Classes(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
    If cProject.UserControlsCount > 0 And mblnIncludeUserControls = True Then
        For i = 0 To cProject.UserControlsCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.UserControls(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
    If cProject.PropertyPagesCount > 0 And mblnIncludePropertyPages = True Then
        For i = 0 To cProject.PropertyPagesCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.PropertyPages(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
    If cProject.DesignersCount > 0 And mblnIncludeDesigners = True Then
        For i = 0 To cProject.DesignersCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.Designers(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
    If cProject.UserDocumentsCount > 0 And mblnIncludeUserDocuments = True Then
        For i = 0 To cProject.UserDocumentsCount - 1
            strOutput = strOutput & FileOnly(ExtractFile(cProject.UserDocuments(i).Filename, "")) & ".html" & vbCrLf
        Next
    End If
Next

strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[ALIAS]" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[MAP]" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[TEXT POPUPS]" & vbCrLf

intFileNum = FreeFile

Open mstrOutputPath & IIf(Right$(mstrOutputPath, 1) = "\", "", "\") & ExtractName(mstrFile) & ".HHP" For Output As #intFileNum
Print #intFileNum, strOutput
Close #intFileNum

End Sub

Public Sub ParseGroup(ByVal pstrFile As String)

Dim intFileNum As Integer
Dim strLine As String
Dim blnOK As Boolean
Dim cProject As clsProject
Dim astrSort() As String
Dim astrProjects() As String
Dim i As Long

ReDim astrProjects(0)

On Error GoTo Handler

intFileNum = FreeFile

mstrFile = pstrFile
mstrPath = Left$(mstrFile, InStrRev(mstrFile, "\"))

Open pstrFile For Input As #intFileNum

Do While Not EOF(intFileNum)
    Line Input #intFileNum, strLine
    If Left$(UCase$(strLine), 11) = "VBGROUP 5.0" Then blnOK = True
    If Left$(UCase$(strLine), 7) = "PROJECT" Then
        If UBound(astrProjects) < mintNumProjects Then
            ReDim Preserve astrProjects(UBound(astrProjects) + 1)
        End If
        astrProjects(mintNumProjects) = AfterEqual(strLine)
        mintNumProjects = mintNumProjects + 1
    End If
    If Left$(UCase$(strLine), 14) = "STARTUPPROJECT" Then
        If UBound(astrProjects) < mintNumProjects Then
            ReDim Preserve astrProjects(UBound(astrProjects) + 1)
        End If
        astrProjects(mintNumProjects) = AfterEqual(strLine)
        mintNumProjects = mintNumProjects + 1
    End If
Loop

Close #intFileNum

ReDim astrSort(UBound(astrProjects))
For i = 0 To UBound(astrProjects)
    astrSort(i) = ExtractName(astrProjects(i))
Next i
Call SortList(astrSort(), astrProjects())

For i = 0 To UBound(astrProjects)
    If Trim$(astrProjects(i)) <> "" Then
        Set cProject = New clsProject
        cProject.FilePath = mstrPath
        cProject.Filename = astrProjects(i)
        cProject.OutputPath = mstrOutputPath
        cProject.GroupItem = mstrFile
        cProject.IncludeNAVBar = mblnIncludeNAVBar
        cProject.IncludeAPI = mblnIncludeAPI
        cProject.IncludeCounts = mblnIncludeCounts
        cProject.IncludeDeclarations = mblnIncludeDeclarations
        cProject.IncludeEvents = mblnIncludeEvents
        cProject.IncludeReferences = mblnIncludeReferences
        cProject.IncludeSubs = mblnIncludeSubs
        cProject.IncludeTypes = mblnIncludeTypes
        cProject.IncludeForms = mblnIncludeForms
        cProject.IncludeClasses = mblnIncludeClasses
        cProject.IncludeDesigners = mblnIncludeDesigners
        cProject.IncludeModules = mblnIncludeModules
        cProject.IncludeRelatedDocs = mblnIncludeRelatedDocs
        cProject.IncludeUserControls = mblnIncludeUserControls
        cProject.IncludeUserDocuments = mblnIncludeUserDocuments
        cProject.IncludePropertyPages = mblnIncludePropertyPages
        cProject.IncludeAttributes = mblnIncludeAttributes
        cProject.IncludeVersionInfo = mblnIncludeVersionInfo
        cProject.FileOutputType = mintOutputType
        cProject.StyleSheetFile = mstrStyleSheetFile
        mcProjects.Add cProject
        Set cProject = Nothing
    End If
Next

For Each cProject In mcProjects
    cProject.ParseVBPFile
    mlngNumCodeLines = mlngNumCodeLines + cProject.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cProject.CommentLineCount
Next

Exit Sub

Handler:
Err.Raise Err.Number, IIf(Left$(Err.Source, 3) <> "cls", "clsGroup.ParseGroup", Err.Source), Err.Description

End Sub

Public Sub SaveHTML()

Dim intFileNum As Integer
Dim cProject As clsProject

intFileNum = FreeFile

Open mstrOutputPath & IIf(Right$(mstrOutputPath, 1) = "\", "", "\") & FileOnly(ExtractFile(mstrFile, mstrPath)) & ".html" For Output As #intFileNum

Print #intFileNum, AddHTMLHeader
Print #intFileNum, AddHTMLBody
Print #intFileNum, AddHTMLFooter

Close #intFileNum

For Each cProject In mcProjects
    cProject.SaveHTML
Next

If mblnOutputStyleSheet = True Then
    Call OutputStyleSheet(mstrOutputPath, mstrStyleSheetFile)
End If

If mintOutputType = HTMLHelp Then
    Call OutputHTMLHelpHHC
    Call OutputHTMLHelpHHP
End If

End Sub

