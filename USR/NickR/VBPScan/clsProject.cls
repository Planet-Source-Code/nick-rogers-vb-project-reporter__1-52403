VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mstrHelpTitle As String
Private mcForms As Collection
Private mcModules As Collection
Private mcClasses As Collection
Private mcUserControls As Collection
Private mcReferences As Collection
Private mcObjects As Collection
Private mcRelatedDocs As Collection
Private mcDesigners As Collection
Private mcUserDocuments As Collection
Private mcPropertyPages As Collection
Private mstrVBPPath As String
Private mstrVBPName As String
Private mstrOutputPath As String
Private mintNumForms As Integer
Private mintNumObjects As Integer
Private mintNumUserControls As Integer
Private mintNumReferences As Integer
Private mintNumModules As Integer
Private mintNumClasses As Integer
Private mintNumRelatedDocs As Integer
Private mintNumDesigners As Integer
Private mintNumUserDocuments As Integer
Private mintNumPropertyPages As Integer
Private mstrTitle As String
Private mstrMajor As String
Private mstrMinor As String
Private mstrRevision As String
Private mstrDescription As String
Private mstrVersionInfo(5) As String
Private mstrEXEName As String
Private mintOutputType As OutputType
Private mblnIncludeNAVBar As Boolean
Private mblnOutputStyleSheet As Boolean
Private mstrGroupItem As String
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

Public Property Get Classes(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcClasses.Count Then Exit Property
Set Classes = mcClasses(plngItem + 1)
End Property

Public Property Get ClassesCount() As Long
ClassesCount = mintNumClasses
End Property

Public Property Get CodeLineCount() As Long
CodeLineCount = mlngNumCodeLines
End Property

Public Property Get CommentLineCount() As Long
CommentLineCount = mlngNumCommentLines
End Property

Public Property Get Designers(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcDesigners.Count Then Exit Property
Set Designers = mcDesigners(plngItem + 1)
End Property

Public Property Get DesignersCount() As Long
DesignersCount = mintNumDesigners
End Property

Public Property Let Filename(ByVal pstrFile As String)
mstrVBPName = ExtractFile(pstrFile, mstrVBPPath)
End Property
Public Property Get Filename() As String
Filename = mstrVBPName
End Property

Public Property Let FileOutputType(ByVal pintType As OutputType)
mintOutputType = pintType
End Property

Public Property Let FilePath(ByVal pstrPath As String)
mstrVBPPath = pstrPath
End Property

Public Property Get Forms(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcForms.Count Then Exit Property
Set Forms = mcForms(plngItem + 1)
End Property

Public Property Get FormsCount() As Long
FormsCount = mintNumForms
End Property

Public Property Let GroupItem(ByVal pstrData As String)
mstrGroupItem = pstrData
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

Public Property Get Modules(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcModules.Count Then Exit Property
Set Modules = mcModules(plngItem + 1)
End Property

Public Property Get ModulesCount() As Long
ModulesCount = mintNumModules
End Property

Public Property Let OutputPath(ByVal pstrData As String)
mstrOutputPath = pstrData
End Property

Public Property Let OutputStyleSheetFile(ByVal pblnData As Boolean)
mblnOutputStyleSheet = pblnData
End Property

Public Property Get PropertyPages(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcPropertyPages.Count Then Exit Property
Set PropertyPages = mcPropertyPages(plngItem + 1)
End Property

Public Property Get PropertyPagesCount() As Long
PropertyPagesCount = mintNumPropertyPages
End Property

Public Property Let StyleSheetFile(ByVal pstrData As String)
mstrStyleSheetFile = pstrData
End Property

Public Property Get UserControls(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcUserControls.Count Then Exit Property
Set UserControls = mcUserControls(plngItem + 1)
End Property

Public Property Get UserControlsCount() As Long
UserControlsCount = mintNumUserControls
End Property

Public Property Get UserDocuments(ByVal plngItem As Long) As clsFile
If (plngItem + 1) > mcUserDocuments.Count Then Exit Property
Set UserDocuments = mcUserDocuments(plngItem + 1)
End Property

Public Property Get UserDocumentsCount() As Long
UserDocumentsCount = mintNumUserDocuments
End Property

Private Function AddHTMLHeader() As String

Dim strOutput As String

strOutput = "<html>" & vbCrLf
strOutput = strOutput & "<head>" & vbCrLf
strOutput = strOutput & "<title>Project " & mstrTitle & "</title>" & vbCrLf
strOutput = strOutput & "</head>" & vbCrLf
strOutput = strOutput & "<link rel=""stylesheet"" type=""text/css"" href=""" & IIf(mstrStyleSheetFile <> "", FileOnly(mstrStyleSheetFile), "general.css") & """>" & vbCrLf
strOutput = strOutput & "<body>" & vbCrLf
strOutput = strOutput & "<h1>Project: " & mstrTitle & "</h1><br>" & vbCrLf
strOutput = strOutput & "<b>EXE/DLL Name: " & mstrEXEName & "</b><br>" & vbCrLf
strOutput = strOutput & "<b>Version: " & mstrMajor & "." & mstrMinor & "." & mstrRevision & "</b><br>" & vbCrLf
strOutput = strOutput & "<b>Report Date: " & Format$(Now, "dd mmm yyyy") & "</b><P>" & vbCrLf
strOutput = strOutput & "<a href=#summary></a>" & vbCrLf
strOutput = strOutput & "<h3>Project Summary</h3>" & vbCrLf
strOutput = strOutput & "<table class=""INTROPAGE"">" & vbCrLf
strOutput = strOutput & "<tr>" & IIf(mintNumForms > 0, "<td class=""INTROHEADER"">Forms</td>", "") & _
    IIf(mintNumModules > 0, "<td class=""INTROHEADER"">Modules</td>", "") & _
    IIf(mintNumClasses > 0, "<td class=""INTROHEADER"">Classes</td>", "") & _
    IIf(mintNumUserControls > 0, "<td class=""INTROHEADER"">User Controls</td>", "") & _
    IIf(mintNumPropertyPages > 0, "<td class=""INTROHEADER"">Property Pages</td>", "") & _
    IIf(mintNumDesigners > 0, "<td class=""INTROHEADER"">Designers</td>", "") & _
    IIf(mintNumUserDocuments > 0, "<td class=""INTROHEADER"">User Documents</td>", "") & _
    IIf(mintNumObjects > 0, "<td class=""INTROHEADER"">Objects</td>", "") & _
    IIf(mintNumReferences > 0, "<td class=""INTROHEADER"">References</td>", "") & _
    IIf(mintNumRelatedDocs > 0, "<td class=""INTROHEADER"">Related Docs</td>", "") & "</td>" & vbCrLf
strOutput = strOutput & "<tr>" & IIf(mintNumForms > 0, "<td class=""INTROCELL"">" & mintNumForms & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumModules > 0, "<td class=""INTROCELL"">" & mintNumModules & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumClasses > 0, "<td class=""INTROCELL"">" & mintNumClasses & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumUserControls > 0, "<td class=""INTROCELL"">" & mintNumUserControls & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumPropertyPages > 0, "<td class=""INTROCELL"">" & mintNumPropertyPages & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumDesigners > 0, "<td class=""INTROCELL"">" & mintNumDesigners & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumUserDocuments > 0, "<td class=""INTROCELL"">" & mintNumUserDocuments & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumObjects > 0, "<td class=""INTROCELL"">" & mintNumObjects & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumReferences > 0, "<td class=""INTROCELL"">" & mintNumReferences & "</td>", "") & vbCrLf
strOutput = strOutput & IIf(mintNumRelatedDocs > 0, "<td class=""INTROCELL"">" & mintNumRelatedDocs & "</td>", "") & "</tr></table><p>" & vbCrLf

If Trim$(mstrDescription) <> "" Then
    strOutput = strOutput & "Project Description: " & mstrDescription & "<p>" & vbCrLf
End If

If mblnIncludeCounts = True Then
    strOutput = strOutput & "Number of code lines: " & mlngNumCodeLines & "<br>" & vbCrLf
    strOutput = strOutput & "Number of comment lines: " & mlngNumCommentLines & "<p>" & vbCrLf
End If

If Trim$(Join(mstrVersionInfo, "")) <> "" And mblnIncludeVersionInfo = True Then
    strOutput = strOutput & "<b>Version Information</b><ul>" & vbCrLf
    If Trim$(mstrVersionInfo(0)) <> "" Then
        strOutput = strOutput & "Comments: " & mstrVersionInfo(0) & "<br>" & vbCrLf
    End If
    If Trim$(mstrVersionInfo(1)) <> "" Then
        strOutput = strOutput & "Company Name: " & mstrVersionInfo(1) & "<br>" & vbCrLf
    End If
    If Trim$(mstrVersionInfo(2)) <> "" Then
        strOutput = strOutput & "File Description: " & mstrVersionInfo(2) & "<br>" & vbCrLf
    End If
    If Trim$(mstrVersionInfo(3)) <> "" Then
        strOutput = strOutput & "Legal Copyright: " & mstrVersionInfo(3) & "<br>" & vbCrLf
    End If
    If Trim$(mstrVersionInfo(4)) <> "" Then
        strOutput = strOutput & "Legal Trademarks: " & mstrVersionInfo(4) & "<br>" & vbCrLf
    End If
    If Trim$(mstrVersionInfo(5)) <> "" Then
        strOutput = strOutput & "Product Name: " & mstrVersionInfo(5) & "<br>" & vbCrLf
    End If
    strOutput = strOutput & "</ul><p>" & vbCrLf
End If

AddHTMLHeader = strOutput

End Function

Private Function AddHTMLBody() As String

Dim i As Integer
Dim cFile As clsFile
Dim cItem As clsItem
Dim strOutput As String

strOutput = ""

If mstrGroupItem <> "" Then
    strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(mstrGroupItem, mstrVBPPath)) & ".html'>Home</a>" & vbCrLf
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumForms > 0 And mblnIncludeForms = True Then
    strOutput = strOutput & "<h3>Forms</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcForms
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumModules > 0 And mblnIncludeModules = True Then
    strOutput = strOutput & "<h3>Modules</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcModules
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumClasses > 0 And mblnIncludeClasses = True Then
    strOutput = strOutput & "<h3>Classes</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcClasses
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumUserControls > 0 And mblnIncludeUserControls = True Then
    strOutput = strOutput & "<h3>User Controls</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcUserControls
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumPropertyPages > 0 And mblnIncludePropertyPages = True Then
    strOutput = strOutput & "<h3>Property Pages</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcPropertyPages
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumDesigners > 0 And mblnIncludeDesigners = True Then
    strOutput = strOutput & "<h3>Designers</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcDesigners
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumUserDocuments > 0 And mblnIncludeUserDocuments = True Then
    strOutput = strOutput & "<h3>User Documents</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcUserDocuments
        strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumRelatedDocs > 0 And mblnIncludeRelatedDocs = True Then
    strOutput = strOutput & "<h3>Related Docs</h3><a href=''></a><p>" & vbCrLf
    For Each cItem In mcRelatedDocs
        strOutput = strOutput & cItem.ItemReference & "<BR>" & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumObjects > 0 And mblnIncludeReferences = True Then
    strOutput = strOutput & "<h3>Objects</h3><a href=''></a><p>" & vbCrLf
    For Each cItem In mcObjects
        strOutput = strOutput & cItem.ItemReference & "<BR>" & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumReferences > 0 And mblnIncludeReferences = True Then
    strOutput = strOutput & "<h3>References</h3><a href=''></a><p>" & vbCrLf
    For Each cItem In mcReferences
        strOutput = strOutput & cItem.ItemReference & "<BR>" & vbCrLf
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

AddHTMLBody = strOutput

End Function

Private Function AddHTMLFooter() As String

AddHTMLFooter = "</body></html>" & vbCrLf

End Function

Private Sub Class_Initialize()

Set mcForms = New Collection
Set mcClasses = New Collection
Set mcModules = New Collection
Set mcUserControls = New Collection
Set mcObjects = New Collection
Set mcReferences = New Collection
Set mcRelatedDocs = New Collection
Set mcDesigners = New Collection
Set mcUserDocuments = New Collection
Set mcPropertyPages = New Collection

mintOutputType = HTML

End Sub

Private Function NAVBarHTML(ByVal pstrSkipFile As String) As String

Dim i As Integer
Dim strOutput As String
Dim cFile As clsFile

strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(mstrVBPName, mstrVBPPath)) & ".html'>Home</a>" & vbCrLf
strOutput = strOutput & "<p>" & vbCrLf

If mintNumForms > 0 And mblnIncludeForms = True Then
    strOutput = strOutput & "<h3>Forms</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcForms
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumModules > 0 And mblnIncludeModules = True Then
    strOutput = strOutput & "<h3>Modules</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcModules
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumClasses > 0 And mblnIncludeClasses = True Then
    strOutput = strOutput & "<h3>Classes</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcClasses
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumUserControls > 0 And mblnIncludeUserControls = True Then
    strOutput = strOutput & "<h3>User Controls</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcUserControls
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumPropertyPages > 0 And mblnIncludePropertyPages = True Then
    strOutput = strOutput & "<h3>Property Pages</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcPropertyPages
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumDesigners > 0 And mblnIncludeDesigners = True Then
    strOutput = strOutput & "<h3>Designers</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcDesigners
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

If mintNumUserDocuments > 0 And mblnIncludeUserDocuments = True Then
    strOutput = strOutput & "<h3>User Documents</h3><a href=''></a><p>" & vbCrLf
    For Each cFile In mcUserDocuments
        If cFile.Filename = pstrSkipFile Then
            strOutput = strOutput & ExtractName(cFile.Filename) & " <BR> " & vbCrLf
        Else
            strOutput = strOutput & "<a href='" & FileOnly(ExtractFile(cFile.Filename, mstrVBPPath)) & ".html'>" & ExtractName(cFile.Filename) & "</a> <BR> " & vbCrLf
        End If
    Next
    strOutput = strOutput & "<p>" & vbCrLf
End If

NAVBarHTML = strOutput

End Function

Private Sub OutputHTMLHelpHHC()

Dim intFileNum As Integer
Dim strOutput As String
Dim cFile As clsFile

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
strOutput = strOutput & "    <param name=""Name"" value=""" & mstrTitle & """>" & vbCrLf
strOutput = strOutput & "    <param name=""Local"" value=""" & FileOnly(ExtractFile(mstrVBPName, "")) & ".html"">" & vbCrLf
strOutput = strOutput & "  </object>" & vbCrLf

If mcForms.Count > 0 And mblnIncludeForms = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""Forms"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcForms
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If

If mcModules.Count > 0 And mblnIncludeModules = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""Modules"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcModules
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If
    
If mcClasses.Count > 0 And mblnIncludeClasses = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""Classes"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcClasses
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If

If mcUserControls.Count > 0 And mblnIncludeUserControls = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""User Controls"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcUserControls
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If
    
If mcPropertyPages.Count > 0 And mblnIncludePropertyPages = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""Property Pages"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcPropertyPages
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If
    
If mcDesigners.Count > 0 And mblnIncludeDesigners = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""Designers"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcDesigners
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If
    
If mcUserDocuments.Count > 0 And mblnIncludeUserDocuments = True Then
    strOutput = strOutput & "    <ul>" & vbCrLf
    strOutput = strOutput & "      <li><object type=""text/sitemap"">" & vbCrLf
    strOutput = strOutput & "        <param name=""Name"" value=""User Documents"">" & vbCrLf
    strOutput = strOutput & "      </object>" & vbCrLf
    strOutput = strOutput & "      <ul>" & vbCrLf
    For Each cFile In mcUserDocuments
        strOutput = strOutput & "    <li><object type=""text/sitemap"">" & vbCrLf
        strOutput = strOutput & "      <param name=""Name"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & """>" & vbCrLf
        strOutput = strOutput & "      <param name=""Local"" value=""" & FileOnly(ExtractFile(cFile.Filename, "")) & ".html"">" & vbCrLf
        strOutput = strOutput & "    </object>" & vbCrLf
    Next
    strOutput = strOutput & "      </ul>" & vbCrLf
    strOutput = strOutput & "    </ul>" & vbCrLf
End If
    
strOutput = strOutput & "</ul>" & vbCrLf
strOutput = strOutput & "</html>" & vbCrLf

intFileNum = FreeFile

Open mstrOutputPath & IIf(Right$(mstrOutputPath, 1) = "\", "", "\") & mstrTitle & ".HHC" For Output As #intFileNum
Print #intFileNum, strOutput
Close #intFileNum

End Sub

Private Sub OutputHTMLHelpHHP()

Dim strOutput As String
Dim cFile As clsFile
Dim intFileNum As Integer

strOutput = strOutput & "[OPTIONS]" & vbCrLf
strOutput = strOutput & "Compiled File=" & mstrTitle & ".chm" & vbCrLf
strOutput = strOutput & "Title=" & mstrHelpTitle & vbCrLf
strOutput = strOutput & "Contents File=" & mstrTitle & ".hhc" & vbCrLf
strOutput = strOutput & "Index File=" & vbCrLf
strOutput = strOutput & "Default topic=" & FileOnly(ExtractFile(mstrVBPName, mstrVBPPath)) & ".html" & vbCrLf
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
strOutput = strOutput & "NewWindow=""" & mstrHelpTitle & """,""" & mstrTitle & ".hhc"","""","""","""",,,,,0x420,0,0x200e,[0,0,900,600],0x0,0x0,,0,0,0" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[FILES]" & vbCrLf
strOutput = strOutput & FileOnly(ExtractFile(mstrVBPName, mstrVBPPath)) & ".html" & vbCrLf

If mintNumForms > 0 And mblnIncludeForms = True Then
    For Each cFile In mcForms
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If
If mintNumModules > 0 And mblnIncludeModules = True Then
    For Each cFile In mcModules
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If
If mintNumClasses > 0 And mblnIncludeClasses = True Then
    For Each cFile In mcClasses
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If
If mintNumUserControls > 0 And mblnIncludeUserControls = True Then
    For Each cFile In mcUserControls
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If
If mintNumPropertyPages > 0 And mblnIncludePropertyPages = True Then
    For Each cFile In mcPropertyPages
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If
If mintNumDesigners > 0 And mblnIncludeDesigners = True Then
    For Each cFile In mcDesigners
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If
If mintNumUserDocuments > 0 And mblnIncludeUserDocuments = True Then
    For Each cFile In mcUserDocuments
        strOutput = strOutput & FileOnly(ExtractFile(cFile.Filename, "")) & ".html" & vbCrLf
    Next
End If

strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[ALIAS]" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[MAP]" & vbCrLf
strOutput = strOutput & "" & vbCrLf
strOutput = strOutput & "[TEXT POPUPS]" & vbCrLf

intFileNum = FreeFile

Open mstrOutputPath & IIf(Right$(mstrOutputPath, 1) = "\", "", "\") & mstrTitle & ".HHP" For Output As #intFileNum
Print #intFileNum, strOutput
Close #intFileNum

End Sub

Public Sub ParseVBPFile(Optional ByVal pstrFilename As String)

Dim astrForm() As String
Dim astrClass() As String
Dim astrModule() As String
Dim astrObject() As String
Dim astrUserControl() As String
Dim astrReference() As String
Dim astrRelatedDocs() As String
Dim astrDesigner() As String
Dim astrUserDocument() As String
Dim astrPropertyPage() As String

Dim intInFile As Integer
Dim strData As String
Dim i As Integer
Dim astrSort() As String
Dim cFile As clsFile
Dim cItem As clsItem

ReDim astrForm(0)
ReDim astrClass(0)
ReDim astrModule(0)
ReDim astrObject(0)
ReDim astrUserControl(0)
ReDim astrReference(0)
ReDim astrRelatedDocs(0)
ReDim astrDesigner(0)
ReDim astrUserDocument(0)
ReDim astrPropertyPage(0)

If pstrFilename <> "" Then
    mstrVBPName = pstrFilename
End If
mstrVBPPath = Left$(mstrVBPName, InStrRev(mstrVBPName, "\"))

intInFile = FreeFile

On Error GoTo Handler

Open mstrVBPName For Input As #intInFile

mintNumForms = 0
mintNumObjects = 0
mintNumUserControls = 0
mintNumReferences = 0
mintNumModules = 0
mintNumClasses = 0
mintNumRelatedDocs = 0
mstrTitle = ""
mstrMajor = ""
mstrMinor = ""
mstrRevision = ""

Do While Not EOF(intInFile)
    Line Input #intInFile, strData

    If InStr(strData, "Name=") = 1 Then
        mstrTitle = AfterEqual(strData)
    End If
    
    If InStr(strData, "ExeName32=") = 1 Then
        mstrEXEName = AfterEqual(strData)
    End If
    
    If InStr(strData, "MajorVer=") > 0 Then
        mstrMajor = AfterEqual(strData)
    End If
    
    If InStr(strData, "MinorVer=") > 0 Then
        mstrMinor = AfterEqual(strData)
    End If
    
    If InStr(strData, "RevisionVer=") > 0 Then
        mstrRevision = AfterEqual(strData)
    End If
    
    If InStr(strData, "Description=") = 1 Then
        mstrDescription = AfterEqual(strData)
    End If
    
    If InStr(strData, "VersionComments=") > 0 Then
        mstrVersionInfo(0) = AfterEqual(strData)
    End If
    
    If InStr(strData, "VersionCompanyName=") > 0 Then
        mstrVersionInfo(1) = AfterEqual(strData)
    End If
    
    If InStr(strData, "VersionFileDescription=") > 0 Then
        mstrVersionInfo(2) = AfterEqual(strData)
    End If
    
    If InStr(strData, "VersionLegalCopyright=") > 0 Then
        mstrVersionInfo(3) = AfterEqual(strData)
    End If
    
    If InStr(strData, "VersionLegalTrademarks=") > 0 Then
        mstrVersionInfo(4) = AfterEqual(strData)
    End If
    
    If InStr(strData, "VersionProductName=") > 0 Then
        mstrVersionInfo(5) = AfterEqual(strData)
    End If
    
    If InStr(strData, "Form=") = 1 Then
        If UBound(astrForm) < mintNumForms Then
            ReDim Preserve astrForm(UBound(astrForm) + 1)
        End If
        astrForm(mintNumForms) = AfterEqual(strData)
        mintNumForms = mintNumForms + 1
    End If
    
    If InStr(strData, "UserControl=") = 1 Then
        If UBound(astrUserControl) < mintNumUserControls Then
            ReDim Preserve astrUserControl(UBound(astrUserControl) + 1)
        End If
        astrUserControl(mintNumUserControls) = AfterEqual(strData)
        mintNumUserControls = mintNumUserControls + 1
    End If
    
    If InStr(strData, "PropertyPage=") = 1 Then
        If UBound(astrPropertyPage) < mintNumPropertyPages Then
            ReDim Preserve astrPropertyPage(UBound(astrPropertyPage) + 1)
        End If
        astrPropertyPage(mintNumPropertyPages) = AfterEqual(strData)
        mintNumPropertyPages = mintNumPropertyPages + 1
    End If
    
    If InStr(strData, "UserDocument=") = 1 Then
        If UBound(astrUserDocument) < mintNumUserDocuments Then
            ReDim Preserve astrUserDocument(UBound(astrUserDocument) + 1)
        End If
        astrUserDocument(mintNumUserDocuments) = AfterEqual(strData)
        mintNumUserDocuments = mintNumUserDocuments + 1
    End If
    
    If InStr(strData, "Designer=") = 1 Then
        If UBound(astrDesigner) < mintNumDesigners Then
            ReDim Preserve astrDesigner(UBound(astrDesigner) + 1)
        End If
        astrDesigner(mintNumDesigners) = AfterEqual(strData)
        mintNumDesigners = mintNumDesigners + 1
    End If
    
    If InStr(strData, "Object=") = 1 Then
        If UBound(astrObject) < mintNumObjects Then
            ReDim Preserve astrObject(UBound(astrObject) + 1)
        End If
        astrObject(mintNumObjects) = AfterEqual(strData)
        mintNumObjects = mintNumObjects + 1
    End If
    
    If InStr(strData, "Reference=") = 1 Then
        If UBound(astrReference) < mintNumReferences Then
            ReDim Preserve astrReference(UBound(astrReference) + 1)
        End If
        astrReference(mintNumReferences) = AfterEqual(strData)
        mintNumReferences = mintNumReferences + 1
    End If
    
    If InStr(strData, "RelatedDoc=") = 1 Or InStr(strData, "ResFile32=") = 1 Then
        If UBound(astrRelatedDocs) < mintNumRelatedDocs Then
            ReDim Preserve astrRelatedDocs(UBound(astrRelatedDocs) + 1)
        End If
        astrRelatedDocs(mintNumRelatedDocs) = RemoveQuotes(AfterEqual(strData))
        mintNumRelatedDocs = mintNumRelatedDocs + 1
    End If
    
    If InStr(strData, "Module=") = 1 Then
        If UBound(astrModule) < mintNumModules Then
            ReDim Preserve astrModule(UBound(astrModule) + 1)
        End If
        astrModule(mintNumModules) = AfterEqual(strData)
        mintNumModules = mintNumModules + 1
    End If
    
    If InStr(strData, "Class=") = 1 Then
        If UBound(astrClass) < mintNumClasses Then
            ReDim Preserve astrClass(UBound(astrClass) + 1)
        End If
        astrClass(mintNumClasses) = AfterEqual(strData)
        mintNumClasses = mintNumClasses + 1
    End If
Loop

Close #intInFile

ReDim astrSort(UBound(astrForm))
For i = 0 To UBound(astrForm)
    astrSort(i) = ExtractName(astrForm(i))
Next i
Call SortList(astrSort(), astrForm())

ReDim astrSort(UBound(astrModule))
For i = 0 To UBound(astrModule)
    astrSort(i) = ExtractName(astrModule(i))
Next i
Call SortList(astrSort(), astrModule())

ReDim astrSort(UBound(astrClass))
For i = 0 To UBound(astrClass)
    astrSort(i) = ExtractName(astrClass(i))
Next i
Call SortList(astrSort(), astrClass())

ReDim astrSort(UBound(astrUserControl))
For i = 0 To UBound(astrUserControl)
    astrSort(i) = ExtractName(astrUserControl(i))
Next i
Call SortList(astrSort(), astrUserControl())

ReDim astrSort(UBound(astrPropertyPage))
For i = 0 To UBound(astrPropertyPage)
    astrSort(i) = ExtractName(astrPropertyPage(i))
Next i
Call SortList(astrSort(), astrPropertyPage())

ReDim astrSort(UBound(astrUserDocument))
For i = 0 To UBound(astrUserDocument)
    astrSort(i) = ExtractName(astrUserDocument(i))
Next i
Call SortList(astrSort(), astrUserDocument())

ReDim astrSort(UBound(astrDesigner))
For i = 0 To UBound(astrDesigner)
    astrSort(i) = ExtractName(astrDesigner(i))
Next i
Call SortList(astrSort(), astrDesigner())

For i = 0 To UBound(astrForm)
    If Trim$(astrForm(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "Form: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrForm(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcForms.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrModule)
    If Trim$(astrModule(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "Module: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrModule(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcModules.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrClass)
    If Trim$(astrClass(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "Class: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrClass(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcClasses.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrUserControl)
    If Trim$(astrUserControl(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "User Control: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrUserControl(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcUserControls.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrPropertyPage)
    If Trim$(astrPropertyPage(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "Property Page: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrPropertyPage(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcPropertyPages.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrUserDocument)
    If Trim$(astrUserDocument(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "User Document: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrUserDocument(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcUserDocuments.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrDesigner)
    If Trim$(astrDesigner(i)) <> "" Then
        Set cFile = New clsFile
        cFile.FileType = "Designer: "
        cFile.FilePath = mstrVBPPath
        cFile.Filename = astrDesigner(i)
        cFile.ProjectTitle = mstrTitle
        cFile.ProjectVersionMajor = mstrMajor
        cFile.ProjectVersionMinor = mstrMinor
        cFile.ProjectVersionRevision = mstrRevision
        cFile.StyleSheetFile = mstrStyleSheetFile
        cFile.IncludeAPI = mblnIncludeAPI
        cFile.IncludeCounts = mblnIncludeCounts
        cFile.IncludeDeclarations = mblnIncludeDeclarations
        cFile.IncludeEvents = mblnIncludeEvents
        cFile.IncludeReferences = mblnIncludeReferences
        cFile.IncludeSubs = mblnIncludeSubs
        cFile.IncludeTypes = mblnIncludeTypes
        cFile.IncludeAttributes = mblnIncludeAttributes
        mcDesigners.Add cFile
        Set cFile = Nothing
    End If
Next i

For i = 0 To UBound(astrReference)
    If Trim$(astrReference(i)) <> "" Then
        Set cItem = New clsItem
        cItem.ItemReference = astrReference(i)
        mcReferences.Add cItem
        Set cItem = Nothing
    End If
Next i

For i = 0 To UBound(astrObject)
    If Trim$(astrObject(i)) <> "" Then
        Set cItem = New clsItem
        cItem.ItemReference = astrObject(i)
        mcObjects.Add cItem
        Set cItem = Nothing
    End If
Next i

For i = 0 To UBound(astrRelatedDocs)
    If Trim$(astrRelatedDocs(i)) <> "" Then
        Set cItem = New clsItem
        cItem.ItemReference = astrRelatedDocs(i)
        mcRelatedDocs.Add cItem
        Set cItem = Nothing
    End If
Next i

If mblnIncludeNAVBar = True Then
    For Each cFile In mcForms
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
    For Each cFile In mcModules
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
    For Each cFile In mcClasses
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
    For Each cFile In mcUserControls
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
    For Each cFile In mcPropertyPages
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
    For Each cFile In mcUserDocuments
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
    For Each cFile In mcDesigners
        cFile.IncludeNAVBar = True
        cFile.NAVBar = NAVBarHTML(cFile.Filename)
    Next
End If

For Each cFile In mcForms
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
For Each cFile In mcClasses
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
For Each cFile In mcModules
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
For Each cFile In mcUserControls
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
For Each cFile In mcPropertyPages
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
For Each cFile In mcUserDocuments
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
For Each cFile In mcDesigners
    cFile.ParseForm
    mlngNumCodeLines = mlngNumCodeLines + cFile.CodeLineCount
    mlngNumCommentLines = mlngNumCommentLines + cFile.CommentLineCount
Next
    
Exit Sub

Handler:
Err.Raise Err.Number, IIf(Left$(Err.Source, 3) <> "cls", "clsProject.ParseVBPFile", Err.Source), Err.Description
    
End Sub

Public Sub SaveHTML()

Dim intFileNum As Integer
Dim cFile As clsFile

intFileNum = FreeFile

Open mstrOutputPath & IIf(Right$(mstrOutputPath, 1) = "\", "", "\") & FileOnly(ExtractFile(mstrVBPName, mstrVBPPath)) & ".html" For Output As #intFileNum

Print #intFileNum, AddHTMLHeader
Print #intFileNum, AddHTMLBody
Print #intFileNum, AddHTMLFooter

Close #intFileNum

If mintNumForms > 0 And mblnIncludeForms = True Then
    For Each cFile In mcForms
        cFile.SaveHTML mstrOutputPath
    Next
End If
If mintNumModules > 0 And mblnIncludeModules = True Then
    For Each cFile In mcModules
        cFile.SaveHTML mstrOutputPath
    Next
End If
If mintNumClasses > 0 And mblnIncludeClasses = True Then
    For Each cFile In mcClasses
        cFile.SaveHTML mstrOutputPath
    Next
End If
If mintNumUserControls > 0 And mblnIncludeUserControls = True Then
    For Each cFile In mcUserControls
        cFile.SaveHTML mstrOutputPath
    Next
End If
If mintNumPropertyPages > 0 And mblnIncludePropertyPages = True Then
    For Each cFile In mcPropertyPages
        cFile.SaveHTML mstrOutputPath
    Next
End If
If mintNumDesigners > 0 And mblnIncludeDesigners = True Then
    For Each cFile In mcDesigners
        cFile.SaveHTML mstrOutputPath
    Next
End If
If mintNumUserDocuments > 0 And mblnIncludeUserDocuments = True Then
    For Each cFile In mcUserDocuments
        cFile.SaveHTML mstrOutputPath
    Next
End If

If mblnOutputStyleSheet = True Then
    Call OutputStyleSheet(mstrOutputPath, mstrStyleSheetFile)
End If

If mintOutputType = HTMLHelp Then
    Call OutputHTMLHelpHHC
    Call OutputHTMLHelpHHP
End If

End Sub

