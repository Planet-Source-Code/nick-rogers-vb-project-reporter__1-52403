VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Project Reporter"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstInclude 
      Height          =   1185
      Left            =   1290
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   840
      Width           =   5625
   End
   Begin VB.ListBox lstOptions 
      Height          =   1185
      Left            =   1290
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   2070
      Width           =   5625
   End
   Begin VB.CommandButton cmdSS 
      Caption         =   "…"
      Height          =   255
      Left            =   6540
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3330
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   885
      Left            =   5670
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5070
      Width           =   1275
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Default         =   -1  'True
      Height          =   885
      Left            =   4290
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5070
      Width           =   1275
   End
   Begin VB.CommandButton cmdGetFolder 
      Caption         =   "…"
      Height          =   255
      Left            =   6540
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   510
      Width           =   345
   End
   Begin VB.TextBox txtOutput 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Top             =   480
      Width           =   5625
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "…"
      Height          =   255
      Left            =   6540
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   345
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   150
      Width           =   5625
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   360
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSS 
      Height          =   315
      Left            =   1290
      TabIndex        =   6
      Text            =   "[default style sheet]"
      Top             =   3300
      Width           =   5625
   End
   Begin VB.Frame fraHelp 
      Height          =   1335
      Left            =   1290
      TabIndex        =   15
      Top             =   3660
      Width           =   5625
      Begin VB.TextBox txtHelpTitle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   270
         Width           =   4515
      End
      Begin VB.CommandButton cmdHelpCompiler 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5100
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   930
         Width           =   345
      End
      Begin VB.TextBox txtHelpCompiler 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   900
         Width           =   4515
      End
      Begin VB.Label lblInfo 
         Caption         =   "Help title"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Help Compiler Location"
         Enabled         =   0   'False
         Height          =   285
         Left            =   150
         TabIndex        =   16
         Top             =   660
         Width           =   1935
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "VB components to include:"
      Height          =   465
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   870
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Caption         =   "HTML Help:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   3750
      Width           =   2295
   End
   Begin VB.Label lblInfo 
      Caption         =   "Style sheet:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   3330
      Width           =   2295
   End
   Begin VB.Label lblInfo 
      Caption         =   "Output options:"
      Height          =   585
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   2100
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Caption         =   "Output path:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   510
      Width           =   2295
   End
   Begin VB.Label lblInfo 
      Caption         =   "Project:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   210
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------------
' VB Project Reporter (Originally VBPScan)
' Documents VB Projects and creates HTML output
' Original program by kerlin (www.kerlinsoftworks.com)
' Enhanced and expanded by Nick Rogers (rogersn@ozemail.com.au)
' Registry class from Steve McMahon (www.vbaccelerator.com)
' -------------------------------------------------------------------------------------------------
' Problem Fix:
'   Corrected error reported by alberto, where files not stored in the project folder will create problems
'   Fixed by doing a full path search, and resolving the ".." entries that might appear
'   Corrected error output created by having an array return parameter
'   Corrected error output created by having character variable declarations (eg MyVar!)
'   Corrected misreporting of functions/subs/properties where text is included in string information
'   Corrected some out of bounds and type mismatch errors (5/4/04)
' New Features:
'   Added "head" tag to HTML files
'   Added support for .VBG (Project group) files
'   Added support for .DSR (Designers) files
'   Added support for .DOB (User Document) files
'   Added support for .PAG (Property page) files
'   Added support for related documents
'   Added support for customised style sheet files
'   Disabled link to current page in navigation bar
'   Added ability to create HTML Help files (if HTML Help Workshop is installed)
'   Re-wrote program to use classes
'   Added code/comment line counts
'   Added declarations section (for module/form level variable declarations)
'   Added processing for API declares
'   Added header comments to HTML
'   Added processing for Types, Enums, Events
'   Added ability to turn off various bits of the output
'   Added processing for procedure attributes
'
' The HTML Help Workshop can be found at this URL:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/htmlhelp/html/hwMicrosoftHTMLHelpDownloads.asp

Option Explicit
Option Compare Text

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub cmdFile_Click()

On Error Resume Next

' set up the command dialog box to select a log file
dlgOpen.DialogTitle = "Select Visual Basic File"
dlgOpen.Flags = cdlOFNCreatePrompt & cdlOFNOverwritePrompt
dlgOpen.CancelError = True
dlgOpen.Filter = "Visual Basic Files (*.VBG;*.VBP)|*.VBG;*.VBP|All files (*.*)|*.*"
If txtFile.Text <> "" Then
    dlgOpen.Filename = txtFile.Text
End If

' display the open dialog
dlgOpen.ShowOpen

' if there were no errors, update the label
If Err = 0 Then
    txtFile.Text = dlgOpen.Filename
    txtOutput.Text = Left$(dlgOpen.Filename, InStrRev(dlgOpen.Filename, "\") - 1)
End If

On Error GoTo 0

End Sub

Private Sub cmdGetFolder_Click()

txtOutput.Text = GetFolder("Output Folder", Me)

End Sub

Private Sub cmdHelpCompiler_Click()

On Error Resume Next

' set up the command dialog box to select a log file
dlgOpen.DialogTitle = "Select Help Compiler"
dlgOpen.Flags = cdlOFNCreatePrompt & cdlOFNOverwritePrompt
dlgOpen.CancelError = True
dlgOpen.Filter = "EXE Files (*.EXE)|*.EXE|All files (*.*)|*.*"
If txtHelpCompiler.Text <> "" Then
    dlgOpen.Filename = txtHelpCompiler.Text
End If

' displace the open dialog
dlgOpen.ShowOpen

' if there were no errors, update the label
If Err = 0 Then
    txtHelpCompiler.Text = dlgOpen.Filename
End If

On Error GoTo 0

End Sub

Private Sub cmdRun_Click()

Dim ctl As Control
Dim cProject As clsProject
Dim cGroup As clsGroup
Dim strError As String

On Error GoTo Handler

If txtFile.Text = "" Then
    MsgBox "You must select a Visual Basic Project file.", vbExclamation, "VB Project File Missing"
    txtFile.SetFocus
    Exit Sub
End If

If txtOutput.Text = "" Then
    MsgBox "You must select a location for the output files.", vbExclamation, "Output Path Missing"
    txtOutput.SetFocus
    Exit Sub
End If

If FileExists(txtFile.Text) = False Then
    MsgBox "The Visual Basic file you specified does not exist. Please select a valid file.", vbExclamation, "Invalid File"
    txtFile.SetFocus
    Exit Sub
End If

If FileExists(txtOutput.Text & IIf(Right$(txtOutput.Text, 1) = "\", "", "\")) = False Then
    MsgBox "The output path you specified does not exist. Please select a valid path.", vbExclamation, "Invalid Path"
    txtOutput.SetFocus
    Exit Sub
End If

Screen.MousePointer = vbHourglass

For Each ctl In Me.Controls
    If TypeOf ctl Is CommandButton Then
        ctl.Enabled = False
    End If
    If TypeOf ctl Is TextBox Then
        ctl.Enabled = False
    End If
    If TypeOf ctl Is CheckBox Then
        ctl.Enabled = False
    End If
    If TypeOf ctl Is ListBox Then
        ctl.Enabled = False
    End If
Next
Me.Refresh

If InStr(LCase$(txtFile.Text), ".vbg") > 0 Then
    ' when a group file has been specified
    Set cGroup = New clsGroup
    cGroup.HelpTitle = txtHelpTitle.Text
    If lstOptions.Selected(2) = True Then
        cGroup.FileOutputType = HTMLHelp
    Else
        cGroup.FileOutputType = HTML
    End If
    cGroup.IncludeAPI = lstOptions.Selected(5)
    cGroup.IncludeCounts = lstOptions.Selected(9)
    cGroup.IncludeDeclarations = lstOptions.Selected(3)
    cGroup.IncludeEvents = lstOptions.Selected(6)
    cGroup.IncludeReferences = lstOptions.Selected(8)
    cGroup.IncludeSubs = lstOptions.Selected(7)
    cGroup.IncludeTypes = lstOptions.Selected(4)
    cGroup.IncludeNAVBar = lstOptions.Selected(0)
    cGroup.IncludeAttributes = lstOptions.Selected(10)
    cGroup.IncludeVersionInfo = lstOptions.Selected(11)
    cGroup.IncludeForms = lstInclude.Selected(0)
    cGroup.IncludeClasses = lstInclude.Selected(2)
    cGroup.IncludeDesigners = lstInclude.Selected(5)
    cGroup.IncludeModules = lstInclude.Selected(1)
    cGroup.IncludeRelatedDocs = lstInclude.Selected(7)
    cGroup.IncludeUserControls = lstInclude.Selected(3)
    cGroup.IncludeUserDocuments = lstInclude.Selected(6)
    cGroup.IncludePropertyPages = lstInclude.Selected(4)
    cGroup.OutputStyleSheetFile = lstOptions.Selected(1)
    cGroup.OutputPath = txtOutput.Text
    If Left$(txtSS.Text, 1) <> "[" Then
        cGroup.StyleSheetFile = txtSS.Text
    Else
        cGroup.StyleSheetFile = ""
    End If
    cGroup.ParseGroup txtFile.Text
    cGroup.SaveHTML
    Set cGroup = Nothing
Else
    ' when a project file has been specified
    Set cProject = New clsProject
    cProject.HelpTitle = txtHelpTitle.Text
    If lstOptions.Selected(2) = True Then
        cProject.FileOutputType = HTMLHelp
    Else
        cProject.FileOutputType = HTML
    End If
    cProject.IncludeAPI = lstOptions.Selected(5)
    cProject.IncludeCounts = lstOptions.Selected(9)
    cProject.IncludeDeclarations = lstOptions.Selected(3)
    cProject.IncludeEvents = lstOptions.Selected(6)
    cProject.IncludeReferences = lstOptions.Selected(8)
    cProject.IncludeSubs = lstOptions.Selected(7)
    cProject.IncludeTypes = lstOptions.Selected(4)
    cProject.IncludeNAVBar = lstOptions.Selected(0)
    cProject.IncludeAttributes = lstOptions.Selected(10)
    cProject.IncludeVersionInfo = lstOptions.Selected(11)
    cProject.OutputStyleSheetFile = lstOptions.Selected(1)
    cProject.IncludeForms = lstInclude.Selected(0)
    cProject.IncludeClasses = lstInclude.Selected(2)
    cProject.IncludeDesigners = lstInclude.Selected(5)
    cProject.IncludeModules = lstInclude.Selected(1)
    cProject.IncludeRelatedDocs = lstInclude.Selected(7)
    cProject.IncludeUserControls = lstInclude.Selected(3)
    cProject.IncludeUserDocuments = lstInclude.Selected(6)
    cProject.IncludePropertyPages = lstInclude.Selected(4)
    cProject.OutputPath = txtOutput.Text
    If Left$(txtSS.Text, 1) <> "[" Then
        cProject.StyleSheetFile = txtSS.Text
    Else
        cProject.StyleSheetFile = ""
    End If
    cProject.ParseVBPFile txtFile.Text
    cProject.SaveHTML
    Set cProject = Nothing
End If

For Each ctl In Me.Controls
    If TypeOf ctl Is CommandButton Then
        ctl.Enabled = True
        ' exceptions
        If ctl.Name = "cmdHelpCompiler" And lstOptions.Selected(2) = False Then
            ctl.Enabled = False
        End If
        If ctl.Name = "cmdSS" And lstOptions.Selected(1) = False Then
            ctl.Enabled = False
        End If
    End If
    If TypeOf ctl Is TextBox Then
        ctl.Enabled = True
        ' exceptions
        If ctl.Name = "txtHelpCompiler" And lstOptions.Selected(2) = False Then
            ctl.Enabled = False
        End If
        If ctl.Name = "txtHelpTitle" And lstOptions.Selected(2) = False Then
            ctl.Enabled = False
        End If
        If ctl.Name = "txtSS" And lstOptions.Selected(1) = False Then
            ctl.Enabled = False
        End If
    End If
    If TypeOf ctl Is CheckBox Then
        ctl.Enabled = True
    End If
    If TypeOf ctl Is ListBox Then
        ctl.Enabled = True
    End If
Next
Me.Refresh

Screen.MousePointer = vbDefault

' compile a HTML help file
If lstOptions.Selected(2) = True And txtHelpCompiler.Text <> "" Then
    DoEvents
    CompileHTMLHelp txtOutput.Text & IIf(Right$(txtOutput.Text, 1) <> "\", "\", "") & ExtractName(txtFile.Text) & ".HHP", txtHelpCompiler.Text
End If

MsgBox "The VB Documentation is complete!", vbInformation, "Documentation Complete"

Exit Sub

Handler:
strError = "An error has occurred while creating the documentation." & vbCrLf & Err.Number & ": " & Err.Description
MsgBox strError, vbExclamation, "Error"

End Sub

Private Sub cmdSS_Click()

On Error Resume Next

' set up the command dialog box to select a log file
dlgOpen.DialogTitle = "Select Style Sheet File"
dlgOpen.Flags = cdlOFNCreatePrompt & cdlOFNOverwritePrompt
dlgOpen.CancelError = True
dlgOpen.Filter = "StyleSheet Files (*.CSS)|*.CSS|All files (*.*)|*.*"
If Left$(txtSS.Text, 1) <> "[" Then
    dlgOpen.Filename = txtSS.Text
End If

' displace the open dialog
dlgOpen.ShowOpen

' if there were no errors, update the label
If Err = 0 Then
    txtSS.Text = dlgOpen.Filename
End If

On Error GoTo 0

End Sub

Private Sub Form_Load()

Dim reg As New clsRegistry
Dim strPath As String, i As Long

lstOptions.AddItem "Put NAV Bar on all files"
lstOptions.AddItem "Generate style sheet file"
lstOptions.AddItem "Create HTML help file"
lstOptions.AddItem "Output general declarations"
lstOptions.AddItem "Output type and enum definitions"
lstOptions.AddItem "Output API declarations"
lstOptions.AddItem "Output user control event declarations"
lstOptions.AddItem "Output subs/functions/properties"
lstOptions.AddItem "Output object and reference information"
lstOptions.AddItem "Output code/comment line counts"
lstOptions.AddItem "Output procedure attributes (where defined)"
lstOptions.AddItem "Output project version information"

For i = 0 To lstOptions.ListCount - 1
    If i <> 2 Then
        lstOptions.Selected(i) = True
    End If
Next i
lstOptions.ListIndex = -1

lstInclude.AddItem "Forms"
lstInclude.AddItem "Modules"
lstInclude.AddItem "Classes"
lstInclude.AddItem "User Controls"
lstInclude.AddItem "Property Pages"
lstInclude.AddItem "Designers"
lstInclude.AddItem "User Documents"
lstInclude.AddItem "Related Documents"

For i = 0 To lstInclude.ListCount - 1
    lstInclude.Selected(i) = True
Next i
lstInclude.ListIndex = -1

reg.ClassKey = HKEY_CURRENT_USER
reg.SectionKey = "Software\Microsoft\HTML Help Workshop"

' find out if the HTML Help Workshop has been installed
reg.ValueKey = "InstallDir"

strPath = Trim$(reg.Value)

If strPath = "" Then
    fraHelp.Visible = False
    lblInfo(4).Visible = False
Else
    txtHelpCompiler.Text = strPath & IIf(Right$(strPath, 1) = "\", "", "\") & "HHC.EXE"
End If

Set reg = Nothing

End Sub

Private Sub lstOptions_ItemCheck(Item As Integer)

Dim obj As Control

Select Case Item
Case 0
    If lstOptions.Selected(2) = True And lstOptions.Selected(0) = True Then
        lstOptions.Selected(2) = False
    End If
Case 1
    txtSS.Enabled = lstOptions.Selected(1)
    cmdSS.Enabled = lstOptions.Selected(1)
Case 2
    If fraHelp.Visible = False And lstOptions.Selected(2) = True Then
        lstOptions.Selected(2) = False
    End If
    If lstOptions.Selected(2) = True And lstOptions.Selected(0) = True Then
        lstOptions.Selected(0) = False
    End If
Case 3
    If lstOptions.Selected(3) = False Then
        lstOptions.Selected(4) = False
    End If
Case 4
    If lstOptions.Selected(3) = False And lstOptions.Selected(4) = True Then
        lstOptions.Selected(3) = True
    End If
End Select

On Error Resume Next
For Each obj In Me.Controls
    If obj.Name <> "dlgOpen" Then
        If obj.Container.Name = "fraHelp" Then
            obj.Enabled = lstOptions.Selected(2)
        End If
    End If
Next
On Error GoTo 0

End Sub

