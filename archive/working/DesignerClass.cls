VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'========================================================================
'
' Sample file to create a Contensive Aggregate Object
'
'========================================================================
'
'Private Main As ccWeb3.MainClass
Private Csv As Object
Private Main As Object
Private ApplicationName As String
Private issueList As String
Private filesFetchedList As String
'
'=================================================================================
'   Execute Method, v3.4 Interface
'=================================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
    '
    Exit Function
ErrorTrap:
    HandleClassError ("Execute")
End Function
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(MainObject As Object)
    On Error GoTo ErrorTrap
    '
    ApplicationName = "unknown"
    Set Main = MainObject
    ApplicationName = Main.ApplicationName
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassError("Init")
End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Joomla As New JoomlaClass
    Dim WorkingPathPage As String
    Dim JoomlaFilename As String
    Dim Styles As String
    Dim StyleSN As Long
    Dim TemplateBodyHTML As String
    Dim DefaultLink As String
    Dim BuildVersion As String
    Dim ManageStyles As Boolean
    Dim ImportLink As String
    Dim CS As Long
    Dim Argument1 As String
    Dim Argument2 As String
    Dim ToolPageAsm As Object
    Dim Button As String
    Dim RQS As String
    Dim QS As String
    Dim FormID As Long
    Dim Copy As String
    Dim TemplateName As String
    Dim TemplateID As Long
    Dim IsNameOK As Boolean
    Dim Hint As String
    '
    Set ToolPageAsm = CreateObject("ccPageAsm.ToolPageClass")
    '
    Hint = "100"
    If Not (Main Is Nothing) Then
        BuildVersion = Main.GetSiteProperty("BuildVersion")
        Argument1 = Main.GetAggrOption("Argument1", OptionString)
        Argument2 = Main.GetAggrOption("Argument2", OptionString)
        RQS = Main.RefreshQueryString
        Button = Main.GetStreamText("button")
        FormID = Main.GetStreamInteger(RequestNameFormID)
        TemplateID = kmaEncodeInteger(Main.GetVisitProperty(VisitPropertyTemplateID))
        '
        If Button <> "" Then
            Hint = "200"
            '
            ' Process buttons
            '
FormID = FormImportOnePage
            Select Case FormID
                Case FormImportOnePage
                    '
                    ' Process Import form
                    '
                    If Button = ButtonCancel Then
                        '
                        ' cancel back to root form
                        '
                        FormID = FormRoot
                    Else
                        TemplateName = Main.GetStreamText(RequestNameTemplateName)
                        TemplateID = CreateNewTemplate(TemplateName)
                        If TemplateID = 0 Then
                            Call Main.AddUserError("There was a problem creating the page Template [" & TemplateName & "]. Select a different name, or use the 'Open' tool to edit the existing template.")
                            TemplateName = ""
                        Else
                            ImportLink = Main.GetStreamText(RequestNameImportLink)
                            If ImportLink = "" Then
                                Call Main.AddUserError("To import a template, enter a URL.")
                            Else
                                ManageStyles = True
                                If Not Main.IsUserError() Then
                                    Call ImportTemplate(TemplateID, ImportLink, ManageStyles, BuildVersion)
                                End If
                            End If
                        End If
                    End If
'                Case FormImport
'                    '
'                    ' Process Import form
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' cancel back to root form
'                        '
'                        FormID = FormRoot
'                    Else
'                        ImportLink = Main.GetStreamText(RequestNameImportLink)
'                        ManageStyles = True
'                        'ManageStyles = Main.GetStreamBoolean(RequestNameManageStyles)
'                        If ImportLink = "" Then
'                            Call Main.AddUserError("To import a template, enter a URL.")
'                        End If
'                        If Not Main.IsUserError() Then
'                            If TemplateID = 0 Then
'                                TemplateID = CreateNewTemplate(TemplateName)
'                                CS = Main.OpenCSContentRecord("page templates", TemplateID)
'                                If Main.IsCSOK(CS) Then
'                                    Call Main.SetCS(CS, "name", "Template " & TemplateID)
'                                End If
'                                Call Main.CloseCS(CS)
'                            End If
'                            If TemplateID <> 0 Then
'                                Call ImportTemplate(TemplateID, ImportLink, ManageStyles, BuildVersion)
'                                FormID = FormEdit
'                            End If
'                        End If
'                    End If
'                Case FormImportJoomla
'                    '
'                    ' Process Import form
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' cancel back to root form
'                        '
'                        FormID = FormRoot
'                    Else
'                        JoomlaFilename = Main.GetStreamText(RequestNameJoomlaFile)
'                        WorkingPathPage = "temp\" & JoomlaFilename
'                        Call Main.ProcessFormInputFile(RequestNameJoomlaFile, WorkingPathPage)
'                        If Not Main.IsUserError() Then
'                            If TemplateID = 0 Then
'                                TemplateID = CreateNewTemplate(TemplateName)
'                                CS = Main.OpenCSContentRecord("page templates", TemplateID)
'                                If Main.IsCSOK(CS) Then
'                                    Call Main.SetCS(CS, "name", "Template " & TemplateID)
'                                End If
'                                Call Main.CloseCS(CS)
'                            End If
'                            If TemplateID <> 0 Then
'                                Call Joomla.ImportJoomlaTemplate(Main, TemplateID, WorkingPathPage)
'                                FormID = FormEdit
'                            End If
'                        End If
'                    End If
'                Case FormCreateNew
'                    '
'                    ' Process new form
'                    '
'
'                    TemplateName = Main.GetStreamText(RequestNameTemplateName)
'                    TemplateID = CreateNewTemplate(TemplateName)
'                    If TemplateID = 0 Then
'                        Call Main.AddUserError("There was a problem creating the page Template [" & TemplateName & "]. Select a different name, or use the 'Open' tool to edit the existing template.")
'                        TemplateName = ""
'                    Else
'                        FormID = FormEdit
'                    End If
'                Case FormOpen
'                    '
'                    '
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' Cancel exits the getcontent with blank
'                        '
'                        FormID = FormEmpty
'                    ElseIf (Button = ButtonOK) Then
'                        '
'                        ' Open the template selected
'                        '
'                        TemplateID = Main.GetStreamInteger(RequestNameTemplateID)
'                        Call Main.SetVisitProperty(VisitPropertyTemplateID, TemplateID)
'                        FormID = FormEdit
'                    End If
'                Case FormEdit
'                    '
'                    '
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' Cancel exits the getcontent with blank
'                        '
'                        FormID = FormEmpty
'                    ElseIf (Button = ButtonSave) Or (Button = ButtonOK) Then
'                        '
'                        ' Perform Save
'                        '
'                        TemplateID = kmaEncodeInteger(Main.GetVisitProperty(VisitPropertyTemplateID))
'                        CS = Main.OpenCSContentRecord("Page Templates", TemplateID)
'                        If Main.IsCSOK(CS) Then
'                            Call Main.SetCS(CS, "bodyhtml", Main.GetStreamText(RequestNameBodyHTML))
'                        End If
'                        Call Main.CloseCS(CS)
'                        If Button = ButtonOK Then
'                            FormID = FormEmpty
'                        End If
'                    End If
'                Case FormFileView
'                    '
'                    '
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' Cancel exits the getcontent with blank
'                        '
'                        FormID = FormEmpty
'                    End If
'                Case FormSiteStyles
'                    '
'                    '
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' Cancel exits the getcontent with blank
'                        '
'                        FormID = FormEmpty
'                    ElseIf (Button = ButtonSave) Or (Button = ButtonOK) Then
'                        '
'                        ' Perform Save
'                        '
'                        Call Main.SaveVirtualFile("templates\styles.css", Main.GetStreamText(RequestNameSiteStyles))
'                        If Main.GetStreamBoolean(RequestNameSiteStylesInline) Then
'                            '
'                            ' Inline Styles
'                            '
'                            Call Main.SetSiteProperty("StylesheetSerialNumber", "0")
'                        Else
'                            '
'                            ' Linked Styles
'                            ' Bump the Style Serial Number so next fetch is not cached
'                            '
'                            StyleSN = kmaEncodeInteger(Main.GetSiteProperty("StylesheetSerialNumber", "0"))
'                            StyleSN = StyleSN + 1
'                            Call Main.SetSiteProperty("StylesheetSerialNumber", StyleSN)
'                            '
'                            ' Save new public stylesheet
'                            '
'                            Call Main.SaveVirtualFile("templates\Public" & StyleSN & ".css", Main.GetStyleSheetProcessed)
'                            Call Main.SaveVirtualFile("templates\Admin" & StyleSN & ".css", Main.GetStyleSheetDefault)
'                        End If
'                        If Button = ButtonOK Then
'                            FormID = FormEmpty
'                        End If
'                    End If
'                Case FormTemplateStyles
'                    '
'                    '
'                    '
'                    If Button = ButtonCancel Then
'                        '
'                        ' Cancel exits the getcontent with blank
'                        '
'                        FormID = FormEmpty
'                    ElseIf (Button = ButtonSave) Or (Button = ButtonOK) Then
'                        '
'                        ' Perform Save
'                        '
'                        Styles = Main.GetStreamText(RequestNameTemplateStyles)
'                        CS = Main.OpenCSContentRecord("page templates", TemplateID)
'                        If Main.IsCSOK(CS) Then
'                            Call Main.SetCS(CS, "StylesFilename", Styles)
'                        End If
'                        Call Main.CloseCS(CS)
'                        If Button = ButtonOK Then
'                            FormID = FormEmpty
'                        End If
'                    End If
                Case Else
                    FormID = FormRoot
            End Select
        End If
        '
        ' Get Forms
        '
FormID = FormImportOnePage
        Hint = "500"
        Call Main.AddRefreshQueryString(RequestNameFormID, FormID)
        Select Case FormID
            Case FormImportOnePage
                '
                ' simple one page form -- enter the name for the template and enter URL
                '
                DefaultLink = Main.GetStreamText(RequestNameImportLink)
                If DefaultLink <> "" Then
                    Call Main.SetVisitProperty("TemplateImporterLastLink", DefaultLink)
                Else
                    DefaultLink = Main.GetVisitProperty("TemplateImporterLastLink", "http://www.contensive.com")
                End If
                Copy = "" _
                    & "<p>&nbsp;</p>" _
                    & "<p>To create a new template, enter a template name and the URL where the page can be found. The template name must be unique in your website.</p>" _
                    & ""
                Call ToolPageAsm.Init(Main)
                Call ToolPageAsm.AddButton(ButtonCancel)
                Call ToolPageAsm.AddButton(ButtonImportTemplates)
                Call ToolPageAsm.AddElement("BodyHeader", "Create a new template")
                Call ToolPageAsm.AddElement("Body", Copy)
                Call ToolPageAsm.AddEditRow("Name", Main.GetFormInputText(RequestNameTemplateName, Main.GetStreamText(RequestNameTemplateName)), "Enter the name of the new template.")
                Call ToolPageAsm.AddEditRow("URL", Main.GetFormInputText(RequestNameImportLink, DefaultLink), "Enter the URL where the new template can be found.")
                'Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
                GetContent = ToolPageAsm.GetPage()
'            Case FormEmpty
'                '
'                ' exit with empty form
'                '
'                GetContent = ""
'            Case FormFileView
'                '
'                ' File View
'                '
'                Copy = "<p>Manage the files in your web directory.</p>"
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The Template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonApply)
'                Call ToolPageAsm.AddElement("BodyHeader", "Manage your web files")
'                Call ToolPageAsm.AddElement("Body", Copy & GetFileView())
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                'Call AddToolPanel_Tools(ToolPageAsm, RQS, TemplateID)
'                GetContent = ToolPageAsm.GetPage()
'            Case FormSiteStyles
'                '
'                ' Site Styles
'                '
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The Template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonSave)
'                Call ToolPageAsm.AddButton(ButtonOK)
'                Call ToolPageAsm.AddElement("BodyHeader", "Manage your site styles")
'                '
'                Copy = Main.GetFormInputCheckBox(RequestNameSiteStylesInline, (Main.GetSiteProperty("StylesheetSerialNumber", "0") = "0"))
'                Call ToolPageAsm.AddEditRow("InLine Styles", Copy, "When checked, the site styles will be included directly in your pages.")
'                '
'                Copy = Main.ReadFile(Main.PhysicalFilePath & "templates\styles.css")
'                Copy = Main.GetFormInputTextExpandable(RequestNameSiteStyles, Copy, 20)
'                Call ToolPageAsm.AddEditRow("Site Styles", Copy, "")
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                'Call AddToolPanel_Tools(ToolPageAsm, RQS, TemplateID)
'                GetContent = ToolPageAsm.GetPage()
'            Case FormTemplateStyles
'                '
'                ' Template Exclusive Styles
'                '
'                CS = Main.OpenCSContentRecord("page templates", TemplateID)
'                If Main.IsCSOK(CS) Then
'                    Styles = Main.GetCS(CS, "Stylesfilename")
'                    TemplateName = Main.GetCS(CS, "name")
'                    If TemplateName = "" Then
'                        TemplateName = "Template # " & TemplateID
'                        Call Main.ExecuteSQL("default", "update cctemplates set name=" & KmaEncodeSQLText(TemplateName) & " where id=" & TemplateID)
'                    End If
'                End If
'                Call Main.CloseCS(CS)
'                '
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonSave)
'                Call ToolPageAsm.AddButton(ButtonOK)
'                Call ToolPageAsm.AddElement("BodyHeader", "Manage the styles for this template")
'                '
'                Copy = Main.GetFormInputTextExpandable(RequestNameTemplateStyles, Styles, 20)
'                Call ToolPageAsm.AddEditRow("Template Styles", Copy, "These styles only apply to this template, and are only used on the website for pages using this template.")
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                'Call AddToolPanel_Tools(ToolPageAsm, RQS, TemplateID)
'                GetContent = ToolPageAsm.GetPage()
'            Case FormImport
'                '
'                ' Edit a template
'                '
'                DefaultLink = Main.GetStreamText(RequestNameImportLink)
'                If DefaultLink <> "" Then
'                    Call Main.SetVisitProperty("TemplateImporterLastLink", DefaultLink)
'                Else
'                    DefaultLink = Main.GetVisitProperty("TemplateImporterLastLink", "http://www.contensive.com")
'                End If
'                ManageStyles = True
'                'ManageStyles = Main.GetStreamBoolean(RequestNameManageStyles)
'                '
'                Copy = "<p>To import a page template, enter the URL of the page. The HTML from the page you reference will by saved into the current template.</p>"
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonBeginImport)
'                Call ToolPageAsm.AddElement("BodyHeader", "Import a template")
'                Call ToolPageAsm.AddElement("Body", Copy)
'                Call ToolPageAsm.AddEditRow("URL", Main.GetFormInputText(RequestNameImportLink, DefaultLink), "Enter the URL where the new template can be found.")
'                'Call ToolPageAsm.AddEditRow("Manage Stylesheets", Main.GetFormInputCheckBox(RequestNameManageStyles, ManageStyles), "When checked, the import will attempt to convert the stylesheets in the new template to Shared Styles. Import statements in the stylesheet will be handled by including shared styles. Although this results in a more managable configuration, it may not be compatible with some stylesheet techniques, such as conditional imports. If not checked, stylesheets will be handled as files and stored in the appropriate folder of your website just as files are stored.")
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                'Call AddToolPanel_Tools(ToolPageAsm, RQS, TemplateID)
'                GetContent = ToolPageAsm.GetPage()
'            Case FormOpen
'                '
'                ' Open a template
'                '
'                Copy = "" _
'                    & "<p>&nbsp;</p>" _
'                    & "<p>Select the template you wish to edit. When you are done, hit edit to begin editing. No changes will be made until you press Save or OK on the edit form.</p>" _
'                    & "<p>" & Main.GetFormInputSelect(RequestNameTemplateID, TemplateID, "page Templates") & "</p>" _
'                    & ""
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonOK)
'                Call ToolPageAsm.AddElement("BodyHeader", "Edit a template")
'                Call ToolPageAsm.AddElement("Body", Copy)
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                GetContent = ToolPageAsm.GetPage()
'            Case FormEdit
'                '
'                ' Edit a template
'                '
'                CS = Main.OpenCSSQL("default", "select name,bodyhtml from cctemplates where id=" & TemplateID)
'                If Main.IsCSOK(CS) Then
'                    TemplateBodyHTML = Main.GetCS(CS, "bodyhtml")
'                    TemplateName = Main.GetCS(CS, "name")
'                    If TemplateName = "" Then
'                        TemplateName = "Template # " & TemplateID
'                        Call Main.ExecuteSQL("default", "update cctemplates set name=" & KmaEncodeSQLText(TemplateName) & " where id=" & TemplateID)
'                    End If
'                End If
'                Call Main.CloseCS(CS)
'                Copy = Main.GetTemplateEditor2(TemplateID, RequestNameBodyHTML, TemplateBodyHTML)
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonSave)
'                Call ToolPageAsm.AddButton(ButtonOK)
'                Call ToolPageAsm.AddElement("BodyHeader", "Edit template " & TemplateName)
'                Call ToolPageAsm.AddElement("Body", Copy)
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                'Call AddToolPanel_Tools(ToolPageAsm, RQS, TemplateID)
'                GetContent = ToolPageAsm.GetPage()
'            Case FormCreateNew
'                '
'                ' Create a new template
'                '
'                Copy = "" _
'                    & "<p>&nbsp;</p>" _
'                    & "<p>To create a new template, enter a template name and hit create. The template name must be unique in your website.</p>" _
'                    & ""
'                Call ToolPageAsm.Init(Main)
'                'Call ToolPageAsm.AddHeader("Template Importer", "The template Importer lets you import, create and modify Page Templates.")
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Call ToolPageAsm.AddButton(ButtonCreate)
'                Call ToolPageAsm.AddElement("BodyHeader", "Create a new template")
'                Call ToolPageAsm.AddElement("Body", Copy)
'                Call ToolPageAsm.AddEditRow("Name", Main.GetFormInputText(RequestNameTemplateName, Main.GetStreamText(RequestNameTemplateName)), "Enter the name of the new template.")
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                GetContent = ToolPageAsm.GetPage()
'            Case Else
'                Hint = "900"
'                '
'                ' Root Form
'                '
'                Copy = "" _
'                    & "<p>&nbsp;</p>" _
'                    & "<p>Welcome to the Template Importer. To begin, choose one of the following actions.</p>" _
'                    & "<ul>" _
'                    & "<li>To import a template from an existing website, click <b>Import</b></li>" _
'                    & "<li>To start working on a new template, click <b>New</b></li>" _
'                    & "<li>To edit an existing template, click <b>Open</b></li>" _
'                    & "</ul>" _
'                    & ""
'                Call ToolPageAsm.Init(Main)
'                Hint = "901"
'                Call ToolPageAsm.AddHeader("Template Importer", "")
'                'Call ToolPageAsm.AddHeader("Template Importer", "The template Importer lets you import, create and modify Page Templates.")
'                Hint = "902"
'                Call ToolPageAsm.AddButton(ButtonCancel)
'                Hint = "903"
'                Call ToolPageAsm.AddElement("BodyHeader", "Template Importer Home")
'                Hint = "904"
'                Call ToolPageAsm.AddElement("Body", Copy)
'                Hint = "905"
'                Call AddToolPanel_Files(ToolPageAsm, RQS, TemplateID, FormID)
'                Hint = "906"
'                GetContent = ToolPageAsm.GetPage()
        End Select
    End If
    '
    Exit Function
ErrorTrap:
    HandleClassError ("GetContent, Hint=[" & Hint & "]")
End Function
'
'=================================================================================
'   Add the tools panel
'=================================================================================
'
Private Sub AddToolPanel_Files(ToolPageAsm As Object, RQS As String, TemplateID As Long, CurrentFormID As Long)
    On Error GoTo ErrorTrap
    '
    Dim QS As String
    '
    QS = ModifyQueryString(RQS, RequestNameFormID, FormCreateNew, True)
    Call ToolPageAsm.AddTool("File", "New", "?" & QS)
    '
    QS = ModifyQueryString(RQS, RequestNameFormID, FormOpen, True)
    Call ToolPageAsm.AddTool("File", "Open", "?" & QS)
    '
    QS = ModifyQueryString(RQS, RequestNameFormID, FormFileView, True)
    Call ToolPageAsm.AddTool("File", "View Web Files", "?" & QS)
    '
    If TemplateID <> 0 Then
        QS = ModifyQueryString(RQS, RequestNameFormID, FormEdit, True)
        Call ToolPageAsm.AddTool("Tools", "Edit Template", "?" & QS)
        '
        QS = ModifyQueryString(RQS, RequestNameFormID, FormSiteStyles, True)
        Call ToolPageAsm.AddTool("Tools", "Edit Site Styles", "?" & QS)
        '
        QS = ModifyQueryString(RQS, RequestNameFormID, FormTemplateStyles, True)
        Call ToolPageAsm.AddTool("Tools", "Edit Template Styles", "?" & QS)
        '
        QS = ModifyQueryString(RQS, RequestNameFormID, FormImport, True)
        Call ToolPageAsm.AddTool("Tools", "Import Website", "?" & QS)
        '
    End If
    '
    '
    Exit Sub
ErrorTrap:
    HandleClassError ("AddToolPanel_Files")
End Sub
'
'=================================================================================
'   Add the tools panel
'=================================================================================
'
Private Sub AddToolPanel_Tools(ToolPageAsm As ToolPageClass, RQS As String, TemplateID As Long)
    On Error GoTo ErrorTrap
    '
    Dim QS As String
    '
    If TemplateID <> 0 Then
        QS = ModifyQueryString(RQS, RequestNameFormID, FormImport, True)
        Call ToolPageAsm.AddTool("Tools", "Import Website", "?" & QS)
    End If
    '
    Exit Sub
ErrorTrap:
    HandleClassError ("AddToolPanel_Tools")
End Sub
'
'=================================================================================
'   Handle errors from this class
'=================================================================================
'
Private Sub HandleClassError(MethodName As String, Optional NotTrap As Boolean, Optional ResumeNext As Boolean)
    '
    Call HandleError("DesignerClass(" & ApplicationName & ")", MethodName, Err.Number, Err.Source, Err.Description, (Not NotTrap), ResumeNext)
    '
End Sub
'
'
'
Private Function CreateNewTemplate(TemplateName As String) As Long
    On Error GoTo ErrorTrap
    '
    '
    Dim CS As Long
    Dim IsNameOK As Boolean
    Dim TemplateID As Long
    '
    If TemplateName = "" Then
        Call Main.AddUserError("you must select a unique, non-blank name for your template to beging. Select a different name, or use the 'Open' tool to edit the existing template.")
    Else
        CS = Main.OpenCSContent("Page Templates", "name=" & KmaEncodeSQLText(TemplateName))
        IsNameOK = Not Main.IsCSOK(CS)
        Call Main.CloseCS(CS)
        '
        If Not IsNameOK Then
            Call Main.AddUserError("The name [" & TemplateName & "] is already used by another template. Select a different name, or use the 'Open' tool to edit the existing template.")
            TemplateName = ""
        Else
            CS = Main.InsertCSRecord("Page Templates")
            If Main.IsCSOK(CS) Then
                Call Main.SetCS(CS, "name", TemplateName)
                TemplateID = Main.GetCSInteger(CS, "ID")
                Call Main.SetVisitProperty(VisitPropertyTemplateID, TemplateID)
            End If
            Call Main.CloseCS(CS)
        End If
    End If
    CreateNewTemplate = TemplateID
    '
    Exit Function
ErrorTrap:
    HandleClassError ("CreateNewTemplate")
End Function
'
'
'
Private Sub ImportTemplate(TemplateID As Long, Link As String, ManageStyles As Boolean, BuildVersion As String)
    On Error GoTo ErrorTrap
    '
    '
    Dim CS As Long
    Dim HTTP As New HTTPClass
    Dim Doc As String
    Dim BasePath As String
    Dim DocBody As String
    Dim DocHead As String
    Dim SourceHost As String
    Dim Pos As Long
    '
    BasePath = GetBasePath(Link)
    SourceHost = ""
    If Link = "" Then
        Main.AddUserError ("The source link can not be blank")
    Else
        SourceHost = Link
        If InStr(1, SourceHost, "://") = 0 Then
            SourceHost = "http://" & SourceHost
        End If
        
        Pos = InStr(1, SourceHost, "://")
        If Pos > 0 Then
            Pos = InStr(Pos + 3, SourceHost, "/")
            If Pos > 0 Then
                SourceHost = Mid(SourceHost, 1, Pos - 1)
            End If
        End If
        If SourceHost = "" Then
            Main.AddUserError ("The host name of the link could not be determined [" & HTTP.ResponseHeader & "]")
        Else
            Doc = HTTP.GetURL(Link)
            If Doc = "" Then
                Main.AddUserError ("The document requested returned [" & HTTP.ResponseHeader & "]")
            Else
                Call ImportTemplate_Convert(Doc, BasePath, SourceHost, TemplateID, ManageStyles, BuildVersion)
                Call Main.ClearPageTemplateCache
            End If
        End If
    End If
    '
    '
    Exit Sub
ErrorTrap:
    HandleClassError ("ImportTemplate")
End Sub
'
'   parse the page associated with the given Doc, return any errors
'       fills the Doc with all tags found
'
Private Function ImportTemplate_Convert(Doc As String, BasePath As String, SourceHost As String, TemplateID As Long, ManageStyles As Boolean, BuildVersion As String)
    On Error GoTo ErrorTrap
    '
    Dim ImportBasePath As String
    Dim StyleTag As String
    Dim BodyTag As String
    Dim CS As Long
    Dim Pos As Long
    Dim posStart As Long
    Dim posEnd As Long
    Dim LinkType As String
    Dim StyleSheet As String
    Dim Link As String
    Dim Position As Long
    Dim ElementCount As Long
    Dim TagCount As Long
    Dim TagName As String
    Dim TagContent As String
    Dim TagHTTPEquiv As String
    Dim TagAlt As String
    Dim TagHeight As String
    Dim TagWidth As String
    Dim TagFace As String
    Dim FontFound As Boolean
    '
    Dim kmaParse As kmaHTML.ParseClass
    Dim ElementPointer As Long
    Dim HTMLCount As Long
    Dim HTMLFound As Boolean
    Dim BodyCount As Long
    Dim FrameCount As Long
    Dim TableCount As Long
    Dim TDCount As Long
    Dim TRCount As Long
    Dim SelectCount As Long
    '
    Dim FontsUsedPointer As Long
    Dim FontsUsedCount As Long
    Dim FontsUsed() As String
    Dim Output As New FastString.FastStringClass
    Dim ElementText As String
    Dim RootRelativeLink As String
    Dim TagDone As Boolean
    Dim PositionSOL As Long
    Dim PositionEOL As Long
    Dim PositionEqual As Long
    Dim DocConverted As String
    Dim IsInHead As Boolean
    Dim OtherHeadTags As String
    Dim HTTP As HTTPClass
    'Dim HTTP As New HTTPClass
    '
    Call AppendLog("methodname", "Analyze tags, body [" & Len(Doc) & "] bytes")
    '
    ' patch for wysiwyg editor deleting adjacent empty span tags
    '
    Doc = Replace(Doc, "<span><span>", "<span class=""fpo""><span class=""fpo"">", 1, 99, vbTextCompare)
    '
    Set kmaParse = New ParseClass
    Call kmaParse.Load(Doc)
    ElementPointer = 0
    FontsUsedCount = 0
    ElementCount = kmaParse.ElementCount
    '
    Call AppendLog("methodname", "Analyze tags, [" & ElementCount & "] elements found")
    '
    Do While ElementPointer < ElementCount
        ElementText = kmaParse.Text(ElementPointer)
        If kmaParse.IsTag(ElementPointer) Then
            TagCount = TagCount + 1
            TagName = kmaParse.TagName(ElementPointer)
            Select Case UCase(TagName)
                Case "HEAD"
                    IsInHead = True
                Case "/HEAD"
                    IsInHead = False
                Case "FORM"
                    Link = kmaParse.ElementAttribute(ElementPointer, "action")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "TD"
                    Link = kmaParse.ElementAttribute(ElementPointer, "Background")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "BODY"
                    Link = kmaParse.ElementAttribute(ElementPointer, "Background")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                    BodyTag = ElementText
                Case "BASE"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        BasePath = GetBasePath(Link)
                        ElementText = ""
                    End If
                Case "A"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                        'End If
                    End If
                Case "META"
                    '
                    TagHTTPEquiv = kmaParse.ElementAttribute(ElementPointer, "HTTPEquiv")
                    TagContent = kmaParse.ElementAttribute(ElementPointer, "Content")
                    If UCase(TagHTTPEquiv = "REFRESH") And (TagContent <> "") Then
                        Link = UCase(TagContent)
                        Position = InStr(1, Link, "URL")
                        If Position <> 0 Then
                            Position = Position + 3
                            Do While Mid(Link, Position, 1) = " "
                                Position = Position + 1
                                DoEvents
                            Loop
                            If Mid(Link, Position, 1) = "=" Then
                                Position = Position + 1
                                Do While Mid(Link, Position, 1) = " "
                                    Position = Position + 1
                                    DoEvents
                                Loop
                                Link = Trim(Mid(TagContent, Position))
                                If IsLinkToThisHost(SourceHost, Link) Then
                                    RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                                    ElementText = Replace(ElementText, Link, RootRelativeLink)
                                End If
                            End If
                        End If
                    End If
                    If IsInHead Then
                        OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                    End If
                        
                Case "AREA"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "IMG"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "EMBED"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "FRAMESET"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "FRAME"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                Case "LINK"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                    LinkType = kmaParse.ElementAttribute(ElementPointer, "TYPE")
                    If (IsLinkToThisHost(SourceHost, Link)) And (LCase(LinkType) = "text/css") Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        If ManageStyles Then
                            '
                            ' css styles - convert to relative styles and comment out element
                            '
                            Set HTTP = New HTTPClass
                            StyleSheet = HTTP.GetURL(SourceHost & RootRelativeLink)
                            HTTP.Interrupt
                            Set HTTP = Nothing
                            ImportBasePath = RootRelativeLink
                            Pos = InStrRev(ImportBasePath, "/")
                            If Pos > 0 Then
                                ImportBasePath = Mid(ImportBasePath, 1, Pos)
                            End If
                            StyleSheet = ConvertStyles_HandleFileReferences(StyleSheet, ImportBasePath, SourceHost, "", TemplateID)
                            Call SaveSharedStyle(TemplateID, kmaDecodeURL(RootRelativeLink), StyleSheet)
                            ElementText = "<!-- " & ElementText & " -->"
                        Else
                            '
                            ' inmanaged css or other files, download them and convert link
                            '
                            Call ConvertStyleLink_HandleImports(Link, BasePath, SourceHost, "", TemplateID)
                            Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & kmaDecodeURL(RootRelativeLink))
                            ElementText = Replace(ElementText, Link, RootRelativeLink)
                        End If
                    End If
                    If IsInHead Then
                        OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                    End If
                Case "SCRIPT"
                    '
                    Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                    If IsLinkToThisHost(SourceHost, Link) Then
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & kmaDecodeURL(RootRelativeLink))
                        ElementText = Replace(ElementText, Link, RootRelativeLink)
                    End If
                    '
                    ' Skip to the </SCRIPT> TAG
                    '
                    If InStr(1, ElementText, "/>") = 0 Then
                        '
                        ' Find the end of the tag
                        '
                        TagDone = False
                        Do While (Not TagDone) And (ElementPointer < ElementCount)
                            '
                            ' Get the next segment
                            '
                            ElementPointer = ElementPointer + 1
                            ElementText = ElementText & kmaParse.Text(ElementPointer)
                            If kmaParse.IsTag(ElementPointer) Then
                                '
                                ' Process a tag (should just be </SCRIPT>, but go until it is
                                '
                                TagCount = TagCount + 1
                                TagDone = (kmaParse.TagName(ElementPointer) = "/" & TagName)
                            End If
                        Loop
                    End If
                    If IsInHead Then
                        OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                    End If
                Case "STYLE"
                    '
                    ' Skip to the </Style> TAG, get the stylesheet between for processing
                    '
                    TagDone = False
                    Do While (Not TagDone) And (ElementPointer < ElementCount)
                        '
                        ' Process the next segment
                        '
                        ElementText = kmaParse.Text(ElementPointer)
                        If kmaParse.IsTag(ElementPointer) Then
                            '
                            ' Process a tag (should just be </SCRIPT>, but go until it is
                            '
                            TagCount = TagCount + 1
                            TagDone = (kmaParse.TagName(ElementPointer) = "/" & TagName)
                        End If
                        StyleTag = StyleTag & ElementText
                        If Not TagDone Then
                            Output.Add ElementText
                            ElementPointer = ElementPointer + 1
                        End If
                    Loop
                    If IsInHead Then
                        StyleTag = ConvertStyles_HandleFileReferences(StyleTag, BasePath, SourceHost, "", TemplateID)
                        OtherHeadTags = OtherHeadTags & vbCrLf & StyleTag
                    End If
                Case "INPUT"
                    '
                    If UCase(kmaParse.ElementAttribute(ElementPointer, "TYPE")) = "IMAGE" Then
                        Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                        If IsLinkToThisHost(SourceHost, Link) Then
                            RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                            Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & Replace(kmaDecodeURL(RootRelativeLink), "/", "\"))
                            ElementText = Replace(ElementText, Link, RootRelativeLink)
                        End If
                    End If
                Case "TITLE", "/TITLE"
                    ElementText = ""
                Case "SPAN"
                    '
                    ' wysiwyg editor auto deletes empty spans, all a class if there is not one
                    '
                    If LCase(kmaParse.ElementAttribute(ElementPointer, "class")) = "" Then
                        ElementText = Replace(ElementText, ">", " class=""emptySpanPlaceHolder"">")
                    End If
                Case Else
                    If IsInHead Then
                        OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                    End If
            End Select
        End If
        Output.Add ElementText
        ElementPointer = ElementPointer + 1
    Loop
    DocConverted = Output.Text
    '
    CS = Main.OpenCSContentRecord("Page Templates", TemplateID)
    If Main.IsCSOK(CS) Then
        Call Main.SetCS(CS, "bodyhtml", GetTagInnerHTML(DocConverted, "body", False))
        If BuildVersion > "3.3.994" Then
            If OtherHeadTags <> "" Then
                If Mid(OtherHeadTags, 1, 2) = vbCrLf Then
                OtherHeadTags = Mid(OtherHeadTags, 3)
                End If
            End If
            Call Main.SetCS(CS, "OtherHeadTags", OtherHeadTags)
            Call Main.SetCS(CS, "BodyTag", BodyTag)
        End If
    End If
    Call Main.CloseCS(CS)
    Call Main.ClearPageTemplateCache
    
    ImportTemplate_Convert = DocConverted
    Call AppendLog("methodname", "Analyze tags, [" & TagCount & "] tags found")
    '
    Set kmaParse = Nothing
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassError("ImportTemplate_Convert", True)
    Resume Next
    'Set kmaParse = Nothing
End Function
'
'=========================================================================================================
'   Get the path of the source Link
'
'   http://gcm.brandeveolve.com/logo-Main.jpg with to /
'   http://gcm.brandeveolve.com/images/logo-Main.jpg with to /images/
'   /images/logo-Main.jpg with to /images/
'
'=========================================================================================================
'
Private Function GetBasePath(Link As String) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim Pos As Long
    Dim LoopCnt As Long
    Dim posStart As Long
    '
    GetBasePath = Link
    If InStr(1, Link, "/") = 1 Then
        '
        '   case /images/logo-Main.jpg with to /images/
        '
        Pos = 1
        Do While (Pos > 0) And LoopCnt < 100
            GetBasePath = Mid(Link, 1, Pos)
            Pos = InStr(Pos + 1, Link, "/")
        Loop
    ElseIf InStr(1, Link, "://") <> 0 Then
        '
        '   case http://gcm.brandeveolve.com/images/logo-Main.jpg with any BasePath  to /images/logo-Main.jpg
        '
        GetBasePath = "/"
        Pos = InStr(1, Link, "://")
        If Pos > 0 Then
            Pos = InStr(Pos + 3, Link, "/")
            If Pos > 0 Then
                posStart = Pos
                Do While (Pos > 0) And LoopCnt < 100
                    GetBasePath = Mid(Link, posStart, Pos)
                    Pos = InStr(Pos + 1, Link, "/")
                Loop
                Pos = InStrRev(GetBasePath, "/")
                If Pos < Len(GetBasePath) Then
                    GetBasePath = Mid(GetBasePath, 1, Pos)
                End If
            End If
        End If
    Else
        '
        '   unknown case
        '
        GetBasePath = "/"
    End If
    
    '
    '
    Exit Function
ErrorTrap:
    HandleClassError ("GetBasePath")
End Function
'
'
'
Private Sub SaveSharedStyle(TemplateID As Long, Name As String, Styles As String)
    On Error GoTo ErrorTrap
    '
    '
    Dim CS As Long
    Dim SharedStylesID As Long
    '
    CS = Main.OpenCSContent("Shared Styles", "name=" & KmaEncodeSQLText(Name))
    If Not Main.IsCSOK(CS) Then
        Call Main.CloseCS(CS)
        CS = Main.InsertCSRecord("Shared Styles")
    End If
    If Main.IsCSOK(CS) Then
        SharedStylesID = Main.GetCS(CS, "ID")
        Call Main.SetCS(CS, "name", Name)
        Call Main.SetCS(CS, "stylefilename", Styles)
    End If
    Call Main.CloseCS(CS)
    If SharedStylesID <> 0 And TemplateID <> 0 Then
        CS = Main.OpenCSContent("Shared Styles Template Rules", "(styleid=" & SharedStylesID & ")and(TemplateID=" & TemplateID & ")")
        If Not Main.IsCSOK(CS) Then
            Call Main.CloseCS(CS)
            CS = Main.InsertCSRecord("Shared Styles Template Rules")
        End If
        If Main.IsCSOK(CS) Then
            Call Main.SetCS(CS, "styleid", SharedStylesID)
            Call Main.SetCS(CS, "TemplateID", TemplateID)
        End If
        Call Main.CloseCS(CS)
    End If

    '
    Exit Sub
ErrorTrap:
    HandleClassError ("SaveSharedStyle")
End Sub
'
'
'
Private Function ConvertStyles_HandleFileReferences(StyleSheet As String, BasePath As String, SourceHost As String, BlockRootRelativeLinkList As String, TemplateID As Long) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim PosURLStart As Long
    Dim PosURLEnd As Long
    Dim URLCommand As String
    Dim Value As String
    Dim Loopcnt2 As Long
    Dim Name As String
    Dim NameValue As String
    Dim PtrStart As Long
    Dim PtrEnd As Long
    Dim Line As String
    Dim Lines() As String
    Dim LineCnt As Long
    Dim LinePtr As Long
    Dim Ptr As Long
    Dim PosImport As Long
    Dim posStart As Long
    Dim posEnd As Long
    Dim Link As String
    Dim RootRelativeLink As String
    Dim LoopCnt As Long
    Dim ImportedStyle As String
    Dim HTTP As New HTTPClass
    Dim Copy As String
Dim ImportStyles As String

    '
    PosImport = 1
    Copy = StyleSheet
    '
    ' convert imports
    '
    Do While (PosImport <> 0) And LoopCnt < 100
        PosImport = InStr(PosImport, Copy, "@import", vbTextCompare)
        If PosImport <> 0 Then
            '
            ' style includes an import -- convert filename and load the file
            '
Dim URLPosStart As Long
Dim URLPosEnd As Long
            URLPosStart = InStr(PosImport, Copy, "url", vbTextCompare)
            If URLPosStart <> 0 Then
                posStart = InStr(URLPosStart, Copy, "(", vbTextCompare)
                If posStart <> 0 Then
                    posStart = posStart + 1
                    URLPosEnd = InStr(posStart, Copy, ")", vbTextCompare)
                    If URLPosEnd <> 0 Then
                        posEnd = URLPosEnd - 1
                        Link = Mid(Copy, posStart, posEnd - posStart + 1)
                        Link = Trim(Link)
                        If Left(Link, 1) = """" And Right(Link, 1) = """" Then
                            Link = Mid(Link, 2, Len(Link) - 2)
                        End If
                        If Left(Link, 1) = "'" And Right(Link, 1) = "'" Then
                            Link = Mid(Link, 2, Len(Link) - 2)
                        End If
                        
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        
Dim ImportBasePath As String
Dim PathPos As Long

ImportBasePath = RootRelativeLink
PathPos = InStrRev(ImportBasePath, "/")
If PathPos > 0 Then
    ImportBasePath = Mid(ImportBasePath, 1, PathPos)
End If
                        
                        
                        
                        ImportStyles = HTTP.GetURL(SourceHost & RootRelativeLink)
                        ImportStyles = ConvertStyles_HandleFileReferences(ImportStyles, ImportBasePath, SourceHost, BlockRootRelativeLinkList, TemplateID)
                        Call SaveSharedStyle(TemplateID, kmaDecodeURL(RootRelativeLink), ImportStyles)
                        
                        'RootRelativeLink = ConvertStyleLink_HandleImports(Link, BasePath, SourceHost, BlockRootRelativeLinkList)
Dim PosImportEnd As Long
Dim PosLineEnd As Long
Dim PosSemi As Long

PosImportEnd = URLPosEnd
PosLineEnd = InStr(PosImportEnd, Copy, vbLf)
PosSemi = InStr(PosImportEnd, Copy, ";")
If PosSemi < PosLineEnd Then
    PosImportEnd = PosSemi
Else
    PosImportEnd = PosLineEnd
End If
                        Copy = Mid(Copy, 1, PosImport - 1) & Mid(Copy, PosImportEnd + 1)
                        'PosImport = URLPosStart
                    End If
                End If
            End If
        End If
        LoopCnt = LoopCnt + 1
    Loop
    '
    ' convert other url() arguments
    '
    Lines = Split(Copy, "}")
    LineCnt = UBound(Lines) + 1
    If LineCnt > 0 Then
        For LinePtr = 0 To LineCnt - 1
            Line = Lines(LinePtr)
            PtrStart = InStrRev(Line, "{")
            LoopCnt = 0
            Do While PtrStart <> 0 And LoopCnt < 100
                PtrStart = PtrStart + 1
                PtrEnd = InStr(PtrStart, Line, ";")
                If PtrEnd = 0 Then
                    NameValue = Mid(Line, PtrStart)
                Else
                    NameValue = Mid(Line, PtrStart, PtrEnd - PtrStart + 1)
                End If
                'If InStr(1, NameValue, "page-top", vbTextCompare) <> 0 Then
                '    NameValue = NameValue
                'End If
                '
                ' determine if a URL in in this namevalue
                '
                PosURLStart = InStr(1, NameValue, "url(", vbTextCompare)
                If PosURLStart <> 0 Then
                    PosURLStart = PosURLStart + 4
                    PosURLEnd = InStr(PosURLStart, NameValue, ")")
                    If PosURLEnd <> 0 Then
                        Link = Mid(NameValue, PosURLStart, PosURLEnd - PosURLStart)
                        Link = trimAll(Link)
                        If Left(Link, 1) = """" And Right(Link, 1) = """" Then
                            Link = Mid(Link, 2, Len(Link) - 2)
                        End If
                        If Left(Link, 1) = "'" And Right(Link, 1) = "'" Then
                            Link = Mid(Link, 2, Len(Link) - 2)
                        End If
                        '
                        ' download the file and replace it in the source
                        '
                        'ImportStyles = ""
                        'RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        'ImportStyles = HTTP.GetURL(SourceHost & RootRelativeLink)
                        'ImportStyles = ConvertStyles_HandleFileReferences(ImportStyles, BasePath, SourceHost, BlockRootRelativeLinkList, TemplateID)
                        'Call SaveSharedStyle(TemplateID, kmaDecodeURL(RootRelativeLink), ImportStyles)
                        
Dim filePathPage As String
                        RootRelativeLink = ConvertLinkToRootRelative(Link, BasePath)
                        filePathPage = Replace(RootRelativeLink, "/", "\")
                        If Left(filePathPage, 1) = "\" Then
                            filePathPage = Mid(filePathPage, 2)
                        End If
                        Call GetURLToFile(SourceHost & RootRelativeLink, Main.PhysicalWWWPath & filePathPage)
                        NameValue = Replace(NameValue, Link, RootRelativeLink, 1, -1, vbTextCompare)
                        Line = Mid(Line, 1, PtrStart - 1) & NameValue & Mid(Line, PtrEnd + 1)
                    End If
                End If
                PtrStart = PtrEnd
                'If PtrEnd > 0 Then
                '    PtrStart = PtrEnd + 1
                'End If
                LoopCnt = LoopCnt + 1
            Loop
            If Ptr <> 0 Then
                
            End If
            Lines(LinePtr) = Line
        Next
    End If
    If Join(Lines, "}") <> Copy Then
        Copy = Copy
    End If
    Copy = Join(Lines, "}")
    '
    ' Done
    '
    ConvertStyles_HandleFileReferences = Copy
    '
    '
    Exit Function
ErrorTrap:
    HandleClassError ("ConvertStyles_HandleFileReferences")
End Function
'
'
'
Private Function ConvertStyleLink_HandleImports(Link As String, BasePath As String, SourceHost As String, BlockRootRelativeLinkList As String, TemplateID As Long) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim Pos As Long
    Dim ImportedStyle As String
    Dim HTTP As New HTTPClass
    Dim Filename As String
    Dim ConvertedLink As String
    Dim ImportLink As String
    Dim LinkPath As String
    '
    ConvertedLink = ConvertLinkToRootRelative(Link, BasePath)
    ConvertStyleLink_HandleImports = ConvertedLink
    If InStr(1, BlockRootRelativeLinkList, ConvertedLink, vbTextCompare) = 0 Then
        ImportLink = SourceHost & ConvertedLink
        ImportedStyle = HTTP.GetURL(ImportLink)
        Pos = InStrRev(ConvertedLink, "/")
        If Pos > 0 Then
            LinkPath = Mid(ConvertedLink, 1, Pos)
        End If
        'If Mid(Linkpath, 1, 1) = "/" Then
        '    Linkpath = Mid(Linkpath, 2)
        'End If
        ImportedStyle = ConvertStyles_HandleFileReferences(ImportedStyle, LinkPath, SourceHost, BlockRootRelativeLinkList & "," & ConvertedLink, TemplateID)
        Filename = Main.PhysicalWWWPath & Replace(ConvertedLink, "/", "\")
        Call Main.SaveFile(Filename, ImportedStyle)
    End If
    '
    Exit Function
ErrorTrap:
    HandleClassError ("ConvertStyleLink_HandleImports")
End Function
'
'
'
Private Sub GetURLToFile(SrcLink As String, DstFilename As String)
    On Error GoTo ErrorTrap
    '
    Dim HTTP As New HTTPClass
    Dim FS As New kmaFileSystem3.FileSystemClass
    Dim TempFilename As String
    Dim Pos As Long
    Dim HTTPStatus As String
    Dim ErrDescription As String
    '
    If InStr(1, vbCrLf & filesFetchedList & vbCrLf, vbCrLf & SrcLink & vbCrLf, vbTextCompare) = 0 Then
        filesFetchedList = filesFetchedList & vbCrLf & SrcLink
        
        DstFilename = Replace(DstFilename, "/", "\")
        Pos = InStrRev(DstFilename, "\")
        If Pos > 0 Then
            TempFilename = Main.PhysicalFilePath & "temp\" & Mid(DstFilename, Pos + 1)
            On Error Resume Next
            Set HTTP = Nothing
            Set HTTP = New HTTPClass
            Call HTTP.GetURLToFile(SrcLink, TempFilename)
            ErrDescription = Err.Description
            If Err.Number <> 0 Then
                Call Main.AddUserError("Error loading file [" & SrcLink & "], " & ErrDescription)
            Else
                'HTTPStatus = HTTP.SocketResponse
                HTTPStatus = KmaGetLine(HTTP.ResponseHeader)
                If InStr(1, HTTPStatus, "200") = 0 Then
                    Call Main.AddUserError("Error loading file [" & SrcLink & "], " & HTTPStatus)
                Else
                    If FS.CheckFile(TempFilename) Then
                        Call Main.DeleteFile(DstFilename)
                        Call FS.CopyFile(TempFilename, DstFilename)
                        If FS.CheckFile(DstFilename) Then
                            Call Main.DeleteFile(TempFilename)
                        End If
                    End If
                End If
            End If
        End If
    Else
        filesFetchedList = filesFetchedList
    End If
    '
    '
    Exit Sub
ErrorTrap:
    HandleClassError ("GetURLToFile")
End Sub
'
'   temporary until a Contensive Build can be installed with the public
'
Private Function GetFileView() As String
    On Error GoTo ErrorTrap
    '
    Dim obj As Object
    Dim OptionString As String
    '
    Set obj = CreateObject("aoFileview.FileViewClass")
    If IsObject(obj) Then
        OptionString = "" _
            & vbCrLf & "BaseFolder" _
            & vbCrLf & "AllowEdit=yes" _
            & vbCrLf & "AllowNavigation=yes" _
            & vbCrLf & "FileSystem=Root"
        Call obj.Init(Main)
        GetFileView = obj.GetContent(OptionString)
        '
        ' remove form
        '
        GetFileView = Replace(GetFileView, "<form ", "<xform ", 1, -1, vbTextCompare)
        GetFileView = Replace(GetFileView, "</form ", "</xform ", 1, -1, vbTextCompare)
        '
        ' remove buttons
        '
        GetFileView = Replace(GetFileView, "type=submit", "type=hidden", 1, -1, vbTextCompare)
        GetFileView = Replace(GetFileView, "type=""submit", "type=""hidden", 1, -1, vbTextCompare)
    End If
    '
    Exit Function
ErrorTrap:
    HandleClassError ("GetFileView")
End Function
'
'========================================================================================================
'   ConvertLinkToRootRelative
'
'   /images/logo-Main.jpg with any Basepath to /images/logo-Main.jpg
'   http://gcm.brandeveolve.com/images/logo-Main.jpg with any BasePath  to /images/logo-Main.jpg
'   images/logo-Main.jpg with Basepath '/' to /images/logo-Main.jpg
'   logo-Main.jpg with Basepath '/images/' to /images/logo-Main.jpg
'
'========================================================================================================
'
Private Function ConvertLinkToRootRelative(Link As String, BasePath As String) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim Pos As Long
    Dim workingbase As String
    Dim WorkingLink As String
    '
    workingbase = BasePath
    WorkingLink = Link
    ConvertLinkToRootRelative = WorkingLink
    If InStr(1, WorkingLink, "/") = 1 Then
        '
        '   case /images/logo-Main.jpg with any Basepath to /images/logo-Main.jpg
        '
    ElseIf InStr(1, WorkingLink, "://") <> 0 Then
        '
        '   case http://gcm.brandeveolve.com/images/logo-Main.jpg with any BasePath  to /images/logo-Main.jpg
        '
        Pos = InStr(1, WorkingLink, "://")
        If Pos > 0 Then
            Pos = InStr(Pos + 3, WorkingLink, "/")
            If Pos > 0 Then
                ConvertLinkToRootRelative = Mid(WorkingLink, Pos)
            Else
                '
                ' This is just the domain name, RootRelative is the root
                '
                ConvertLinkToRootRelative = "/"
            End If
        End If
    Else
        '
        '   case images/logo-Main.jpg with Basepath '/' to /images/logo-Main.jpg
        '   case logo-Main.jpg with Basepath '/images/' to /images/logo-Main.jpg
        '
        Do While Left(WorkingLink, 3) = "../"
            If Len(WorkingLink) > 3 Then
                WorkingLink = Mid(WorkingLink, 4)
            Else
                WorkingLink = ""
            End If
            If Right(workingbase, 1) = "/" Then
                workingbase = Mid(workingbase, 1, Len(workingbase) - 1)
            End If
            Pos = InStrRev(workingbase, "/")
            If Pos > 0 Then
                workingbase = Left(workingbase, Pos - 1)
            End If
            workingbase = workingbase & "/"
            
        Loop
        
        ConvertLinkToRootRelative = workingbase & WorkingLink
    End If
    '
    '
    Exit Function
ErrorTrap:
    HandleClassError ("ConvertLinkToRootRelative")
End Function
'
'
'
Private Function trimAll(Source As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Ptr As Long
    Dim test As Long
    '
    trimAll = Source
    For Ptr = 1 To Len(trimAll)
        test = Asc(Mid(trimAll, Ptr, 1))
        If (test > 32) And (test <= 128) Then
            Exit For
        End If
        
    Next
    trimAll = Mid(trimAll, Ptr)
    For Ptr = Len(trimAll) To 1 Step -1
        test = Asc(Mid(trimAll, Ptr, 1))
        If (test > 32) And (test <= 128) Then
            Exit For
        End If
    Next
    trimAll = Mid(trimAll, 1, Ptr)
    '
    Exit Function
ErrorTrap:
    HandleClassError ("trimAll")
End Function
