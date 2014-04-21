
Option Explicit On

Imports Contensive.BaseClasses

Namespace Contensive.addons.themeManager
    Public Class templateImporterClass
        Inherits BaseClasses.AddonBaseClass
        '
        Private Const debugging = False
        '
        Public Const RequestNameFormID = "formid"
        Public Const RequestNameTemplateID = "templateid"
        Public Const RequestNameTemplateName = "templatename"
        Public Const RequestNameBodyHTML = "TemplateBody"
        Public Const RequestNameImportLink = "ImportLink"
        Public Const RequestNameManageStyles = "ManageStyles"
        Public Const RequestNameSiteStylesInline = "TemplateImporterSiteStylesInline"
        Public Const RequestNameSiteStyles = "TemplateImporterSiteStyles"
        Public Const RequestNameTemplateStyles = "TemplateImporterTemplateStyles"
        Public Const RequestNameJoomlaFile = "JoomlaFile"
        Public Const RequestNameImportMethod = "importMethod"
        Public Const importMethodToFile = "toFile"
        Public Const importMethodToRecord = "toRecord"
        '
        Public Const VisitPropertyTemplateID = "TemplateImporterTemplateID"
        '
        Public Const FormEmpty = -1              ' return empty form and Admin will exit the addon
        Public Const FormRoot = 1
        Public Const FormCreateNew = 2
        Public Const FormEdit = 3
        Public Const FormImport = 4
        Public Const FormOpen = 5
        Public Const FormFileView = 6
        Public Const FormSiteStyles = 7
        Public Const FormTemplateStyles = 8
        Public Const FormImportJoomla = 9
        Public Const FormImportOnePage = 10
        Public Const FormDone = 11
        '
        Public Const ButtonCreate = " Create New Template "
        Public Const ButtonBeginEditing = " Begin Editing "
        Public Const ButtonBeginImport = " Begin Import "
        Public Const ButtonCancel = " Cancel "
        '
        Private ApplicationName As String
        Private issueList As String
        Private filesFetchedList As String
        '
        '=================================================================================
        '   Aggregate Object Interface
        '=================================================================================
        '
        Public Overrides Function Execute(ByVal cp As CPBaseClass) As Object
            Dim body As String = ""
            Dim Hint As String = ""
            Try
                Dim man As New adminFramework.pageWithNavClass
                '
                Dim buttonBar As String = ""
                '
                Dim templateFilename As String
                Dim DefaultLink As String
                Dim BuildVersion As String
                Dim ManageStyles As Boolean
                Dim ImportLink As String
                'Dim ToolPageAsm As Object
                Dim Button As String
                Dim RQS As String
                Dim FormID As Integer
                Dim TemplateName As String
                Dim TemplateID As Integer
                Dim importMethod As String
                '
                'ToolPageAsm = CreateObject("ccPageAsm.ToolPageClass")
                '
                Hint = "100"
                ApplicationName = cp.Site.Name
                '
                BuildVersion = cp.Site.GetProperty("buildversion")
                RQS = cp.Doc.RefreshQueryString
                Button = cp.Doc.Var("button")
                '
                ' set defaults
                '
                FormID = FormImportOnePage
                importMethod = importMethodToFile
                DefaultLink = cp.Visit.GetProperty("TemplateImporterLastLink", "http://www.contensive.com")
                TemplateName = cp.Visit.GetProperty("TemplateImporterLastTemplateName", "")
                '
                If Button <> "" Then
                    If Button = ButtonCancel Then
                        Return ""
                    End If
                    FormID = cp.Utils.EncodeInteger(cp.Doc.Var(RequestNameFormID))
                    TemplateID = cp.Utils.EncodeInteger(cp.Visit.GetProperty((VisitPropertyTemplateID)))
                    '
                    ' Process buttons
                    '
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
                                DefaultLink = cp.Doc.Var(RequestNameImportLink)
                                Call cp.Visit.SetProperty("TemplateImporterLastLink", DefaultLink)
                                TemplateName = cp.Doc.Var(RequestNameTemplateName)
                                Call cp.Visit.SetProperty("TemplateImporterLastTemplateName", TemplateName)
                                importMethod = cp.Doc.Var(RequestNameImportMethod)
                                If importMethod = "" Then
                                    importMethod = importMethodToFile
                                End If
                                templateFilename = TemplateName
                                templateFilename = templateFilename.Replace(" ", "-")
                                templateFilename = templateFilename.Replace(" ", "-")
                                templateFilename = "template_" & templateFilename & ".html"
                                If cp.File.fileExists(cp.Site.PhysicalWWWPath & templateFilename) Then
                                    '
                                    ' template is already in use
                                    '
                                    Call cp.UserError.Add("The template file [" & templateFilename & "] is already in use. Please select another templatename.")
                                Else
                                    TemplateID = CreateNewTemplate(cp, TemplateName)
                                    If TemplateID = 0 Then
                                        Call cp.UserError.Add("There was a problem creating the page Template [" & TemplateName & "]. Select a different name, or use the 'Open' tool to edit the existing template.")
                                        TemplateName = ""
                                    Else
                                        ImportLink = cp.Doc.Var(RequestNameImportLink)
                                        If ImportLink = "" Then
                                            Call cp.UserError.Add("To import a template, enter a URL.")
                                        Else
                                            ManageStyles = True
                                            If cp.UserError.OK Then
                                                Try
                                                    If (cp.File.ReadVirtual("templates\styles.css") <> "") Then
                                                        '
                                                        ' show warning if there is a site stylesheet
                                                        '
                                                        Call cp.UserError.Add("This site contains a stylesheet in the site styles. These styles may interfere with the new template. Copy this styleshet to a shared stylesheet and associate it to the templates that need it.")
                                                    End If
                                                    Call ImportTemplate(cp, TemplateID, ImportLink, ManageStyles, BuildVersion, templateFilename, importMethod)
                                                    Call cp.Cache.Clear("page templates")
                                                    FormID = FormDone
                                                Catch ex As Exception
                                                    '
                                                    ' stay on this form and let htem try again
                                                    '
                                                    Call cp.UserError.Add("There was an unexpected problem during the template import. The error message was [" & ex.Message & "]")
                                                    FormID = FormID
                                                End Try
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Case Else
                            FormID = FormImportOnePage
                    End Select
                End If
                '
                ' Get Forms
                '
                Hint = "500"
                Call cp.Doc.AddRefreshQueryString(RequestNameFormID, FormID)
                Select Case FormID
                    Case FormDone
                        '
                        ' simple one page form -- enter the name for the template and enter URL
                        '
                        Dim form As New adminFramework.formSimpleClass
                        '
                        form.title = "Template Importer"
                        form.description = "<p>The template was imported. Any issues will be listed here.</p>"
                        form.body &= cp.Html.div("<a href=""" & cp.Site.GetProperty("adminURL") & "?af=4&cid=" & cp.Content.GetID("page templates") & "&id=" & TemplateID & """>Edit the new template</a>", , "tiFormCaption")
                        form.body &= cp.Html.div("<a href=""" & cp.Site.GetProperty("adminURL") & "?cid=" & cp.Content.GetID("shared Styles") & """>View Shared Styles</a>", , "tiFormCaption")
                        body = form.getHtml(cp)

                        'body &= cp.Html.h1("Template Importer", , "tiTitle")
                        'body &= cp.Html.p("The template was imported. Any issues will be listed here.", , "tiDescription")
                        'body &= cp.Html.div("<a href=""" & cp.Site.GetProperty("adminURL") & "?af=4&cid=" & cp.Content.GetID("page templates") & "&id=" & TemplateID & """>Edit the new template</a>", , "tiFormCaption")
                        'body &= cp.Html.div("<a href=""" & cp.Site.GetProperty("adminURL") & "?cid=" & cp.Content.GetID("shared Styles") & """>View Shared Styles</a>", , "tiFormCaption")
                    Case Else
                        '
                        ' FormImportOnePage - simple one page form -- enter the name for the template and enter URL
                        '
                        Dim form As New adminFramework.formNameValueRowsClass
                        '
                        form.addFormButton(ButtonCancel)
                        form.addFormButton(ButtonBeginImport)
                        form.title = "Template Importer"
                        form.description = "<p>To create a new template, enter a template name and the URL where the page can be found. The template name must be unique in your website.</p>"
                        '
                        'buttonBar &= cp.Html.Button("button", ButtonCancel)
                        'buttonBar &= cp.Html.Button("button", ButtonBeginImport)
                        '
                        'body &= cp.Html.h1("Template Importer", , "tiTitle")
                        'body &= cp.Html.p("To create a new template, enter a template name and the URL where the page can be found. The template name must be unique in your website.", , "tiDescription")
                        '
                        form.addRow()
                        form.rowName = "Template Name"
                        form.rowValue = cp.Html.InputText(RequestNameTemplateName, TemplateName)
                        '
                        form.addRow()
                        form.rowName = "Source URL"
                        form.rowValue = cp.Html.InputText(RequestNameImportLink, DefaultLink)
                        '
                        form.addRow()
                        form.rowName = "Import Method"
                        form.rowValue = "" _
                            & cp.Html.div(cp.Html.RadioBox(RequestNameImportMethod, importMethodToFile, importMethod) & "Save to File") _
                            & cp.Html.div(cp.Html.RadioBox(RequestNameImportMethod, importMethodToRecord, importMethod) & "Save to Record") _
                            & ""
                        'body &= cp.Html.ul("<p>Select a template import Method.</p>" _
                        '    & cp.Html.li(cp.Html.RadioBox(RequestNameImportMethod, importMethodToFile, importMethod) & "Save to File") _
                        '    & cp.Html.li(cp.Html.RadioBox(RequestNameImportMethod, importMethodToRecord, importMethod) & "Save to Record") _
                        '    & "")
                        '
                        'body &= cp.Html.div("Enter the name of the new template.", , "tiFormCaption")
                        'body &= cp.Html.div(cp.Html.InputText(RequestNameTemplateName, TemplateName), , "tiFormInput")
                        ''
                        'body &= cp.Html.div("Enter the URL where the new template can be found", , "tiFormCaption")
                        'body &= cp.Html.div(cp.Html.InputText(RequestNameImportLink, DefaultLink), , "tiFormInput")
                        '
                        'ns.body = body
                        body = form.getHtml(cp)
                        'body &= cp.Html.div(buttonBar, , "tiButtonBar")
                        'body = cp.Html.Form(body)
                        '
                        ' 
                        '
                End Select
                man.title = "Template Importer"
                man.body = body
                body = man.getHtml(cp)
                cp.Doc.AddHeadStyle(man.styleSheet)
            Catch ex As Exception
                HandleClassError(cp, ex, "GetContent", "trap, Hint=[" & Hint & "]")
                body = "<p>There was an unknown problem with the template importer. The message returned was [" & ex.Message & "]."
            End Try
            Return body
        End Function
        '
        '=================================================================================
        '   Handle errors from this class
        '=================================================================================
        '
        Private Sub HandleClassError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByRef MethodName As String, ByVal description As String)
            '
            Call cp.Site.ErrorReport(ex, "Error in templateImport." & MethodName & ", " & description)
            '
        End Sub
        '
        '
        '
        Private Function CreateNewTemplate(ByVal cp As CPBaseClass, ByRef TemplateName As String) As Integer
            Try
                Dim IsNameOK As Boolean
                Dim TemplateID As Integer
                Dim cs As CPCSBaseClass = cp.CSNew
                '
                If TemplateName = "" Then
                    Call cp.UserError.Add("you must select a unique, non-blank name for your template to beging. Select a different name, or use the 'Open' tool to edit the existing template.")
                Else
                    cs.Open("Page Templates", "name=" & cp.Db.EncodeSQLText(TemplateName), , , "id")
                    IsNameOK = Not cs.OK
                    cs.Close()
                    '
                    If Not IsNameOK Then
                        Call cp.UserError.Add("The name [" & TemplateName & "] is already used by another template. Select a different name, or use the 'Open' tool to edit the existing template.")
                        TemplateName = ""
                    Else
                        Call cs.Insert("Page Templates")
                        If cs.OK Then
                            Call cs.SetField("name", TemplateName)
                            TemplateID = cs.GetInteger("ID")
                            Call cp.Visit.SetProperty(VisitPropertyTemplateID, TemplateID)
                        End If
                        Call cs.Close()
                    End If
                End If
                CreateNewTemplate = TemplateID

            Catch ex As Exception
                HandleClassError(cp, ex, "CreateNewTemplate", "trap")
            End Try
        End Function
        '
        '
        '
        Private Sub ImportTemplate(ByVal cp As CPBaseClass, ByRef TemplateID As Integer, ByRef Link As String, ByRef ManageStyles As Boolean, ByRef BuildVersion As String, ByRef templateFilename As String, ByVal importMethod As String)
            Try
                '
                Dim Doc As String
                Dim BasePath As String
                Dim SourceHost As String
                Dim Pos As Integer
                Dim tempPath As String = cp.Site.PhysicalFilePath & "temp\"
                Dim webClient As System.Net.WebClient = New System.Net.WebClient()
                '
                Call cp.File.Save(tempPath & "tempfile.txt", "delete this file")
                Call cp.File.Delete(tempPath & "tempfile.txt")
                '
                BasePath = GetBasePath(cp, Link)
                SourceHost = ""
                If Link = "" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object cp.UserError.Add. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cp.UserError.Add("The source link can not be blank")
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
                        cp.UserError.Add("The host name of the link could not be determined [" & Link & "]")
                    Else
                        appendDebugLog(cp, "downloading template [" & Link & "]")
                        Doc = webClient.DownloadString(Link)
                        If Doc = "" Then
                            cp.UserError.Add("The request for [" & Link & "] returned empty")
                        Else
                            Doc = ImportTemplate_Convert(cp, Doc, BasePath, SourceHost, TemplateID, ManageStyles, BuildVersion, tempPath, importMethod, templateFilename)
                            cp.Cache.Clear("page templates")
                        End If
                    End If
                End If
            Catch ex As Exception
                HandleClassError(cp, ex, "ImportTemplate", "trap")
            End Try
        End Sub
        '
        '   parse the page associated with the given Doc, return any errors
        '       fills the Doc with all tags found
        '
        Private Function ImportTemplate_Convert(ByVal cp As CPBaseClass, ByRef Doc As String, ByRef BasePath As String, ByRef SourceHost As String, ByRef TemplateID As Integer, ByRef ManageStyles As Boolean, ByRef BuildVersion As String, ByVal tempPath As String, ByVal importMethod As String, ByRef templateFilename As String) As String
            Dim returnVal As String = ""
            Dim hint As String = "enter"
            Try
                'Dim GetTagInnerHTML As Object
                'Dim IsLinkToThisHost As Object
                Dim localRootRelativeLink As String
                Dim ImportBasePath As String
                Dim StyleTag As String = ""
                Dim BodyTag As String = ""
                Dim Pos As Integer
                Dim LinkType As String
                Dim StyleSheet As String
                Dim Link As String
                Dim Position As Integer
                Dim ElementCount As Integer
                Dim TagCount As Integer
                Dim TagName As String
                Dim TagContent As String
                Dim TagHTTPEquiv As String
                Dim ElementPointer As Integer
                Dim FontsUsedCount As Integer
                Dim ElementText As String
                Dim remoteRootRelativeLink As String
                Dim TagDone As Boolean
                Dim DocConverted As String
                Dim IsInHead As Boolean
                Dim OtherHeadTags As String = ""
                Dim virtualfilename As String
                Dim inlineStyle As String
                Dim inlineStyleNew As String
                '
                'Dim Output As FastString.FastStringClass = New FastString.FastStringClass
                Dim out As String = ""
                Dim kmaParse As kmaHTML.ParseClass = New kmaHTML.ParseClass
                Dim webClient As System.Net.WebClient = New System.Net.WebClient()
                Dim cs As CPCSBaseClass = cp.CSNew
                '
                Doc = Replace(Doc, "<span><span>", "<span class=""fpo""><span class=""fpo"">", 1, 99, CompareMethod.Text)
                Call kmaParse.Load(Doc)
                ElementPointer = 0
                FontsUsedCount = 0
                ElementCount = kmaParse.ElementCount
                hint &= ",elementcount=" & ElementCount
                '
                Do While ElementPointer < ElementCount
                    ElementText = kmaParse.Text(ElementPointer)
                    If kmaParse.IsTag(ElementPointer) Then
                        TagCount = TagCount + 1
                        TagName = kmaParse.TagName(ElementPointer)
                        hint &= ",tag=" & TagName
                        Select Case UCase(TagName)
                            Case "HEAD"
                                IsInHead = True
                            Case "/HEAD"
                                IsInHead = False
                            Case "FORM"
                                Link = kmaParse.ElementAttribute(ElementPointer, "action")
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                End If
                            Case "TD"
                                Link = kmaParse.ElementAttribute(ElementPointer, "Background")
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                End If
                            Case "BODY"
                                Link = kmaParse.ElementAttribute(ElementPointer, "Background")
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                    'remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                    ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                    ''If virtualfilename.Substring(0, 1) = "\" Then
                                    ''    virtualfilename = virtualfilename.Substring(1)
                                    ''End If
                                    'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                End If
                                BodyTag = ElementText
                            Case "BASE"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    BasePath = GetBasePath(cp, Link)
                                    ElementText = ""
                                End If
                            Case "A"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                    'End If
                                End If
                            Case "META"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                TagHTTPEquiv = kmaParse.ElementAttribute(ElementPointer, "HTTPEquiv")
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                TagContent = kmaParse.ElementAttribute(ElementPointer, "Content")
                                If UCase(CStr(TagHTTPEquiv = "REFRESH")) And (TagContent <> "") Then
                                    Link = UCase(TagContent)
                                    Position = InStr(1, Link, "URL")
                                    If Position <> 0 Then
                                        Position = Position + 3
                                        Do While Mid(Link, Position, 1) = " "
                                            Position = Position + 1
                                            'System.Windows.Forms.Application.DoEvents()
                                        Loop
                                        If Mid(Link, Position, 1) = "=" Then
                                            Position = Position + 1
                                            Do While Mid(Link, Position, 1) = " "
                                                Position = Position + 1
                                                'System.Windows.Forms.Application.DoEvents()
                                            Loop
                                            Link = Trim(Mid(TagContent, Position))
                                            'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            If IsLinkToThisHost(cp, SourceHost, Link) Then
                                                remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                                localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                                ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                            End If
                                        End If
                                    End If
                                End If
                                If IsInHead Then
                                    OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                                End If

                            Case "AREA"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                End If
                            Case "IMG"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                    'remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                    ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                    ''If virtualfilename.Substring(0, 1) = "\" Then
                                    ''    virtualfilename = virtualfilename.Substring(1)
                                    ''End If
                                    'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                End If
                            Case "EMBED"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                    '    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    '    virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                    '    'virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                    '    'If virtualfilename.Substring(0, 1) = "\" Then
                                    '    '    virtualfilename = virtualfilename.Substring(1)
                                    '    'End If
                                    '    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    '    ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                End If
                            Case "FRAMESET"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                    'remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                    ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                    ''If virtualfilename.Substring(0, 1) = "\" Then
                                    ''    virtualfilename = virtualfilename.Substring(1)
                                    ''End If
                                    'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                End If
                            Case "FRAME"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                    'remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                    ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                    ''If virtualfilename.Substring(0, 1) = "\" Then
                                    ''    virtualfilename = virtualfilename.Substring(1)
                                    ''End If
                                    'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                End If
                            Case "LINK"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "HREF")
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                LinkType = kmaParse.ElementAttribute(ElementPointer, "TYPE")
                                hint &= ",link=" & Link & " type=" & LinkType
                                If (IsLinkToThisHost(cp, SourceHost, Link)) And (LCase(LinkType) = "text/css") Then
                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    hint &= ",100"
                                    If ManageStyles Then
                                        '
                                        ' css styles - convert to relative styles and comment out element
                                        '
                                        hint &= ",110"
                                        appendDebugLog(cp, "downloading style [" & SourceHost & remoteRootRelativeLink & "]")
                                        StyleSheet = webClient.DownloadString(SourceHost & remoteRootRelativeLink)
                                        ImportBasePath = remoteRootRelativeLink
                                        Pos = InStrRev(ImportBasePath, "/")
                                        If Pos > 0 Then
                                            ImportBasePath = Mid(ImportBasePath, 1, Pos)
                                        End If
                                        StyleSheet = ConvertStyles_HandleFileReferences(cp, StyleSheet, ImportBasePath, SourceHost, "", TemplateID, tempPath)
                                        'UPGRADE_WARNING: Couldn't resolve default property of object cp.Utils.DecodeUrl(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        Call SaveSharedStyle(cp, TemplateID, cp.Utils.DecodeUrl(remoteRootRelativeLink), StyleSheet)
                                        ElementText = "<!-- " & ElementText & " -->"
                                    Else
                                        '
                                        ' inmanaged css or other files, download them and convert link
                                        '
                                        hint &= ",120"
                                        Call ConvertStyleLink_HandleImports(cp, Link, BasePath, SourceHost, "", TemplateID, tempPath)
                                        '
                                        localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                        virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                        Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                        ElementText = Replace(ElementText, Link, localRootRelativeLink)
                                        'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                        ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                        ''If virtualfilename.Substring(0, 1) = "\" Then
                                        ''    virtualfilename = virtualfilename.Substring(1)
                                        ''End If
                                        'hint &= ",130"
                                        'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                        'hint &= ",140"
                                        'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                    End If
                                End If
                                hint &= ",150"
                                If IsInHead Then
                                    OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                                End If
                            Case "SCRIPT"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                                'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If IsLinkToThisHost(cp, SourceHost, Link) Then

                                    remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                    virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                    Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    ElementText = Replace(ElementText, Link, localRootRelativeLink)

                                    'remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                    'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                    ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                    ''If virtualfilename.Substring(0, 1) = "\" Then
                                    ''    virtualfilename = virtualfilename.Substring(1)
                                    ''End If
                                    'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                    'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
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
                                        'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        ElementText = ElementText & kmaParse.Text(ElementPointer)
                                        'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.IsTag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        If kmaParse.IsTag(ElementPointer) Then
                                            '
                                            ' Process a tag (should just be </SCRIPT>, but go until it is
                                            '
                                            TagCount = TagCount + 1
                                            'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.TagName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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

                                        out &= ElementText
                                        'Output.Add(ElementText)
                                        ElementPointer = ElementPointer + 1
                                    End If
                                Loop
                                If IsInHead Then
                                    StyleTag = ConvertStyles_HandleFileReferences(cp, StyleTag, BasePath, SourceHost, "", TemplateID, tempPath)
                                    OtherHeadTags = OtherHeadTags & vbCrLf & StyleTag
                                End If
                            Case "INPUT"
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If UCase(kmaParse.ElementAttribute(ElementPointer, "TYPE")) = "IMAGE" Then
                                    'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    Link = kmaParse.ElementAttribute(ElementPointer, "SRC")
                                    'UPGRADE_WARNING: Couldn't resolve default property of object IsLinkToThisHost(cp,SourceHost, Link). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    If IsLinkToThisHost(cp, SourceHost, Link) Then
                                        remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                        localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                        virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                        Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                        ElementText = Replace(ElementText, Link, localRootRelativeLink)

                                        'remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                        'virtualfilename = convertRootRelativeLinkToPathFilename(cp, remoteRootRelativeLink)
                                        ''virtualfilename = Replace(cp.Utils.DecodeUrl(RootRelativeLink), "/", "\")
                                        ''If virtualfilename.Substring(0, 1) = "\" Then
                                        ''    virtualfilename = virtualfilename.Substring(1)
                                        ''End If
                                        'Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                        'ElementText = Replace(ElementText, Link, remoteRootRelativeLink)
                                    End If
                                End If
                            Case "TITLE"
                                '
                                ' Skip to the </title> TAG
                                '
                                TagDone = False
                                Do While (Not TagDone) And (ElementPointer < ElementCount)
                                    '
                                    ' Process the next segment
                                    '
                                    ElementText = kmaParse.Text(ElementPointer)
                                    If kmaParse.IsTag(ElementPointer) Then
                                        TagCount = TagCount + 1
                                        TagDone = (kmaParse.TagName(ElementPointer) = "/" & TagName)
                                    End If
                                    If Not TagDone Then
                                        ElementPointer = ElementPointer + 1
                                    End If
                                Loop
                                ElementText = ""
                            Case "SPAN"
                                '
                                ' wysiwyg editor auto deletes empty spans, all a class if there is not one
                                '
                                'UPGRADE_WARNING: Couldn't resolve default property of object kmaParse.ElementAttribute. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If LCase(kmaParse.ElementAttribute(ElementPointer, "class")) = "" Then
                                    ElementText = Replace(ElementText, ">", " class=""emptySpanPlaceHolder"">")
                                End If
                            Case Else
                                If IsInHead Then
                                    OtherHeadTags = OtherHeadTags & vbCrLf & ElementText
                                End If
                        End Select
                        inlineStyle = kmaParse.ElementAttribute(ElementPointer, "style")
                        If inlineStyle <> "" Then
                            inlineStyleNew = ConvertStyles_HandleFileReferences(cp, inlineStyle, BasePath, SourceHost, "", TemplateID, tempPath)
                            If inlineStyle = inlineStyleNew Then
                                Call appendDebugLog(cp, "inline style [" & inlineStyle & "] not changed")
                            Else
                                ElementText = ElementText.Replace(inlineStyle, inlineStyleNew)
                                Call appendDebugLog(cp, "inline style [" & inlineStyle & "] replaced with [" & inlineStyleNew & "]")
                            End If
                        End If
                    End If
                    out &= ElementText
                    'Output.Add(ElementText)
                    ElementPointer = ElementPointer + 1
                Loop
                DocConverted = out
                'DocConverted = Output.Text
                '
                cs.Open("Page Templates", "id=" & cp.Db.EncodeSQLNumber(TemplateID))
                If cs.OK Then
                    If importMethod = importMethodToRecord Then
                        Call cs.SetField("bodyhtml", GetTagInnerHTML(cp, DocConverted, "body", False))
                    ElseIf importMethod = importMethodToFile Then
                        Call cs.SetField("bodyhtml", "{% import """ & templateFilename & """ %}")
                        Call cp.File.Save(cp.Site.PhysicalWWWPath & templateFilename, DocConverted)
                    End If
                    If BuildVersion > "3.3.994" Then
                        If OtherHeadTags <> "" Then
                            If Mid(OtherHeadTags, 1, 2) = vbCrLf Then
                                OtherHeadTags = Mid(OtherHeadTags, 3)
                            End If
                        End If
                        Call cs.SetField("OtherHeadTags", OtherHeadTags)
                        Call cs.SetField("BodyTag", BodyTag)
                    End If
                End If
                Call cs.Close()
                Call cp.Cache.Clear("page templates")
                returnVal = DocConverted
                kmaParse = Nothing
            Catch ex As Exception
                Call HandleClassError(cp, ex, "ImportTemplate_Convert", "trap, hint=" & hint)
            End Try
            Return returnVal
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
        Private Function GetBasePath(ByVal cp As CPBaseClass, ByRef Link As String) As String
            Dim returnVal As String = ""
            Try
                '
                '
                Dim Pos As Integer
                Dim LoopCnt As Integer
                Dim posStart As Integer
                '
                returnVal = Link
                If InStr(1, Link, "/") = 1 Then
                    '
                    '   case /images/logo-Main.jpg with to /images/
                    '
                    Pos = 1
                    Do While (Pos > 0) And LoopCnt < 100
                        returnVal = Mid(Link, 1, Pos)
                        Pos = InStr(Pos + 1, Link, "/")
                    Loop
                ElseIf InStr(1, Link, "://") <> 0 Then
                    '
                    '   case http://gcm.brandeveolve.com/images/logo-Main.jpg with any BasePath  to /images/logo-Main.jpg
                    '
                    returnVal = "/"
                    Pos = InStr(1, Link, "://")
                    If Pos > 0 Then
                        Pos = InStr(Pos + 3, Link, "/")
                        If Pos > 0 Then
                            posStart = Pos
                            Do While (Pos > 0) And LoopCnt < 100
                                returnVal = Mid(Link, posStart, Pos)
                                Pos = InStr(Pos + 1, Link, "/")
                            Loop
                            Pos = InStrRev(returnVal, "/")
                            If Pos < Len(returnVal) Then
                                returnVal = Mid(returnVal, 1, Pos)
                            End If
                        End If
                    End If
                Else
                    '
                    '   unknown case
                    '
                    returnVal = "/"
                End If
            Catch ex As Exception
                HandleClassError(cp, ex, "GetBasePath", "trap")
            End Try
            Return returnVal
        End Function
        '
        '
        '
        Private Sub SaveSharedStyle(ByVal cp As CPBaseClass, ByRef TemplateID As Integer, ByRef Name As String, ByRef Styles As String)
            Try
                Dim SharedStylesID As Integer
                Dim cs As CPCSBaseClass = cp.CSNew
                '
                cs.Open("Shared Styles", "name=" & cp.Db.EncodeSQLText(Name))
                'UPGRADE_WARNING: Couldn't resolve default property of object Main.IsCSOK. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not cs.OK Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Main.CloseCS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call cs.Close()
                    'UPGRADE_WARNING: Couldn't resolve default property of object Main.InsertCSRecord. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cs.Insert("Shared Styles")
                End If
                If cs.OK Then
                    SharedStylesID = cs.GetInteger("ID")
                    Call cs.SetField("name", Name)
                    Call cs.SetField("stylefilename", Styles)
                End If
                cs.Close()
                If SharedStylesID <> 0 And TemplateID <> 0 Then
                    cs.Open("Shared Styles Template Rules", "(styleid=" & SharedStylesID & ")and(TemplateID=" & TemplateID & ")")
                    If Not cs.OK Then
                        Call cs.Close()
                        cs.Insert("Shared Styles Template Rules")
                    End If
                    If cs.OK Then
                        Call cs.SetField("styleid", SharedStylesID)
                        Call cs.SetField("TemplateID", TemplateID)
                    End If
                    Call cs.Close()
                End If
            Catch ex As Exception
                HandleClassError(cp, ex, "SaveSharedStyle", "trap")
            End Try
        End Sub
        '
        '
        '
        Private Function ConvertStyles_HandleFileReferences(ByVal cp As CPBaseClass, ByRef StyleSheet As String, ByRef BasePath As String, ByRef SourceHost As String, ByRef BlockRootRelativeLinkList As String, ByRef TemplateID As Integer, ByRef tempPath As String) As String
            Dim returnVal As String = ""
            Try
                '
                '
                'Dim destPathFilename As String
                Dim PosURLStart As Integer
                Dim PosURLEnd As Integer
                Dim NameValue As String
                Dim PtrStart As Integer
                Dim PtrEnd As Integer
                Dim Line As String
                Dim Lines() As String
                Dim LineCnt As Integer
                Dim LinePtr As Integer
                Dim Ptr As Integer
                Dim PosImport As Integer
                Dim posStart As Integer
                Dim posEnd As Integer
                Dim Link As String
                Dim RootRelativeLink As String
                Dim LoopCnt As Integer
                Dim Copy As String
                Dim ImportStyles As String
                Dim URLPosStart As Integer
                Dim URLPosEnd As Integer
                Dim ImportBasePath As String
                Dim PathPos As Integer
                Dim PosImportEnd As Integer
                Dim PosLineEnd As Integer
                Dim PosSemi As Integer
                Dim virtualfilename As String
                Dim isInlineStyle As Boolean = False
                Dim remoteRootRelativeLink As String
                Dim localRootRelativeLink As String
                '
                Dim webClient As System.Net.WebClient = New System.Net.WebClient()
                '
                PosImport = 1
                Copy = StyleSheet
                '
                Do While (PosImport <> 0) And LoopCnt < 100
                    PosImport = InStr(PosImport, Copy, "@import", CompareMethod.Text)
                    If PosImport <> 0 Then
                        '
                        ' style includes an import -- convert filename and load the file
                        '
                        URLPosStart = InStr(PosImport, Copy, "url", CompareMethod.Text)
                        If URLPosStart <> 0 Then
                            posStart = InStr(URLPosStart, Copy, "(", CompareMethod.Text)
                            If posStart <> 0 Then
                                posStart = posStart + 1
                                URLPosEnd = InStr(posStart, Copy, ")", CompareMethod.Text)
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

                                    RootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)


                                    ImportBasePath = RootRelativeLink
                                    PathPos = InStrRev(ImportBasePath, "/")
                                    If PathPos > 0 Then
                                        ImportBasePath = Mid(ImportBasePath, 1, PathPos)
                                    End If
                                    '
                                    appendDebugLog(cp, "downloading style [" & SourceHost & RootRelativeLink & "]")
                                    ImportStyles = webClient.DownloadString(SourceHost & RootRelativeLink)
                                    ImportStyles = ConvertStyles_HandleFileReferences(cp, ImportStyles, ImportBasePath, SourceHost, BlockRootRelativeLinkList, TemplateID, tempPath)
                                    Call SaveSharedStyle(cp, TemplateID, cp.Utils.DecodeUrl(RootRelativeLink), ImportStyles)

                                    'RootRelativeLink = ConvertStyleLink_HandleImports(Link, BasePath, SourceHost, BlockRootRelativeLinkList)

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
                If Copy.IndexOf("{") = -1 Then
                    ' inline style - fake a single styleselector
                    isInlineStyle = True
                    Copy = "{" & Copy & "}"
                End If
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
                            PosURLStart = InStr(1, NameValue, "url(", CompareMethod.Text)
                            Dim loopPtr As Integer
                            loopPtr = 0
                            Do
                                If (PosURLStart <> 0) Then
                                    Call appendDebugLog(cp, "url found in style namevalue [" & NameValue & "]")
                                    PosURLStart = PosURLStart + 4
                                    PosURLEnd = InStr(PosURLStart, NameValue, ")")
                                    If PosURLEnd <> 0 Then
                                        Link = Mid(NameValue, PosURLStart, PosURLEnd - PosURLStart)
                                        Link = trimAll(cp, Link)
                                        If (Left(Link, 1) = """" And Right(Link, 1) = """") Or (Left(Link, 1) = "'" And Right(Link, 1) = "'") Then
                                            PosURLStart += 1
                                            PosURLEnd -= 1
                                            Link = Mid(Link, 2, Len(Link) - 2)
                                        End If
                                        '
                                        remoteRootRelativeLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                                        localRootRelativeLink = convertRemoteToLocalLink(cp, remoteRootRelativeLink)
                                        virtualfilename = convertRootRelativeLinkToPathFilename(cp, localRootRelativeLink)
                                        Call cp.File.DeleteVirtual(virtualfilename)
                                        Call GetURLToFile(cp, SourceHost & remoteRootRelativeLink, cp.Site.PhysicalWWWPath & virtualfilename, tempPath)
                                        NameValue = Mid(NameValue, 1, PosURLStart - 1) & localRootRelativeLink & Mid(NameValue, PosURLEnd)
                                        'NameValue = Replace(NameValue, Link, localRootRelativeLink, 1, -1, CompareMethod.Text)
                                    End If
                                    PosURLStart = InStr(PosURLStart, NameValue, "url(", CompareMethod.Text)
                                End If
                                loopPtr += 1
                            Loop While (PosURLStart <> 0) And loopPtr < 100
                            If PtrEnd = 0 Then
                                Line = Mid(Line, 1, PtrStart - 1) & NameValue
                            Else
                                Line = Mid(Line, 1, PtrStart - 1) & NameValue & Mid(Line, PtrEnd + 1)
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
                If isInlineStyle Then
                    'remove the fake styleselector for inline styles
                    Copy = Copy.Substring(1, Copy.Length - 2)
                End If
                '
                ' Done
                '
                returnVal = Copy
            Catch ex As Exception
                HandleClassError(cp, ex, "ConvertStyles_HandleFileReferences", "trap")
            End Try
            Return returnVal
        End Function
        '
        '
        '
        Private Function ConvertStyleLink_HandleImports(ByVal cp As CPBaseClass, ByRef Link As String, ByRef BasePath As String, ByRef SourceHost As String, ByRef BlockRootRelativeLinkList As String, ByRef TemplateID As Integer, ByRef tempPath As String) As String
            Dim returnVal As String = ""
            Try
                '
                '
                Dim Pos As Integer
                Dim ImportedStyle As String
                Dim Filename As String
                Dim ConvertedLink As String
                Dim ImportLink As String
                Dim LinkPath As String = ""
                Dim webClient As System.Net.WebClient = New System.Net.WebClient()
                '
                ConvertedLink = ConvertLinkToRootRelative(cp, Link, BasePath)
                returnVal = ConvertedLink
                If InStr(1, BlockRootRelativeLinkList, ConvertedLink, CompareMethod.Text) = 0 Then
                    ImportLink = SourceHost & ConvertedLink
                    'UPGRADE_WARNING: Couldn't resolve default property of object HTTP.GetURL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

                    appendDebugLog(cp, "downloading style [" & ImportLink & "]")
                    ImportedStyle = webClient.DownloadString(ImportLink)
                    Pos = InStrRev(ConvertedLink, "/")
                    If Pos > 0 Then
                        LinkPath = Mid(ConvertedLink, 1, Pos)
                    End If
                    'If Mid(Linkpath, 1, 1) = "/" Then
                    '    Linkpath = Mid(Linkpath, 2)
                    'End If
                    ImportedStyle = ConvertStyles_HandleFileReferences(cp, ImportedStyle, LinkPath, SourceHost, BlockRootRelativeLinkList & "," & ConvertedLink, TemplateID, tempPath)
                    'UPGRADE_WARNING: Couldn't resolve default property of object cp.Site.PhysicalWWWPath. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Filename = cp.Site.PhysicalWWWPath & Replace(ConvertedLink, "/", "\")
                    'UPGRADE_WARNING: Couldn't resolve default property of object Main.SaveFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cp.File.Save(Filename, ImportedStyle)
                End If
            Catch ex As Exception
                HandleClassError(cp, ex, "ConvertStyleLink_HandleImports", "trap")
            End Try
            Return returnVal
        End Function
        '
        '
        '
        Private Sub GetURLToFile(ByVal cp As CPBaseClass, ByRef SrcLink As String, ByRef dstPathFilename As String, ByRef tempPath As String)
            Try
                Dim tempPathFilename As String
                Dim Pos As Integer = 0
                'Dim ErrDescription As String
                Dim webclient As System.Net.WebClient = New System.Net.WebClient
                Dim dstPathFilenameFixed As String
                '
                If InStr(1, vbCrLf & filesFetchedList & vbCrLf, vbCrLf & SrcLink & vbCrLf, CompareMethod.Text) <> 0 Then
                    '
                    ' already downloaded
                    '
                    filesFetchedList = filesFetchedList
                ElseIf (InStr(SrcLink, "/cclib/", CompareMethod.Text) <> 0) Then
                    '
                    ' reserved path
                    '
                    Call cp.UserError.Add("The resource " & SrcLink & " could not be imported because the path /cclib is reserved.")
                ElseIf (InStr(SrcLink, "/admin/", CompareMethod.Text) <> 0) Then
                    '
                    ' reserved path
                    '
                    Call cp.UserError.Add("The resource " & SrcLink & " could not be imported because the path /admin is reserved.")
                ElseIf (InStr(SrcLink, "/" & cp.Site.PageDefault, CompareMethod.Text) <> 0) Then
                    '
                    ' reserved path
                    '
                    Call cp.UserError.Add("The resource " & SrcLink & " could not be imported because the page " & "/" & cp.Site.PageDefault & " is reserved.")
                Else
                    filesFetchedList = filesFetchedList & vbCrLf & SrcLink
                    dstPathFilenameFixed = Replace(dstPathFilename, "/", "\")
                    Pos = InStrRev(dstPathFilenameFixed, "\")
                    If Pos > 0 Then
                        tempPathFilename = tempPath & Mid(dstPathFilenameFixed, Pos + 1)
                        appendDebugLog(cp, "downloading [" & SrcLink & "] to [" & tempPathFilename & "]")
                        Try
                            webclient.DownloadFile(SrcLink, tempPathFilename)
                        Catch ex As Exception
                            Call cp.UserError.Add("Error loading file [" & SrcLink & "], " & ex.Message)
                            Exit Sub
                        End Try
                        If cp.File.fileExists(tempPathFilename) Then
                            Try
                                Call cp.File.Delete(dstPathFilenameFixed)
                                Call cp.File.Save(dstPathFilenameFixed, "createpath")
                                Call cp.File.Delete(dstPathFilenameFixed)
                                Call System.IO.File.Copy(tempPathFilename, dstPathFilenameFixed)
                                Call cp.File.Delete(tempPathFilename)
                            Catch ex As Exception
                                Call cp.UserError.Add("Error loading file [" & SrcLink & "], " & ex.Message)
                                Exit Sub
                            End Try
                        End If
                    End If
                End If

            Catch ex As Exception
                HandleClassError(cp, ex, "GetURLToFile", "trap")

            End Try
        End Sub
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
        Private Function ConvertLinkToRootRelative(ByVal cp As CPBaseClass, ByRef Link As String, ByRef BasePath As String) As String
            Dim returnVal As String = ""
            Try
                '
                '
                Dim Pos As Integer
                Dim workingbase As String
                Dim WorkingLink As String
                '
                workingbase = BasePath
                WorkingLink = Link
                returnVal = WorkingLink
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
                            returnVal = Mid(WorkingLink, Pos)
                        Else
                            '
                            ' This is just the domain name, RootRelative is the root
                            '
                            returnVal = "/"
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

                    returnVal = workingbase & WorkingLink
                End If
            Catch ex As Exception
                HandleClassError(cp, ex, "ConvertLinkToRootRelative", "trap")
            End Try
            Return returnVal
        End Function
        '
        '
        '
        Private Function trimAll(ByVal cp As CPBaseClass, ByRef Source As String) As String
            Dim returnVal As String = ""
            Try
                '
                Dim Ptr As Integer
                Dim test As Integer
                '
                returnVal = Source
                For Ptr = 1 To Len(returnVal)
                    test = Asc(Mid(returnVal, Ptr, 1))
                    If (test > 32) And (test <= 128) Then
                        Exit For
                    End If

                Next
                returnVal = Mid(returnVal, Ptr)
                For Ptr = Len(returnVal) To 1 Step -1
                    test = Asc(Mid(returnVal, Ptr, 1))
                    If (test > 32) And (test <= 128) Then
                        Exit For
                    End If
                Next
                returnVal = Mid(returnVal, 1, Ptr)
            Catch ex As Exception
                HandleClassError(cp, ex, "trimAll", "trap")
            End Try
            Return returnVal
        End Function
        '
        ' returns true of the link is a valid link on the source host
        '
        Public Function IsLinkToThisHost(ByVal cp As CPBaseClass, ByRef Host As String, ByRef Link As String) As Boolean
            Try
                '
                Dim LinkHost As String
                Dim Pos As Integer
                '
                IsLinkToThisHost = False
                If Trim(Link) = "" Then
                    '
                    ' Blank is not a link
                    '
                    IsLinkToThisHost = False
                ElseIf InStr(1, Link, "://") <> 0 Then
                    '
                    ' includes protocol, may be link to another site
                    '
                    LinkHost = LCase(Link)
                    Pos = 1
                    Pos = InStr(Pos, LinkHost, "://")
                    If Pos > 0 Then
                        Pos = InStr(Pos + 3, LinkHost, "/")
                        If Pos > 0 Then
                            LinkHost = Mid(LinkHost, 1, Pos - 1)
                        End If
                        IsLinkToThisHost = (LCase(Host) = LinkHost)
                        If Not IsLinkToThisHost Then
                            '
                            ' try combinations including/excluding www.
                            '
                            If InStr(1, LinkHost, "www.", CompareMethod.Text) <> 0 Then
                                '
                                ' remove it
                                '
                                LinkHost = Replace(LinkHost, "www.", "", 1, -1, CompareMethod.Text)
                                IsLinkToThisHost = (LCase(Host) = LinkHost)
                            Else
                                '
                                ' add it
                                '
                                LinkHost = Replace(LinkHost, "://", "://www.", 1, -1, CompareMethod.Text)
                                IsLinkToThisHost = (LCase(Host) = LinkHost)
                            End If
                        End If
                    End If
                ElseIf InStr(1, Link, "#") = 1 Then
                    '
                    ' Is a bookmark, not a link
                    '
                    IsLinkToThisHost = False
                Else
                    '
                    ' all others are links on the source
                    '
                    IsLinkToThisHost = True
                End If
                If Not IsLinkToThisHost Then
                    Link = Link
                End If
            Catch ex As Exception
                HandleClassError(cp, ex, "IsLinkToThisHost", "trap")
            End Try
        End Function
        '
        '========================================================================================================
        '
        ' Finds all tags matching the input, and concatinates them into the output
        ' does NOT account for nested tags, use for body, script, style
        '
        ' ReturnAll - if true, it returns all the occurances, back-to-back
        '
        '========================================================================================================
        '
        Public Function GetTagInnerHTML(ByVal cp As CPBaseClass, ByRef PageSource As String, ByRef Tag As String, ByRef ReturnAll As Boolean) As String
            Dim returnVal As String = ""
            Try
                '
                Dim TagStart As Integer
                Dim TagEnd As Integer
                Dim LoopCnt As Integer
                'Dim WB As String
                Dim Pos As Integer
                'Dim posEnd As Integer
                Dim CommentPos As Integer
                Dim ScriptPos As Integer
                '
                Pos = 1
                Do While (Pos > 0) And (LoopCnt < 100)
                    TagStart = InStr(Pos, PageSource, "<" & Tag, CompareMethod.Text)
                    If TagStart = 0 Then
                        Pos = 0
                    Else
                        '
                        ' tag found, skip any comments that start between current position and the tag
                        '
                        CommentPos = InStr(Pos, PageSource, "<!--")
                        If (CommentPos <> 0) And (CommentPos < TagStart) Then
                            '
                            ' skip comment and start again
                            '
                            Pos = InStr(CommentPos, PageSource, "-->")
                        Else
                            ScriptPos = InStr(Pos, PageSource, "<script")
                            If (ScriptPos <> 0) And (ScriptPos < TagStart) Then
                                '
                                ' skip comment and start again
                                '
                                Pos = InStr(ScriptPos, PageSource, "</script")
                            Else
                                '
                                ' Get the tags innerHTML
                                '
                                TagStart = InStr(TagStart, PageSource, ">", CompareMethod.Text)
                                Pos = TagStart
                                If TagStart <> 0 Then
                                    TagStart = TagStart + 1
                                    TagEnd = InStr(TagStart, PageSource, "</" & Tag, CompareMethod.Text)
                                    If TagEnd <> 0 Then
                                        returnVal = returnVal & Mid(PageSource, TagStart, TagEnd - TagStart)
                                    End If
                                End If
                            End If
                        End If
                        LoopCnt = LoopCnt + 1
                        If ReturnAll Then
                            TagStart = InStr(TagEnd, PageSource, "<" & Tag, CompareMethod.Text)
                        Else
                            TagStart = 0
                        End If
                    End If
                Loop
            Catch ex As Exception
                HandleClassError(cp, ex, "GetTagInnerHTML", "trap")
            End Try
            Return returnVal
        End Function
        '
        '
        '
        Private Sub appendDebugLog(ByVal cp As CPBaseClass, ByVal message As String)
            If debugging Then
                Call cp.File.AppendVirtual("templateImportDebug.log", vbCrLf & Now & vbTab & message)
            End If
        End Sub
        '
        '
        '
        Private Function convertRootRelativeLinkToPathFilename(ByVal cp As CPBaseClass, ByVal src As String) As String
            Dim working As String = src
            '
            working = Replace(cp.Utils.DecodeUrl(working), "/", "\")
            If working.Substring(0, 1) = "\" Then
                working = working.Substring(1)
            End If
            Return working
            '
        End Function
        '
        '
        '
        Private Function convertRemoteToLocalLink(ByVal cp As CPBaseClass, ByVal remoteRootRelativeLink As String) As String
            Dim working As String = remoteRootRelativeLink
            Dim workingExt As String
            Dim pPos As Integer
            Dim qPos As Integer
            Dim pos As Integer
            Dim workingLeft As String
            Dim workingRight As String
            '
            pos = -1
            workingExt = ""
            qPos = working.IndexOf("?")
            pPos = working.IndexOf("#")
            If (qPos >= 0) And (pPos >= 0) Then
                If qPos > pPos Then
                    pos = pPos
                Else
                    pos = qPos
                End If
            ElseIf (pPos >= 0) Then
                pos = pPos
            ElseIf (qPos >= 0) Then
                pos = qPos
            End If
            If pos >= 0 Then
                workingRight = working.Substring(pos)
                workingLeft = working.Substring(0, pos)
                workingExt = ""
                pos = workingLeft.LastIndexOf(".")
                If pos >= 0 Then
                    workingExt = workingLeft.Substring(pos)
                    workingLeft = workingLeft.Substring(0, pos)
                End If
                workingRight = workingRight.Replace("?", "-")
                workingRight = workingRight.Replace("#", "-")
                working = workingLeft & "-" & workingRight & workingExt
            End If
            '
            Return working
            '
        End Function
        '
        '
        '
    End Class
End Namespace
