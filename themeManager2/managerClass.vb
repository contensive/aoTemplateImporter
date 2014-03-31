
Imports Contensive.BaseClasses

Namespace Contensive.addons.themeManager
    '
    Public Class managerClass
        Inherits AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            '
            Try
                Dim body As String = ""
                Dim managerMacros As New managerMacrosClass
                Dim ManagerQuickImport As New ManagerQuickImportClass
                Dim rqs As String = CP.Doc.RefreshQueryString
                Call CP.Doc.AddHeadJavascript("var themeManagerFrameRqs='" & CP.Utils.EncodeJavascript(rqs) & "';")
                Dim home As New adminFramework.formSimpleClass
                '
                '
                Dim manager As New adminFramework.pageWithNavClass
                Dim rightNow As Date = getRightNow(CP)
                Dim srcFormId As Integer = CP.Doc.GetInteger(rnSrcFormId)
                Dim dstFormId As Integer = CP.Doc.GetInteger(rnDstFormId)
                '
                ' process form
                '
                If (srcFormId <> 0) Then
                    Select Case srcFormId
                        Case formIdMacroMin To formIdMacroMax
                            '
                            '
                            '
                            dstFormId = managerMacros.processForm(CP, srcFormId, rqs, rightNow)
                        Case formIdToolsMin To formIdToolsMax
                            '
                            '
                            '
                            dstFormId = ManagerQuickImport.processForm(CP, srcFormId, rqs, rightNow)
                    End Select
                End If
                '
                ' get form
                '
                manager.navCaption = "Macros"
                manager.navLink = "?" & CP.Utils.ModifyQueryString(rqs, rnDstFormId, formIdMacroList)
                '
                manager.addNav()
                manager.navCaption = "Quick Import"
                manager.navLink = "?" & CP.Utils.ModifyQueryString(rqs, rnDstFormId, formIdToolsQuickImport)
                '
                Select Case dstFormId
                    Case formIdToolsQuickImport
                        '
                        '
                        '
                        manager.setActiveNav("Quick Import")
                        body = ManagerQuickImport.getForm(CP, dstFormId, rqs, rightNow)
                        body = CP.Html.div(body, , , "themeManagerQuickImport")
                        manager.body = body
                    Case formIdMacroList
                        '
                        ' default is account list
                        '
                        manager.setActiveNav("Macros")
                        body = managerMacros.getForm(CP, dstFormId, rqs, rightNow)
                        body = CP.Html.div(body, , , "themeManagerMacros")
                        manager.body = body
                    Case Else
                        '
                        ' home/splash
                        '
                        home.title = "Theme Manager"
                        home.description = "<p>Use this tool to import and manage themes, including page templates, layouts, copy records, javascript and css</p>"
                        manager.body = home.getHtml(CP)
                End Select
                '
                'Assemble
                '
                manager.title = "Theme Manager"
                CP.Doc.AddHeadStyle(manager.styleSheet)
                returnHtml = manager.getHtml(CP)
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in aoManagerTemplate.adminClass.execute")
            End Try
            Return returnHtml
        End Function

    End Class
End Namespace
