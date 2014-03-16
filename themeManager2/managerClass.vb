
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
                Dim managerSampleC As New ManagerQuickImportClass
                Dim rqs As String = CP.Doc.RefreshQueryString
                Call CP.Doc.AddHeadJavascript("var msBaseRqs='" & rqs & "';")
                Dim adminUrl As String = CP.Site.GetProperty("adminUrl")
                Dim accountManagerUrl As String = adminUrl & "?addonGuid={B2290F18-3477-449C-BBCF-D0FD44E2B677}"
                '
                Dim manager As New adminFramework.pageWithNavClass
                Dim rightNow As Date = getRightNow(CP)
                Dim srcFormId As Integer = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnSrcFormId))
                Dim dstFormId As Integer = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnDstFormId))
                '
                ' process form
                '
                If (srcFormId <> 0) Then
                    Select Case srcFormId
                        Case formIdSampleAMin To formIdSampleAMax
                            '
                            '
                            '
                            dstFormId = managerMacros.processForm(CP, srcFormId, rqs, rightNow)
                        Case formIdToolsQuickImport
                            '
                            '
                            '
                            dstFormId = ManagerQuickImport.processForm(CP, srcFormId, rqs)
                    End Select
                End If
                '
                ' get form
                '
                manager.navCaption = "Macros"
                manager.navLink = "?" & CP.Utils.ModifyQueryString(rqs, rnDstFormId, formIdSampleAList)
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
                        body = CP.Html.div(body, , , "managerQuickImport")
                        manager.body = body
                    Case Else
                        '
                        ' default is account list
                        '
                        manager.setActiveNav("Macros")
                        body = managerMacros.getForm(CP, dstFormId, rqs, rightNow)
                        body = CP.Html.div(body, , , "managerMacros")
                        manager.body = body
                End Select
                '
                'Assemble
                '
                manager.title = "Manager Sample"
                CP.Doc.AddHeadStyle(manager.styleSheet)
                returnHtml = manager.getHtml(CP)
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in aoManagerTemplate.adminClass.execute")
            End Try
            Return returnHtml
        End Function

    End Class
End Namespace
