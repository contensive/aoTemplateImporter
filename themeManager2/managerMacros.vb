
Imports Contensive.BaseClasses

Namespace Contensive.addons.themeManager
    '
    Public Class managerMacrosClass
        '
        '
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As Integer
            '
            Dim nextFormId As Integer = srcFormId
            Dim managerSampleADetails As New managerSampleADetailsClass
            Dim managerSampleAContentList As New managerSampleADetailListClass
            Dim managerSampleAList As New managerMacroListClass
            '
            Try
                '
                ' process form
                '
                If (srcFormId <> 0) Then
                    Select Case srcFormId
                        Case formIdMacroList
                            '
                            '
                            '
                            nextFormId = managerSampleAList.processForm(cp, srcFormId, rqs, rightNow)
                        Case formIdMacroDetailList
                            '
                            '
                            '
                            nextFormId = managerSampleAContentList.processForm(cp, srcFormId, rqs, rightNow)
                        Case formIdMacroDetails
                            '
                            ' account details
                            '
                            nextFormId = managerSampleADetails.processForm(cp, srcFormId, rqs, rightNow)
                    End Select
                End If
            Catch ex As Exception
                errorReport(ex, cp, "processForm")
            End Try
            Return nextFormId
        End Function
        '
        '
        '
        Friend Function getForm(ByVal CP As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As Object
            Dim content As String = ""
            Dim body As String = ""
            Dim button As String = CP.Doc.GetProperty(rnButton)
            Dim managerSampleADetails As New managerSampleADetailsClass
            Dim managerSampleADetailList As New managerSampleADetailListClass
            Dim managerMacroList As New managerMacroListClass
            Dim userId As Integer
            Dim rqsTabs As String
            Dim tabList As String = ""
            Dim tabbedContent As New adminFramework.contentWithTabsClass
            Dim macroBody As String = ""
            '
            Try
                If (dstFormId = formIdMacroList) Or (dstFormId = 0) Then
                    '
                    ' account list form
                    '
                    body = managerMacroList.getForm(CP, dstFormId, rqs, rightNow)
                Else
                    '
                    userId = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnUserId))
                    rqsTabs = rqs
                    rqsTabs = CP.Utils.ModifyQueryString(rqsTabs, rnUserId, userId)
                    '
                    tabbedContent.addTab()
                    tabbedContent.tabCaption = "Details"
                    tabbedContent.tabLink = "?" & CP.Utils.ModifyQueryString(rqsTabs, rnDstFormId, formIdMacroDetails)
                    '
                    tabbedContent.addTab()
                    tabbedContent.tabCaption = "Detail List"
                    tabbedContent.tabLink = "?" & CP.Utils.ModifyQueryString(rqsTabs, rnDstFormId, formIdMacroDetailList)
                    '
                    ' get form
                    '
                    rqs = CP.Utils.ModifyQueryString(rqs, rnUserId, userId)
                    rqs = CP.Utils.ModifyQueryString(rqs, rnDstFormId, dstFormId)
                    '
                    Select Case dstFormId
                        Case formIdMacroDetails
                            '
                            ' Account Details
                            '
                            tabbedContent.setActiveTab("Details")
                            macroBody = managerSampleADetails.getForm(CP, dstFormId, rqs, rightNow)
                        Case formIdMacroDetailList
                            '
                            '
                            '
                            tabbedContent.setActiveTab("Detail List")
                            macroBody = managerSampleADetailList.getForm(CP, dstFormId, rqs, rightNow)
                    End Select
                    macroBody = CP.Html.div(macroBody, "", "", "themeManagerMacroBody")
                    tabbedContent.body = macroBody
                    tabbedContent.title = "User: " & CP.Content.GetRecordName("people", userId)
                    body = tabbedContent.getHtml(CP)
                End If
            Catch ex As Exception
                Call errorReport(ex, CP, "getForm")
            End Try
            '
            ' return body
            '
            Return body
        End Function
        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in managerSampleAClass." & method)
        End Sub
    End Class
End Namespace
