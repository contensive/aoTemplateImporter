
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
            Dim managerSampleAList As New managerSampleAListClass
            '
            Try
                '
                ' process form
                '
                If (srcFormId <> 0) Then
                    Select Case srcFormId
                        Case formIdSampleAList
                            '
                            '
                            '
                            nextFormId = managerSampleAList.processForm(cp, srcFormId, rqs)
                        Case formIdSampleADetailList
                            '
                            '
                            '
                            nextFormId = managerSampleAContentList.processForm(cp, srcFormId, rqs)
                        Case formIdSampleADetails
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
            Dim managerSampleAList As New managerSampleAListClass
            Dim userId As Integer
            Dim rqsTabs As String
            Dim tabList As String = ""
            Dim tabbedContent As New adminFramework.contentWithTabsClass
            '
            Try
                If (dstFormId = formIdSampleAList) Or (dstFormId = 0) Then
                    '
                    ' account list form
                    '
                    body = managerSampleAList.getForm(CP, dstFormId, rqs)
                Else
                    '
                    userId = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnUserId))
                    rqsTabs = rqs
                    rqsTabs = CP.Utils.ModifyQueryString(rqsTabs, rnUserId, userId)
                    '
                    tabbedContent.addTab()
                    tabbedContent.tabCaption = "Details"
                    tabbedContent.tabLink = "?" & CP.Utils.ModifyQueryString(rqsTabs, rnDstFormId, formIdSampleADetails)
                    '
                    tabbedContent.addTab()
                    tabbedContent.tabCaption = "Detail List"
                    tabbedContent.tabLink = "?" & CP.Utils.ModifyQueryString(rqsTabs, rnDstFormId, formIdSampleADetailList)
                    '
                    ' get form
                    '
                    rqs = CP.Utils.ModifyQueryString(rqs, rnUserId, userId)
                    rqs = CP.Utils.ModifyQueryString(rqs, rnDstFormId, dstFormId)
                    '
                    Select Case dstFormId
                        Case formIdSampleADetails
                            '
                            ' Account Details
                            '
                            tabbedContent.setActiveTab("Details")
                            tabbedContent.body = managerSampleADetails.getForm(CP, dstFormId, rqs, rightNow)
                        Case formIdSampleADetailList
                            '
                            '
                            '
                            tabbedContent.setActiveTab("Detail List")
                            tabbedContent.body = managerSampleADetailList.getForm(CP, dstFormId, rqs, rightNow)
                    End Select
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
