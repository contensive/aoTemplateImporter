
Imports Contensive.BaseClasses
Imports adminFramework

Namespace Contensive.addons.themeManager
    '
    Public Class ManagerQuickImportClass
        '
        '
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String) As Integer
            '
            Dim nextFormId As Integer = 0
            Try
                '
                ' process ajax buttons and return to list
                '
                nextFormId = formIdToolsList
            Catch ex As Exception
                '
                '
                '
                errorReport(ex, cp, "processForm")
            End Try
            Return nextFormId
        End Function
        '
        '
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As String
            Dim returnHtml As String = ""
            '
            Try
                Dim block As CPBlockBaseClass = cp.BlockNew()
                Dim body As CPBlockBaseClass = cp.BlockNew()
                Dim row As CPBlockBaseClass = cp.BlockNew()
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim rowList As String = ""
                Dim sql As String = ""
                Dim rowPtr As Integer = 0
                Dim accountLink As String = ""
                Dim qs As String = ""
                Dim orgListAllowActiveOnly As Boolean
                Dim report As reportListClass
                Dim editLink As String = ""
                Dim adminUrl As String = ""
                Dim sqlNow = cp.Db.EncodeSQLDate(rightNow)
                Dim dateExpires As Date = Date.MinValue
                '
                orgListAllowActiveOnly = cp.Utils.EncodeBoolean(cp.Visit.GetProperty("memberSmart orgListAllowActiveOnly", "1"))
                '
                '
                '
                report = New reportListClass(cp)
                report.title = "Quick Import"
                If orgListAllowActiveOnly Then
                    report.description = "All active organization memberships."
                Else
                    report.description = "All organizations."
                End If
                '
                report.addColumn()
                report.columnCaption = "&nbsp;"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Name"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "City"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "State"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                cs.OpenSQL("organizations")
                adminUrl = cp.Site.GetProperty("adminUrl") & "?cid=" & cp.Content.GetID("organizations") & "&af=4&id="
                Do While cs.OK()
                    editLink = "<a href=""" & adminUrl & cs.GetInteger("id").ToString & """>" & cs.GetText("name") & "</a>"
                    '
                    report.addRow()
                    report.setCell(editLink)
                    report.setCell(cs.GetText("name"))
                    report.setCell(cs.GetText("city"))
                    report.setCell(cs.GetText("state"))
                    cs.GoNext()
                Loop
                cs.Close()
                '
                report.htmlLeftOfTable = "" _
                    & cr & "<div>filters</div>" _
                    & cr & cp.Html.CheckBox("orgListAllowActiveOnly", orgListAllowActiveOnly, , "orgListAllowActiveOnly") & "Active&nbsp;Only" _
                    & ""
                cp.Doc.AddHeadJavascript("" _
                    & cr & "jQuery(document).ready(function(){" _
                    & cr2 & "jQuery('#abAccountAddButton').click(function(){" _
                    & cr2 & "window.location='" & cp.Site.GetProperty("adminUrl") & "?af=4&id=0&cid=" & cp.Content.GetID("organizations") & "';" _
                    & cr2 & "return false;" _
                    & cr2 & "});" _
                    & cr & "})" _
                    & "")
                returnHtml = report.getHtml(cp)
            Catch ex As Exception
                '
                '
                '
                errorReport(ex, cp, "getForm")
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.adminOrgListClass." & method)
        End Sub
    End Class
End Namespace
