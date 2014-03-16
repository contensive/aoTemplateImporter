Imports Contensive.BaseClasses
Imports adminFramework

Namespace Contensive.addons.themeManager
    '
    Public Class managerSampleAListClass
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
                nextFormId = formIdSampleAList
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
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String) As String
            Dim block As CPBlockBaseClass = cp.BlockNew()
            Dim body As CPBlockBaseClass = cp.BlockNew()
            Dim row As CPBlockBaseClass = cp.BlockNew()
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim rowList As String = ""
            Dim sql As String = ""
            Dim rowPtr As Integer = 0
            Dim nameLink As String = ""
            Dim qs As String = ""
            Dim userId As Integer
            Dim report As reportListClass
            Dim s As String = ""
            Dim adminUrl As String = ""
            Dim js As String = ""
            Dim dateExpiration As Date = #12:00:00 AM#
            Dim dateExpirationText As String = ""
            Dim accountListMembershipStatusId As Integer
            Dim val As String
            Dim rightNow As Date = Now()
            Dim rightNowDate As Date = rightNow.Date
            Dim rightNowDateSql As String = cp.Db.EncodeSQLDate(rightNowDate)
            '
            Try
                report = New reportListClass(cp)
                report.title = "User List (managerSampleAListClass)"
                '
                report.columnCaption = "row"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "ID"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Name"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "First Name"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Last Name"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Email"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Phone"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                cs.Open("people", , , , , 50, 1)
                Do While cs.OK()
                    userId = cs.GetInteger("Id")
                    qs = rqs
                    qs = cp.Utils.ModifyQueryString(qs, rnDstFormId, formIdSampleADetails)
                    qs = cp.Utils.ModifyQueryString(qs, rnUserId, userId)
                    nameLink = "<a href=""?" & qs & """>" & cs.GetText("name") & "</a>"
                    '
                    report.addRow()
                    report.setCell(rowPtr + 1)
                    report.setCell(userId.ToString)
                    report.setCell(nameLink)
                    report.setCell(cs.GetText("firstname"))
                    report.setCell(cs.GetText("lastname"))
                    report.setCell(cs.GetText("email"))
                    report.setCell(cs.GetText("phone"))
                    rowPtr += 1
                    cs.GoNext()
                Loop
                cs.Close()
                '
                val = accountListMembershipStatusId.ToString()
                report.htmlLeftOfTable = "" _
                    & cr & "<div class=""mmFilterTitle"">filters</div>" _
                    & ""
                adminUrl = cp.Site.GetProperty("adminUrl") _
                    & "?af=4" _
                    & "&id=0" _
                    & "&cid=" & cp.Content.GetID("people") & "" _
                    & ""
                js = "" _
                    & cr & "jQuery(document).ready(function(){" _
                    & cr2 & "jQuery('#abAddButton').click(function(){" _
                    & cr2 & "window.location='" & adminUrl & "';" _
                    & cr2 & "return false;" _
                    & cr2 & "});" _
                    & cr & "})" _
                    & ""
                cp.Doc.AddHeadJavascript(js)
                s = report.getHtml(cp)
            Catch ex As Exception
                '
                '
                '
                errorReport(ex, cp, "getForm")
            End Try
            Return s
        End Function
        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.adminListClass." & method)
        End Sub
    End Class
End Namespace
