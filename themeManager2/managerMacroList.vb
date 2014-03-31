Imports Contensive.BaseClasses
Imports adminFramework

Namespace Contensive.addons.themeManager
    '
    Public Class managerMacroListClass
        '
        '
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String) As Integer
            '
            Dim nextFormId As Integer = 0
            Try
                Dim button = cp.Doc.GetText("button")
                If button <> "" Then
                    Select Case button
                        Case buttonAdd
                            '
                            ' add button should be handled by ajax
                            '
                    End Select
                End If
                '
                ' process ajax buttons and return to list
                '
                nextFormId = formIdMacroList
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
                report.title = "Import Macros"
                '
                report.columnCaption = "row"
                report.columnCaptionClass = afwStyles.afwTextAlignRight & " " & afwStyles.afwWidth50px
                report.columnCellClass = afwStyles.afwTextAlignRight
                '
                report.addColumn()
                report.columnCaption = "ID"
                report.columnCaptionClass = afwStyles.afwTextAlignRight & " " & afwStyles.afwWidth50px
                report.columnCellClass = afwStyles.afwTextAlignRight
                '
                report.addColumn()
                report.columnCaption = "Execute"
                report.columnCaptionClass = afwStyles.afwTextAlignCenter & " " & afwStyles.afwWidth100px
                report.columnCellClass = afwStyles.afwTextAlignCenter
                '
                report.addColumn()
                report.columnCaption = "Name"
                report.columnCaptionClass = afwStyles.afwTextAlignLeft
                report.columnCellClass = afwStyles.afwTextAlignLeft
                '
                cs.Open("Theme Import Macros", , , , , 50, 1)
                Do While cs.OK()
                    userId = cs.GetInteger("Id")
                    qs = rqs
                    qs = cp.Utils.ModifyQueryString(qs, rnDstFormId, formIdMacroDetails)
                    qs = cp.Utils.ModifyQueryString(qs, rnUserId, userId)
                    nameLink = "<a href=""?" & qs & """>" & cs.GetText("name") & "</a>"
                    '
                    report.addRow()
                    report.setCell(rowPtr + 1)
                    report.setCell(userId.ToString)
                    report.setCell("<a class=""afwButton"" id=""asdf"">execute</a>")
                    report.setCell(nameLink)
                    rowPtr += 1
                    cs.GoNext()
                Loop
                cs.Close()
                '
                val = accountListMembershipStatusId.ToString()
                report.htmlLeftOfTable = "" _
                    & cr & "<div class=""mmFilterTitle"">filters</div>" _
                    & ""
                '
                ' add button
                '
                Call report.addFormButton(buttonAdd, "button", "tmAddButton")
                adminUrl = cp.Site.GetProperty("adminUrl") _
                    & "?af=4" _
                    & "&id=0" _
                    & "&cid=" & cp.Content.GetID("theme import macros") & "" _
                    & ""
                js = "" _
                    & cr & "jQuery(document).ready(function(){" _
                    & cr2 & "jQuery('#tmAddButton').click(function(){" _
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
