Imports Contensive.BaseClasses
Imports adminFramework

Namespace Contensive.addons.themeManager
    '
    Public Class managerSampleADetailListClass
        '
        '
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String) As Integer
            Dim nextFormId As Integer = srcFormId
            Dim cs As CPCSBaseClass = cp.CSNew
            Dim body As CPBlockBaseClass = cp.BlockNew
            Dim button As String = cp.Doc.GetProperty(rnButton)
            Dim userId As Integer = cp.Doc.GetProperty(rnUserId)
            Try
                '
                ' always process because srcform was provided - return here after processing
                '
                If button = buttonSave Or button = buttonOK Then
                    '
                End If
                If button = buttonCancel Or button = buttonOK Then
                    nextFormId = formIdMacroList
                Else
                    nextFormId = srcFormId
                End If
            Catch ex As Exception
                errorReport(ex, cp, "processForm")
            End Try
            Return nextFormId
        End Function
        '
        '
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As String
            Dim returnHtml As String = ""
            Try
                Dim report As reportListClass
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim usertId As Integer = cp.Doc.GetProperty(rnUserId)
                Dim rowPtr As Integer = 1
                '
                '
                '
                report = New reportListClass(cp)
                report.title = "Visits (managerSampleADetailListClass)"
                report.description = "All visits by this user."
                '
                report.columnCaption = "Row"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Visit"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                report.addColumn()
                report.columnCaption = "Date"
                report.columnCaptionClass = ""
                report.columnCellClass = ""
                '
                cs.Open("visits", "(memberid=" & usertId & ")")
                Do While (cs.OK)
                    report.addRow()
                    report.setCell(rowPtr)
                    report.setCell(cs.GetInteger("id"))
                    report.setCell(cs.GetDate("dateAdded"))
                    rowPtr += 1
                    cs.GoNext()
                Loop
                returnHtml = report.getHtml(cp)
            Catch ex As Exception
                errorReport(ex, cp, "getForm")
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.adminAccountPurchaseHistoryClass." & method)
        End Sub
    End Class
End Namespace
