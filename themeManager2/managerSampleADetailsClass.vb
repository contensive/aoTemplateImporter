
Imports Contensive.BaseClasses

Namespace Contensive.addons.themeManager
    '
    Public Class managerSampleADetailsClass
        '
        '
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As Integer
            Dim returnNextFormId As Integer = 0
            '
            Try
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim body As CPBlockBaseClass = cp.BlockNew
                Dim updateAccount As Boolean = False
                Dim button As String = cp.Doc.GetText(rnButton)
                Dim userId As Integer = cp.Doc.GetInteger(rnUserId)
                '
                ' always process because srcform was provided - return here after processing
                '
                If button = buttonSave Or button = buttonOK Then
                    Call cs.Open("people", "id=" & userId)
                    If cs.OK Then
                        Call cs.SetFormInput("name")
                        Call cs.SetFormInput("firstName")
                        Call cs.SetFormInput("lastName")
                        Call cs.SetFormInput("email")
                    End If
                    Call cs.Close()
                End If
                If button = buttonCancel Or button = buttonOK Then
                    returnNextFormId = formIdMacroList
                Else
                    returnNextFormId = formIdMacroDetails
                End If
            Catch ex As Exception
                '
                '
                '
                errorReport(ex, cp, "processForm")
            End Try
            Return returnNextFormId
        End Function
        '
        '
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As String
            Dim s As String = ""
            '
            Try
                Dim body As CPBlockBaseClass = cp.BlockNew()
                Dim userId As Integer = cp.Doc.GetInteger(rnUserId)
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim hidden As String = ""
                Dim form As String = ""
                Dim buttons As String = ""
                Dim nvForm As New adminFramework.formNameValueRowsClass
                Dim addLink As String = cp.Doc.GetProperty("adminUrl") & "?af=4&cid=" & cp.Content.GetID("people")
                '
                Call cp.Doc.AddRefreshQueryString(rnUserId, userId)
                Call cp.Doc.AddRefreshQueryString(rnDstFormId, dstFormId)
                '
                nvForm.title = "User Details (managerSampleADetailsClass)"
                nvForm.addFormButton(buttonCancel)
                nvForm.addFormButton(buttonSave)
                nvForm.addFormButton(buttonOK)
                nvForm.addFormHidden(rnSrcFormId, dstFormId)
                '
                Call cs.Open("people", "id=" & userId)
                If Not cs.OK Then
                    body.Load("<p>The user you requested could not be found.</p>")
                Else
                    '
                    nvForm.addRow()
                    nvForm.rowName = "Name"
                    nvForm.rowValue = cp.Html.InputText("name", cs.GetText("name"))
                    '
                    nvForm.addRow()
                    nvForm.rowName = "First Name"
                    nvForm.rowValue = cp.Html.InputText("firstname", cs.GetText("firstname"))
                    '
                    nvForm.addRow()
                    nvForm.rowName = "Last Name"
                    nvForm.rowValue = cp.Html.InputText("lastname", cs.GetText("lastname"))
                    '
                    nvForm.addRow()
                    nvForm.rowName = "Email"
                    nvForm.rowValue = cp.Html.InputText("email", cs.GetText("email"))
                End If
                Call cs.Close()
                s = nvForm.getHtml(cp)
            Catch ex As Exception
                errorReport(ex, cp, "getForm")
            End Try
            Return s
        End Function
        '
        '
        '
        Friend Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.adminAccountDetailsClass." & method)
        End Sub
    End Class
End Namespace
