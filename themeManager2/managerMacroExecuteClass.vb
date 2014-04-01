
Imports Contensive.BaseClasses
Imports adminFramework

Namespace Contensive.addons.themeManager
    '
    Public Class managerMacroExecuteClass
        '
        '
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightNow As Date) As Integer
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
                nextFormId = srcFormId
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
            Try
                Dim form As New formSimpleClass
                '
                form.title = "Macro Execute"
                form.body = "This is the body"
                returnHtml = form.getHtml(cp)
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
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.adminListClass." & method)
        End Sub
    End Class
End Namespace
