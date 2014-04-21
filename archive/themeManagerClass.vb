
Option Explicit On

Imports Contensive.BaseClasses

Namespace Contensive.addons
    Public Class themeManagerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Const rnSrcFormId As String = "srcFormId"
        Const rnDstFormId As String = "dstFormId"
        '
        Const formIdSectionAMin As Integer = 100
        Const formIdSectionAMax As Integer = 199
        '
        Const formIdSectionBMin As Integer = 200
        Const formIdSectionBMax As Integer = 299
        '
        Const formIdSectionCMin As Integer = 300
        Const formIdSectionCMax As Integer = 399
        '
        '=================================================================================
        '   Aggregate Object Interface
        '=================================================================================
        '
        Public Overrides Function Execute(ByVal cp As CPBaseClass) As Object
            Dim returnHtml As String = "Theme Manager"
            '
            Try
                Dim manager As New adminFramework.pageWithNavClass
                Dim rightNow As Date = New Date()
                Dim srcFormId As Integer = cp.Doc.GetInteger(rnSrcFormId)
                Dim dstFormId As Integer = cp.Doc.GetInteger(rnDstFormId)
                Dim rqs As String = cp.Doc.RefreshQueryString
                Dim contentBody As String = "default content message"
                '
                If srcFormId <> 0 Then
                    Select Case srcFormId
                        Case formIdSectionAMin To formIdSectionAMax
                            '
                            ' section processing returns the dstFormId
                            '
                            dstFormId = srcFormId
                        Case formIdSectionBMin To formIdSectionBMax
                            '
                            ' section processing returns the dstFormId
                            '
                            dstFormId = srcFormId
                        Case formIdSectionCMin To formIdSectionCMax
                            '
                            ' section processing returns the dstFormId
                            '
                            dstFormId = srcFormId
                            ' never an else case in processing section
                    End Select
                End If
                '
                manager.addNav()
                manager.navCaption = "Section A"
                manager.navLink = "?" & cp.Utils.ModifyQueryString(rqs, rnDstFormId, formIdSectionAMin)
                '
                manager.addNav()
                manager.navCaption = "Section B"
                manager.navLink = "?" & cp.Utils.ModifyQueryString(rqs, rnDstFormId, formIdSectionBMin)
                '
                manager.addNav()
                manager.navCaption = "Section C"
                manager.navLink = "?" & cp.Utils.ModifyQueryString(rqs, rnDstFormId, formIdSectionCMin)
                '
                Select Case dstFormId
                    Case formIdSectionBMin To formIdSectionBMax
                        '
                        ' section get returns html
                        '
                        contentBody = "Section B Content"
                    Case formIdSectionCMin To formIdSectionCMax
                        '
                        ' section get returns html
                        '
                        contentBody = "Section C Content"
                    Case Else
                        '
                        ' else default section 
                        '
                        contentBody = "Section A Content"
                End Select
                '
                ' Assemble body
                '   wrap the contentBody with the 'frame' - a div with a unique ID used as a destination for ajax handlers
                '
                manager.body = cp.Html.div(contentBody, , , "managerFrame")
                returnHtml = manager.getHtml(cp)
            Catch ex As Exception
                HandleClassError(cp, ex, "Execute", "trap")
            End Try
            Return returnHtml
        End Function
        '
        '=================================================================================
        '   Handle errors from this class
        '=================================================================================
        '
        Private Sub HandleClassError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByRef MethodName As String, ByVal description As String)
            '
            Call cp.Site.ErrorReport(ex, "Error in themeManagerClass." & MethodName & ", " & description)
            '
        End Sub
    End Class
End Namespace
