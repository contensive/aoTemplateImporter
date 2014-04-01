
Imports Contensive.BaseClasses

Namespace Contensive.addons.themeManager
    Public Class ajaxHtmlClass
        Inherits AddonBaseClass
        '
        ' Ajax Handler - Remote Method
        '   returns content of inner classes to the contains originally created around them
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            '
            Try
                Dim macroList As New managerMacroListClass
                Dim macroExecute As New managerMacroExecuteClass
                Dim srcFormId As Integer = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnSrcFormId))
                Dim dstFormId As Integer = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnDstFormId))
                Dim rqs As String = CP.Doc.GetProperty("baseRqs")
                Dim rightNow As Date = getRightNow(CP)
                '
                ' process form
                '
                If (srcFormId <> 0) Then
                    Select Case srcFormId
                        Case formIdMacroList
                            '
                            '
                            '
                            dstFormId = macroList.processForm(CP, srcFormId, rqs, rightNow)
                        Case formIdMacroExecute
                            '
                            '
                            '
                            dstFormId = macroExecute.processForm(CP, srcFormId, rqs, rightNow)
                    End Select
                End If
                '
                ' get form
                '
                Select Case dstFormId
                    Case formIdMacroList
                        '
                        '
                        '
                        returnHtml = macroList.getForm(CP, dstFormId, rqs, rightNow)
                    Case formIdMacroExecute
                        '
                        '
                        '
                        returnHtml = macroExecute.getForm(CP, dstFormId, rqs, rightNow)
                End Select
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in aoManagerTemplate.adminClass.execute")
            End Try
            '
            ' assemble body
            '
            Return returnHtml
        End Function
    End Class
End Namespace
