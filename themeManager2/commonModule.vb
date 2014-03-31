
Imports Contensive.BaseClasses
Namespace Contensive.addons.themeManager

    Public Module commonModule
        '
        Public Const cr As String = vbCrLf & vbTab
        Public Const cr2 As String = cr & vbTab
        Public Const cr3 As String = cr2 & vbTab
        '
        Public Const buttonOK As String = " OK "
        Public Const buttonSave As String = " Save "
        Public Const buttonCancel As String = " Cancel "
        Public Const buttonAdd As String = " Add "
        '
        Public Const rnUserId As String = "userId"
        '
        Public Const rnSrcFormId As String = "srcFormId"
        Public Const rnDstFormId As String = "dstFormId"
        Public Const rnButton As String = "button"
        '
        ' Home Form
        '
        Public Const formIdHome As Integer = 1
        '
        ' typeA Forms
        '
        Public Const formIdMacroMin As Integer = 110
        Public Const formIdMacroList As Integer = 110
        Public Const formIdMacroDetails As Integer = 121
        Public Const formIdMacroDetailList As Integer = 122
        Public Const formIdMacroMax As Integer = 129
        '
        '
        ' typeB Forms
        '
        Public Const formIdToolsMin As Integer = 130
        Public Const formIdToolsQuickImport As Integer = 130
        Public Const formIdToolsQuickImportDone As Integer = 131
        Public Const formIdToolsMax As Integer = 139
        '
        ' typeC Forms
        '
        'Public Const formIdToolsMin As Integer = 140
        'Public Const formIdToolsList As Integer = 140
        'Public Const formIdToolsMax As Integer = 149
        '
        ' Settings
        '
        Public Const formIdSettings As Integer = 170
        '
        ' Reference -- the admin framework styles for table columns
        '
        '/*
        '* Manager table cell widths
        '*/
        '#afw .afwWidth10 { width: 10% }
        '#afw .afwWidth20 { width: 20% }
        '#afw .afwWidth30 { width: 30% }
        '#afw .afwWidth40 { width: 40% }
        '#afw .afwWidth50 { width: 50% }
        '#afw .afwWidth60 { width: 60% }
        '#afw .afwWidth70 { width: 70% }
        '#afw .afwWidth80 { width: 80% }
        '#afw .afwWidth90 { width: 90% }
        '#afw .afwWidth100 { width: 100% }
        '/*
        '*/
        '#afw .afwWidth10px { width: 10px }
        '#afw .afwWidth20px { width: 20px }
        '#afw .afwWidth30px { width: 30px }
        '#afw .afwWidth40px { width: 40px }
        '#afw .afwWidth50px { width: 50px }
        '#afw .afwWidth60px { width: 60px }
        '#afw .afwWidth70px { width: 70px }
        '#afw .afwWidth80px { width: 80px }
        '#afw .afwWidth90px { width: 90px }

        '#afw .afwWidth100px { width: 100px }
        '#afw .afwWidth200px { width: 200px }
        '#afw .afwWidth300px { width: 300px }
        '#afw .afwWidth400px { width: 400px }
        '#afw .afwWidth500px { width: 500px }
        '/*
        '*/
        '#afw .afwMaginLeft100px { margin-left: 100px }
        '#afw .afwMaginLeft200px { margin-left: 200px }
        '#afw .afwMaginLeft300px { margin-left: 300px }
        '#afw .afwMaginLeft400px { margin-left: 400px }
        '#afw .afwMaginLeft500px { margin-left: 500px }
        '/*
        '*/
        '#afw .afwTextAlignRight { text-align:right }
        '#afw .afwTextAlignLeft { text-align:left }
        '#afw .afwTextAlignCenter { text-align:center }
        '    '
        Public Function toJSON(ByVal value As String) As String
            Dim s As String = value
            Try
                '
                s = s.Replace("""", "\""")
                s = s.Replace(vbCrLf, "\n")
                s = s.Replace(vbCr, "\n")
                s = s.Replace(vbLf, "\n")
                '
            Catch ex As Exception
                s = value
            End Try
            Return s
        End Function
        '
        '
        '
        Friend Function buffDate(ByVal sourceDate As Date) As String
            Dim returnValue As String
            '
            If sourceDate < #1/1/1900# Then
                returnValue = ""
            Else
                returnValue = sourceDate.ToShortDateString
            End If
            Return returnValue

        End Function
        '
        '
        '
        Friend Function getRightNow(ByVal cp As Contensive.BaseClasses.CPBaseClass) As Date
            Dim returnValue As Date = Date.Now()
            Try
                '
                ' change 'sample' to the name of this collection
                '
                Dim testString As String = cp.Site.GetProperty("Sample Manager Test Mode Date", "")
                If testString <> "" Then
                    returnValue = encodeMinDate(cp.Utils.EncodeDate(testString))
                    If returnValue = Date.MinValue Then
                        returnValue = Date.Now()
                    End If
                End If
            Catch ex As Exception
            End Try
            Return returnValue
        End Function
        '
        '
        '
        Friend Function encodeMinDate(ByVal sourceDate As Date) As Date
            Dim returnValue As Date = sourceDate
            If returnValue < #1/1/1900# Then
                returnValue = Date.MinValue
            End If
            Return returnValue
        End Function
        '
        '
        '
        '
        Friend Sub appendLog(ByVal cp As CPBaseClass, ByVal logMessage As String)
            Dim nowDate As Date = Date.Now.Date()
            Dim logFilename As String = nowDate.Year & nowDate.Month.ToString("D2") & nowDate.Day.ToString("D2") & ".log"
            Call cp.File.CreateFolder(cp.Site.PhysicalInstallPath & "\logs\managerSample")
            Call cp.Utils.AppendLog("managerSample\" & logFilename, logMessage)
        End Sub
    End Module
End Namespace
