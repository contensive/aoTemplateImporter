
Imports Contensive.BaseClasses
Namespace Contensive.addons.themeManager

    Public Module commonModule
        '
        Public Const cr As String = vbCrLf & vbTab
        Public Const cr2 As String = cr & vbTab
        Public Const cr3 As String = cr2 & vbTab
        '
        ' Get Layout,Get www File,Get Content File,Get Inner,Get Outer,Set Inner,Set Outer,Set Layout,Set www File,Set Content File,Set Template
        '
        Public Structure themeImportMacroInstructions
            Public Const loadLayout As Integer = 1
            Public Const loadWwwFile As Integer = 2
            Public Const loadContentFile As Integer = 3
            Public Const getInner As Integer = 4
            Public Const getOuter As Integer = 5
            Public Const setInner As Integer = 6
            Public Const setOuter As Integer = 7
            Public Const setlayout As Integer = 8
            Public Const saveWwwFile As Integer = 9
            Public Const saveContentFile As Integer = 10
            Public Const saveTemplateBody As Integer = 11
            Public Const saveTemplateHead As Integer = 12
            Public Const saveTemplateBodyTag As Integer = 13
            Public Const append As Integer = 14
            Public Const saveCopy As Integer = 15
            Public Const loadCopy As Integer = 16
            Public Const savePage As Integer = 17
            Public Const loadPage As Integer = 18
        End Structure
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
        Public Const formIdMacroExecute As Integer = 111
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
        '
        '
        '
        Friend Function executeMacro(ByVal cp As CPBaseClass, ByVal macroId As Integer, ByRef return_progressMessage As String) As Boolean
            Dim returnOK As Boolean
            Try
                Dim registerNames(100) As String
                Dim registerValues(100) As String
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim csWork As CPCSBaseClass = cp.CSNew()
                Dim blockWork As CPBlockBaseClass = cp.BlockNew()
                Dim src As String = ""
                Dim dst As String = ""
                Dim selector As String = ""
                Dim regName As String = ""
                Dim regCnt As Integer = 0
                Dim regPtr As Integer
                Dim srcRegPtr As Integer = 0
                Dim dstRegPtr As Integer = 0
                Dim srcValue As String = ""
                Dim dstValue As String = ""
                '
                return_progressMessage = ""
                If cs.Open("theme import macros", "id=" & macroId) Then

                End If
                Call cs.Close()
                '
                If cs.Open("theme import macro lines", "themeImportMacroId=" & macroId, "sortorder,id", , , 999) Then
                    Do
                        src = cs.GetText("source")
                        dst = cs.GetText("destination")
                        selector = cs.GetText("selector")
                        Select Case cs.GetInteger("instructionId")
                            Case themeImportMacroInstructions.loadPage
                                return_progressMessage &= "<br>Load Page, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        If Not csWork.Open("page content", "name=" & cp.Db.EncodeSQLText(src)) Then
                                            return_progressMessage &= "<br>***** page content record not found"
                                        Else
                                            srcValue = csWork.GetText("copyFilename")
                                            If selector <> "" Then
                                                Call blockWork.OpenFile(srcValue)
                                                srcValue = blockWork.GetInner(selector)
                                            End If
                                            registerValues(dstRegPtr) = srcValue
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.savePage
                                return_progressMessage &= "<br>Save Page, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If selector <> "" Then
                                            blockWork.Load(srcValue)
                                            srcValue = blockWork.GetInner(selector)
                                        End If
                                        If Not csWork.Open("page content", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                            Call csWork.Close()
                                            Call csWork.Insert("page content")
                                            Call csWork.SetField("name", dst)
                                        End If
                                        If csWork.OK Then
                                            Call csWork.SetField("copyFilename", srcValue)
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadCopy
                                return_progressMessage &= "<br>Load Copy, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        If Not csWork.Open("copy content", "name=" & cp.Db.EncodeSQLText(src)) Then
                                            return_progressMessage &= "<br>***** copy content record not found"
                                        Else
                                            srcValue = csWork.GetText("copyFilename")
                                            If selector <> "" Then
                                                Call blockWork.OpenFile(srcValue)
                                                srcValue = blockWork.GetInner(selector)
                                            End If
                                            registerValues(dstRegPtr) = srcValue
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.saveCopy
                                '
                                '
                                '
                                return_progressMessage &= "<br>Save Copy, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If selector <> "" Then
                                            blockWork.Load(srcValue)
                                            srcValue = blockWork.GetInner(selector)
                                        End If
                                        If Not csWork.Open("copy content", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                            Call csWork.Close()
                                            Call csWork.Insert("copy content")
                                            Call csWork.SetField("name", dst)
                                        End If
                                        If csWork.OK Then
                                            Call csWork.SetField("copyFilename", srcValue)
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.append
                                '
                                '
                                '
                                return_progressMessage &= "<br>Append, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                        End If
                                        If selector = "" Then
                                            '
                                            ' simple append
                                            '
                                            registerValues(dstRegPtr) &= srcValue
                                        Else
                                            '
                                            ' append inner
                                            '
                                            blockWork.Load(srcValue)
                                            dstValue = blockWork.GetInner(selector) & src
                                            Call blockWork.SetInner(selector, dstValue)
                                            dstValue = blockWork.GetHtml()
                                            registerValues(dstRegPtr) = dstValue
                                        End If
                                    End If
                                End If
                            Case themeImportMacroInstructions.getInner
                                '
                                '
                                '
                                return_progressMessage &= "<br>Get Inner, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                        End If
                                        blockWork.Load(srcValue)
                                        srcValue = blockWork.GetInner(selector)
                                        registerValues(dstRegPtr) = srcValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.getOuter
                                '
                                '
                                '
                                return_progressMessage &= "<br>Get Outer, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                        End If
                                        blockWork.Load(srcValue)
                                        srcValue = blockWork.GetOuter(selector)
                                        registerValues(dstRegPtr) = srcValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.setOuter
                                '
                                '
                                '
                                return_progressMessage &= "<br>Set Outer, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                        End If
                                        dstValue = registerValues(dstRegPtr)
                                        blockWork.Load(dstValue)
                                        Call blockWork.SetOuter(selector, srcValue)
                                        dstValue = blockWork.GetHtml()
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.setInner
                                '
                                '
                                '
                                return_progressMessage &= "<br>Set Inner, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                        End If
                                        dstValue = registerValues(dstRegPtr)
                                        blockWork.Load(dstValue)
                                        Call blockWork.SetInner(selector, srcValue)
                                        dstValue = blockWork.GetHtml()
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadLayout
                                '
                                ' load wwwfile
                                '
                                return_progressMessage &= "<br>Load Layout, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        If Not csWork.Open("layouts", "name=" & cp.Db.EncodeSQLText(src)) Then
                                            return_progressMessage &= "<br>***** layout content record not found"
                                        Else
                                            srcValue = csWork.GetText("layout")
                                            If selector <> "" Then
                                                Call blockWork.OpenFile(srcValue)
                                                srcValue = blockWork.GetInner(selector)
                                            End If
                                            registerValues(dstRegPtr) = srcValue
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadWwwFile
                                '
                                ' load wwwfile
                                '
                                return_progressMessage &= "<br>Load www File, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If regPtr >= 0 Then
                                        src = cp.Site.PhysicalWWWPath & src
                                        srcValue = cp.File.Read(src)
                                        If selector <> "" Then
                                            Call blockWork.OpenFile(src)
                                            srcValue = blockWork.GetInner(selector)
                                        End If
                                        registerValues(regPtr) = srcValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.saveTemplateBody
                                '
                                '
                                '
                                return_progressMessage &= "<br>Save Template Body, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If selector <> "" Then
                                            blockWork.Load(srcValue)
                                            srcValue = blockWork.GetInner(selector)
                                        End If
                                        If Not csWork.Open("page templates", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                            Call csWork.Close()
                                            Call csWork.Insert("page templates")
                                            Call csWork.SetField("name", dst)
                                        End If
                                        If csWork.OK Then
                                            Call csWork.SetField("BodyHTML", srcValue)
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.saveTemplateHead
                                '
                                '
                                '
                                return_progressMessage &= "<br>SAve Template Head, src=[" & src & "], selector=[" & selector & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If selector <> "" Then
                                            blockWork.Load(srcValue)
                                            srcValue = blockWork.GetInner(selector)
                                        End If
                                        If Not csWork.Open("page templates", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                            Call csWork.Close()
                                            Call csWork.Insert("page templates")
                                            Call csWork.SetField("name", dst)
                                        End If
                                        If csWork.OK Then
                                            Call csWork.SetField("OtherHeadTags", srcValue)
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case Else

                                '
                                '
                                '
                        End Select
                        '
                        '
                        '
                        Call cs.GoNext()
                    Loop While cs.OK()
                End If
                Call cs.Close()
                Call cp.Cache.ClearAll()
                '
                return_progressMessage &= "<br>Execute Completed Successfully."
            Catch ex As Exception

            End Try
            Return returnOK
        End Function
        '
        '
        '
        Private Function getRegPtr(ByRef regCnt As Integer, ByRef registernames() As String, ByVal regName As String, createIfNotFound As Boolean) As Integer
            Dim regPtr As Integer = 0
            '
            regName = regName.Trim
            If regCnt > 0 Then
                For regPtr = 0 To regCnt - 1
                    If (registernames(regPtr) = regName) Then
                        Exit For
                    End If
                Next
            End If
            If regPtr >= regCnt Then
                '
                ' not found
                '
                If createIfNotFound Then
                    regPtr = regCnt
                    regCnt += 1
                    registernames(regPtr) = regName
                Else
                    regPtr = -1
                End If
            End If
            Return regPtr
        End Function
    End Module
End Namespace
