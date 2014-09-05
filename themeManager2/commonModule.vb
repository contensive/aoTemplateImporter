
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
            Public Const savelayout As Integer = 8
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
            Public Const findReplace As Integer = 19
            Public Const setHref As Integer = 20
            Public Const setSrc As Integer = 21
            Public Const setClass As Integer = 22
            Public Const setId As Integer = 23
            Public Const insertAfter As Integer = 24
            Public Const saveStyle As Integer = 25
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
                Dim find As String = ""
                Dim replace As String = ""
                Dim regName As String = ""
                Dim regCnt As Integer = 0
                Dim regPtr As Integer
                Dim srcRegPtr As Integer = 0
                Dim dstRegPtr As Integer = 0
                Dim srcValue As String = ""
                Dim dstValue As String = ""
                Dim replacePtr As Integer = 0
                Dim replaceValue As String = ""

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
                        find = cs.GetText("find")
                        replace = cs.GetText("replace")
                        Select Case cs.GetInteger("instructionId")
                            Case themeImportMacroInstructions.saveWwwFile
                                '
                                ' saves register [src], optionally selects inner [find], to www file=[dst]
                                '
                                return_progressMessage &= "<br>Save File, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src = "") Then
                                    return_progressMessage &= ", ERROR: source can not be empty"
                                ElseIf (dst = "") Then
                                    return_progressMessage &= ", ERROR: destination can not be empty"
                                Else
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr < 0 Then
                                        srcValue = src
                                        return_progressMessage &= ", source is liternal"
                                    Else
                                        srcValue = registerValues(regPtr)
                                        return_progressMessage &= ", source is register"
                                    End If
                                    If find = "" Then
                                        dstValue = srcValue
                                    Else
                                        blockWork.Load(srcValue)
                                        dstValue = blockWork.GetInner(find)
                                        return_progressMessage &= ", getInner on source"
                                    End If
                                    Call cp.File.Save(cp.Site.PhysicalWWWPath & dst, dstValue)
                                End If
                            Case themeImportMacroInstructions.saveStyle
                                '
                                ' saves register [src], optionally selects inner [find], to shared style=[dst]
                                '
                                return_progressMessage &= "<br>Save Style, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src = "") Then
                                    return_progressMessage &= ", ERROR: source can not be empty"
                                ElseIf (dst = "") Then
                                    return_progressMessage &= ", ERROR: destination can not be empty"
                                Else
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr < 0 Then
                                        srcValue = src
                                        return_progressMessage &= ", source is liternal"
                                    Else
                                        srcValue = registerValues(regPtr)
                                        return_progressMessage &= ", source is register"
                                    End If
                                    If find = "" Then
                                        dstValue = srcValue
                                    Else
                                        blockWork.Load(srcValue)
                                        dstValue = blockWork.GetInner(find)
                                        return_progressMessage &= ", getInner on source"
                                    End If
                                    If Not csWork.Open("shared styles", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                        return_progressMessage &= ", create new Layout"
                                        Call csWork.Close()
                                        Call csWork.Insert("shared styles")
                                        Call csWork.SetField("name", dst)
                                        Call csWork.SetField("alwaysInclude", "1")
                                    End If
                                    If csWork.OK Then
                                        Call csWork.SetField("StyleFilename", dstValue)
                                    End If
                                    Call csWork.Close()
                                End If
                            Case themeImportMacroInstructions.savelayout
                                '
                                ' saves register [src], optionally selects inner [find], to layout name=[dst]
                                '
                                return_progressMessage &= "<br>Save Layout, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src = "") Then
                                    return_progressMessage &= ", ERROR: source can not be empty"
                                ElseIf (dst = "") Then
                                    return_progressMessage &= ", ERROR: destination can not be empty"
                                Else
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr < 0 Then
                                        srcValue = src
                                        return_progressMessage &= ", source is liternal"
                                    Else
                                        srcValue = registerValues(regPtr)
                                        return_progressMessage &= ", source is register"
                                    End If
                                    If find = "" Then
                                        dstValue = srcValue
                                    Else
                                        blockWork.Load(srcValue)
                                        dstValue = blockWork.GetInner(find)
                                        return_progressMessage &= ", getInner on source"
                                    End If
                                    If Not csWork.Open("layouts", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                        return_progressMessage &= ", create new Layout"
                                        Call csWork.Close()
                                        Call csWork.Insert("layouts")
                                        Call csWork.SetField("name", dst)
                                    End If
                                    If csWork.OK Then
                                        Call csWork.SetField("layout", dstValue)
                                    End If
                                    Call csWork.Close()
                                End If
                            Case themeImportMacroInstructions.insertAfter
                                '
                                ' insert [replace] right after element [find] of html [src]
                                '
                                return_progressMessage &= "<br>insertAfter, src=[" & src & "], find=[" & find & "], replace=[" & replace & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        dstValue = registerValues(dstRegPtr)
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
                                        replacePtr = getRegPtr(regCnt, registerNames, replace, False)
                                        If replacePtr >= 0 Then
                                            '
                                            ' replace is a register
                                            '
                                            replaceValue = registerValues(replacePtr)
                                        Else
                                            '
                                            ' replace is a literal
                                            '
                                            replaceValue = replace
                                        End If
                                        If find = "" Then
                                            '
                                            ' simple append
                                            '
                                            dstValue = srcValue & replaceValue
                                        Else
                                            '
                                            ' append inner
                                            '
                                            blockWork.Load(srcValue)
                                            Call blockWork.SetOuter(find, blockWork.GetOuter(find) & replaceValue)
                                            dstValue = blockWork.GetHtml()
                                        End If
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If


                            Case themeImportMacroInstructions.findReplace
                                '
                                ' does find [find] and replace [replace] from src to dst
                                ' if find="" dst is unchanged
                                '
                                return_progressMessage &= "<br>FindReplace, src=[" & src & "], find=[" & find & "], replace=[" & replace & "], dst=[" & dst & "]"
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
                                        replacePtr = getRegPtr(regCnt, registerNames, replace, False)
                                        If replacePtr >= 0 Then
                                            '
                                            ' replace is a register
                                            '
                                            replaceValue = registerValues(replacePtr)
                                        Else
                                            '
                                            ' replace is a literal
                                            '
                                            replaceValue = replace
                                        End If
                                        If find <> "" Then
                                            dstValue = dstValue.Replace(find, replaceValue)
                                            registerValues(dstRegPtr) = dstValue
                                        End If
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadPage
                                '
                                ' load page content where name=[src], optionally selects inner [find], saves in register [dst]
                                '
                                return_progressMessage &= "<br>Load Page, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        If Not csWork.Open("page content", "name=" & cp.Db.EncodeSQLText(src)) Then
                                            return_progressMessage &= "<br>***** page content record not found"
                                        Else
                                            srcValue = csWork.GetText("copyFilename")
                                            If find = "" Then
                                                dstValue = srcValue
                                            Else
                                                Call blockWork.OpenFile(srcValue)
                                                dstValue = blockWork.GetInner(find)
                                            End If
                                            registerValues(dstRegPtr) = dstValue
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.savePage
                                '
                                ' saves register [src], optionally selects inner [find], to page content name=[dst]
                                '
                                return_progressMessage &= "<br>Save Page, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If srcRegPtr >= 0 Then
                                        srcValue = registerValues(srcRegPtr)
                                        If find = "" Then
                                            dstValue = srcValue
                                        Else
                                            blockWork.Load(srcValue)
                                            dstValue = blockWork.GetInner(find)
                                        End If
                                        If Not csWork.Open("page content", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                            Call csWork.Close()
                                            Call csWork.Insert("page content")
                                            Call csWork.SetField("name", dst)
                                        End If
                                        If csWork.OK Then
                                            Call csWork.SetField("copyFilename", dstValue)
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadCopy
                                '
                                ' reads copy content where name=[src], optionally selects inner [find], saves in register [dst]
                                '
                                return_progressMessage &= "<br>Load Copy, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        If Not csWork.Open("copy content", "name=" & cp.Db.EncodeSQLText(src)) Then
                                            return_progressMessage &= "<br>***** copy content record not found"
                                        Else
                                            srcValue = csWork.GetText("copyFilename")
                                            If find = "" Then
                                                dstValue = srcValue
                                            Else
                                                Call blockWork.OpenFile(srcValue)
                                                dstValue = blockWork.GetInner(find)
                                            End If
                                            registerValues(dstRegPtr) = dstValue
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.saveCopy
                                '
                                ' saves register [src], optionally selects inner [find], to copy content name=[dst]
                                '
                                return_progressMessage &= "<br>Save Copy, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If find = "" Then
                                            dstValue = srcValue
                                        Else
                                            blockWork.Load(srcValue)
                                            dstValue = blockWork.GetInner(find)
                                        End If
                                        If Not csWork.Open("copy content", "name=" & cp.Db.EncodeSQLText(dst)) Then
                                            Call csWork.Close()
                                            Call csWork.Insert("copy content")
                                            Call csWork.SetField("name", dst)
                                        End If
                                        If csWork.OK Then
                                            Call csWork.SetField("copyFilename", dstValue)
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.append
                                '
                                ' appends [replace] to child of [find] of html [src]
                                '
                                return_progressMessage &= "<br>Append, src=[" & src & "], find=[" & find & "], replace=[" & replace & "], dst=[" & dst & "]"
                                If (dst = "") Then
                                    return_progressMessage &= ", ERROR: destination is required"
                                Else
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr < 0 Then
                                        return_progressMessage &= ", ERROR: destination register not created"
                                    Else
                                        dstValue = registerValues(dstRegPtr)
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                            return_progressMessage &= ", source is a register"
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                            return_progressMessage &= ", source is a literal"
                                        End If
                                        replacePtr = getRegPtr(regCnt, registerNames, replace, False)
                                        If replacePtr >= 0 Then
                                            '
                                            ' replace is a register
                                            '
                                            replaceValue = registerValues(replacePtr)
                                            return_progressMessage &= ", replace is a register"
                                        Else
                                            '
                                            ' replace is a literal
                                            '
                                            replaceValue = replace
                                            return_progressMessage &= ", replace is a literal"
                                        End If
                                        If find = "" Then
                                            '
                                            ' simple append
                                            '
                                            dstValue = srcValue & replaceValue
                                            return_progressMessage &= ", append replace to source, save to destination"
                                        Else
                                            '
                                            ' append inner
                                            '
                                            blockWork.Load(srcValue)
                                            Call blockWork.SetInner(find, blockWork.GetInner(find) & replaceValue)
                                            dstValue = blockWork.GetHtml()
                                            return_progressMessage &= ", append replace as last child node of selector in source, save to destination"
                                        End If
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.getInner
                                '
                                ' gets [src], optionally selects inner [find], to register [dst]
                                '
                                return_progressMessage &= "<br>Get Inner, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
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
                                        If find <> "" Then
                                            blockWork.Load(srcValue)
                                            dstValue = blockWork.GetInner(find)
                                        Else
                                            dstValue = srcValue
                                        End If
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.getOuter
                                '
                                ' gets [src], optionally selects outer [find], to register [dst]
                                '
                                return_progressMessage &= "<br>Get Outer, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
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
                                        srcValue = blockWork.GetOuter(find)
                                        registerValues(dstRegPtr) = srcValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.setOuter
                                '
                                ' saves [src] to outer [find] of register [dst]
                                '
                                return_progressMessage &= "<br>Set Outer, src=[" & src & "], find=[" & find & "], replace=[" & replace & "], dst=[" & dst & "]"
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
                                        replacePtr = getRegPtr(regCnt, registerNames, replace, False)
                                        If replacePtr >= 0 Then
                                            '
                                            ' replace is a register
                                            '
                                            replaceValue = registerValues(replacePtr)
                                        Else
                                            '
                                            ' replace is a literal
                                            '
                                            replaceValue = replace
                                        End If
                                        dstValue = srcValue
                                        If find <> "" Then
                                            blockWork.Load(dstValue)
                                            Call blockWork.SetOuter(find, replaceValue)
                                            dstValue = blockWork.GetHtml()
                                        End If
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.setInner
                                '
                                ' saves [src] to inner [find] of register [dst]
                                '
                                return_progressMessage &= "<br>Set Inner, src=[" & src & "], find=[" & find & "], replace=[" & replace & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        srcRegPtr = getRegPtr(regCnt, registerNames, src, False)
                                        If srcRegPtr >= 0 Then
                                            '
                                            ' src is a register
                                            '
                                            srcValue = registerValues(srcRegPtr)
                                            return_progressMessage &= ", src is a register"
                                        Else
                                            '
                                            ' src is literal
                                            '
                                            srcValue = src
                                            return_progressMessage &= ", src is literal"
                                        End If
                                        replacePtr = getRegPtr(regCnt, registerNames, replace, False)
                                        If replacePtr >= 0 Then
                                            '
                                            ' replace is a register
                                            '
                                            replaceValue = registerValues(replacePtr)
                                            return_progressMessage &= ", replace  is a register"
                                        Else
                                            '
                                            ' replace is a literal
                                            '
                                            replaceValue = replace
                                            return_progressMessage &= ", replace is literal"
                                        End If
                                        dstValue = srcValue
                                        If find <> "" Then
                                            blockWork.Load(dstValue)
                                            Call blockWork.SetInner(find, replaceValue)
                                            dstValue = blockWork.GetHtml()
                                        End If
                                        registerValues(dstRegPtr) = dstValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadLayout
                                '
                                ' load layout where name=[src], optionally selects inner [find], saves in register [dst]
                                '
                                return_progressMessage &= "<br>Load Layout, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    dstRegPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If dstRegPtr >= 0 Then
                                        If Not csWork.Open("layouts", "name=" & cp.Db.EncodeSQLText(src)) Then
                                            return_progressMessage &= "<br>***** layout content record not found"
                                        Else
                                            srcValue = csWork.GetText("layout")
                                            If find <> "" Then
                                                Call blockWork.OpenFile(srcValue)
                                                srcValue = blockWork.GetInner(find)
                                            End If
                                            registerValues(dstRegPtr) = srcValue
                                        End If
                                        Call csWork.Close()
                                    End If
                                End If
                            Case themeImportMacroInstructions.loadWwwFile
                                '
                                ' load www file where name=[src], optionally selects inner [find], saves in register [dst]
                                '
                                return_progressMessage &= "<br>Load www File, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, dst, True)
                                    If regPtr >= 0 Then
                                        src = cp.Site.PhysicalWWWPath & src
                                        srcValue = cp.File.Read(src)
                                        If find <> "" Then
                                            Call blockWork.OpenFile(src)
                                            srcValue = blockWork.GetInner(find)
                                        End If
                                        registerValues(regPtr) = srcValue
                                    End If
                                End If
                            Case themeImportMacroInstructions.saveTemplateBody
                                '
                                ' saves register [src], optionally selects inner [find], to template body where name=[dst]
                                '
                                return_progressMessage &= "<br>Save Template Body, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If find <> "" Then
                                            blockWork.Load(srcValue)
                                            srcValue = blockWork.GetInner(find)
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
                                ' saves register [src], optionally selects inner [find], to template head where name=[dst]
                                '
                                return_progressMessage &= "<br>Save Template Head, src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
                                If (src <> "") And (dst <> "") Then
                                    regPtr = getRegPtr(regCnt, registerNames, src, False)
                                    If regPtr >= 0 Then
                                        srcValue = registerValues(regPtr)
                                        If find <> "" Then
                                            blockWork.Load(srcValue)
                                            srcValue = blockWork.GetInner(find)
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
                                return_progressMessage &= "<br>Unknown Command, instructionId=[" & cs.GetInteger("instructionId") & "], src=[" & src & "], find=[" & find & "], dst=[" & dst & "]"
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
        '
        Friend Function cpVisitGetText(ByVal cp As CPBaseClass, ByVal propertyName As String, Optional ByVal defaultPropertyValue As String = "") As String
            Dim returnString As String = ""
            Try
                returnString = cp.Visit.GetText(propertyName, defaultPropertyValue)
                If (returnString Is Nothing) Then
                    returnString = ""
                End If
            Catch ex As Exception
                returnString = ""

            End Try
            Return returnString
        End Function
    End Module
End Namespace
