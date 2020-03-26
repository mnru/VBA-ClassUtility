Attribute VB_Name = "classUtil"
Option Base 0

Function disposeProc(tp, cmpn, procName, Optional knd = 0, Optional sCode = "")
    'tp "get","del","replace"
    Dim cmp
    Dim lineStart, lineDef, lineContent, lineEnd, defcnt, linecnt, i
    Dim xdef, xcnt, xend
    disposeProc = ""
    tp = LCase(tp)
    If tp = "del" Then sCode = ""
    Set cmp = ActiveWorkbook.VBProject.VBComponents(cmpn)
    With cmp.CodeModule
        linecnt = .ProcCountLines(procName, knd)
        lineStart = .ProcStartLine(procName, knd)
        lineDef = .procBodyLine(procName, knd)
        lineEnd = lineStart + linecnt - 1
        For i = 1 To linecnt - 1
            str0 = Trim(.Lines(lineEnd, 1))
            If str0 = "End Function" Or str0 = "End Sub" Or str0 = "End Property" Then
                Exit For
            Else
                lineEnd = lineEnd - 1
            End If
        Next i
        defcnt = 0
        Do While lineDef + defcnt < lineEnd
            strLine = Trim(.Lines(lineDef + defcnt, 1))
            defcnt = defcnt + 1
            If Not strLine Like "* _" Then
                Exit Do
            End If
        Loop
        lineContent = lineDef + defcnt
        If tp = "get" Then
            xdef = .Lines(lineDef, lineContent - lineDef)
            xcnt = .Lines(lineContent, lineEnd - lineContent)
            xend = .Lines(lineEnd, 1)
            disposeProc = Array(xdef, xcnt, xend)
            Exit Function
        End If
        On Error Resume Next
        If tp = "del" Or tp = "replace" Then
            Call .DeleteLines(lineContent, lineEnd - lineContent)
            If tp = "replace" Then
                Call .InsertLines(lineContent, sCode)
            End If
        End If
        On Error GoTo 0
    End With
    Set cmp = Nothing
End Function

Function getNoOptionLine(modn)
    Dim ret, i
    ret = 0
    Set cmp = ActiveWorkbook.VBProject.VBComponents(modn)
    With cmp.CodeModule
        dcln = .CountOfDeclarationLines
        For i = 1 To decln
            sLine = Trim(.Lines(i, 1))
            If sLine Like "Option " Then
                ret = ret + 1
            Else
                Exit For
            End If
        Next i
    End With
    getNoOptionLine = ret
End Function

Sub mkInterFace(ifcn As String, ParamArray ArgClsns())
    clsns = ArgClsns
    clsn = clsns(0)
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(clsn))
    Set cmps = ActiveWorkbook.VBProject.VBComponents
    Set sCmp = cmps(clsn)
    Set tCmp = cmps.Add(2)
    tCmp.Name = ifcn
    Call cpCode(clsn, ifcn, "all")
    fncs = getModProcs(ifcn)
    For Each fnc In fncs(0).keys
        Call disposeProc("del", ifcn, fnc)
    Next
    For Each prp In fncs(1).keys
        For Each knd In fncs(1)(prp)
            Call disposeProc("del", ifcn, prp, knd)
        Next
    Next
End Sub

Sub mkSubClass(sclsn As String, ifcn As String, ParamArray ArgClsns())
    clsns = ArgClsns
    clsn = clsns(0)
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(clsn))
    Set cmps = ActiveWorkbook.VBProject.VBComponents
    Set sCmp = cmps(clsn)
    Set tCmp = cmps.Add(2)
    tCmp.Name = sclsn
    Call cpCode(clsn, sclsn, "all")
    fncs = getModProcs(ifcn)
End Sub

Sub cpCode(smodn, tmodn, Optional part = "all")
    Set sCmp = ActiveWorkbook.VBProject.VBComponents(smodn)
    With sCmp.CodeModule
        Select Case LCase(part)
            Case "all"
                sCode = .Lines(1, .CountOfLines)
            Case "dcl"
                sCode = .Lines(1, .CountOfDeclarationLines)
            Case "prc"
                sCode = .Lines(.CountOfDeclarationLines + 1, .CountOfLines)
            Case Else
        End Select
    End With
    Set tCmp = ActiveWorkbook.VBProject.VBComponents(tmodn)
    If sCode <> "" Then
        tCmp.CodeModule.AddFromString sCode
    End If
    Set sCmp = Nothing
    Set tCmp = Nothing
End Sub

Function getModProcs(modn As String)
    bn = ActiveWorkbook.Name
    Dim procName
    Dim procLineNum  As Long
    Dim linecnt   As Long
    Dim fncDic
    Dim prpDic
    Set fncDic = CreateObject("Scripting.Dictionary")
    Set prpDic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set cmp = ActiveWorkbook.VBProject.VBComponents(modn)
    With cmp.CodeModule
        If .CountOfLines > 0 Then
            procName = ""
            For linecnt = .CountOfDeclarationLines + 1 To .CountOfLines
                If procName <> .ProcOfLine(linecnt, 0) Then
                    procName = .ProcOfLine(linecnt, 0)
                    procLineNum = tryToGetProcLineNum(cmp, procName, 0)
                    If procLineNum <> 0 Then
                        Call fncDic.Add(procName, 0)
                    Else
                        If Not prpDic.exists(procName) Then
                            Call prpDic.Add(procName, New Collection)
                            For knd = 1 To 3
                                procLineNum = tryToGetProcLineNum(cmp, procName, knd)
                                If procLineNum <> 0 Then
                                    prpDic(procName).Add knd
                                End If
                            Next knd
                        End If
                    End If
                End If
            Next linecnt
        End If
    End With
    getModProcs = Array(fncDic, prpDic)
    Set cmp = Nothing
End Function

Private Function tryToGetProcLineNum(cmp, procName, Optional knd = 0)
    On Error Resume Next
    ret = 0
    ret = cmp.CodeModule.ProcCountLines(procName, knd)
    tryToGetProcLineNum = ret
End Function

Private Function lenAry(ary)
    lenAry = UBound(ary) - LBound(ary) + 1
End Function

Private Function defaultInterfaceName(clsn As String)
    n = InStr(clsn, "_")
    If n > 0 Then
        ifc = Left(clsn, n - 1)
    Else
        ifc = "I" & clsn
    End If
    defaultInterfaceName = ifc
End Function

Function isLexicallyProc(sLine, pos, n)
    Dim n0, c1, c2
    Dim ret
    n0 = Len(sLine)
    ret = True
    If pos + n > n0 Then ret = False
    If n > 1 Then
        c1 = Mid(lStr, n - 1, 1)
        If c1 <> " " And c1 <> "(" Then
            ret = False
        End If
    End If
    pos2 = pos + n - 1
    If pos2 < n0 Then
        c2 = Mid(lStr, pos2, 1)
        If c2 <> " " And c2 <> "," And c2 <> "(" And c2 <> ")" Then
            ret = False
        End If
    End If
    isLexicallyProc = ret
End Function

Function addIfcPreFix(sLine, ifc, pos)
    Dim ret
    If pos = 1 Then
        ret = ifc & "_" & sLine
    Else
        ret = Left(sLine, pos - 1) & ifc & "_" & Right(sLine, Len(sLine) - pos + 1)
    End If
    addIfcPreFix = ret
End Function

Function delIfcPreFix(sLine, ifc)
    Dim ret
    ret = Replace(sLine, ifc & "_", "")
    ret = ifc & "_" & sLine
    delIfcPreFix = ret
End Function

Function mkPrpStatement(x, tp, symbol)
    symbol = LCase(symbol)
    Dim sc, gls, o, v
    Dim ret(1 To 3)
    sc = IIf(symbol Like "*_", "Public", "Private")
    gls = UCase(Left(symbol, 1)) & "et"
    o = IIf(gls = "Set" Or InStr(symbol, "o") > 0, "Set ", "")
    v = IIf(gls = "Let" Or (gls = "Set" And InStr(symbol, "v") > 0), "ByVal ", "")
    tpp = IIf(tp = "", "", " As " & tp)
    tmp = sc & " Property " & gls & " " & x
    If gls = "Get" Then
        tmp = tmp & "()" & tpp
    Else
        tmp = tmp & "(" & v & x & "_" & tpp & ")"
    End If
    ret(1) = tmp
    tmp = o
    If gls = "Get" Then
        tmp = tmp & x & " = m_" & x
    Else
        tmp = tmp & "m_" & x & " = " & x & "_"
    End If
    ret(2) = tmp
    ret(3) = "End Property"
    mkPrpStatement = Join(ret, vbCrLf)
End Function

Sub mkPrep()
    Dim cmp
    Dim sLine
    Dim i, n
    Set cmp = Application.VBE.SelectedVBComponent
    Debug.Print cmp.Name
    With cmp.CodeModule
        For i = 1 To .CountOfDeclarationLines
            sLine = .Lines(i, 1)
            n = InStr(sLine, "'")
            If n = 0 Then GoTo endfor
            ary1 = Split(Trim(Left(sLine, n - 1)))
            ary2 = Split(Trim(Right(sLine, Len(sLine) - n)), ",")
            If lenAry(ary1) <> 2 And lenAry(ary1) <> 4 Then GoTo endfor
            If lenAry(ary1) = 4 And Trim(ary1(2)) <> "As" Then GoTo endfor
            s1 = Trim(ary1(0))
            s2 = Trim(ary1(1))
            If s1 <> "Dim" And s1 = "Private" <> s1 = "Public" Then GoTo endfor
            If lenAry(ary1) = 2 Then
                s4 = ""
            Else
                s4 = Trim(ary1(3))
            End If
            If Left(s2, 2) = "m_" Then s2 = Right(s2, Len(s2) - 2)
            For Each elm In ary2
                s = Trim(elm)
                If s <> "" Then
                    .AddFromString (mkPrp(s2, s4, s))
                End If
            Next elm
endfor:
        Next i
    End With
End Sub
