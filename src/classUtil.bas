Attribute VB_Name = "classUtil"
Sub addPrp(nm, tp, Optional isset = False)
End Sub

Sub insertIfcPrefix(ifcn)
End Sub

Function prcReplace(sLine, prc, ifc)
    Dim ret
    Dim i, np, ni, nl, pos
    np = Len(prc)
    ni = Len(ifc)
    nl = Len(sLine)
    ret = sLine
    pos = 1
    Do
        If pos > nl Then Exit Do
        i = InStr(pos, ret, prc)
        If i = 0 Then Exit Do
        If isIfcProc(ret, i, np) Then
            ret = preFixIfc(ret, ifc, i)
            pos = i + ni + 2
        Else
            pos = i + np
        End If
    Loop
    prcReplace = ret
End Function

Function isIfcProc(sLine, pos, n)
    Dim n0, c1, c2
    Dim ret
    n0 = Len(sLine)
    ret = True
    If pos + n > n0 Then ret = False
    If n > 1 Then
        c1 = Mid(lStr, n - 1, 1)
        If c1 <> " " And c1 <> "(" And c1 <> ")" Then
            ret = False
        End If
    End If
    pos2 = pos + n - 1
    If pos2 < n0 Then
        c2 = Mid(lStr, pos2, 1)
        If c2 <> " " And c2 <> "(" And c2 <> ")" Then
            ret = False
        End If
    End If
    isIfcProc = ret
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

Function ridIfcPreFix(sLine, ifc)
    Dim ret
    ret = Replace(sLine, ifc & "_", "")
    ret = ifc & "_" & sLine
    ridIfcPreFix = ret
End Function

Sub testcode()
    Call mkInterFace(, "LogWriter")
End Sub

Sub testcode1()
    Call mkSubClass("SpecialLogWriter", , "LogWriter")
End Sub

Sub mkInterFace(Optional ifcn As String = "", Optional clsn As String)
    If ifcn = "" Then ifcn = mkInterfaceName(clsn)
    Set cmps = ActiveWorkbook.VBProject.VBComponents
    Set sCmp = cmps(clsn)
    Set tCmp = cmps.Add(2)
    tCmp.Name = ifcn
    Call cpCode(sCmp, tCmp, "all")
    fncs = getModProcs(ifcn)
    For Each fnc In fncs(0)
        Call delProcCode(fnc, 0, ifcn)
    Next
    For Each prp In fncs(1).keys
        For Each knd In fncs(1)(prp)
            Call delProcCode(prp, knd, ifcn)
        Next
    Next
End Sub

Sub mkSubClass(sclsn As String, Optional ifcn As String = "", Optional clsn As String)
    If ifcn = "" Then ifcn = mkInterfaceName(clsn)
    Set cmps = ActiveWorkbook.VBProject.VBComponents
    Set sCmp = cmps(clsn)
    Set tCmp = cmps.Add(2)
    tCmp.Name = sclsn
    Call cpCode(sCmp, tCmp, "all")
    fncs = getModProcs(ifcn)
End Sub

Function mkInterfaceName(clsn As String)
    n = InStr(clsn, "_")
    If n > 0 Then
        ifc = Left(clsn, n - 1)
    Else
        ifc = "I" & clsn
    End If
    mkInterfaceName = ifc
End Function

Sub cpCode(sCmp, tCmp, Optional part = "all")
    With sCmp.CodeModule
        Select Case LCase(part)
            Case "all"
                scode = .Lines(1, .CountOfLines)
            Case "dcl"
                scode = .Lines(1, .CountOfDeclarationLines)
            Case "prc"
                scode = .Lines(.CountOfDeclarationLines + 1, .CountOfLines)
            Case Else
        End Select
    End With
    If scode <> "" Then
        tCmp.CodeModule.AddFromString scode
    End If
End Sub

Sub delProcCode(procName, knd, cmpn)
    Set cmp = ActiveWorkbook.VBProject.VBComponents(cmpn)
    With cmp.CodeModule
        lineDef = .procBodyLine(procName, knd)
        lineEnd = .ProcCountLines(procName, knd) + .ProcStartLine(procName, knd)
        cnt = 0
        Do While lineDef + cnt < lineEnd
            strLine = Trim(.Lines(lineDef + cnt, 1))
            cnt = cnt + 1
            If Not strLine Like "* _" Then
                Exit Do
            End If
        Loop
        If lineEnd - lineDef - cnt - 1 > 0 Then
            Call .DeleteLines(lineDef + cnt, lineEnd - lineDef - cnt - 1)
        End If
    End With
    Set cmp = Nothing
End Sub

Function getModProcs(modn As String)
    bn = ActiveWorkbook.Name
    Dim procName
    Dim procLineNum       As Long
    Dim lineCnt           As Long
    Dim fncClc
    Dim prpDic
    Set fncClc = New Collection
    Set prpDic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set cmp = ActiveWorkbook.VBProject.VBComponents(modn)
    With cmp.CodeModule
        If .CountOfLines > 0 Then
            mdlType = getModType(cmp)
            procName = ""
            For lineCnt = .CountOfDeclarationLines + 1 To .CountOfLines
                If procName <> .ProcOfLine(lineCnt, 0) Then
                    procName = .ProcOfLine(lineCnt, 0)
                    procLineNum = tryToGetProcLineNum(cmp, procName, 0)
                    If procLineNum <> 0 Then
                        fncClc.Add (procName)
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
            Next lineCnt
        End If
    End With
    getModProcs = Array(fncClc, prpDic)
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
