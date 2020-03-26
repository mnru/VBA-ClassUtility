Attribute VB_Name = "testCode"
Sub test()
    strA = Join(Split("g,f,e,a,b", ","), vbCrLf)
    Call disposeProc("replace", "Logwriter", "myoutput", , strA)
End Sub

Sub testCode0()
    Call mkInterFace(, "LogWriter")
End Sub

Sub testcode1()
    Call mkSubClass("SpecialLogWriter", , "LogWriter")
End Sub

Sub testprp()
    Debug.Print mkPrpStatement("abc", "Long", "g")
    Debug.Print mkPrpStatement("abc", "Long", "g_")
    Debug.Print mkPrpStatement("abc", "Long", "s_")
    Debug.Print mkPrpStatement("abc", "Long", "s")
    Debug.Print mkPrpStatement("abc", "Long", "sov")
    Debug.Print mkPrpStatement("abc", "Long", "l")
    Debug.Print mkPrpStatement("abc", "", "g")
    Debug.Print mkPrpStatement("abc", "", "g_")
    Debug.Print mkPrpStatement("abc", "", "s_")
    Debug.Print mkPrpStatement("abc", "", "s")
    Debug.Print mkPrpStatement("abc", "", "sov")
    Debug.Print mkPrpStatement("abc", "", "l")
End Sub
