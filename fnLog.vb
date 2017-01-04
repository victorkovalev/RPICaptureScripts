''' ===================================================================================================================
'''                                         START LOGGING FUNCTIONS
''' ===================================================================================================================
''' <summary>
'''     Logging Functions, helpfull to make shorter calls from scripts and support indenting.
'''     
'''     Sample Function:
'''         Function fnGetValueInFieldList(strClassName As String, strFieldName As String, strValue As String, sglDistance As Single) As String
'''             On Error GoTo lbl_error
'''         
'''             fnLog(CDRTypeInfo, "fnGetValueInFieldList Start strClassName [" & strClassName & "] strFieldName [" & strFieldName & "] strValue [" & strValue & "]")
'''             fnLogIndentIncrease()
'''             
'''             'DO Some work
'''         lbl_end:
'''             fnLogIndentDecrease()
'''             fnLog(CDRTypeInfo, "fnGetValueInFieldList End Returning [" & CStr(fnGetValueInFieldList) & "]")
'''             Exit Function
'''         lbl_error:
'''             fnLogIndentDecrease()
'''             fnLog(CDRTypeError, "fnGetValueInFieldList Error [" & Err.Description & "]")
'''         End Function
''' </summary>
''' <remarks>Global variables to store the level of indent with logging</remarks>
Global gLogIndent As Integer
Global gLogSuspend As Boolean

''' <summary>
'''     Function can be used to suspend verbose logging.
'''     For example when loading configuration, you can have a call to fnLogSuspend to pause logging, then resume when load is complete
''' </summary>
''' <param name="pLogSuspend">Boolean true/false</param>
Public Function fnLogSuspend(pLogSuspend As Boolean)
    On Error GoTo lbl_error
    gLogSuspend = pLogSuspend

    Exit Function
lbl_error:
    fnLog(CDRTypeError, "fnLogSuspend Error [" & Err.Description & "]")
End Function

''' <summary>
'''     Resets the log indent level to 0
''' </summary>
Public Function fnLogIndentReset()
    On Error GoTo lbl_error
    gLogIndent = 0

    Exit Function
lbl_error:
    fnLog(CDRTypeError, "fnLogIndentReset Error [" & Err.Description & "]")
End Function

''' <summary>
'''     Increases the log indent by 1 level (2 spaces)
''' </summary>
Public Function fnLogIndentIncrease()
    On Error GoTo lbl_error
    gLogIndent = gLogIndent + 1

    If gLogIndent > 20 Then
        gLogIndent = 20
    End If

    Exit Function
lbl_error:
    fnLog(CDRTypeError, "fnLogIndentIncrease Error [" & Err.Description & "]")
End Function

''' <summary>
'''     Decreases the log indent by 1 level (2 spaces)
''' </summary>
Public Function fnLogIndentDecrease()
    On Error GoTo lbl_error
    gLogIndent = gLogIndent - 1

    If gLogIndent < 0 Then
        gLogIndent = 0
    End If

    Exit Function
lbl_error:
    fnLog(CDRTypeError, "fnLogIndentDecrease Error [" & Err.Description & "]")
End Function

''' <summary>
'''     Logs the message at specific level (CDRTypeError, CDRTypeWarning, CDRTypeInfo)
''' </summary>
''' <param name="pType">CDRTypeError, CDRTypeWarning, CDRTypeInfo</param>
''' <param name="pMessage">Message to print</param>
Public Function fnLog(pType As CDRMessageType, pMessage As String)
    On Error GoTo lbl_error

    ' Do not log info messages if logging is suspended
    If gLogSuspend = True And pType = CDRMessageType.CDRTypeInfo Then
        GoTo lbl_end
    End If

    Project.LogScriptMessageEx(pType, CDRSeverityLogFileOnly, Space(gLogIndent * 2) & pMessage)
lbl_end:
    Exit Function
lbl_error:
    Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnLogIndentDecrease Error [" & Err.Description & "]")
End Function

''' ===================================================================================================================
'''                                         END LOGGING FUNCTIONS
''' ===================================================================================================================