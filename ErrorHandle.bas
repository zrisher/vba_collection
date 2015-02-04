Attribute VB_Name = "ErrorHandle"
'*** Provides default error handling
'*** Zach Risher 8/5/2012

Option Explicit



'can use conditional compliation for debug prints, very cool
'#Const blncDevelopMode = True
'
'Sub Xyz()
'  #If developMode Then
'    Debug.Print "something"
'  #End If
'End Sub

Function LogError(ByVal lngErrNumber As Long, ByVal strErrDescription As String, _
    strCallingProc As String, Optional vParameters, Optional bShowUser As Boolean = True) As Boolean
    ' Purpose: Generic error handler.
    ' Arguments: lngErrNumber - value of Err.Number
    ' strErrDescription - value of Err.Description
    ' strCallingProc - name of sub|function that generated the error.
    ' vParameters - optional string: List of parameters to record.
    ' bShowUser - optional boolean: If False, suppresses display.
    ' Adapted from original version of author Allen Browne, allen@allenbrowne.com
        
On Error GoTo Err_LogError


Dim strMsg As String      ' String for display in MsgBox


Select Case lngErrNumber
Case 0
    Debug.Print strCallingProc & " called error 0."
Case 2501                ' Cancelled
    'Do nothing.
Case 3314, 2101, 2115    ' Can't save.
    If bShowUser Then
        strMsg = "Record cannot be saved at this time." & vbCrLf & _
            "Complete the entry, or press <Esc> to undo."
        MsgBox strMsg, vbExclamation, strCallingProc
    End If
    
Case 5151 'word can't find template
        If bShowUser Then
        strMsg = "Cannot locate template file." & vbCrLf & _
            "Please ensure it is located in the same folder as this spreadsheet."
        MsgBox strMsg, vbExclamation, strCallingProc
    End If
    
Case Else
    If bShowUser Then
        strMsg = "Error " & lngErrNumber & ": " & strErrDescription
        MsgBox strMsg, vbExclamation, strCallingProc
    End If
    
    'error logging
    'todo: code to write to errors sheet
'    Set rst = CurrentDb.OpenRecordset("tLogError", , dbAppendOnly)
'    rst.AddNew
'        rst![ErrNumber] = lngErrNumber
'        rst![ErrDescription] = Left$(strErrDescription, 255)
'        rst![ErrDate] = Now()
'        rst![CallingProc] = strCallingProc
'        rst![UserName] = CurrentUser()
'        rst![ShowUser] = bShowUser
'        If Not IsMissing(vParameters) Then
'            rst![Parameters] = Left(vParameters, 255)
'        End If
End Select
    
    LogError = True
    



Exit_LogError:
    Set rst = Nothing
    Exit Function

Err_LogError:
    strMsg = "An unexpected situation arose in your program." & vbCrLf & _
        "Please write down the following details:" & vbCrLf & vbCrLf & _
        "Calling Proc: " & strCallingProc & vbCrLf & _
        "Error Number " & lngErrNumber & vbCrLf & strErrDescription & vbCrLf & vbCrLf & _
        "Unable to record because Error " & Err.Number & vbCrLf & Err.Description
    MsgBox strMsg, vbCritical, "LogError()"
    Resume Exit_LogError
    
    
End Function

