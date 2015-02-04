Attribute VB_Name = "Project"
Public Const gcHandleErrors As Boolean = False

Public Const gcWORKBOOK_CONTROL_FILENAME = "interview debrief tool.xlsm"
Public Const gcWORKBOOK_CONTROL_FILEPATH = "C:\path\to\file\" & _
    "rest\of\long\path"
Public Const gcWORKBOOK_DB_FILENAME = "interview debrief data.xlsm"
Public Const gcWORKBOOK_DB_FILEPATH = "C:\path\to\file\" & _
    "rest\of\long\path"

Public Const gcCOMPAT_MAXROWS = gcXL2003_MAXROWS
Public Const gcCOMPAT_MAXCOLS = gcXL2003_MAXCOLS


Option Explicit
Const cModulename = "ModuleName" 'for error handling




'------------------------------------------------------------------------------
' @module
' @desc
' @version
' @authors
' @created
' @last-updated
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' @sub
' @desc
'------------------------------------------------------------------------------

'------------------@EHSBS Standard Error Handling start block------------------
' Dim cSubName As String
' cSubName = "SubName"
' If gcHandleErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0
' 'Debug.Print cModulename & "." & cSubName & " Commencing"
'------------------@EHSBE Standard Error Handling start block------------------


    ' Subcomment
    '--------------------------------------------------------------------------
    
            ' Subcomment2
            '------------------------------------------------------------------


'-------------------@EHEBS Standard Error Handling end block-------------------
'PROC_EXIT:
'  'Debug.Print cModulename & "." & cSubName & " Done"
'  Exit Function
'
'PROC_ERR:
'  MsgBox "Error " & Err.Number & " from " & Err.Source & " : " & Err.Description
'  Resume PROC_EXIT
'-------------------@EHEBE Standard Error Handling end block-------------------

