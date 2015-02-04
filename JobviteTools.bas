Attribute VB_Name = "JobviteTools"
'------------------------------------------------------------------------------
' @module JobviteTools
' @desc Provides tools to export from jobvite into a DB
' @version 0.4
' @author Zach Risher
'------------------------------------------------------------------------------
Option Explicit
Const cModulename = "JobviteTools"

Const cCntrlFileName = "interview debrief tool.xlsm"
Const cCntrlFilePath = "C:\path\to\file\" & _
    "rest\of\long\path"

Const cDBFileName = "interview debrief data play.xlsm"
Const cDBFilePath = "C:\path\to\file\" & _
    "rest\of\long\path"
    
Const cJVEVAL1_FileName = "JVEVAL1.xlsx"
Const cJVEVAL1_FilePath = "C:\path\to\file\" & _
    "rest\of\long\path"

Const INPUT_TYPE_JVEVAL1 = 1
Const INPUT_JVEVAL1_COL_FIRST = 1
Const INPUT_JVEVAL1_COL_LAST = 2
Dim dicJVEVAL_AssessText As Dictionary

Const cJVELEMENT_TYPE_INTERVIEW = 1
Const cINTERVIEW_COL_FIRST = 1
Const cINTERVIEW_COL_INTERVIEWER = 1
Const cINTERVIEW_COL_SUBMITTER = 2
Const cINTERVIEW_COL_SUBMISSIONDATE = 3
Const cINTERVIEW_COL_CANDIDATE = 4
Const cINTERVIEW_COL_REQUISITION = 5
Const cINTERVIEW_COL_ASSESSMENTS = 6
Const cINTERVIEW_COL_LAST = 6



Const cJVELEMENT_TYPE_ASSESSMENT = 2
Const cASSESSMENT_COL_FIRST = 1
Const cASSESSMENT_COL_TOPIC = 1
Const cASSESSMENT_COL_RATING = 2
Const cASSESSMENT_COL_EXPLANATION = 3
Const cASSESSMENT_COL_LAST = 3


Const DB_NAMES_INTERVIEWS = "interviews"
Const DB_NAMES_INTERVIEWS_USER = "user_id_Users"
Const DB_NAMES_INTERVIEWS_CANDIDATE = "candidate_id_Candidates"

Const DB_NAMES_USERS = "interviews"
Const DB_NAMES_USERS_ID = "user_id"
Const DB_NAMES_USERS_NAME = "name"

Const DB_NAMES_CANDIDATES = "candidates"
Const DB_NAMES_CANDIDATES_ID = "candidate_id"
Const DB_NAMES_CANDIDATES_NAME = "name"







'------------------------------------------------------------------------------
' @sub ExportToDB
' @desc Takes data from excel sheets, runs corresponding DB commands, and
'   writes any info/errors to that worksheet
' @scope Public
'------------------------------------------------------------------------------
Public Sub ExportToDB(Optional SourceWorkSheet As Worksheet)
'------------------@EHSBS Standard Error Handling start block------------------
 Dim cSubName As String
 cSubName = "ExportToDB"
 If gcHandleErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0
 'Debug.Print cModulename & "." & cSubName & " Commencing"
'------------------@EHSBE Standard Error Handling start block------------------


    ' can't use ActiveSheet to set optional variable in proc def
    If SourceWorkSheet Is Nothing Then
        Set SourceWorkSheet = ActiveSheet
    End If

    ' read data from ws
    Dim varReadArr As Variant
    varReadArr = WorksheetToArray(SourceWorkSheet, , INPUT_JVEVAL1_COL_FIRST, , INPUT_JVEVAL1_COL_LAST)

    ' turn array into commands
    Dim varCommands As Variant
    varCommands = JVReadArrayToDBCommands(varReadArr)
    
    'process DB commands
    Dim xldbDB As New ExcelDB
    Dim varDBReturns As Variant
    Call xldbDB.Load(cDBFileName, cDBFilePath)
    varDBReturns = xldbDB.Execute(varCommands)
    
    'return DB results to worksheet
    Dim varReturnArrays As Variant
    Dim varCurRet2 As Variant
    Dim varCurRet As Variant
    Dim intCurRet As Integer
    Dim intFinalRet As Integer
    intCurRet = 1
    intFinalRet = 1 + 27 * UBound(varCommands(1)) + 1
    ReDim varReturnArrays(1 To 100)
    
    varReturnArrays(intCurRet) = varReadArr
    intCurRet = intCurRet + 1
    
    For Each varCurRet2 In varCommands
        varReturnArrays(intCurRet) = varCurRet2(1)
        intCurRet = intCurRet + 1
          
        If Not TypeName(varCurRet2(1)(6)) = "String" Then
            For Each varCurRet In varCurRet2(1)(6)
            varReturnArrays(intCurRet) = varCurRet
            intCurRet = intCurRet + 1
            Next
        End If
    Next
    

    varReturnArrays(intCurRet) = varDBReturns
    intCurRet = intCurRet + 1
    
    Dim wshDest2 As Worksheet
    Set wshDest2 = Workbooks(cCntrlFileName).Worksheets("Dest2")
    If wshDest2 Is Nothing Then
        Set wshDest2 = Workbooks(cCntrlFileName).Worksheets.Add
        wshDest2.Name = "Dest2"
    End If
    Call TwoDArraysToWorksheet(varReturnArrays, wshDest2)


'-------------------@EHEBS Standard Error Handling end block-------------------
PROC_EXIT:
  'Debug.Print cModulename & "." & cSubName & " Done"
  Exit Sub

PROC_ERR:
  MsgBox "Error " & Err.Number & " from " & Err.Source & " : " & Err.Description
  Resume PROC_EXIT
'-------------------@EHEBE Standard Error Handling end block-------------------
End Sub


'------------------------------------------------------------------------------
' @Function JVReadArrayToDBCommands
' @desc Takes an array of data from jobvite, parses it for recognized elements,
'   and returns corresponding DB commands
' @scope Private
'------------------------------------------------------------------------------
Function JVReadArrayToDBCommands(ReadArray As Variant) As Variant

    Dim intReadRow As Integer
    Dim strCurCell As String
    
    Dim lngLastFoundStartRow As Long
    Dim intLastFoundInputType As Integer
    
    Dim varSubReadArray As Variant
    
    Dim varCommands As Variant
    ReDim varCommands(1 To UBound(ReadArray, 1))
    Dim lngCommandCount As Long
    
    For intReadRow = 1 To UBound(ReadArray, 1)
        
        'check if this cell contains the start of an element
        strCurCell = CStr(ReadArray(intReadRow, 1))
        Select Case Left$(strCurCell, 13)
            
            Case "Evaluation by" 'Found an element - JVEVAL1
            
                ' process any pending reads, of which we've now found the end
                If lngLastFoundStartRow > 0 Then
                    varSubReadArray = GetSub2DArray(ReadArray, _
                                        lngLastFoundStartRow, _
                                        LBound(ReadArray, 2), _
                                        intReadRow - 1, _
                                        UBound(ReadArray, 2) _
                                        )
                    lngCommandCount = lngCommandCount + 1
'                    Debug.Print "setting varcommands(" & lngCommandCount&; ") as" & JVSubArrayToDBCommands( _
'                                                    varSubReadArray, _
'                                                    intLastFoundInputType _
'                                                    )
                    varCommands(lngCommandCount) = JVSubArrayToDBCommands( _
                                                    varSubReadArray, _
                                                    intLastFoundInputType _
                                                    )

                End If
                
                ' mark current eval for processing, as we've found its start
                lngLastFoundStartRow = intReadRow
                intLastFoundInputType = INPUT_TYPE_JVEVAL1
            
            Case Else 'not an element start

        End Select
        
    Next
    
    ' process any pending reads, of which we've now found the end
    If lngLastFoundStartRow > 0 Then
        varSubReadArray = GetSub2DArray(ReadArray, _
                            lngLastFoundStartRow, _
                            LBound(ReadArray, 2), _
                            intReadRow - 1, _
                            UBound(ReadArray, 2) _
                            )
        lngCommandCount = lngCommandCount + 1
        varCommands(lngCommandCount) = JVSubArrayToDBCommands( _
                                        varSubReadArray, _
                                        intLastFoundInputType _
                                        )
'                            Debug.Print "setting varcommands(" & lngCommandCount&; ") as" & JVSubArrayToDBCommands( _
'                                                    varSubReadArray, _
'                                                    intLastFoundInputType _
'                                                    )
    End If
    
    ReDim Preserve varCommands(1 To lngCommandCount)
    JVReadArrayToDBCommands = varCommands

    
End Function


'------------------------------------------------------------------------------
' @Function JVSubArrayToDBCommands
' @desc Takes an array of a single evaluation element from Jobvite
'   and returns corresponding DB commands
' @scope Private
'------------------------------------------------------------------------------
Function JVSubArrayToDBCommands(SubArray As Variant, InputType As Integer) As Variant


    Dim varReturnCommands As Variant
    ReDim varReturnCommands(1 To UBound(SubArray, 1))
    Dim lngCommandCount As Long
    
    Select Case InputType
        Case INPUT_TYPE_JVEVAL1
        
            lngCommandCount = lngCommandCount + 1
            varReturnCommands(lngCommandCount) = JVSubArrayToElements_JVEVAL1(SubArray)
            
        Case Else
    End Select
    
    ReDim Preserve varReturnCommands(1 To lngCommandCount)
    JVSubArrayToDBCommands = varReturnCommands
    
End Function


'------------------------------------------------------------------------------
' @Function
' @desc
' @scope
'------------------------------------------------------------------------------
Function JVSubArrayToElements_JVEVAL1(SubArray As Variant) As Variant

    Dim lngCurArrStartRow As Long
    lngCurArrStartRow = LBound(SubArray, 1)
    Dim lngCurArrEndRow As Long
    lngCurArrEndRow = UBound(SubArray, 1)
    Dim lngCurArrHeight As Long
    lngCurArrHeight = lngCurArrEndRow - lngCurArrStartRow + 1
    
    Dim varInterview As Variant
    ReDim varInterview(cINTERVIEW_COL_FIRST To cINTERVIEW_COL_LAST)
    
    Dim varAssessments As Variant
    ReDim varAssessments(1 To lngCurArrHeight)
    Dim lngAssessCount As Long
    
    If dicJVEVAL_AssessText Is Nothing Then
        Set dicJVEVAL_AssessText = LoadEvalDic(OpenWorkbook(cJVEVAL1_FileName, cJVEVAL1_FilePath).Worksheets("form text"), INPUT_TYPE_JVEVAL1)
    End If

    Dim lngCurRow As Long
    lngCurRow = lngCurArrStartRow
    
    Dim varCurCell As Variant
    Dim varCurCellRight As Variant
    Dim varCurCellDown As Variant
    Dim varCurAssess As Variant
    
    'parse row 1
    varCurCell = SplitSequential(SanitizeHTML(SubArray(lngCurRow, 1)), Array("Evaluation by ", " Completed by "))
    varInterview(cINTERVIEW_COL_INTERVIEWER) = varCurCell(1)
    If UBound(varCurCell) > 1 Then varInterview(cINTERVIEW_COL_SUBMITTER) = varCurCell(2)
    'parse row 2
    lngCurRow = lngCurRow + 1
    varCurCell = SplitSequential(SanitizeHTML(SubArray(lngCurRow, 1)), Array("Submitted: "))
    varInterview(cINTERVIEW_COL_SUBMISSIONDATE) = CDate(varCurCell(1))
    'parse row 5
    lngCurRow = lngCurRow + 3
    varInterview(cINTERVIEW_COL_CANDIDATE) = SanitizeHTML(SubArray(lngCurRow, 2))
    'parse row 6
    lngCurRow = lngCurRow + 1
    varInterview(cINTERVIEW_COL_REQUISITION) = SanitizeHTML(SubArray(lngCurRow, 2))
    'parse evals
    Do While lngCurRow < lngCurArrEndRow
        varCurCell = SanitizeHTML(SubArray(lngCurRow, 1))
        If dicJVEVAL_AssessText.Exists(varCurCell) Then
            varCurCellRight = SanitizeHTML(SubArray(lngCurRow, 2))
            varCurCellDown = SanitizeHTML(SubArray(lngCurRow + 1, 1))
            ReDim varCurAssess(cASSESSMENT_COL_FIRST To cASSESSMENT_COL_LAST)
            varCurAssess(cASSESSMENT_COL_TOPIC) = dicJVEVAL_AssessText.Item(varCurCell)
            varCurAssess(cASSESSMENT_COL_RATING) = RemoveUnspecified(varCurCellRight)
            varCurAssess(cASSESSMENT_COL_EXPLANATION) = RemoveUnspecified(varCurCellDown)
            
            If Not varCurAssess(cASSESSMENT_COL_RATING) = "NULL" _
                Or Not varCurAssess(cASSESSMENT_COL_EXPLANATION) = "NULL" Then
                lngAssessCount = lngAssessCount + 1
                varAssessments(lngAssessCount) = varCurAssess
            End If
            
            lngCurRow = lngCurRow + 2
        Else
            Debug.Print "Unrecognized Crit Topic:"; varCurCell
            lngCurRow = lngCurRow + 1
        End If
    Loop
    
    'final return
    If lngAssessCount > 0 Then
    ReDim Preserve varAssessments(1 To lngAssessCount)
    varInterview(cINTERVIEW_COL_ASSESSMENTS) = varAssessments
    Else
    varInterview(cINTERVIEW_COL_ASSESSMENTS) = "NULL"
    End If
    
    JVSubArrayToElements_JVEVAL1 = varInterview
  
End Function





'------------------------------------------------------------------------------
' @Function
' @desc
' @scope
'------------------------------------------------------------------------------
Function LoadEvalDic(Worksheet As Worksheet, EvalType As Integer) As Dictionary
    Dim varReadArr As Variant
    varReadArr = WorksheetToArray(Worksheet, 2, 4, , 6)
    
    Dim dicReturnDic As New Dictionary
    
    Dim lngCurArrStartRow As Long
    lngCurArrStartRow = LBound(varReadArr, 1)
    Dim lngCurArrEndRow As Long
    lngCurArrEndRow = UBound(varReadArr, 1)
    Dim lngCurArrHeight As Long
    lngCurArrHeight = lngCurArrEndRow - lngCurArrStartRow + 1
    
    Dim lngCurRow As Long
    
    For lngCurRow = lngCurArrStartRow To lngCurArrEndRow
    If EvalType = INPUT_TYPE_JVEVAL1 Then
    Select Case varReadArr(lngCurRow, 2)
        Case "Ignore"
        Case "Assessment"
            Call AddToDic(dicReturnDic, varReadArr(lngCurRow, 1), varReadArr(lngCurRow, 3))
        Case "Else"
    End Select
    End If
    Next lngCurRow
    
    Set LoadEvalDic = dicReturnDic
End Function


'------------------------------------------------------------------------------
' @Function
' @desc
' @scope
'------------------------------------------------------------------------------
Function RemoveUnspecified(ByVal StringToCheck As String) As String
    Select Case StringToCheck
    Case "Not specified.", "", "na", "NA", "n/a", "N/A"
        RemoveUnspecified = "NULL"
    Case Else
        RemoveUnspecified = StringToCheck
    End Select
End Function

