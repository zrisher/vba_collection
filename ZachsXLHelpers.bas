Attribute VB_Name = "ZachsXLHelpers"
'------------------------------------------------------------------------------
' @module Zach's XL Helpers
' @desc A bunch of functions that excel should come with, but were fun to
' program anyway
' @version 0.4
' @created 3-21-2013
' @last-updated 3-26-2013
'------------------------------------------------------------------------------
Option Explicit
Const cModulename = "ZachsVBAHelpers"

Public Const gcXL2003_MAXROWS = 65536
Public Const gcXL2007_MAXROWS = 1048576
Public Const gcXL2003_MAXCOLS = 256
Public Const gcXL2007_MAXCOLS = 16384

Public Const gcVBA_MAXARRAYDIM = 6000

' Reference code
'------------------------------------------------------------------------------

'MonthlyWB = Application.GetOpenFilename( _
'FileFilter:="Microsoft Excel Workbooks, *.xls; *.xlsx", Title:="Open Workbook")
'
'Workbooks(FileName).Close
'Close Without saving changes
'Workbooks.close(filepathasstring,false)
'
'Close and Save Changes
'Workbooks.close(filepathasstring,true)
'
'Set ID = ThisWorkbook.VBProject.References
'ID.AddFromGuid "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}", 2, 5
'
'Chr(32)
'Asc(c)
'
'Dim clipboard As MSForms.DataObject
'Set clipboard = New MSForms.DataObject
'clipboard.SetText outputText
'clipboard.PutInClipboard
'
'Split(outputstep, ")", 2)(0)


'------------------------------------------------------------------------------
' @function SheetExists
' @desc Does the sheet with provided name exist in provided workbook?
'------------------------------------------------------------------------------
Public Function SheetExists(SheetName As String, _
                            Workbook As Workbook) As Boolean
  SheetExists = False
  Dim wshCur As Worksheet
  
  For Each wshCur In Workbook.Worksheets
    If SheetName = wshCur.Name Then
      SheetExists = True
      GoTo PROC_END
    End If
  Next wshCur
  
PROC_END:
End Function


'------------------------------------------------------------------------------
' @function WorksheetToArray
' @desc Returns an array of the data within provided workbook and rows/cols
' @scope Public
'------------------------------------------------------------------------------
Public Function WorksheetToArray(WorksheetToRead As Worksheet, _
                                    Optional FirstRow As Long = 1, _
                                    Optional FirstCol As Integer = 1, _
                                    Optional LastRow As Long, _
                                    Optional LastCol As Integer) _
                                    As Variant()
                                    
    If LastRow = 0 Then 'uninitialized
        LastRow = WorksheetToRead.Cells(gcCOMPAT_MAXROWS, FirstCol).End(xlUp).Row
    End If
    
    If LastCol = 0 Then 'uninitialized
        LastCol = WorksheetToRead.Cells(gcCOMPAT_MAXCOLS, FirstCol).End(xlToLeft).Row
    End If
    
    With WorksheetToRead
        WorksheetToArray = .Range(.Cells(FirstRow, FirstCol), .Cells(LastRow, LastCol))
    End With
    
End Function


'------------------------------------------------------------------------------
' @sub TwoDArraysToWorksheet
' @desc Writes arrays to a worksheet (usually for debugging)
' @scope Public
'------------------------------------------------------------------------------
Public Sub TwoDArraysToWorksheet(OneDArrayOf2DArrays As Variant, _
                                    WorksheetToWrite As Worksheet, _
                                    Optional BorderChar As String = "*")
    WorksheetToWrite.UsedRange.Clear
                                    
    Dim varCurArr As Variant
    Dim lngCurArrCount As Long
    Dim lngCurRow As Long
    Dim intCurCol As Integer
    
    Dim intCurArrDim As Integer
    Dim lngCurArrStartRow As Long
    Dim lngCurArrEndRow As Long
    Dim intCurArrStartCol As Integer
    Dim intCurArrEndCol As Integer
    Dim lngCurArrHeight As Long
    Dim intCurArrWidth As Integer
    
    Dim rngWrite As Range
    
    lngCurRow = 1
    For Each varCurArr In OneDArrayOf2DArrays
        intCurArrDim = GetNumArrayDim(varCurArr)
        If intCurArrDim > 0 Then
            With WorksheetToWrite
            
                If intCurArrDim > 1 Then
                    lngCurArrStartRow = LBound(varCurArr, 1)
                    lngCurArrEndRow = UBound(varCurArr, 1)
                    intCurArrStartCol = LBound(varCurArr, 2)
                    intCurArrEndCol = UBound(varCurArr, 2)
                Else
                    lngCurArrStartRow = 1
                    lngCurArrEndRow = 1
                    intCurArrStartCol = LBound(varCurArr, 1)
                    intCurArrEndCol = UBound(varCurArr, 1)
                End If
                lngCurArrHeight = lngCurArrEndRow - lngCurArrStartRow + 1
                intCurArrWidth = intCurArrEndCol - intCurArrStartCol + 1

                    
                'write descr line
                lngCurArrCount = lngCurArrCount + 1
                .Cells(lngCurRow, 1) = _
                    "Array #" & lngCurArrCount & _
                    ", Rows " & lngCurArrStartRow & " to " & lngCurArrEndRow & _
                    ", Cols " & intCurArrStartCol & " to " & intCurArrEndCol
                lngCurRow = lngCurRow + 1
                
                'write border
                .Cells(lngCurRow, 1).Resize(lngCurArrHeight + 2, _
                                            intCurArrWidth + 2) _
                                            = BorderChar
                
                'write array content
                Set rngWrite = .Cells(lngCurRow + 1, 2).Resize(lngCurArrHeight, _
                                                intCurArrWidth)
                rngWrite.NumberFormat = "@"

                On Error GoTo BadWriteValues
                rngWrite = varCurArr
                On Error GoTo 0
                GoTo FinishedWriting
                
BadWriteValues:
                If intCurArrDim = 1 Then
                    For intCurCol = intCurArrStartCol To intCurArrEndCol
                        If IsArray(varCurArr(intCurCol)) Then
                            varCurArr(intCurCol) = "Array"
                        Else
                            varCurArr(intCurCol) = CStr(varCurArr(intCurCol))
                        End If
                    Next
                Else
                    For lngCurRow = lngCurArrStartRow To lngCurArrEndRow
                        For intCurCol = intCurArrStartCol To intCurArrEndCol
                            If IsArray(varCurArr(lngCurRow, intCurCol)) Then
                                varCurArr(lngCurRow, intCurCol) = "Array"
                            Else
                                varCurArr(lngCurRow, intCurCol) = CStr(varCurArr(lngCurRow, intCurCol))
                            End If
                        Next
                    Next
                End If

                rngWrite = varCurArr
                Resume FinishedWriting
                
FinishedWriting:
                lngCurRow = lngCurRow + lngCurArrHeight + 3
    
            End With
        Else
            WorksheetToWrite.Cells(lngCurRow, 1) = "Non-array variant found, skipping."
            lngCurRow = lngCurRow + 2
        End If
    Next
                                        
End Sub
'------------------------------------------------------------------------------
' @function OpenWorkbook
' @desc Returns an array of the data within provided workbook and rows/cols
' @scope Public
'------------------------------------------------------------------------------
Public Function OpenWorkbook(fileName As String, FilePath As String) As Workbook
    Dim wkbkReturn As Workbook
    On Error Resume Next
    Set wkbkReturn = Workbooks(fileName)
    On Error GoTo 0
    If wkbkReturn Is Nothing Then
        If Not Right(FilePath, 1) = "/" Then FilePath = FilePath & "/"
        Set wkbkReturn = Workbooks.Open(FilePath & fileName)
    End If
    Set OpenWorkbook = wkbkReturn
End Function


'------------------------------------------------------------------------------
' @function SplitSequential
' @desc Returns an array of elements of a string between delimiters
' @scope Public
'------------------------------------------------------------------------------
Public Function SplitSequential(ByVal StringToParse As String, _
                                StringDelimiters As Variant _
                                ) As String()

    Dim strCutString As String
    strCutString = StringToParse
    
    Dim lngCutStart As Long
    lngCutStart = 0
    Dim lngCutLen As Long
    
    Dim strReturns() As String
    ReDim strReturns(1 To UBound(StringDelimiters, 1) + 1)
    Dim lngReturnCount As Long
    lngReturnCount = 0
    
    ' Go through each delimiter and cut the string into array cells
    '--------------------------------------------------------------------------
    Dim varDelimiter As Variant
    For Each varDelimiter In StringDelimiters
    
        lngCutStart = InStr(1, strCutString, varDelimiter, vbBinaryCompare)
        
        If lngCutStart > 0 Then 'found the delimiter
        
            If lngCutStart > 1 Then 'there is text to left of delimiter
                lngReturnCount = lngReturnCount + 1
                strReturns(lngReturnCount) = Left(strCutString, lngCutStart - 1)
            End If
            
            'save text to right of delimiter as new string to inspect
            lngCutLen = Len(varDelimiter)
            strCutString = Right(strCutString, Len(strCutString) - (lngCutLen + lngCutStart - 1))
            
        End If
        
    Next
    
    'if there's any remaining string, add to returns
    If Len(strCutString) > 1 Then
    lngReturnCount = lngReturnCount + 1
    strReturns(lngReturnCount) = strCutString
    End If
    
    'remove any empty spaces in the return array
    ReDim Preserve strReturns(1 To lngReturnCount)
    SplitSequential = strReturns

End Function


'------------------------------------------------------------------------------
' @function SanitizeHTML
' @desc Replaces HTML chars that VBA doesn't recognize with things it can.
' This probably has a long way to go.
' @scope Public
'------------------------------------------------------------------------------
Public Function SanitizeHTML(ByVal HTML As String) As String
    SanitizeHTML = Replace(HTML, Chr(160), Chr(32))
End Function


'------------------------------------------------------------------------------
' @function AddToDic
' @desc Checks if an element exists in a dictionary before adding
' originally had customing adding function for strings when they existed,
' might want to extend to vars
' @scope Public
'------------------------------------------------------------------------------
Public Sub AddToDic(ByRef Dictionary As Dictionary, ByVal Key As String, Value As Variant)
    
    If Dictionary.Exists(Key) Then
    Else
        Dictionary.Add Key, Value
    End If
    
End Sub


'------------------------------------------------------------------------------
' @function Nz
' @desc
' @scope Public
'------------------------------------------------------------------------------
Public Function Nz(VariantToInspect As Variant, Optional ReplaceEmptiesWith As Variant = " ") As Variant
    If IsEmpty(VariantToInspect) Then
        Debug.Print "Is empty"
        Nz = ReplaceEmptiesWith
    Else
        Debug.Print "Wasn't empty"
        Debug.Print CStr(VariantToInspect)
        Nz = VariantToInspect
    End If
End Function


    
    








