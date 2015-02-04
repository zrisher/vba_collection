Attribute VB_Name = "ZachsArrayHelpers"
'------------------------------------------------------------------------------
' @module Zach's Array Helpers
' @desc A bunch of functions that arrays should (and probably do) come with,
' but were fun to program anyway
' @version 0.4
' @created 3-21-2013
' @last-updated 3-26-2013
'------------------------------------------------------------------------------
Option Explicit
Const cModulename = "ZachsVBAHelpers"


'------------------------------------------------------------------------------
' @function ArrayMerge1
' @desc Custom array merging, I think this might not be needed anymore
' @scope Public
'------------------------------------------------------------------------------
Public Function ArrayMerge1(Array1 As Variant, Array2 As Variant) As Variant
    Dim returnArray As Variant
    
    ReDim returnArray(LBound(Array1) To UBound(Array1), 0 To 1)
    
    For intCount = LBound(Array1) To UBound(Array1)
        returnArray(intCount, 0) = Array1(intCount)
        returnArray(intCount, 1) = Array2(intCount)
    Next

ArrayMerge1 = returnArray

End Function
        
        
'------------------------------------------------------------------------------
' @function Print2DArray
' @desc Debug.Prints out a 2D Array
' @scope Public
'------------------------------------------------------------------------------
Public Function Print2DArray(ArrayToPrint As Variant)
    Dim intCountR As Integer
    Dim intCountC As Integer
    Dim strReturnLine As String
    For intCountR = LBound(ArrayToPrint, 1) To UBound(ArrayToPrint, 1)
        For intCountC = LBound(ArrayToPrint, 2) To UBound(ArrayToPrint, 2)
            strReturnLine = strReturnLine & ArrayToPrint(intCountR, intCountC) & ","
        Next
        Debug.Print strReturnLine
        strReturnLine = ""
    Next
End Function
 

'------------------------------------------------------------------------------
' @function GetSub2DArray
' @desc Returns an array from the inside of another
' @scope Public
' @author Inspired by code from BitCoinBetter on
'   http://stackoverflow.com/questions/175170/how-do-i-slice-an-array-in-excel-vba
'------------------------------------------------------------------------------
Public Function GetSub2DArray(varArray As Variant, _
                            Optional ByVal StartRow As Long, _
                            Optional ByVal StartCol As Integer, _
                            Optional ByVal EndRow As Long, _
                            Optional ByVal EndCol As Integer, _
                            Optional ByVal Height As Long, _
                            Optional ByVal Width As Integer _
                            ) As Variant

    If StartRow = 0 Then StartRow = LBound(varArray, 1)
    If StartCol = 0 Then StartCol = LBound(varArray, 2)
    If EndRow = 0 Then
        If Height = 0 Then
            EndRow = UBound(varArray, 1)
        Else
            EndRow = StartRow + Height - 1
        End If
    End If
    If EndCol = 0 Then
        If Width = 0 Then
            EndCol = UBound(varArray, 2)
        Else
            EndCol = StartCol + Width - 1
        End If
    End If
    
    Dim varReturnArr As Variant
    ReDim varReturnArr(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    Dim lngCurRow As Long
    Dim intCurCol As Integer
    For lngCurRow = StartRow To EndRow
        For intCurCol = StartCol To EndCol
            varReturnArr(lngCurRow - StartRow + 1, intCurCol - StartCol + 1) = varArray(lngCurRow, intCurCol)
        Next
    Next
    
    GetSub2DArray = varReturnArr
    
End Function


'------------------------------------------------------------------------------
' @function
' @desc
' @scope
' @author
'------------------------------------------------------------------------------
Function GetNumArrayDim(ArrayToInspect As Variant) As Integer

    Dim intCurDim As Integer
    Dim intReturnDim As Integer
    Dim lngErrorCheck As Long
    On Error GoTo FinalDimension
    
    For intCurDim = 1 To gcVBA_MAXARRAYDIM
      lngErrorCheck = LBound(ArrayToInspect, intCurDim)
    Next intCurDim
    intReturnDim = intCurDim
    
PROC_EXIT:
    GetNumArrayDim = intReturnDim
    Exit Function
    
FinalDimension:
    intReturnDim = intCurDim - 1
    GoTo PROC_EXIT
    
End Function

Function ArrayElementsAre(ArrayToCheck As Variant, VarTypeName As String) As Boolean
    If IsArray(ArrayToCheck) Then
        Dim varCur As Variant
        Dim blnIsOnlyExpected As String
        blnIsOnlyExpected = True
        For Each varCur In ArrayToCheck
            If Not TypeName(varCur) = VarTypeName Then
                blnIsOnlyExpected = False
                Exit For
            End If
        Next
        ArrayElementsAre = blnIsOnlyExpected
    Else
        ArrayElementsAre = (TypeName(ArrayToCheck) = VarTypeName)
End If


End Function
