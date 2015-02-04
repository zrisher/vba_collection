Attribute VB_Name = "ZachsArrayHelpers_Test"
Function GetNumArrayDim_Test() As Boolean
    Dim varArrayDim0 As Variant
    Dim varArrayDim1(1 To 10) As Variant
    Dim varArrayDim2(1 To 3, 0 To 5) As Variant
    Dim varArrayDim5 As Variant
    ReDim varArrayDim5(1 To 2, 1 To 4, 3 To 5, 2 To 4, 1 To 2)
    Debug.Print "GetNumArrayDim(varArrayDim0): " & GetNumArrayDim(varArrayDim0)
    Debug.Print "GetNumArrayDim(varArrayDim1): " & GetNumArrayDim(varArrayDim1)
    Debug.Print "GetNumArrayDim(varArrayDim2): " & GetNumArrayDim(varArrayDim2)
    Debug.Print "GetNumArrayDim(varArrayDim5): " & GetNumArrayDim(varArrayDim5)
End Function

Function TestThisShit() As Boolean
    Dim varTestArr(1 To 3) As Variant
    Dim varTestArrInside(1 To 10) As Variant
    varTestArr(1) = 1
    varTestArr(2) = "two"
    varTestArr(3) = varTestArrInside
'    Dim varCur As Variant
'    For Each varCur In varTestArr
'        If IsArray(varCur) Then varCur = "array"
'        Debug.Print varCur
'    Next
'        For Each varCur In varTestArr
'
'        Debug.Print varCur
'    Next
    Dim intCur As Integer
    For intCur = 1 To 3
        If IsArray(varTestArr(intCur)) Then varTestArr(intCur) = "array"
        'Debug.Print varTestArr(intCur)
    Next
        For intCur = 1 To 3
        Debug.Print varTestArr(intCur)
    Next
End Function

Function ArrayElementsAre_test() As Boolean
    Dim varTest(1 To 2) As Variant
    Dim strtest(1 To 2) As String
    varTest(1) = ""
    varTest(2) = ""
    Dim strtest2(1 To 2) As String
    Debug.Print ArrayElementsAre(varTest, "String")
    Debug.Print ArrayElementsAre(strtest, "String")
    Debug.Print ArrayElementsAre(strtest2, "String")
End Function
