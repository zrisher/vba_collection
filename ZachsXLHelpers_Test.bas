Attribute VB_Name = "ZachsXLHelpers_Test"
Sub runtests()
MsgBox test_SheetExists_P1 & test_SheetExists_N1
End Sub


Function test_SheetExists_P1()
test_SheetExists_P1 = False
Set wksh = ActiveWorkbook.Worksheets.Add
wksh.Name = "Test WorkSheet"
test_SheetExists_P1 = SheetExists(wksh.Name, ActiveWorkbook)
Application.DisplayAlerts = False
Worksheets(wksh.Name).Delete
Application.DisplayAlerts = True
Debug.Print "test_SheetExists_P1: " & test_SheetExists_P1
End Function

Function test_SheetExists_N1()
test_SheetExists_N1 = Not SheetExists("Whowouldevernameaworksheetsomethingthislong12342q3dre4r", ThisWorkbook)
End Function

Function test_OneDArrayToString() As Boolean

    Dim TestItem() As String
    ReDim TestItem(1 To 3) As String
    TestItem(1) = "Pickle"
    TestItem(2) = "Pineapple"
    TestItem(3) = "Papaya"
    Dim TestItem2() As Variant
    TestItem2 = Array(1, "text", True)
    TestShoppingList = Join(TestItem, ", ")
    Debug.Print (TestShoppingList)
    
End Function

Function ArrayMerge1_PTest()
ArrayMerge1_PTest = False

Dim My_Array
My_Array = Array("one", "two", "three")

Dim Array1
Dim Array2
Dim CombinedArray(1 To 3, 1 To 2) As Variant

Array1 = Array("fieldname1", "fieldname2", "fieldname3")
Array2 = Array("Johhny Cash", 1, "Alcatraz")
CombinedArray(1, 1) = Array1(0)
CombinedArray(1, 2) = Array2(0)
CombinedArray(2, 1) = Array1(1)
CombinedArray(2, 2) = Array2(1)
CombinedArray(3, 1) = Array1(2)
CombinedArray(3, 2) = Array2(2)

Dim Finalarray As Variant
Finalarray = ArrayMerge1(Array1, Array2)

Print2DArray (Finalarray)



End Function

Function test_OpenWorkbook() As Boolean
Dim strFilePath As String
strFilePath = "C:\path\to\file\"
Dim strFileName As String
strFileName = "interview debrief data.xlsm"
Dim wkbkResult As Workbook
Set wkbkResult = OpenWorkbook(strFileName, strFilePath)
End Function

Function test_SplitSequential() As Boolean
    Dim strToCut As String
    Dim strCutters() As String
    Dim Result() As String
    
    strToCut = "first part CUT second part CUT!@#$$%^&*()ME third part CUT ME HERE fourth part"
    ReDim strCutters(1 To 3)
    strCutters(1) = " CUT "
    strCutters(2) = " CUT!@#$$%^&*()ME "
    strCutters(3) = " CUT ME HERE "
    Result = SplitSequential(strToCut, strCutters)
    test_SplitSequential = (Join(Result, ",") = "first part,second part,third part,fourth part")
    Debug.Print Join(Result, ",")
    
    strToCut = "first part CUT second part CUT!@#$$%^&*()ME third part CUT ME HERE fourth part"
    ReDim strCutters(1 To 3)
    strCutters(1) = ""
    strCutters(2) = " "
    strCutters(3) = " CUT ME HERE "
    Result = SplitSequential(strToCut, strCutters)
    test_SplitSequential = (Join(Result, ",") = "first,part CUT second part CUT!@#$$%^&*()ME third part,fourth part,")
    Debug.Print Join(Result, ",")
    
End Function

Function test_TwoDArraysToWorksheet() As Boolean
    Dim varArray1 As Variant
        ReDim varArray1(1 To 3, 1 To 5)
        varArray1(1, 1) = "vararray1 first"
        varArray1(3, 5) = "vararray1 last"
    Dim varArray2 As Variant
        ReDim varArray2(0 To 3, 0 To 3)
        varArray2(0, 0) = "varArray2 first"
        varArray2(3, 1) = "varArray2 last"
        varArray2(2, 2) = "varArray2 mid"
        varArray2(1, 3) = "varArray2 last"
        varArray2(3, 3) = "varArray2 last"
    Dim varArray3 As Variant
    Dim varArray4 As Variant
    Dim varArray5 As Variant
        ReDim varArray5(1 To 3, 1 To 5)
        varArray5(1, 1) = "vararray5 first"
        varArray5(3, 5) = "vararray5 last"
    Dim varArrOfArrays As Variant
        ReDim varArrOfArrays(1 To 5)
        varArrOfArrays(1) = varArray1
        varArrOfArrays(2) = varArray2
        varArrOfArrays(3) = varArray3
        varArrOfArrays(4) = varArray4
        varArrOfArrays(5) = varArray5
    Dim wshWorksheet As Worksheet
    Set wshWorksheet = Workbooks("interview debrief tool.xlsm").Worksheets("Dest2")
    Call TwoDArraysToWorksheet(varArrOfArrays, wshWorksheet)
End Function


