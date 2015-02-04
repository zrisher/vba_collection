Attribute VB_Name = "JobviteTools_Test"
Function ExportToDB_Test() As Boolean
    Dim wshSource As Worksheet
    Set wshSource = Workbooks(gcWORKBOOK_CONTROL_FILENAME).Worksheets("upload from jv")
    Call ExportToDB(wshSource)
End Function
