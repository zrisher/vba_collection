Attribute VB_Name = "Sqlite3Helpers"
'------------------------------------------------------------------------------
' @module SQLite3 Helper Functions
' @authors Adapted by Zach Risher From code included in the SQLiteForExcel
'   dist, http://sqliteforexcel.codeplex.com/
'------------------------------------------------------------------------------

#If Win64 Then
Public Function SQLite3ExecuteNonQuery(ByVal dbHandle As LongPtr, ByVal SqlCommand As String) As Long
    Dim stmtHandle As LongPtr
#Else
Public Function SQLite3ExecuteNonQuery(ByVal dbHandle As Long, ByVal SqlCommand As String) As Long
    Dim stmtHandle As Long
#End If
    
    SQLite3PrepareV2 dbHandle, SqlCommand, stmtHandle
    SQLite3Step stmtHandle
    SQLite3Finalize stmtHandle
    
    SQLite3ExecuteNonQuery = SQLite3Changes(dbHandle)
End Function

#If Win64 Then
Public Sub SQLite3ExecuteQuery(ByVal dbHandle As LongPtr, ByVal sqlQuery As String)
    Dim stmtHandle As LongPtr
#Else
Public Sub SQLite3ExecuteQuery(ByVal dbHandle As Long, ByVal sqlQuery As String)
    Dim stmtHandle As Long
#End If
    ' Dumps a query to the debug window. No error checking
    
    Dim RetVal As Long

    RetVal = SQLite3PrepareV2(dbHandle, sqlQuery, stmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(stmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns stmtHandle
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Move to next row
    RetVal = SQLite3Step(stmtHandle)
    Do While RetVal = SQLITE_ROW
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns stmtHandle
        RetVal = SQLite3Step(stmtHandle)
    Loop

    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(stmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
End Sub


#If Win64 Then
Public Sub PrintColumns(ByVal stmtHandle As LongPtr)
#Else
Public Sub PrintColumns(ByVal stmtHandle As Long)
#End If
    Dim colCount As Long
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    
    Dim i As Long
    
    colCount = SQLite3ColumnCount(stmtHandle)
    Debug.Print "Column count: " & colCount
    For i = 0 To colCount - 1
        colName = SQLite3ColumnName(stmtHandle, i)
        colType = SQLite3ColumnType(stmtHandle, i)
        colTypeName = TypeName(colType)
        colValue = ColumnValue(stmtHandle, i, colType)
        Debug.Print "Column " & i & ":", colName, colTypeName, colValue
    Next
End Sub

#If Win64 Then
Public Sub PrintParameters(ByVal stmtHandle As LongPtr)
#Else
Public Sub PrintParameters(ByVal stmtHandle As Long)
#End If
    Dim paramCount As Long
    Dim paramName As String
    
    Dim i As Long
    
    paramCount = SQLite3BindParameterCount(stmtHandle)
    Debug.Print "Parameter count: " & paramCount
    For i = 1 To paramCount
        paramName = SQLite3BindParameterName(stmtHandle, i)
        Debug.Print "Parameter " & i & ":", paramName
    Next
End Sub

Public Function TypeName(ByVal SQLiteType As Long) As String
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            TypeName = "INTEGER"
        Case SQLITE_FLOAT:
            TypeName = "FLOAT"
        Case SQLITE_TEXT:
            TypeName = "TEXT"
        Case SQLITE_BLOB:
            TypeName = "BLOB"
        Case SQLITE_NULL:
            TypeName = "NULL"
    End Select
End Function

#If Win64 Then
Public Function ColumnValue(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Public Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function
