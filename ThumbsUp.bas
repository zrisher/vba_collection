Attribute VB_Name = "ThumbsUp"
'------------------------------------------------------------------------------
' @module
' @desc
' @version
' @created
' @last-updated
'------------------------------------------------------------------------------
Option Explicit

Public ThumbsUp_Path As String
Public SQLite3DLL_Path As String
Public SQLite3DLL_x64_Path As String
Public ThumbsUpDB_Path As String
Public TestFile_Name As String
Public TestBackup_Name As String
Public WorkingFile_Name As String

Sub SetFilePaths()
    ThumbsUp_Path = ThisWorkbook.Path
    SQLite3DLL_Path = ThumbsUp_Path & "\lib"
    SQLite3DLL_x64_Path = SQLite3DLL_Path & "\x64"
    ThumbsUpDB_Path = ThumbsUp_Path & "\db"
    TestFile_Name = "TestSQLiteDB.db3"
    TestBackup_Name = "TestSQLiteBackup.db3"
    WorkingFile_Name = "WorkingDB.db3"
End Sub

'------------------------------------------------------------------------------
' @sub
' @desc
'------------------------------------------------------------------------------
Public Sub ThumbsUp_Init()
    
    Dim InitReturn As Long
    
    Call SetFilePaths
    Call Sqlite3Demo.AllTests

    #If Win64 Then
        InitReturn = SQLite3Initialize(SQLite3DLL_x64_Path)
    #Else
        InitReturn = SQLite3Initialize(SQLite3DLL_Path)
    #End If

    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        Exit Sub
    Else
        Debug.Print "It works!" '"SQLite initialized."
    End If

    SQLite3Free
    
End Sub
    

Sub ExportFromXLDBToDB()

'CREATE TABLE If not exists Users (id integer primary key autoincrement notnull, name text COLLATE NOCASE , team_id integer, email text)

    Debug.Print "----- ExportFromXLDBToDB Start -----"
    
    ' START Insert Time
    Dim testStart As Date
    testStart = Now()
    
    Call SetFilePaths
    
    'open the workbook
    Dim wbkXLDB As Workbook
    Set wbkXLDB = OpenWorkbook("interview debrief data play.xlsm", ThumbsUp_Path)

    
    ' Open the database - getting a DbHandle back
    Dim strDBFilename As String
    strDBFilename = ThumbsUpDB_Path & WorkingFile_Name
    
    #If Win64 Then
        Dim lngDBHandle As LongPtr
        Dim lngStmtHandle As LongPtr
    #Else
        Dim lngDBHandle As Long
        Dim lngStmtHandle As Long
    #End If
    
    Dim strLastSQLRet As Long
    strLastSQLRet = SQLite3Open(strDBFilename, lngDBHandle)
    Debug.Print "SQLite3Open returned " & strLastSQLRet
    
    '------------------------
    ' sdfdsfsd
    ' ================
    'SQLite3ExecuteQuery lngDBHandle, "CREATE TABLE MyBigTable (TheId INTEGER, TheDate REAL, TheText TEXT, TheValue REAL)"


'    '------------------------
'    ' Create the table
'    ' ================
'    SQLite3PrepareV2 lngDBHandle, "CREATE TABLE MyBigTable (TheId INTEGER, TheDate REAL, TheText TEXT, TheValue REAL)", lngStmtHandle
'    SQLite3Step lngStmtHandle
'    SQLite3Finalize lngStmtHandle
'
'    '---------------------------
'    ' Add an index
'    ' ================
'    SQLite3PrepareV2 lngDBHandle, "CREATE INDEX idx_MyBigTable_Id_Date ON MyBigTable (TheId, TheDate)", lngStmtHandle
'    SQLite3Step lngStmtHandle
'    SQLite3Finalize lngStmtHandle
'
'    '-------------------
'    ' Begin transaction
'    '==================
'    SQLite3PrepareV2 lngDBHandle, "BEGIN TRANSACTION", lngStmtHandle
'    SQLite3Step lngStmtHandle
'    SQLite3Finalize lngStmtHandle
'
'    '-------------------------
'    ' Prepare an insert statement with parameters
'    ' ===============
'    ' Create the sql statement - getting a StmtHandle back
'    strLastSQLRet = SQLite3PrepareV2(lngDBHandle, "INSERT INTO MyBigTable Values (?, ?, ?, ?)", lngStmtHandle)
'    If strLastSQLRet <> SQLITE_OK Then
'        Debug.Print "SQLite3PrepareV2 returned " & SQLite3ErrMsg(lngDBHandle)
'        Beep
'    End If
'
'    Randomize
'    Dim startDate As Date
'    startDate = DateValue("1 Jan 2000")
'    Dim curDate As Date
'    Dim curValue As Double
'    Dim offset As Long
'
'    Dim i As Long
'    For i = 1 To 100000
'        curDate = startDate + i
'        curValue = Rnd() * 1000
'
'        strLastSQLRet = SQLite3BindInt32(lngStmtHandle, 1, 42000 + i)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Bind returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3BindDate(lngStmtHandle, 2, curDate)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Bind returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3BindText(lngStmtHandle, 3, "The quick brown fox jumped over the lazy dog.")
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Bind returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3BindDouble(lngStmtHandle, 4, curValue)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Bind returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3Step(lngStmtHandle)
'        If strLastSQLRet <> SQLITE_DONE Then
'            Debug.Print "SQLite3Step returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3Reset(lngStmtHandle)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Reset returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'    Next
'
'    ' Finalize (delete) the statement
'    strLastSQLRet = SQLite3Finalize(lngStmtHandle)
'    Debug.Print "SQLite3Finalize returned " & strLastSQLRet
'
'    '-------------------
'    ' Commit transaction
'    '==================
'    ' (I'm re-using the same variable lngStmtHandle for the new statement)
'    SQLite3PrepareV2 lngDBHandle, "COMMIT TRANSACTION", lngStmtHandle
'    SQLite3Step lngStmtHandle
'    SQLite3Finalize lngStmtHandle
'
'    ' STOP Insert Time
'    Debug.Print "Insert Elapsed: " & Format(Now() - testStart, "HH:mm:ss")
'
'    ' START Select  Time
'    testStart = Now()
'
'
'    '-------------------------
'    ' Select statement
'    ' ===============
'    ' Create the sql statement - getting a StmtHandle back
'    ' Now using named parameters!
'    strLastSQLRet = SQLite3PrepareV2(lngDBHandle, "SELECT TheId, datetime(TheDate), TheText, TheValue FROM MyBigTable WHERE TheId = @FindThisId AND TheDate <= @FindThisDate LIMIT 1", lngStmtHandle)
'    Debug.Print "SQLite3PrepareV2 returned " & strLastSQLRet
'
'    Dim paramIndexId As Long
'    Dim paramIndexDate As Long
'
'    paramIndexId = SQLite3BindParameterIndex(lngStmtHandle, "@FindThisId")
'    If paramIndexId = 0 Then
'        Debug.Print "SQLite3BindParameterIndex could not find the Id parameter!"
'        Beep
'    End If
'
'    paramIndexDate = SQLite3BindParameterIndex(lngStmtHandle, "@FindThisDate")
'    If paramIndexDate = 0 Then
'        Debug.Print "SQLite3BindParameterIndex could not find the Date parameter!"
'        Beep
'    End If
'
'    startDate = DateValue("1 Jan 2000")
'
'
'    For i = 1 To 100000
'        offset = i Mod 10000
'        ' Bind the parameters
'        strLastSQLRet = SQLite3BindInt32(lngStmtHandle, paramIndexId, 42000 + 500 + offset)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Bind returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3BindDate(lngStmtHandle, paramIndexDate, startDate + 500 + offset)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Bind returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'
'        strLastSQLRet = SQLite3Step(lngStmtHandle)
'        If strLastSQLRet = SQLITE_ROW Then
'            ' We have access to the result columns here.
'            If offset = 1 Then
'                Debug.Print "At row " & i
'                Debug.Print "------------"
'                PrintColumns lngStmtHandle
'                Debug.Print "============"
'            End If
'        ElseIf strLastSQLRet = SQLITE_DONE Then
'            Debug.Print "No row found"
'        End If
'
'        strLastSQLRet = SQLite3Reset(lngStmtHandle)
'        If strLastSQLRet <> SQLITE_OK Then
'            Debug.Print "SQLite3Reset returned " & strLastSQLRet, SQLite3ErrMsg(lngDBHandle)
'            Beep
'        End If
'    Next
'
'    ' Finalize (delete) the statement
'    strLastSQLRet = SQLite3Finalize(lngStmtHandle)
'    Debug.Print "SQLite3Finalize returned " & strLastSQLRet

    ' STOP Select time
    Debug.Print "Select Elapsed: " & Format(Now() - testStart, "HH:mm:ss")
    
    ' Close the database
    strLastSQLRet = SQLite3Close(lngDBHandle)
    Kill strDBFilename

    Debug.Print "----- ExportFromXLDBToDB End -----"
End Sub


