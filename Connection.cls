VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarMyUsername As String 'local copy
Private mvarMyPassword As String 'local copy
Private mvarMyServer As String 'local copy
Private mvarMyTables As Tables 'local copy
'local variable(s) to hold property value(s)
Private mvarMyProcessList As ProcessLists 'local copy
Private mvarMyServerVersion As String 'local copy
Private mvarMyTimeAlive As Date 'local copy
Private mvarMyVariableState As VariableStates 'local copy
'local variable(s) to hold property value(s)
Private mvarMyHandle As Long 'local copy

Private myRec As API_MYSQL         'valid record to hold ourselves
Private myRec_res As API_MYSQL_RES 'this is we we will store results for ourselves
Private myRec_field As API_MYSQL_FIELD
Private myRec_rows As API_MYSQL_ROWS

Public Property Let MyHandle(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyHandle = 5
    mvarMyHandle = vData
End Property


Public Property Get MyHandle() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyHandle
    MyHandle = mvarMyHandle
End Property




Public Property Set MyVariableState(ByVal vData As VariableStates)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyVariableState = Form1
    Set mvarMyVariableState = vData
End Property


Public Property Get MyVariableState() As VariableStates
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyVariableState
    Set MyVariableState = mvarMyVariableState
End Property



Public Sub MyRefreshInfo()
Dim MyRes As ResultTable

Set MyRes = MyExecute("show full processlist")
MyRes.RT_MoveFirst
Do While Not MyRes.RT_EOF
    With MyProcessList.Add
        .MyID = MyRes("ID")
        .MyCommand = MyRes("Command")
        .MyDB = MyRes("DB")
        .MyHost = MyRes("Host")
        .MyInfo = MyRes("Info")
        .MyState = MyRes("State")
        .MyTime = MyRes("Time")
        .MyUser = MyRes("User")
    End With
    MyRes.RT_MoveNext
Loop

Set MyRes = MyExecute("show variables")
MyRes.RT_MoveFirst
Do While Not MyRes.RT_EOF
    With MyVariableState.Add
        .VarName = MyRes("variable_name")
        .VarValue = MyRes("value")
    End With
    MyRes.RT_MoveNext
Loop

Set MyRes = MyExecute("show status")
MyRes.RT_MoveFirst
Do While Not MyRes.RT_EOF
    With MyVariableState.Add
        .VarName = MyRes("variable_name")
        .VarValue = MyRes("value")
    End With
    MyRes.RT_MoveNext
Loop

End Sub

Public Property Let MyTimeAlive(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyTimeAlive = 5
    mvarMyTimeAlive = vData
End Property


Public Property Get MyTimeAlive() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyTimeAlive
    MyTimeAlive = mvarMyTimeAlive
End Property



Public Property Let MyServerVersion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyServerVersion = 5
    mvarMyServerVersion = vData
End Property


Public Property Get MyServerVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyServerVersion
    MyServerVersion = mvarMyServerVersion
End Property



Public Property Set MyProcessList(ByVal vData As ProcessLists)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyProcessList = Form1
    Set mvarMyProcessList = vData
End Property


Public Property Get MyProcessList() As ProcessLists
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyProcessList
    Set MyProcessList = mvarMyProcessList
End Property



Public Property Set MyTables(ByVal vData As Tables)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyTables = Form1
    Set mvarMyTables = vData
End Property


Public Property Get MyTables() As Tables
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyTables
    Set MyTables = mvarMyTables
End Property



Public Function MyConnect(Optional pMyServer As String, Optional pMyUsername As String, Optional pMyPassword As String) As Boolean

If pMyServer = "" And MyServer = "" Then
    Err.Raise vbObjectError & 1000, App.ProductName, "Servidor não informado"
    Exit Function
End If

If pMyUsername = "" And MyUsername = "" Then
    Err.Raise vbObjectError & 1001, App.ProductName, "Usuário não informado"
    Exit Function
End If

If pMyServer <> "" Then MyServer = pMyServer
If pMyUsername <> "" Then MyUsername = pMyUsername
If pMyPassword <> "" Then MyPassword = pMyPassword

MyHandle = API_mysql_connect(myRec, StrPtr(StrConv(MyServer, vbFromUnicode)), StrPtr(StrConv(MyUsername, vbFromUnicode)), StrPtr(StrConv(MyPassword, vbFromUnicode)))

If MyHandle Then
    CopyMemory myRec, ByVal MyHandle, LenB(myRec) 'again we have to mooo-ve the memory
    MyServerVersion = ptr2str(myRec.server_version)
    MyRefreshInfo
    MyConnect = True
Else
    Err.Raise API_mysql_errno(myRec), App.ProductName, ptr2str(API_mysql_error(myRec))
    MyServerVersion = ""
    MyConnect = False
End If

End Function

Public Sub MyClose()
End Sub

Public Sub MySelectDatabase(MyDatabaseName As String)
Dim Ret As Long
    If MyHandle <> 0 Then
        Ret = API_mysql_select_db(myRec, StrPtr(StrConv(MyDatabaseName, vbFromUnicode)))
        If Ret <> 0 Then Err.Raise API_mysql_errno(myRec), App.ProductName, ptr2str(API_mysql_error(myRec))
    End If
End Sub


Public Function MyExecute(SQL_Command As String, Optional ClientCursor As Boolean = True) As ResultTable
Dim MyResult As New ResultTable
Dim TmpSQL As String
Dim Ret As Long
Dim m_fieldcount As Long
Dim m_rowcount As Long
Dim i As Long
Dim j As Long
Dim s As String
Dim PickUp() As Long
Dim TblName As String

If InStr(UCase(SQL_Command), " FROM ") = 0 And InStr(UCase(SQL_Command), "INSERT") = 0 And InStr(UCase(SQL_Command), "UPDATE") = 0 Then
    TmpSQL = "select * from " & SQL_Command
Else
    If Left(UCase(SQL_Command), 6) = "SELECT" Then
        TblName = Right(SQL_Command, Len(SQL_Command) - InStr(UCase(SQL_Command), "FROM") - 4)
        TblName = Trim(Left(TblName, InStr(TblName, " ")))
    End If
End If


Ret = API_mysql_query(myRec, StrPtr(StrConv(SQL_Command, vbFromUnicode)))
If Ret = 0 Then 'query was good
    
    If ClientCursor Then
        Ret = API_mysql_store_result(myRec)
    Else
        Ret = API_mysql_use_result(myRec)
    End If
    
    If Ret Then
    
        MyRecRes = myRec
        
        MyResult.MyStoreResult Ret, Me, TblName

        Set MyExecute = MyResult
        MyResult.RT_MoveFirst
        
    End If
Else
    Err.Raise API_mysql_errno(myRec), App.ProductName, ptr2str(API_mysql_error(myRec))
End If

End Function


Public Property Let MyServer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyServer = 5
    mvarMyServer = vData
End Property


Public Property Get MyServer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyServer
    MyServer = mvarMyServer
End Property



Public Property Let MyPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyPassword = 5
    mvarMyPassword = vData
End Property


Public Property Get MyPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyPassword
    MyPassword = mvarMyPassword
End Property



Public Property Let MyUsername(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyUsername = 5
    mvarMyUsername = vData
End Property


Public Property Get MyUsername() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyUsername
    MyUsername = mvarMyUsername
End Property



Private Sub Class_Initialize()
Dim lSys As String
Dim lByte() As Byte
Dim lStr As String
Dim Ret As Long

    lStr = Space(255)
    
    Ret = GetSystemDirectory(lStr, Len(lStr))
    lStr = Left(lStr, Ret)
    lStr = IIf(Right(lStr, 1) = "\", lStr, lStr & "\")
    
    If Dir(lStr & "libmysql.dll") = "" Then
        lByte = LoadResData("LIBMYSQL", "DLL")
        Close
        Open lStr & "libmysql.dll" For Binary Shared As #1
        Put #1, , lByte
        Close
    End If
    
    MyHandle = API_mysql_init(myRec)
    If MyHandle Then
        CopyMemory myRec, ByVal MyHandle, LenB(myRec) 'move the memory back from pointer
    Else
        Err.Raise 512, App.ProductName, "Não foi possível inicializar o mecanismo de Banco de Dados::Fatal"
    End If

    Set MyProcessList = New ProcessLists
    Set MyVariableState = New VariableStates

End Sub


Private Sub Class_Terminate()
Dim Ret As Long

Ret = API_mysql_close(myRec)

Set MyVariableState = Nothing
Set MyProcessList = Nothing

End Sub


