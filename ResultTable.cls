VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResultTable"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarRT_EOF As Boolean 'local copy
Private mvarRT_BOF As Boolean 'local copy
Private mvarRT_Field As Fields 'local copy
Dim myRec As API_MYSQL         'valid record to hold ourselves
Dim myRec_res As API_MYSQL_RES 'this is we we will store results for ourselves
Dim myRec_field As API_MYSQL_FIELD
Dim myRec_rows As API_MYSQL_ROWS
'local variable(s) to hold property value(s)
Private mvarRT_RowCount As Long 'local copy
Dim CurReg As Long
Dim pRet()
Dim Populate As Boolean
Public Property Let RT_RowCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RT_RowCount = 5
    mvarRT_RowCount = vData
End Property


Public Property Get RT_RowCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RT_RowCount
    RT_RowCount = mvarRT_RowCount
End Property




Public Sub MyStoreResult(Ret As Long, Conn As MySQLDTA.Connection, TblName As String)
Dim lFld As Field
Dim MyResult As ResultTable
Dim dType As String

If Ret Then
    myRec = MyRecRes
    
    CopyMemory myRec_res, ByVal Ret, LenB(myRec_res)

    m_fieldcount = myRec_res.field_count
    nCol = m_fieldcount
    RT_RowCount = convert642l(myRec_res.row_count)
    ReDim PickUp(1 To m_fieldcount)
    
    If TblName <> "" Then
        Set MyResult = Conn.MyExecute("show fields from " & TblName)
    End If

    For i = 1 To m_fieldcount
        CurReg = 1
        Ret = API_mysql_fetch_field(myRec_res)
        If Ret Then
            CopyMemory myRec_field, ByVal Ret, LenB(myRec_field)
            Set lFld = RT_Field.Add(ptr2str(myRec_field.name))
            lFld.FieldName = ptr2str(myRec_field.name)
            lFld.FieldID = i
            If TblName <> "" Then
                MyResult.RT_MoveFirst
                Do While Not MyResult.RT_EOF
                    If MyResult.RT_Field("Field").Value = lFld.FieldName Then
                        lFld.FieldDefault = MyResult.RT_Field("Default").Value
                        lFld.FieldExtra = MyResult.RT_Field("Extra").Value
                        lFld.FieldNull = IIf(MyResult.RT_Field("Null").Value = "YES", True, False)
                        If InStr(MyResult.RT_Field("Type").Value, "(") = 0 Then
                            dType = MyResult.RT_Field("Type").Value
                        Else
                            dType = Left(MyResult.RT_Field("Type").Value, InStr(MyResult.RT_Field("Type").Value, "(") - 1)
                        End If
                        
                        Select Case UCase(dType)
                            Case "TINYINT"
                                lFld.FieldType = ft_TinyInt
                            Case "SMALLINT"
                                lFld.FieldType = ft_SmallInt
                            Case "MEDIUMINT"
                                lFld.FieldType = ft_MediumInt
                            Case "INT"
                                lFld.FieldType = ft_Int
                            Case "BIGINT"
                                lFld.FieldType = ft_BigInt
                            Case "FLOAT"
                                lFld.FieldType = ft_Float
                            Case "DOUBLE"
                                lFld.FieldType = ft_Double
                            Case "DECIMAL"
                                lFld.FieldType = ft_Decimal
                            Case "DATE"
                                lFld.FieldType = ft_Date
                            Case "DATETIME"
                                lFld.FieldType = ft_DateTime
                            Case "TIMESTAMP"
                                lFld.FieldType = ft_TimeStamp
                            Case "TIME"
                                lFld.FieldType = ft_Time
                            Case "YEAR"
                                lFld.FieldType = ft_Year
                            Case "CHAR"
                                lFld.FieldType = ft_Char
                            Case "VARCHAR"
                                lFld.FieldType = ft_VarChar
                            Case "TINYBLOB"
                                lFld.FieldType = ft_TinyBlob
                            Case "TINYTEXT"
                                lFld.FieldType = ft_TinyText
                            Case "TEXT"
                                lFld.FieldType = ft_Text
                            Case "BLOB"
                                lFld.FieldType = ft_Blob
                            Case "MEDIUMBLOB"
                                lFld.FieldType = ft_MediumBlob
                            Case "MEDIUMTEXT"
                                lFld.FieldType = ft_MediumText
                            Case "LONGBLOB"
                                lFld.FieldType = ft_LongBlob
                            Case "LONGTEXT"
                                lFld.FieldType = ft_LongText
                            Case "ENUM"
                                lFld.FieldType = ft_Enum
                            Case "SET"
                                lFld.FieldType = ft_Set
                        End Select
                        Exit Do
                    End If
                    MyResult.RT_MoveNext
                Loop
            End If
            
        End If
    Next
              
Else
    RT_RowCount = convert642l(myRec.affected_rows)
End If

End Sub


Public Property Set RT_Field(ByVal vData As Fields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.RT_Field = Form1
    Set mvarRT_Field = vData
End Property


Public Property Get RT_Field() As Fields
Attribute RT_Field.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RT_Field
    Set RT_Field = mvarRT_Field
End Property



Public Sub RT_MovePrevious()
Dim Fld As Field
Dim PickUp() As Long
Dim Ret As Long
Dim j As Long
Dim i As Long

If RT_RowCount = 0 Then RT_EOF = True
If RT_RowCount = 0 Then RT_BOF = True

If CurReg = 1 Then
    Err.Raise vbObjectError & 1003, App.ProductName, "Tentativa de ultrapassar primeiro registro"
Else
    CurReg = CurReg - 1
End If

If Not Populate Then
    Populate = True
        
    ReDim pRet(1 To RT_RowCount, 1 To RT_Field.Count)
    ReDim PickUp(1 To RT_Field.Count)

    For j = 1 To RT_RowCount
        Ret = API_mysql_fetch_row(myRec_res)
        If Ret Then
            CopyMemory PickUp(1), ByVal Ret, SIZE_OF_CHAR * RT_Field.Count 'copy it into array so we can pick it up
            For Each Fld In RT_Field
                If Ret Then
                    s = ptr2str(PickUp(Fld.FieldID))
                    pRet(j, Fld.FieldID) = s
                End If
            Next
        End If
    Next
End If

For Each Fld In RT_Field
    If RT_RowCount = 0 Then
        Fld.Value = Empty
    Else
        If CurReg > RT_RowCount Then
            Fld.Value = Empty
        Else
        
            Select Case Fld.FieldType
                Case ft_TinyInt, ft_SmallInt, ft_MediumInt, ft_Int, ft_BigInt, ft_Year
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CLng(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Float, ft_Double, ft_Decimal, ft_TimeStamp
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDbl(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Date, ft_DateTime, ft_Time
                    If IsDate(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDate(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case Else
                    Fld.Value = pRet(CurReg, Fld.FieldID)
            End Select
        End If
    End If
Next

End Sub


Public Sub RT_MoveLast()
Dim Fld As Field
Dim PickUp() As Long
Dim Ret As Long
Dim j As Long
Dim i As Long

CurReg = RT_RowCount
RT_EOF = True

If RT_RowCount = 0 Then RT_EOF = True
If RT_RowCount = 0 Then RT_BOF = True

If Not Populate Then
    Populate = True
        
    ReDim pRet(1 To RT_RowCount, 1 To RT_Field.Count)
    ReDim PickUp(1 To RT_Field.Count)

    For j = 1 To RT_RowCount
        Ret = API_mysql_fetch_row(myRec_res)
        If Ret Then
            CopyMemory PickUp(1), ByVal Ret, SIZE_OF_CHAR * RT_Field.Count 'copy it into array so we can pick it up
            For Each Fld In RT_Field
                If Ret Then
                    s = ptr2str(PickUp(Fld.FieldID))
                    pRet(j, Fld.FieldID) = s
                End If
            Next
        End If
    Next
End If

For Each Fld In RT_Field
    If RT_RowCount = 0 Then
        Fld.Value = Empty
    Else
        If CurReg > RT_RowCount Then
            Fld.Value = Empty
        Else
        
            Select Case Fld.FieldType
                Case ft_TinyInt, ft_SmallInt, ft_MediumInt, ft_Int, ft_BigInt, ft_Year
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CLng(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Float, ft_Double, ft_Decimal, ft_TimeStamp
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDbl(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Date, ft_DateTime, ft_Time
                    If IsDate(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDate(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case Else
                    Fld.Value = pRet(CurReg, Fld.FieldID)
            End Select
        End If
    End If
Next

End Sub


Public Sub RT_MoveFirst()
Dim Fld As Field
Dim PickUp() As Long
Dim Ret As Long
Dim j As Long
Dim i As Long

CurReg = 1
RT_BOF = True

If RT_RowCount = 0 Then RT_EOF = True
If RT_RowCount = 0 Then RT_BOF = True

If Not Populate Then
    Populate = True
    If RT_RowCount > 0 Then ReDim pRet(1 To RT_RowCount, 1 To RT_Field.Count)
    ReDim PickUp(1 To RT_Field.Count)

    For j = 1 To RT_RowCount
        Ret = API_mysql_fetch_row(myRec_res)
        If Ret Then
            CopyMemory PickUp(1), ByVal Ret, SIZE_OF_CHAR * RT_Field.Count 'copy it into array so we can pick it up
            For Each Fld In RT_Field
                If Ret Then
                    s = ptr2str(PickUp(Fld.FieldID))
                    pRet(j, Fld.FieldID) = s
                End If
            Next
        End If
    Next
End If

For Each Fld In RT_Field
    If RT_RowCount = 0 Then
        Fld.Value = Empty
    Else
        If CurReg > RT_RowCount Then
            Fld.Value = Empty
        Else
            Select Case Fld.FieldType
                Case ft_TinyInt, ft_SmallInt, ft_MediumInt, ft_Int, ft_BigInt, ft_Year
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CLng(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Float, ft_Double, ft_Decimal, ft_TimeStamp
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDbl(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Date, ft_DateTime, ft_Time
                    If IsDate(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDate(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case Else
                    Fld.Value = pRet(CurReg, Fld.FieldID)
            End Select
        End If
    End If
Next

End Sub

Public Sub RT_MoveNext()
Dim Fld As Field
Dim PickUp() As Long
Dim Ret As Long
Dim j As Long
Dim i As Long

If RT_RowCount = 0 Then RT_EOF = True
If RT_RowCount = 0 Then RT_BOF = True

If CurReg >= RT_RowCount + 1 Then
    Err.Raise vbObjectError & 1003, App.ProductName, "Tentativa de ultrapassar último registro"
Else
    CurReg = CurReg + 1
End If

If Not Populate Then
    Populate = True
        
    ReDim pRet(1 To RT_RowCount, 1 To RT_Field.Count)
    ReDim PickUp(1 To RT_Field.Count)

    For j = 1 To RT_RowCount
        Ret = API_mysql_fetch_row(myRec_res)
        If Ret Then
            CopyMemory PickUp(1), ByVal Ret, SIZE_OF_CHAR * RT_Field.Count 'copy it into array so we can pick it up
            For Each Fld In RT_Field
                If Ret Then
                    s = ptr2str(PickUp(Fld.FieldID))
                    pRet(j, Fld.FieldID) = s
                End If
            Next
        End If
    Next
End If

For Each Fld In RT_Field
    If RT_RowCount = 0 Then
        Fld.Value = Empty
    Else
        If CurReg > RT_RowCount Then
            Fld.Value = Empty
        Else
            Select Case Fld.FieldType
                Case ft_TinyInt, ft_SmallInt, ft_MediumInt, ft_Int, ft_BigInt, ft_Year
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CLng(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Float, ft_Double, ft_Decimal, ft_TimeStamp
                    If IsNumeric(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDbl(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case ft_Date, ft_DateTime, ft_Time
                    If IsDate(Trim(pRet(CurReg, Fld.FieldID))) Then
                        Fld.Value = CDate(Trim(pRet(CurReg, Fld.FieldID)))
                    Else
                        Fld.Value = Trim(pRet(CurReg, Fld.FieldID))
                    End If
                Case Else
                    Fld.Value = pRet(CurReg, Fld.FieldID)
            End Select
        End If
    End If
Next

End Sub


Public Property Let RT_BOF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RT_BOF = 5
    mvarRT_BOF = vData
End Property


Public Property Get RT_BOF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RT_BOF
    If CurReg = 1 Then mvarRT_BOF = True Else mvarRT_BOF = False
    If RT_RowCount = 0 Then mvarRT_BOF = True
    
    RT_BOF = mvarRT_BOF
End Property



Public Property Let RT_EOF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RT_EOF = 5
    mvarRT_EOF = vData
End Property


Public Property Get RT_EOF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RT_EOF
    
    If CurReg >= RT_RowCount + 1 Then mvarRT_EOF = True Else mvarRT_EOF = False
    If RT_RowCount = 0 Then mvarRT_EOF = True
    RT_EOF = mvarRT_EOF
End Property



Private Sub Class_Initialize()
Set RT_Field = New Fields
End Sub


Private Sub Class_Terminate()
Dim Ret As Long

Ret = API_mysql_free_result(myRec_res)
End Sub


