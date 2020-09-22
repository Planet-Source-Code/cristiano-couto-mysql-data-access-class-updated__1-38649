VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarTableName As String 'local copy
Private mvarTableType As String 'local copy
Public Enum en_row_format
    rf_Fixed = 1
    rf_Dynamic = 2
End Enum
Private mvarRowFormat As en_row_format 'local copy
Private mvarNumRows As Long 'local copy
Private mvarAverageRowLen As Long 'local copy
Private mvarDataLen As Long 'local copy
Private mvarMaxDataLen As Long 'local copy
Private mvarIndexLen As Long 'local copy
Private mvarDataFree As Long 'local copy
Private mvarAutoIncrement As Long 'local copy
Private mvarCreateTime As Date 'local copy
Private mvarUpdateTime As Date 'local copy
Private mvarCheckTime As Date 'local copy
Private mvarCreateOptions As String 'local copy
Private mvarComments As String 'local copy
Private mvarTableFields As Fields 'local copy
Public Property Set TableFields(ByVal vData As Fields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TableFields = Form1
    Set mvarTableFields = vData
End Property


Public Property Get TableFields() As Fields
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TableFields
    Set TableFields = mvarTableFields
End Property



Public Property Let TableComments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comments = 5
    mvarComments = vData
End Property


Public Property Get TableComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comments
    TableComments = mvarComments
End Property



Public Property Let TableCreateOptions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreateOptions = 5
    mvarCreateOptions = vData
End Property


Public Property Get TableCreateOptions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreateOptions
    TableCreateOptions = mvarCreateOptions
End Property



Public Property Let TableCheckTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CheckTime = 5
    mvarCheckTime = vData
End Property


Public Property Get TableCheckTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CheckTime
    TableCheckTime = mvarCheckTime
End Property



Public Property Let TableUpdateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpdateTime = 5
    mvarUpdateTime = vData
End Property


Public Property Get TableUpdateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpdateTime
    TableUpdateTime = mvarUpdateTime
End Property



Public Property Let TableCreateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreateTime = 5
    mvarCreateTime = vData
End Property


Public Property Get TableCreateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreateTime
    TableCreateTime = mvarCreateTime
End Property



Public Property Let TableAutoIncrement(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutoIncrement = 5
    mvarAutoIncrement = vData
End Property


Public Property Get TableAutoIncrement() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutoIncrement
    TableAutoIncrement = mvarAutoIncrement
End Property



Public Property Let TableDataFree(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataFree = 5
    mvarDataFree = vData
End Property


Public Property Get TableDataFree() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataFree
    TableDataFree = mvarDataFree
End Property



Public Property Let TableIndexLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexLen = 5
    mvarIndexLen = vData
End Property


Public Property Get TableIndexLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexLen
    TableIndexLen = mvarIndexLen
End Property



Public Property Let TableMaxDataLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxDataLen = 5
    mvarMaxDataLen = vData
End Property


Public Property Get TableMaxDataLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxDataLen
    TableMaxDataLen = mvarMaxDataLen
End Property



Public Property Let TableDataLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataLen = 5
    mvarDataLen = vData
End Property


Public Property Get TableDataLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataLen
    TableDataLen = mvarDataLen
End Property



Public Property Let TableAverageRowLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AverageRowLen = 5
    mvarAverageRowLen = vData
End Property


Public Property Get TableAverageRowLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AverageRowLen
    TableAverageRowLen = mvarAverageRowLen
End Property



Public Property Let TableNumRows(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumRows = 5
    mvarNumRows = vData
End Property


Public Property Get TableNumRows() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumRows
    TableNumRows = mvarNumRows
End Property



Public Property Let TableRowFormat(ByVal vData As en_row_format)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowFormat = 5
    mvarRowFormat = vData
End Property


Public Property Get TableRowFormat() As en_row_format
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowFormat
    TableRowFormat = mvarRowFormat
End Property




Public Property Let TableType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TableType = 5
    mvarTableType = vData
End Property


Public Property Get TableType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TableType
    TableType = mvarTableType
End Property



Public Property Let TableName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TableName = 5
    mvarTableName = vData
End Property


Public Property Get TableName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TableName
    TableName = mvarTableName
End Property



