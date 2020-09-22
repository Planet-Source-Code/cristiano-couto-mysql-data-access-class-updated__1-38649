VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarFieldName As String 'local copy
Public Enum en_FieldType
    ft_TinyInt = 1
    ft_SmallInt = 2
    ft_MediumInt = 3
    ft_Int = 4
    ft_BigInt = 5
    ft_Float = 6
    ft_Double = 7
    ft_Decimal = 8
    ft_Date = 9
    ft_DateTime = 10
    ft_TimeStamp = 11
    ft_Time = 12
    ft_Year = 13
    ft_Char = 14
    ft_VarChar = 15
    ft_TinyBlob = 16
    ft_TinyText = 17
    ft_Text = 18
    ft_Blob = 19
    ft_MediumBlob = 20
    ft_MediumText = 21
    ft_LongBlob = 22
    ft_LongText = 23
    ft_Enum = 24
    ft_Set = 25
End Enum
Private mvarFieldType As en_FieldType 'local copy
Private mvarFieldNull As Boolean 'local copy
Private mvarFieldDefault As String 'local copy
Private mvarFieldExtra As String 'local copy
'local variable(s) to hold property value(s)
Private mvarValue As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarFieldID As Long 'local copy
Public Property Let FieldID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldID = 5
    mvarFieldID = vData
End Property


Public Property Get FieldID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldID
    FieldID = mvarFieldID
End Property



Public Property Let Value(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


Public Property Set Value(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Value = Form1
    Set mvarValue = vData
End Property


Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    If IsObject(mvarValue) Then
        Set Value = mvarValue
    Else
        Value = mvarValue
    End If
End Property



Public Property Let FieldExtra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldExtra = 5
    mvarFieldExtra = vData
End Property


Public Property Get FieldExtra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldExtra
    FieldExtra = mvarFieldExtra
End Property



Public Property Let FieldDefault(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldDefault = 5
    mvarFieldDefault = vData
End Property


Public Property Get FieldDefault() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldDefault
    FieldDefault = mvarFieldDefault
End Property



Public Property Let FieldNull(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldNull = 5
    mvarFieldNull = vData
End Property


Public Property Get FieldNull() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldNull
    FieldNull = mvarFieldNull
End Property



Public Property Let FieldType(ByVal vData As en_FieldType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldType = 5
    mvarFieldType = vData
End Property


Public Property Get FieldType() As en_FieldType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldType
    FieldType = mvarFieldType
End Property




Public Property Let FieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property



