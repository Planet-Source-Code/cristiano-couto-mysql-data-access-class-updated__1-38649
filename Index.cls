VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarIndexName As String 'local copy
Private mvarIndexPrimary As Boolean 'local copy
Private mvarIndexUnique As Boolean 'local copy
Private mvarIndexFields As Fields 'local copy
Private mvarIndexCollation As String 'local copy
Private mvarIndexCardinality As Long 'local copy
Private mvarIndexSubPart As String 'local copy
Private mvarIndexPacked As String 'local copy
Private mvarIndexComment As String 'local copy
Public Property Let IndexComment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexComment = 5
    mvarIndexComment = vData
End Property


Public Property Get IndexComment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexComment
    IndexComment = mvarIndexComment
End Property



Public Property Let IndexPacked(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexPacked = 5
    mvarIndexPacked = vData
End Property


Public Property Get IndexPacked() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexPacked
    IndexPacked = mvarIndexPacked
End Property



Public Property Let IndexSubPart(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexSubPart = 5
    mvarIndexSubPart = vData
End Property


Public Property Get IndexSubPart() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexSubPart
    IndexSubPart = mvarIndexSubPart
End Property



Public Property Let IndexCardinality(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexCardinality = 5
    mvarIndexCardinality = vData
End Property


Public Property Get IndexCardinality() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexCardinality
    IndexCardinality = mvarIndexCardinality
End Property



Public Property Let IndexCollation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexCollation = 5
    mvarIndexCollation = vData
End Property


Public Property Get IndexCollation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexCollation
    IndexCollation = mvarIndexCollation
End Property



Public Property Set IndexFields(ByVal vData As Fields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.IndexFields = Form1
    Set mvarIndexFields = vData
End Property


Public Property Get IndexFields() As Fields
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexFields
    Set IndexFields = mvarIndexFields
End Property



Public Property Let IndexUnique(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexUnique = 5
    mvarIndexUnique = vData
End Property


Public Property Get IndexUnique() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexUnique
    IndexUnique = mvarIndexUnique
End Property



Public Property Let IndexPrimary(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexPrimary = 5
    mvarIndexPrimary = vData
End Property


Public Property Get IndexPrimary() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexPrimary
    IndexPrimary = mvarIndexPrimary
End Property



Public Property Let IndexName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexName = 5
    mvarIndexName = vData
End Property


Public Property Get IndexName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexName
    IndexName = mvarIndexName
End Property



