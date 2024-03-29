VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VariableState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarVarName As String 'local copy
Private mvarVarValue As String 'local copy
Public Property Let VarValue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VarValue = 5
    mvarVarValue = vData
End Property


Public Property Get VarValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VarValue
    VarValue = mvarVarValue
End Property



Public Property Let VarName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VarName = 5
    mvarVarName = vData
End Property


Public Property Get VarName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VarName
    VarName = mvarVarName
End Property



