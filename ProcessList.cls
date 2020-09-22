VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarMyID As Long 'local copy
Private mvarMyUser As String 'local copy
Private mvarMyHost As String 'local copy
Private mvarMyDB As String 'local copy
Private mvarMyCommand As String 'local copy
Private mvarMyTime As Long 'local copy
Private mvarMyState As String 'local copy
Private mvarMyInfo As String 'local copy
Sub MyKillProcess()

End Sub

Public Property Let MyInfo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyInfo = 5
    mvarMyInfo = vData
End Property


Public Property Get MyInfo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyInfo
    MyInfo = mvarMyInfo
End Property



Public Property Let MyState(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyState = 5
    mvarMyState = vData
End Property


Public Property Get MyState() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyState
    MyState = mvarMyState
End Property



Public Property Let MyTime(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyTime = 5
    mvarMyTime = vData
End Property


Public Property Get MyTime() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyTime
    MyTime = mvarMyTime
End Property



Public Property Let MyCommand(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyCommand = 5
    mvarMyCommand = vData
End Property


Public Property Get MyCommand() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyCommand
    MyCommand = mvarMyCommand
End Property



Public Property Let MyDB(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyDB = 5
    mvarMyDB = vData
End Property


Public Property Get MyDB() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyDB
    MyDB = mvarMyDB
End Property



Public Property Let MyHost(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyHost = 5
    mvarMyHost = vData
End Property


Public Property Get MyHost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyHost
    MyHost = mvarMyHost
End Property



Public Property Let MyUser(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyUser = 5
    mvarMyUser = vData
End Property


Public Property Get MyUser() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyUser
    MyUser = mvarMyUser
End Property



Public Property Let MyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyID = 5
    mvarMyID = vData
End Property


Public Property Get MyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyID
    MyID = mvarMyID
End Property



