VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   2415
   ClientTop       =   1590
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim MyConn As New MySQLDTA.Connection
Dim MyRes As ResultTable
Dim host As String
Dim usr As String
Dim pwt As String
host = "localhost"
usr = "jacouto"
pwt = "laj42290"
If MyConn.MyConnect(host, usr, pwt) = True Then

    MyConn.MySelectDatabase "pe"
    Set MyRes = MyConn.MyExecute("insert into server values('paulo','silva')") 'The result query is set to Result Table Object
    Set MyRes = MyConn.MyExecute("select name, user from server") 'The result query is set to Result Table Object

    Do While Not MyRes.RT_EOF
        MsgBox MyRes("name") 'Here you put the column name
        MyRes.RT_MoveNext
    Loop
End If

End Sub


