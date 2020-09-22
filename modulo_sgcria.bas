Attribute VB_Name = "modulo_sgcria"
Public My As New MySQLDTA.Connection

Public Apache_Dir As String
Public MySQL_Dir As String
Public MySQL_Usr As String
Public MySQL_Pwd As String
Public MySQL_Hst As String

Public Boot_Time As Date
Public Boot_Periodo As Integer
Public Next_Boot As Date

Public MSG_Sessao As Session
Function MailValido(email As String, res_txt As String, res_id_usuario As Long) As Boolean
Dim Tbl As MySQLDTA.ResultTable
Dim Nom() As String

Dim RespSisNome As String
Dim RespSisMail As String
Dim RespSisFone As String

Set Tbl = My.MyExecute("select valor from config where variavel='ADM_SIS_NOME'")
If Tbl.RT_RowCount > 0 Then RespSisNome = Tbl("valor")

Set Tbl = My.MyExecute("select valor from config where variavel='ADM_SIS_EMAIL'")
If Tbl.RT_RowCount > 0 Then RespSisMail = Tbl("valor")

Set Tbl = My.MyExecute("select valor from config where variavel='ADM_SIS_FONE'")
If Tbl.RT_RowCount > 0 Then RespSisFone = Tbl("valor")

Nom = Split(email, "@")
Set Tbl = My.MyExecute("select usr.* from usuarios usr, email_usuario email where email.email_usr='" & email & "' and usr.id_usr=email.id_usr")
If Tbl.RT_RowCount = 0 Then
    res_txt = "Caro " & Nom(0) & "," & vbCrLf & vbCrLf
    res_txt = res_txt & "Você não possui permissão para fazer qualquer tipo de consulta. Por favor, solicite seu cadastramento junto ao responsável pelo sistema. Segue abaixo as informações do responsável do sistema: " & vbCrLf & vbCrLf
    res_txt = res_txt & RespSisNome & vbCrLf & vbCrLf
    res_txt = res_txt & "Email: " & RespSisMail & vbCrLf & vbCrLf
    res_txt = res_txt & "Fone: " & RespSisFone & vbCrLf & vbCrLf
    res_txt = res_txt & vbCrLf & vbCrLf
    res_txt = res_txt & "Atenciosamente"
    MailValido = False
Else
    res_id_usuario = Tbl("id_usr")
    If Tbl("inativo_usr") = 1 Then
        res_txt = "Caro " & Tbl("nome_usr") & "," & vbCrLf & vbCrLf
        res_txt = res_txt & "Você está cadastrado no sistema, mas por motivos de força maior seu cadastro foi desativado. Por favor, solicite a ativação do seu cadastro junto ao responsável pelo sistema. Segue abaixo as informações do responsável do sistema: " & vbCrLf & vbCrLf
        res_txt = res_txt & RespSisNome & vbCrLf & vbCrLf
        res_txt = res_txt & "Email: " & RespSisMail & vbCrLf & vbCrLf
        res_txt = res_txt & "Fone: " & RespSisFone & vbCrLf & vbCrLf
        res_txt = res_txt & vbCrLf & vbCrLf
        res_txt = res_txt & "Atenciosamente"
        MailValido = False
    Else
        MailValido = True
    End If
End If

End Function

Sub MandaEmail(email As String, assunto As String, texto As String, Anexos() As String)
Dim MIndex As Long
Dim NumAnexo As Long

Dim lMSG As MAPI.Message
Dim lFLD As MAPI.Folder
Dim lRCP As MAPI.Recipient
Dim lATC As MAPI.Attachment

NumAnexo = Ub_Vet(Anexos)

With frmMain
    Set lMSG = MSG_Sessao.Inbox.Messages.Add
    
    With lMSG
                
        If NumAnexo > 0 Then
            For x = 0 To NumAnexo
                Set lATC = .Attachments.Add
                lATC.ReadFromFile Anexos(x)
                lATC.Position = .Attachments.Count
                lATC.Type = MAPI.mapiAttachmentType.mapiFileData
            Next
        Else
            
        .Text = texto
        .Subject = assunto
        .Importance = ActMsgHigh
        
        Set lRCP = .Recipients.Add
        lRCP.Name = email
        lRCP.Resolve
        
        .Send
            
        End If
    
    End With
    
End With

End Sub


Sub Pausa(pSegundos As Long, Optional pEvents As Boolean = False)
Dim Tick As Long
Tick = Timer

Do
    If Tick + pSegundos < Timer Then Exit Do
    If pEvents Then DoEvents
Loop

End Sub



Function TranslateEmail(pTXT As String) As String
If Left(pTXT, 4) = "SMTP" Then
    TranslateEmail = Right(pTXT, Len(pTXT) - 5)
Else
    For x = Len(pTXT) To 1 Step -1
        If Mid(pTXT, x, 1) = "=" Then
            TranslateEmail = LCase(Right(pTXT, Len(pTXT) - x)) & "@ambev.com.br"
            Exit For
        End If
    Next
End If
End Function

Function Ub_Vet(pVet) As Long
On Error Resume Next
Ub_Vet = 0
Ub_Vet = UBound(pVet)
End Function


Function Lb_Vet(pVet) As Long
On Error Resume Next
Lb_Vet = 0
Lb_Vet = UBound(pVet)
End Function



