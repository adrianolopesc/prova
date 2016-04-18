<%
'-----------------------------------------------------------------------------
'Nome: CriptografarString
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Funcao para criptografar uma string
'-----------------------------------------------------------------------------
function CriptografarString(Senha) 

    Dim strTempChar
    dim strSenhaCript
    dim i

    strSenhaCript = ""

    For i = 1 To Len(Senha)

        If Asc(Mid(Senha, i, 1)) < 128 Then
            strTempChar = Asc(Mid(Senha, i, 1)) + 128
        ElseIf Asc(Mid(Senha, i, 1)) > 128 Then
            strTempChar = Asc(Mid(Senha, i, 1)) - 128
        End If

        strSenhaCript = strSenhaCript & Chr(strTempChar)

    Next

    CriptografarString = strSenhaCript

End Function

 %>