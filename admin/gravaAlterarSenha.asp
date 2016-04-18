<%
option explicit
%>
<!-- #include file = "../include/configuracoes.asp" -->
<!-- #include file = "../include/checaSessao.asp" -->
<!-- #include file = "../include/funcoesASP/md5.asp" -->
<!-- #include file = "../include/funcoesASP/string.asp" -->
<!-- #include file = "../include/funcoesASP/data.asp" -->
<!-- #include file = "../include/funcoesASP/numero.asp" -->
<!-- #include file = "../include/classesEntidades/cUsuario.asp" -->
<%
dim cdSenhaAtual
dim cdSenha
dim objTbUsuario
dim Mensagem
dim erro
dim idUsuario

Set objTbUsuario = new clsTbUsuario

cdSenhaAtual = request.Form("cdSenhaAtual")
cdSenha = request.Form("cdSenha")
idUsuario = request.Form("idUsuario")

cdSenhaAtual = md5(cdSenhaAtual)

if objTbUsuario.VerificarSenha(session("DsLogin"), cdSenhaAtual) <> 1 AND idUsuario = "" then
    Mensagem = "Senha inválida"
else

    cdSenha = md5(cdSenha)

    objTbUsuario.AlterarSenha idUsuario, cdSenha

    if not isNull(objTbUsuario.strErro) then
	    Mensagem = "Erro:" & objTbUsuario.strErro
    end if 

end if

Set objTbUsuario = nothing

erro = 1

if Mensagem = "" then
    Mensagem = "Registro salvo com sucesso!"
    erro = 0
end if

response.Redirect("alterarSenha.asp?Retorno="&server.URLEncode(Mensagem)&"&erro="&erro&"&idUsuario="&idUsuario)
%>

