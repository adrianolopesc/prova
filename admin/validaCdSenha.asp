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
dim objTbUsuario

cdSenhaAtual = request.QueryString("cdSenhaAtual")

Set objTbUsuario = new clsTbUsuario

cdSenhaAtual = md5(cdSenhaAtual)

if objTbUsuario.VerificarSenha(session("DsLogin"), cdSenhaAtual) <> 1 then
    response.write "false"
else
    response.write "true"
end if

Set objTbUsuario = nothing
%>