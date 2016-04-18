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
dim cdNovaSenha
dim objTbUsuario

cdNovaSenha = request.QueryString("cdSenha")

Set objTbUsuario = new clsTbUsuario

cdNovaSenha = md5(cdNovaSenha)

if objTbUsuario.VerificarSenha(session("DsLogin"), cdNovaSenha) <> 1 then
    response.write "true"
else
    response.write "false"
end if

Set objTbUsuario = nothing
%>