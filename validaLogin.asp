<%
option explicit
%>
<!-- #include file = "include/configuracoes.asp" -->
<!-- #include file = "include/funcoesASP/md5.asp" -->
<!-- #include file = "include/funcoesASP/string.asp" -->
<!-- #include file = "include/funcoesASP/numero.asp" -->
<!-- #include file = "include/classesEntidades/cTentativaAcesso.asp" -->
<!-- #include file = "include/classesEntidades/cUsuario.asp" -->
<%
dim strLogin
dim strSenha
dim objTbUsuario
dim objTbMenuAcao
dim intQtdDiasExpiraSenha
dim dtUltimaAlteracaoSenha
dim objTbTentativaAcesso
dim objTbConfiguracaoSistema
dim objTbUsuarioProjeto

strLogin = request.form("Login") 
strSenha = request.form("Senha")

strSenha = md5(strSenha)

Set objTbUsuario = new clsTbUsuario
Set objConexaoBanco = new clsConexaoBanco
Set objTbTentativaAcesso = new clsTbTentativaAcesso

objTbTentativaAcesso.DsLogin = strLogin
objTbTentativaAcesso.CdIPOrigem = Request.ServerVariables("REMOTE_ADDR")

if objTbUsuario.VerificarSenha(strLogin, strSenha) = 1 then

    objTbTentativaAcesso.FgAcessoRealizado = 1
    objTbTentativaAcesso.Inserir

    Set objTbTentativaAcesso = nothing

    objTbUsuario.AdicionarFiltro "DsLogin", FormatarStringGravacao(strLogin)

    Set objRs = objTbUsuario.Listar()

    Session("IdUsuario") = objRs("IdUsuario")
    Session("NmUsuario") = objRs("NmUsuario")
    Session("DsEmail") = objRs("DsEmail")
    Session("DsLogin") = objRs("DsLogin")

    Set objRs = nothing

    Set objTbUsuario = nothing

    response.Redirect("admin/principal.asp")

else

    objTbTentativaAcesso.FgAcessoRealizado = 0
    objTbTentativaAcesso.Inserir

    Set objTbTentativaAcesso = nothing

    session.Abandon()
    response.Redirect("default.asp?retorno="&server.URLEncode("Login ou senha inv&aacute;lidos.")&"&erro=1")  
end if
%>