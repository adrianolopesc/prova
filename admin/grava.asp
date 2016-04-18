<%option explicit%>
<!-- #include file = "../include/configuracoes.asp" -->
<!-- #include file = "../include/checaSessao.asp" -->
<!-- #include file = "../include/classesEntidades/cSiteMenu.asp" -->
<!-- #include file = "../include/funcoesasp/numero.asp" -->
<!-- #include file = "../include/funcoesasp/string.asp" -->
<!-- #include file = "../include/funcoesasp/data.asp" -->

<%
dim IdSiteMenu
dim DsSiteMenu
dim NrOrdem
dim objTbSiteMenu
dim Mensagem
dim DsConteudo
dim erro
dim fgAtivo

IdSiteMenu = request.form("IdSiteMenu") 
DsSiteMenu = request.form("DsSiteMenu") 
DsConteudo = request.form("DsConteudo") 
NrOrdem = request.form("NrOrdem") 
fgAtivo = request.form("fgAtivo") 

if fgAtivo = "" then
    fgAtivo = "0"
end if

Set objTbSiteMenu = new clsTbSiteMenu

objTbSiteMenu.DsSiteMenu = DsSiteMenu
objTbSiteMenu.NrOrdem = NrOrdem
objTbSiteMenu.DsConteudo = DsConteudo
objTbSiteMenu.fgAtivo = fgAtivo

Set objConexaoBanco = new clsConexaoBanco

if IdSiteMenu = "" then

    objTbSiteMenu.Inserir
    IdSiteMenu = objTbSiteMenu.IdSiteMenu

    if not isNull(objTbSiteMenu.strErro) then
	    Mensagem = "Erro:" & objTbSiteMenu.strErro
    else
        Mensagem = ""
    end if

else

    objTbSiteMenu.Atualizar(IdSiteMenu)

    if not isNull(objTbSiteMenu.strErro) then
	    Mensagem = "Erro:" & objTbSiteMenu.strErro
    else
        Mensagem = ""
    end if

end if

erro = 1

if Mensagem = "" then
    Mensagem = "Registro salvo com sucesso!"
    erro = 0
end if

Set objTbSiteMenu = nothing

Set objConexaoBanco = nothing

response.Redirect("edicao.asp?IdSiteMenu="&IdSiteMenu&"&Retorno="&server.URLEncode(Mensagem)&"&erro="&erro)
%>