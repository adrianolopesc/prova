
<%
class clsTbSiteMenu

	dim strInsertCampos
	dim strNomeCampo
	dim strInsertDados
	dim strErro
	dim strWhere
	dim IdSiteMenu
	dim DsSiteMenu
	dim DsConteudo
	dim DtInclusao
	dim DtAlteracao
	dim IdUsuarioInclusao
	dim IdUsuarioAlteracao
    dim FgAtivo
    dim nrOrdem

	private Sub Class_Initialize()

		strWhere = null
		strErro = null
		IdSiteMenu = null
		DsSiteMenu = null
		DsConteudo = null
		DtInclusao = null
		DtAlteracao = null
		IdUsuarioInclusao = null
		IdUsuarioAlteracao = null
        FgAtivo = null
        nrOrdem = null

	end Sub

	function Inserir()

		dim strInsertCampos
        dim objRs

		if not isnull(IdSiteMenu) then 
			strErro = "Campo IdSiteMenu não pode ser preenchido"
		end if
		if isnull(DsSiteMenu) OR DsSiteMenu = "" then 
			strErro = "Campo DsSiteMenu não foi preenchido"
		end if 
		if isnull(DsConteudo) OR DsConteudo = "" then 
			strErro = "Campo DsConteudo não foi preenchido"
		end if 
        if isnull(nrOrdem) then
			strErro = "Campo nrOrdem não foi preenchido"
        end if

        strInsertCampos = " @IdSiteMenu = NULL"
		strInsertCampos = strInsertCampos & ", @DsSiteMenu = " & FormatarStringGravacao(DsSiteMenu)
		strInsertCampos = strInsertCampos & ", @DsConteudo = " & FormatarStringGravacao(DsConteudo)
		strInsertCampos = strInsertCampos & ", @FgAtivo = " & FormatarStringGravacao(FgAtivo)
		strInsertCampos = strInsertCampos & ", @nrOrdem = " & FormatarStringGravacao(nrOrdem)
		strInsertCampos = strInsertCampos & ", @IdUsuario = " & FormatarNumeroGravacao(session("IdUsuario"), 0)

        if isnull(strErro) then
		    strSQL = "EXEC dbo.spEditSiteMenu " & strInsertCampos
		    Set objRs = objConexaoBanco.QuerySQL(strSQL)
            if not(objRs.eof) then
                if "" & objRs(0) = "" then
                    strErro = "Erro ao inserir o registro."
                else
                    IdSiteMenu = objRs(0)
                end if
            else
                strErro = "Erro ao inserir o registro."
            end if
            objRs.close
            Set objRs = nothing 
        end if

	end function

	function Atualizar(byval IdSiteMenu)

		dim strUpdate
        dim objRs

        strUpdate = " @IdSiteMenu = " & IdSiteMenu
		strUpdate = strUpdate & ", @DsSiteMenu = " & FormatarStringGravacao(DsSiteMenu)
		strUpdate = strUpdate & ", @DsConteudo = " & FormatarStringGravacao(DsConteudo)
		strUpdate = strUpdate & ", @FgAtivo = " & FormatarStringGravacao(FgAtivo)
		strUpdate = strUpdate & ", @nrOrdem = " & FormatarStringGravacao(nrOrdem)
		strUpdate = strUpdate & ", @IdUsuario = " & FormatarNumeroGravacao(session("IdUsuario"), 0)

		strSQL = "EXEC dbo.spEditSiteMenu " & strUpdate
		Set objRs = objConexaoBanco.QuerySQL(strSQL)
        if not(objRs.eof) then
            if "" & objRs(0) = "" then
                strErro = "Erro ao atualizar o registro."
            end if
        else
            strErro = "Erro ao atualizar o registro."
        end if
        objRs.close
        Set objRs = nothing 

	end function

	function Listar()

		dim objRs

		strSQL = "exec dbo.spListarMenus" & strWhere 
		Set objRs = objConexaoBanco.QuerySQL(strSQL)
		Set Listar = objRs

	end function

	function ListarMenusAtivos()

		dim objRs

		strSQL = "exec dbo.spListarMenusAtivos"
		Set objRs = objConexaoBanco.QuerySQL(strSQL)
		Set ListarMenusAtivos = objRs

	end function

	function AdicionarFiltro(byval pstrCampo, byval pstrFiltro)

    	strWhere = strWhere & " @" & pstrCampo & " = " & pstrFiltro 

	end function

end class

%>