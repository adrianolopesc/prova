<%
class clsTbUsuario

	dim strInsertCampos
	dim strNomeCampo
	dim strInsertDados
	dim strErro
	dim strWhere
	dim IdUsuario
	dim NmUsuario
	dim DsEmail
	dim DsLogin
	dim CdSenha
	dim FgAtivo
	dim DtAlteracaoSenha
	dim DtUltimoAcesso
	dim DtInclusao
	dim DtAlteracao
	dim IdUsuarioInclusao
	dim IdUsuarioAlteracao


	private Sub Class_Initialize()

		strWhere = null
		strErro = null
		IdUsuario = null
		NmUsuario = null
		DsEmail = null
		DsLogin = null
		CdSenha = null
		FgAtivo = null
		DtAlteracaoSenha = null
		DtUltimoAcesso = null
		DtInclusao = null
		DtAlteracao = null
		IdUsuarioInclusao = null
		IdUsuarioAlteracao = null

	end Sub

	function Listar()

        dim objRs

		strSQL = "exec dbo.spListarUsuarios" & strWhere 
		Set objRs = objConexaoBanco.QuerySQL(strSQL)
		Set Listar = objRs

	end function

	function AdicionarFiltro(byval pstrCampo, byval pstrFiltro)

    	strWhere = strWhere & " @" & pstrCampo & " = " & FormatarStringGravacao(pstrFiltro)

	end function

    function VerificarSenha(Byval pstrDsLogin, pstrCdSenha)

        dim objRs

        strSQL = "EXEC dbo.spVerificarSenha " & FormatarStringGravacao(pstrDsLogin) & ", " & FormatarStringGravacao(pstrCdSenha)
        Set objRs = objConexaoBanco.QuerySQL(strSQL)
        if not objRs.eof then
            VerificarSenha = cdbl(objRs(0))
        else
            VerificarSenha = 0
        end if
        objRs.close
        Set objRs = nothing

    end function

	function AlterarSenha(byval idUsuario, byVal cdSenha)

		dim strUpdate
        dim objRs

        strUpdate = " @IdUsuario = " & FormatarNumeroGravacao(session("IdUsuario"), 0)
		strUpdate = strUpdate & ", @cdSenha = " & FormatarStringGravacao(cdSenha)

		strSQL = "EXEC dbo.spAlterarSenha " & strUpdate
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


end class


%>