<%
class clsTbTentativaAcesso

	dim strInsertCampos
	dim strNomeCampo
	dim strInsertDados
	dim strErro
	dim IdTentativaAcesso
	dim DsLogin
	dim DtInclusao
	dim FgAcessoRealizado
	dim CdIPOrigem


	private Sub Class_Initialize()

		strErro = null
		IdTentativaAcesso = null
		DsLogin = null
		DtInclusao = null
		FgAcessoRealizado = null
		CdIPOrigem = null

	end Sub

	function Inserir()

		if isnull(DsLogin) OR DsLogin = "" then 
			strErro = "Campo DsLogin não foi preenchido"
		end if
		if isnull(FgAcessoRealizado) OR FgAcessoRealizado = "" then 
			strErro = "Campo FgAcessoRealizado não foi preenchido"
		end if
		if isnull(CdIPOrigem) OR CdIPOrigem = "" then 
			strErro = "Campo CdIPOrigem não foi preenchido"
		end if

		strInsertCampos = " @DsLogin = " & FormatarStringGravacao(DsLogin)
		strInsertCampos = strInsertCampos & ", @FgAcessoRealizado = " & FormatarStringGravacao(FgAcessoRealizado)
		strInsertCampos = strInsertCampos & ", @CdIPOrigem = " & FormatarStringGravacao(CdIPOrigem)

        if isnull(strErro) then
		    strSQL = "EXEC dbo.spInsertTbTentativaAcesso " & strInsertCampos
		    Set objRs = objConexaoBanco.QuerySQL(strSQL)
            if not(objRs.eof) then
                if "" & objRs(0) = "" then
                    strErro = "Erro ao inserir o registro."
                else
                    IdTentativaAcesso = objRs(0)
                end if
            else
                strErro = "Erro ao inserir o registro."
            end if
            objRs.close
            Set objRs = nothing 
        end if


	end function

end class


%>