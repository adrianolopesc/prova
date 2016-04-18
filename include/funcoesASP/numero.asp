<%
'-----------------------------------------------------------------------------
'Nome: FormatarNumeroGravacao
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Formatar um numero no formato 999999.00
'Parametros: dblValor - valor a ser formatado
'            intCD - número de casas decimais
'-----------------------------------------------------------------------------
Public Function FormatarNumeroGravacao(byVal dblValor, byVal intCD)

    dim dblValorRetornoFuncao

	dblValorRetornoFuncao = dblValor

	if trim(" " & dblValorRetornoFuncao) <> "" AND trim(" " & dblValorRetornoFuncao) <> "NULL"  then
	
		if inStr(dblValorRetornoFuncao, "-") > 0 then
		
			dblValorRetornoFuncao = replace(dblValorRetornoFuncao, "-", "")
			dblValorRetornoFuncao = formatnumber(dblValorRetornoFuncao, intCD)
			
			dblValorRetornoFuncao  = replace(dblValorRetornoFuncao, ".", "")
			dblValorRetornoFuncao  = replace(dblValorRetornoFuncao, ",", ".")
			
			dblValorRetornoFuncao = "-" & dblValorRetornoFuncao

		else
		
			dblValorRetornoFuncao = formatnumber(dblValorRetornoFuncao, intCD)
			
			dblValorRetornoFuncao  = replace(dblValorRetornoFuncao, ".", "")
			dblValorRetornoFuncao  = replace(dblValorRetornoFuncao, ",", ".")
		
		end if
		
	else
		dblValorRetornoFuncao = "NULL"
	end if

	FormatarNumeroGravacao = dblValorRetornoFuncao

end function

'-----------------------------------------------------------------------------
'Nome: FormatarNumeroExbicao
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Formatar um numero no formato 999.999,00
'Parametros: dblValor - valor a ser formatado
'            intCD - número de casas decimais
'-----------------------------------------------------------------------------
Public Function FormatarNumeroExbicao(byVal dblValorRetornoFuncao, byVal intCD)

	if "" & dblValorRetornoFuncao <> "" then
	
		if inStr(dblValorRetornoFuncao, "-") > 0 then
			
			dblValorRetornoFuncao = "-" & formatNumber(replace(dblValorRetornoFuncao, "-", ""), intCD)	
		
		else
		
			dblValorRetornoFuncao = formatNumber(0 & dblValorRetornoFuncao, intCD)
		
		end if
	
	else
	
		dblValorRetornoFuncao = ""
	
	end if
	
	FormatarNumeroExbicao = dblValorRetornoFuncao

end function

'-----------------------------------------------------------------------------
'As funcoes abaixo nao estao sendo utilizadas
'-----------------------------------------------------------------------------
Public Function IntToStr(intNumber, intQuant)
	If intQuant = "" Then intQuant = 0
	If Len(intNumber) >= cdbl(intQuant) Then
		IntToStr = left(intNumber, intQuant)
		Exit Function
	End If
	IntToStr = String(intQuant - Len(intNumber), "0") & intNumber
End Function
%>