<%
'-----------------------------------------------------------------------------
'Nome: FormatarStringGravacao
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Formatar a string para ser gravada no banco de dados
'-----------------------------------------------------------------------------
function FormatarStringGravacao(ByVal pstrString)

	if trim("" & pstrString) = "" then
		pstrString = "NULL"
	else
		pstrString = "'" & replace(pstrString, "'", "''") & "'"
	end if

	FormatarStringGravacao = pstrString

end function

'-----------------------------------------------------------------------------
'Nome: FormatarStringBusca
'Autor: Adriano Carvalhaes
'Data: 12/07/2012
'Funcao: Formatar a string para ser buscada no banco de dados
'Parametros: pstrString- string a ser formatada
'pintModo: 1-Colocar % no inicio e final
'pintModo: 2-Colocar % no inicio e final e substituir espa�os por %
'pintModo: 3-Colocar % no final
'pintModo: 4-Colocar % no inicio
'pintModo: 5-N�o colocar %
'-----------------------------------------------------------------------------
function FormatarStringBusca(ByVal pstrString, ByVal pintModo)

    if pintModo = 1 then
        pstrString = "'%" & replace(pstrString, "'", "''") & "%'"
    elseif pintModo = 2 then
        pstrString = "'%" & replace(replace(pstrString, " ", "%"), "'", "''") & "%'"
    elseif pintModo = 3 then
        pstrString = "'" & replace(pstrString, "'", "''") & "%'"
    elseif pintModo = 4 then
        pstrString = "'%" & replace(pstrString, "'", "''") & "'"
    elseif pintModo = 5 then
        pstrString = "'" & replace(pstrString, "'", "''") & "'"
    end if

	FormatarStringBusca = pstrString

end function

'-----------------------------------------------------------------------------
'As funcoes abaixo nao estao sendo utilizadas
'-----------------------------------------------------------------------------

function ConverteAcentos(ByVal pstrString)

	pstrString = replace(pstrString, "&amp;", "&")	
	pstrString = replace(pstrString, "&#170;", "�")	
	pstrString = replace(pstrString, "&#193;", "�")		
	pstrString = replace(pstrString, "&#194;", "�")	
	pstrString = replace(pstrString, "&#195;", "�")
	pstrString = replace(pstrString, "&#196;", "�") 	
	pstrString = replace(pstrString, "&#199;", "�")		
	pstrString = replace(pstrString, "&#201;", "�")
	pstrString = replace(pstrString, "&#202;", "�")	
	pstrString = replace(pstrString, "&#205;", "�")	
	pstrString = replace(pstrString, "&#211;", "�")
	pstrString = replace(pstrString, "&#212;", "�")		
	pstrString = replace(pstrString, "&#213;", "�")	
	pstrString = replace(pstrString, "&#218;", "�")			
	pstrString = replace(pstrString, "&#224;", "�")								
	pstrString = replace(pstrString, "&#225;", "�")							
	pstrString = replace(pstrString, "&#226;", "�")								
	pstrString = replace(pstrString, "&#227;", "�")	
	pstrString = replace(pstrString, "&#231;", "�")					
	pstrString = replace(pstrString, "&#233;", "�")						
	pstrString = replace(pstrString, "&#232;", "�")						
	pstrString = replace(pstrString, "&#234;", "�")	
	pstrString = replace(pstrString, "&#237;", "�")	
	pstrString = replace(pstrString, "&#241;", "�")	
	pstrString = replace(pstrString, "&#243;", "�")		
	pstrString = replace(pstrString, "&#244;", "�")	
	pstrString = replace(pstrString, "&#245;", "�")
	pstrString = replace(pstrString, "&#250;", "�")		
	pstrString = replace(pstrString, "&#252;", "�")	
	pstrString = replace(pstrString, "&#220;", "�")		
	pstrString = replace(pstrString, "&#236;", "�")
	pstrString = replace(pstrString, "&#162;", "�")	 	
	pstrString = replace(pstrString, "&#163;", "�")
	pstrString = replace(pstrString, "&#164;", "�")
	pstrString = replace(pstrString, "&#165;", "�")
	pstrString = replace(pstrString, "&#166;", "�")
	pstrString = replace(pstrString, "&#167;", "�")
	pstrString = replace(pstrString, "&#168;", "�")
	pstrString = replace(pstrString, "&#169;", "�")
	pstrString = replace(pstrString, "&#171;", "�")
	pstrString = replace(pstrString, "&#172;", "�")
	pstrString = replace(pstrString, "&#173;", "�")
	pstrString = replace(pstrString, "&#174;", "�")
	pstrString = replace(pstrString, "&#175;", "�")
	pstrString = replace(pstrString, "&#176;", "�")
	pstrString = replace(pstrString, "&#177;", "�")
	pstrString = replace(pstrString, "&#178;", "�")
	pstrString = replace(pstrString, "&#179;", "�")
	pstrString = replace(pstrString, "&#180;", "�")
	pstrString = replace(pstrString, "&#181;", "�")
	pstrString = replace(pstrString, "&#182;", "�")
	pstrString = replace(pstrString, "&#183;", "�")
	pstrString = replace(pstrString, "&#184;", "�")
	pstrString = replace(pstrString, "&#185;", "�")
	pstrString = replace(pstrString, "&#186;", "�")
	pstrString = replace(pstrString, "&#187;", "�")
	pstrString = replace(pstrString, "&#188;", "�")
	pstrString = replace(pstrString, "&#189;", "�")
	pstrString = replace(pstrString, "&#190;", "�")
	pstrString = replace(pstrString, "&#191;", "�")
	pstrString = replace(pstrString, "&#192;", "�")
	pstrString = replace(pstrString, "&#197;", "�")
	pstrString = replace(pstrString, "&#198;", "�")
	pstrString = replace(pstrString, "&#200;", "�")
	pstrString = replace(pstrString, "&#203;", "�")
	pstrString = replace(pstrString, "&#204;", "�")
	pstrString = replace(pstrString, "&#206;", "�")
	pstrString = replace(pstrString, "&#207;", "�")
	pstrString = replace(pstrString, "&#208;", "�")
	pstrString = replace(pstrString, "&#209;", "�")
	pstrString = replace(pstrString, "&#214;", "�")
	pstrString = replace(pstrString, "&#215;", "�")
	pstrString = replace(pstrString, "&#216;", "�")
	pstrString = replace(pstrString, "&#217;", "�")
	pstrString = replace(pstrString, "&#219;", "�")
	pstrString = replace(pstrString, "&#221;", "�")
	pstrString = replace(pstrString, "&#222;", "�")
	pstrString = replace(pstrString, "&#223;", "�")
	pstrString = replace(pstrString, "&#228;", "�")
	pstrString = replace(pstrString, "&#229;", "�")
	pstrString = replace(pstrString, "&#230;", "�")
	pstrString = replace(pstrString, "&#235;", "�")
	pstrString = replace(pstrString, "&#238;", "�")
	pstrString = replace(pstrString, "&#239;", "�")
	pstrString = replace(pstrString, "&#240;", "�")
	pstrString = replace(pstrString, "&#242;", "�")
	pstrString = replace(pstrString, "&#246;", "�")
	pstrString = replace(pstrString, "&#247;", "�")
	pstrString = replace(pstrString, "&#248;", "�")
	pstrString = replace(pstrString, "&#249;", "�")
	pstrString = replace(pstrString, "&#251;", "�")
	pstrString = replace(pstrString, "&#253;", "�")
	pstrString = replace(pstrString, "&#254;", "�")
	pstrString = replace(pstrString, "&#255;", "�")
	pstrString = replace(pstrString, "&#159;", "�")
	pstrString = replace(pstrString, "&#140;", "�")
	ConverteAcentos = pstrString

end function

function ContaLinhas(pstrTexto)
	
	dim intContLinhas
	dim intLinhasCaracteres
	dim intPosicaoOcorrencia
		
	const cIntcaracteresLinha = 115
	

	do while inStr( pstrTexto, chr(13)) > 0
		
		intPosicaoOcorrencia = inStr( pstrTexto, chr(13))
		pstrTexto = mid(pstrTexto, intPosicaoOcorrencia+1)
		
		intLinhasCaracteres = intPosicaoOcorrencia / cIntcaracteresLinha
		intLinhasCaracteres = fix(intLinhasCaracteres)
		
		intContLinhas = intContLinhas + 1 + intLinhasCaracteres
	loop

	intContLinhas = intContLinhas + 1
	
	intPosicaoOcorrencia = len(pstrTexto)
	intLinhasCaracteres = intPosicaoOcorrencia / cIntcaracteresLinha
	intLinhasCaracteres = fix(intLinhasCaracteres)
	
	intContLinhas = intContLinhas + intLinhasCaracteres
	
	ContaLinhas = intContLinhas

end function

Public function NormalizarNome(BYVal pstrNome)

	dim strRetorno
	dim arrPalavras
	dim intIndice
	dim strPalavra
	
	strRetorno = lcase(replace(replace(replace(trim(pstrNome), ".", ""), "'", ""), chr(34), ""))
	arrPalavras = split(strRetorno, " ")
	for intIndice = 0 to ubound(arrPalavras)
		strPalavra = arrPalavras(intIndice)
		if strPalavra <> "de" and strPalavra <> "da" and strPalavra <> "do" and strPalavra <> "e" and strPalavra <> "dos" then
			arrPalavras(intIndice) = ucase(left(strPalavra, 1)) & mid(strPalavra, 2)
		end if
	next
	NormalizarNome = join(arrPalavras, " ")
	
end function

'Rotina para padronizar nomes de pessoas.
'Autor: Bruno Serafim
'Data : 01/04/2002
'----------------------------------------------------------------------------------------

' :: Function EliminaCaracteresEspeciais: substitui caracteres especiais como "'" ou "=" por espa�os em branco 
' de forma a evitar poss�veis problemas com manipula��o de strings e atualiza��es com banco de dados.

Public Function EliminaCaracteresEspeciais(ByVal Nome)

	Dim conj
	Dim intMax, maxConj
	Dim intIndiceI, intIndiceJ
	
	intMax  = Len(Nome)
	conj = array("/","*","-","+","=","'","!","@","#","$","%","&","(",")","-","_","{","}","[","]","�","�","?","^","~",",",";","<",">","|",".","�","�","�","�","�","�")
	maxConj = Ubound(conj)
	
	for intIndiceI = 1 to intMax
		for intIndiceJ = 0 to maxConj
			if (Mid(Nome, intIndiceI, 1) = conj(intIndiceJ)) then
				Nome = Mid(Nome, 1, intIndiceI-1) & " " & Mid(Nome, intIndiceI+1, intMax)
			end If                                                    
		next  
	next
	EliminaCaracteresEspeciais = Nome  
End Function

'----------------------------------------------------------------------------------------

' :: Function PadronizaNome: padroniza nomes de pessoas, deixando em letras mai�sculas e 
'    eliminando excessos de espa�os em branco.

Public Function EliminaEspacosDuplos(ByVal nome)

	Dim intContador
	Dim intmax
	Dim intPosicao
	
	nome = Trim(nome)
	nome = Ucase(nome)
	intmax  = Len(nome)
	
	for intContador = 1 to intmax
		if (Mid(nome,intContador,1) = " ") then
			intPosicao = intContador + 1
			while (Mid(nome,intPosicao,1) = " ")
				intPosicao = intPosicao + 1
			wend
			nome = Mid(nome,1,intContador) & Mid(nome, intPosicao, intmax)
			intmax  = Len(nome) 
		end if
	next
	EliminaEspacosDuplos = nome
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'      Nome: Fun��o Publica IntToStr(numero, quantidade)
'     Autor: Bruno Oliveira
'      Data: 09/01/03
' Descricao: Formata o numero para string e retorna a quantidade de zeros a esquerda
'   Retorno: string
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Function IntToStr(ByVal intNumber, ByVal intQuant)
	
	If intQuant = "" Then 
		intQuant = 0
	end if
	
	If Len(intNumber) >= cdbl(intQuant) Then
		IntToStr = left(intNumber, intQuant)
		Exit Function
	End If
	IntToStr = String(intQuant - Len(intNumber), "0") & intNumber
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'      Nome: Fun��o Publica FormatToString(strString, intLen)
'     Autor: Bruno Oliveira
'      Data: 09/01/03
' Descricao: Retorna a string de acordo com o tamanho completado de espacos a esquerda
'   Retorno: string
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Function FormatToString(ByVal strString, ByVal intLen)
	':: Validando o tamanho da string
	dim intStringLen
	
	strString = trim(" " & strString)
	
	intStringLen = Len(strString)
	
	':: Definido a quantidade de espacos a serem inseridos, ou se a string deve ser truncada
	If intLen > intStringLen Then
		':: Adicionando os espacos na string
		For i = 1 To (intLen - intStringLen)
			strString = strString & " "
		Next
		FormatToString = strString
	Else   ':: Truncando a string
		FormatToString = Left(strString, intLen)
	End If
	
End Function



public Function ConverteAcentoEntities(texto)
	texto = Replace(texto, "�", "&iexcl;")
	texto = Replace(texto, "�", "&iquest;")
	'texto = Replace(texto, "'", "&apos;")
	 
	texto = Replace(texto, "�", "&aacute;")
	texto = Replace(texto, "�", "&eacute;")
	texto = Replace(texto, "�", "&iacute;")
	texto = Replace(texto, "�", "&oacute;")
	texto = Replace(texto, "�", "&uacute;")
	texto = Replace(texto, "�", "&ntilde;")
	texto = Replace(texto, "�", "&ccedil;")
	 
	texto = Replace(texto, "�", "&Aacute;")
	texto = Replace(texto, "�", "&Eacute;")
	texto = Replace(texto, "�", "&Iacute;")
	texto = Replace(texto, "�", "&Oacute;")
	texto = Replace(texto, "�", "&Uacute;")
	texto = Replace(texto, "�", "&Ntilde;")
	texto = Replace(texto, "�", "&Ccedil;")
	 
	texto = Replace(texto, "�", "&agrave;")
	texto = Replace(texto, "�", "&egrave;")
	texto = Replace(texto, "�", "&igrave;")
	texto = Replace(texto, "�", "&ograve;")
	texto = Replace(texto, "�", "&ugrave;")
	 
	texto = Replace(texto, "�", "&Agrave;")
	texto = Replace(texto, "�", "&Egrave;")
	texto = Replace(texto, "�", "&Igrave;")
	texto = Replace(texto, "�", "&Ograve;")
	texto = Replace(texto, "�", "&Ugrave;")
	 
	texto = Replace(texto, "�", "&auml;")
	texto = Replace(texto, "�", "&euml;")
	texto = Replace(texto, "�", "&iuml;")
	texto = Replace(texto, "�", "&ouml;")
	texto = Replace(texto, "�", "&uuml;")
	 
	texto = Replace(texto, "�", "&Auml;")
	texto = Replace(texto, "�", "&Euml;")
	texto = Replace(texto, "�", "&Iuml;")
	texto = Replace(texto, "�", "&Ouml;")
	texto = Replace(texto, "�", "&Uuml;")
	 
	texto = Replace(texto, "�", "&acirc;")
	texto = Replace(texto, "�", "&ecirc;")
	texto = Replace(texto, "�", "&icirc;")
	texto = Replace(texto, "�", "&ocirc;")
	texto = Replace(texto, "�", "&ucirc;")
	 
	texto = Replace(texto, "�", "&Acirc;")
	texto = Replace(texto, "�", "&Ecirc;")
	texto = Replace(texto, "�", "&Icirc;")
	texto = Replace(texto, "�", "&Ocirc;")
	texto = Replace(texto, "�", "&Ucirc;")
	
	texto = Replace(texto, "�", "&atilde;")
	texto = Replace(texto, "�", "&otilde;")
	 
	texto = Replace(texto, "�", "&Atilde;")
	texto = Replace(texto, "�", "&Etilde;")
	ConverteAcentoEntities = texto
End Function


Function RemoveHTML( strText )
	Dim RegEx

	Set RegEx = New RegExp

	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True

	RemoveHTML = RegEx.Replace(strText, "")
End Function


Public Function ContarCaracteres(ByVal texto, ByVal caracter)

     Dim x

     x = Split(texto, caracter)
     ContarCaracteres = UBound(x)

End Function
%>