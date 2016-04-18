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
'pintModo: 2-Colocar % no inicio e final e substituir espaços por %
'pintModo: 3-Colocar % no final
'pintModo: 4-Colocar % no inicio
'pintModo: 5-Não colocar %
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
	pstrString = replace(pstrString, "&#170;", "ª")	
	pstrString = replace(pstrString, "&#193;", "Á")		
	pstrString = replace(pstrString, "&#194;", "Â")	
	pstrString = replace(pstrString, "&#195;", "Ã")
	pstrString = replace(pstrString, "&#196;", "Ä") 	
	pstrString = replace(pstrString, "&#199;", "Ç")		
	pstrString = replace(pstrString, "&#201;", "É")
	pstrString = replace(pstrString, "&#202;", "Ê")	
	pstrString = replace(pstrString, "&#205;", "Í")	
	pstrString = replace(pstrString, "&#211;", "Ó")
	pstrString = replace(pstrString, "&#212;", "Ô")		
	pstrString = replace(pstrString, "&#213;", "Õ")	
	pstrString = replace(pstrString, "&#218;", "Ú")			
	pstrString = replace(pstrString, "&#224;", "à")								
	pstrString = replace(pstrString, "&#225;", "á")							
	pstrString = replace(pstrString, "&#226;", "â")								
	pstrString = replace(pstrString, "&#227;", "ã")	
	pstrString = replace(pstrString, "&#231;", "ç")					
	pstrString = replace(pstrString, "&#233;", "é")						
	pstrString = replace(pstrString, "&#232;", "è")						
	pstrString = replace(pstrString, "&#234;", "ê")	
	pstrString = replace(pstrString, "&#237;", "í")	
	pstrString = replace(pstrString, "&#241;", "ñ")	
	pstrString = replace(pstrString, "&#243;", "ó")		
	pstrString = replace(pstrString, "&#244;", "ô")	
	pstrString = replace(pstrString, "&#245;", "õ")
	pstrString = replace(pstrString, "&#250;", "ú")		
	pstrString = replace(pstrString, "&#252;", "ü")	
	pstrString = replace(pstrString, "&#220;", "Ü")		
	pstrString = replace(pstrString, "&#236;", "ì")
	pstrString = replace(pstrString, "&#162;", "¢")	 	
	pstrString = replace(pstrString, "&#163;", "£")
	pstrString = replace(pstrString, "&#164;", "¤")
	pstrString = replace(pstrString, "&#165;", "¥")
	pstrString = replace(pstrString, "&#166;", "¦")
	pstrString = replace(pstrString, "&#167;", "§")
	pstrString = replace(pstrString, "&#168;", "¨")
	pstrString = replace(pstrString, "&#169;", "©")
	pstrString = replace(pstrString, "&#171;", "«")
	pstrString = replace(pstrString, "&#172;", "¬")
	pstrString = replace(pstrString, "&#173;", "­")
	pstrString = replace(pstrString, "&#174;", "®")
	pstrString = replace(pstrString, "&#175;", "¯")
	pstrString = replace(pstrString, "&#176;", "°")
	pstrString = replace(pstrString, "&#177;", "±")
	pstrString = replace(pstrString, "&#178;", "²")
	pstrString = replace(pstrString, "&#179;", "³")
	pstrString = replace(pstrString, "&#180;", "´")
	pstrString = replace(pstrString, "&#181;", "µ")
	pstrString = replace(pstrString, "&#182;", "¶")
	pstrString = replace(pstrString, "&#183;", "·")
	pstrString = replace(pstrString, "&#184;", "¸")
	pstrString = replace(pstrString, "&#185;", "¹")
	pstrString = replace(pstrString, "&#186;", "º")
	pstrString = replace(pstrString, "&#187;", "»")
	pstrString = replace(pstrString, "&#188;", "¼")
	pstrString = replace(pstrString, "&#189;", "½")
	pstrString = replace(pstrString, "&#190;", "¾")
	pstrString = replace(pstrString, "&#191;", "¿")
	pstrString = replace(pstrString, "&#192;", "À")
	pstrString = replace(pstrString, "&#197;", "Å")
	pstrString = replace(pstrString, "&#198;", "Æ")
	pstrString = replace(pstrString, "&#200;", "È")
	pstrString = replace(pstrString, "&#203;", "Ë")
	pstrString = replace(pstrString, "&#204;", "Ì")
	pstrString = replace(pstrString, "&#206;", "Î")
	pstrString = replace(pstrString, "&#207;", "Ï")
	pstrString = replace(pstrString, "&#208;", "Ð")
	pstrString = replace(pstrString, "&#209;", "Ñ")
	pstrString = replace(pstrString, "&#214;", "Ö")
	pstrString = replace(pstrString, "&#215;", "×")
	pstrString = replace(pstrString, "&#216;", "Ø")
	pstrString = replace(pstrString, "&#217;", "Ù")
	pstrString = replace(pstrString, "&#219;", "Û")
	pstrString = replace(pstrString, "&#221;", "Ý")
	pstrString = replace(pstrString, "&#222;", "Þ")
	pstrString = replace(pstrString, "&#223;", "ß")
	pstrString = replace(pstrString, "&#228;", "ä")
	pstrString = replace(pstrString, "&#229;", "å")
	pstrString = replace(pstrString, "&#230;", "æ")
	pstrString = replace(pstrString, "&#235;", "ë")
	pstrString = replace(pstrString, "&#238;", "î")
	pstrString = replace(pstrString, "&#239;", "ï")
	pstrString = replace(pstrString, "&#240;", "ð")
	pstrString = replace(pstrString, "&#242;", "ò")
	pstrString = replace(pstrString, "&#246;", "ö")
	pstrString = replace(pstrString, "&#247;", "÷")
	pstrString = replace(pstrString, "&#248;", "ø")
	pstrString = replace(pstrString, "&#249;", "ù")
	pstrString = replace(pstrString, "&#251;", "û")
	pstrString = replace(pstrString, "&#253;", "ý")
	pstrString = replace(pstrString, "&#254;", "þ")
	pstrString = replace(pstrString, "&#255;", "ÿ")
	pstrString = replace(pstrString, "&#159;", "Ÿ")
	pstrString = replace(pstrString, "&#140;", "Œ")
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

' :: Function EliminaCaracteresEspeciais: substitui caracteres especiais como "'" ou "=" por espaços em branco 
' de forma a evitar possíveis problemas com manipulação de strings e atualizações com banco de dados.

Public Function EliminaCaracteresEspeciais(ByVal Nome)

	Dim conj
	Dim intMax, maxConj
	Dim intIndiceI, intIndiceJ
	
	intMax  = Len(Nome)
	conj = array("/","*","-","+","=","'","!","@","#","$","%","&","(",")","-","_","{","}","[","]","º","ª","?","^","~",",",";","<",">","|",".","¹","²","³","£","¢","¬")
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

' :: Function PadronizaNome: padroniza nomes de pessoas, deixando em letras maiúsculas e 
'    eliminando excessos de espaços em branco.

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
'      Nome: Função Publica IntToStr(numero, quantidade)
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
'      Nome: Função Publica FormatToString(strString, intLen)
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
	texto = Replace(texto, "¡", "&iexcl;")
	texto = Replace(texto, "¿", "&iquest;")
	'texto = Replace(texto, "'", "&apos;")
	 
	texto = Replace(texto, "á", "&aacute;")
	texto = Replace(texto, "é", "&eacute;")
	texto = Replace(texto, "í", "&iacute;")
	texto = Replace(texto, "ó", "&oacute;")
	texto = Replace(texto, "ú", "&uacute;")
	texto = Replace(texto, "ñ", "&ntilde;")
	texto = Replace(texto, "ç", "&ccedil;")
	 
	texto = Replace(texto, "Á", "&Aacute;")
	texto = Replace(texto, "É", "&Eacute;")
	texto = Replace(texto, "Í", "&Iacute;")
	texto = Replace(texto, "Ó", "&Oacute;")
	texto = Replace(texto, "Ú", "&Uacute;")
	texto = Replace(texto, "Ñ", "&Ntilde;")
	texto = Replace(texto, "Ç", "&Ccedil;")
	 
	texto = Replace(texto, "à", "&agrave;")
	texto = Replace(texto, "è", "&egrave;")
	texto = Replace(texto, "ì", "&igrave;")
	texto = Replace(texto, "ò", "&ograve;")
	texto = Replace(texto, "ù", "&ugrave;")
	 
	texto = Replace(texto, "À", "&Agrave;")
	texto = Replace(texto, "È", "&Egrave;")
	texto = Replace(texto, "Ì", "&Igrave;")
	texto = Replace(texto, "Ò", "&Ograve;")
	texto = Replace(texto, "Ù", "&Ugrave;")
	 
	texto = Replace(texto, "ä", "&auml;")
	texto = Replace(texto, "ë", "&euml;")
	texto = Replace(texto, "ï", "&iuml;")
	texto = Replace(texto, "ö", "&ouml;")
	texto = Replace(texto, "ü", "&uuml;")
	 
	texto = Replace(texto, "Ä", "&Auml;")
	texto = Replace(texto, "Ë", "&Euml;")
	texto = Replace(texto, "Ï", "&Iuml;")
	texto = Replace(texto, "Ö", "&Ouml;")
	texto = Replace(texto, "Ü", "&Uuml;")
	 
	texto = Replace(texto, "â", "&acirc;")
	texto = Replace(texto, "ê", "&ecirc;")
	texto = Replace(texto, "î", "&icirc;")
	texto = Replace(texto, "ô", "&ocirc;")
	texto = Replace(texto, "û", "&ucirc;")
	 
	texto = Replace(texto, "Â", "&Acirc;")
	texto = Replace(texto, "Ê", "&Ecirc;")
	texto = Replace(texto, "Î", "&Icirc;")
	texto = Replace(texto, "Ô", "&Ocirc;")
	texto = Replace(texto, "Û", "&Ucirc;")
	
	texto = Replace(texto, "ã", "&atilde;")
	texto = Replace(texto, "õ", "&otilde;")
	 
	texto = Replace(texto, "Ã", "&Atilde;")
	texto = Replace(texto, "Õ", "&Etilde;")
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