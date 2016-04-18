<%
'-----------------------------------------------------------------------------
'Nome: FormatarDataExibicao
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Formatar a data no padrao dd/mm/aaaa
'Parametros: dtData - data a ser formatada
'            intFormato: 1 - dd/mm/aaaa
'                        2 - dd/mm/aaaa hh:mm
'                        3 - dd/mm/aaaa hh:mm:ss
'-----------------------------------------------------------------------------
Public Function FormatarDataExibicao(ByVal dtData, ByVal intFormato)

    dim dtRetorno

	If Trim(" " & dtData) <> "" Then
		if isdate(dtData) then
			dtRetorno = Right(0 & Day(dtData), 2) & "/" & Right(0 & Month(dtData), 2) & "/" & Year(dtData)

            if intFormato = 2 then
                dtRetorno = dtRetorno & " " & Right(0 & hour(dtData), 2) & ":" & Right(0 & minute(dtData), 2)
            elseif intFormato = 3 then
                dtRetorno = dtRetorno & " " & Right(0 & hour(dtData), 2) & ":" & Right(0 & minute(dtData), 2) & ":" & Right(0 & second(dtData), 2)
            end if

        else
            dtRetorno = ""
		end if
    else
        dtRetorno = ""
	End If
		
	FormatarDataExibicao = dtRetorno

End Function

'-----------------------------------------------------------------------------
'Nome: FormatarDataGravacao
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Formatar a data no padrao aaaa-mm-dd
'Parametros: dtData - data a ser formatada
'            intFormato: 1 - aaaa-mm-dd
'                        2 - aaaa-mm-dd hh:mm
'                        3 - aaaa-mm-dd hh:mm:ss
'-----------------------------------------------------------------------------
Public Function FormatarDataGravacao(byVal dtData, ByVal intFormato)

    dim dtRetorno

	If Trim(" " & dtData) <> "" Then
		if isdate(dtData) then
			dtRetorno = Year(dtData) & "-" & month(dtData) & "-" & day(dtData)

            if intFormato = 2 then
                dtRetorno = dtRetorno & " " & Right(0 & hour(dtData), 2) & ":" & Right(0 & minute(dtData), 2)
            elseif intFormato = 3 then
                dtRetorno = dtRetorno & " " & Right(0 & hour(dtData), 2) & ":" & Right(0 & minute(dtData), 2) & ":" & Right(0 & second(dtData), 2)
            end if

            dtRetorno = "'" & dtRetorno & "'"

        else
            dtRetorno = "null"
		end if
    else
        dtRetorno = "null"
	End If

	FormatarDataGravacao = dtRetorno

End Function

'-----------------------------------------------------------------------------
'Nome: FormataHoraExibicao
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Formatar a hora no padrao hh:mm
'Parametros: dtData - data a ser formatada
'            intFormato: 1 - hh:mm
'                        2 - hh:mm:ss
'-----------------------------------------------------------------------------
Public Function FormataHoraExibicao(ByVal dtData, ByVal intFormato)

	If Trim(" " & dtData) <> "" Then
		if isdate(dtData) then
			dtData = right("0" & hour(dtData), 2) & ":" & right("0" & minute(dtData), 2)

            if intFormato = 2 then
                dtData = dtData & ":" & right("0" & second(dtData), 2)
            end if
         
        else
            dtData = ""
		end if
    else
        dtData = ""
	End If
		
	FormataHoraExibicao = dtData

End Function


'-----------------------------------------------------------------------------
'Nome: MontarFiltroData
'Autor: Adriano Carvalhaes
'Data: 28/06/2012
'Funcao: Retornar uma string para filtrar um campo de data
'Parametros: pstrCampoData   - nome do campo a ser filtrado
'            pstrDataInicial - data inicial
'            pstrDataFinal   - data final
'-----------------------------------------------------------------------------
function MontarFiltroData(Byval pstrCampoData, Byval pstrDataInicial, Byval pstrDataFinal)

	dim fstrRetorno
	
	if pstrDataInicial <> "" AND pstrDataFinal <> "" then
		fstrRetorno = pstrCampoData & " BETWEEN " & FormatarDataGravacao(pstrDataInicial & " 00:00", 2) & " AND " & FormatarDataGravacao(pstrDataFinal& " 23:59", 2) & " "
	elseif pstrDataInicial <> "" then
		fstrRetorno = pstrCampoData & " >= " & FormatarDataGravacao(pstrDataInicial & " 00:00", 2) 
	elseif pstrDataFinal <> "" then	
		fstrRetorno = pstrCampoData & " <= " & FormatarDataGravacao(pstrDataFinal & " 00:00", 2) 
	else
		fstrRetorno = ""
	end if
	
	MontarFiltroData = fstrRetorno

end function


'###########################
'# DESC : advanced format date
'# %A - AM or PM
'# %a - am or pm
'# %m - Month with leading zeroes (01 - 12)
'# %n - Month without leading zeroes (1 - 12)
'# %F - Month name (January - December)
'# %M - Three letter month name (Jan - Dec)
'# %d - Day with leading zeroes (01 - 31)
'# %j - Day without leading zeroes (1 - 31)
'# %H - Hour with leading zeroes (12 hour)
'# %h - Hour with leading zeroes (24 hour)
'# %G - Hour without leading zeroes (12 hour)
'# %g - Hour without leading zeroes (24 hour)
'# %i - Minute with leading zeroes (01 to 60)
'# %I - Minute without leading zeroes (1 to 60)
'# %s - Second with leading zeroes (01 to 60)
'# %S - Second without leading zeroes (1 to 60)
'# %L - Number of day of week (1 to 7)
'# %l - Name of day of week (Sunday to Saturday)
'# %D - Three letter name of day of week (Sun to Sat)
'# %O - Ordinal suffix (st, nd rd, th)
'# %U - UNIX Timestamp
'# %Y - Four digit year (2003)
'# %y - Two digit year (03)
'###########################
function FormatDate(format, intTimeStamp)
	'############################
	'# prevent month name error
	'############################
	Dim monthname():Redim monthname(12)
	monthname(1) = "January"
	monthname(2) = "February"
	monthname(3) = "March"
	monthname(4) = "April"
	monthname(5) = "May"
	monthname(6) = "June"
	monthname(7) = "July"
	monthname(8) = "August"
	monthname(9) = "September"
	monthname(10) = "October"
	monthname(11) = "November"
	monthname(12) = "December"
	'############################
	dim OrigDateFormat, A
	'############################
	'IntTimeStamp
	'############################
	if not (isnumeric(intTimeStamp)) then
		if isdate(intTimeStamp) then
			intTimeStamp = DateDiff("S", "01/01/1970 00:00:00", intTimeStamp)
		else
			Response.Write("Invalid Date")
			exit function
		end if
	end if
	'############################
	'original format
	'############################
	if (intTimeStamp=0) then
		OrigDateFormat = now()
	else
		OrigDateFormat = DateAdd("s", intTimeStamp, "01/01/1970 00:00:00")
	end if
	OrigDateFormat = trim(OrigDateFormat)
	'############################
	'get current partial format
	'############################
	dim dateDay:dateDay=DatePart("d",OrigDateFormat)
	dim dateMonth:dateMonth=DatePart("m",OrigDateFormat)
	dim dateYear:dateYear=DatePart("yyyy",OrigDateFormat)
	dim dateHour:dateHour=DatePart("h",OrigDateFormat)
	dim dateMinute:dateMinute=DatePart("m",OrigDateFormat)
	dim dateSecond:dateSecond=DatePart("s",OrigDateFormat)
	'############################
	'relpace
	'############################
	format = replace(format, "%Y", right(dateYear, 4))
	format = replace(format, "%y", right(dateYear, 2))
	format = replace(format, "%m", right("00"&dateMonth,2))
	format = replace(format, "%n", cint(dateMonth))
	format = replace(format, "%F", monthname(cint(dateMonth)))
	format = replace(format, "%M", left(monthname(cint(dateMonth)), 3))
	format = replace(format, "%d", right("00"&dateDay,2))
	format = replace(format, "%j", cint(dateDay))
	format = replace(format, "%h", right("00"&dateHour,2))
	format = replace(format, "%g", cint(dateHour))
	'############################
	'12 hours
	'############################
	if (cint(dateHour) > 12) then
		A = "PM"
	else
		A = "AM"
	end if

	format = replace(format, "%A", A)
	format = replace(format, "%a", lcase(A))

	if (A = "PM") then
		format = replace(format, "%H", right("00" & (dateHour-12), 2))
	else
		format = replace(format, "%H", right("00"&(dateHour),2))
	end if

	if (A = "PM") then
		format = replace(format, "%G", cint(dateHour)-12)
	else
		format = replace(format, "%G", cint(dateHour))
	end if
	'############################
	'time
	'############################
	format = replace(format, "%i", right("00"&dateMinute,2))
	format = replace(format, "%I", cint(dateMinute))
	format = replace(format, "%s", right("00"&dateSecond,2))
	format = replace(format, "%S", cint(dateSecond))
	format = replace(format, "%L", WeekDay(OrigDateFormat))
	format = replace(format, "%D", left(WeekDayName(WeekDay(OrigDateFormat)), 3))
	format = replace(format, "%l", WeekDayName(WeekDay(OrigDateFormat)))
	format = replace(format, "%U", intTimeStamp)
	format = replace(format, "11%O", "11th")
	format = replace(format, "1%O", "1st")
	format = replace(format, "12%O", "12th")
	format = replace(format, "2%O", "2nd")
	format = replace(format, "13%O", "13th")
	format = replace(format, "3%O", "3rd")
	format = replace(format, "%O", "th")
	'############################
	'return
	'############################
	FormatDate = format
end function

'-----------------------------------------------------------------------------
'Nome: SegundosParaHoras
'Autor: Adriano Carvalhaes
'Data: 30/04/2014
'Funcao: Transformar segundos em HH:MM:SS
'Parametros: vrSegundos - número em segundos
'-----------------------------------------------------------------------------
Public Function SegundosParaHoras(ByVal vrSegundos)

    dim dtRetorno

	dtRetorno = INT(vrSegundos/3600)&":"&RIGHT("00" & INT((vrSegundos-((INT(vrSegundos/3600))*3600))/60), 2)&":"&RIGHT("00" & vrSegundos-(((INT(vrSegundos/3600))*3600)+((INT((vrSegundos-((INT(vrSegundos/3600))*3600))/60))*60)), 2)
		
	SegundosParaHoras = dtRetorno

End Function

%>