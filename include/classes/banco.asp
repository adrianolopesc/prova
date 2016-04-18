<%
class clsConexaoBanco  
   dim Conexao, bolDebug, intTotalRegistros

    Private Sub Class_Initialize()

        Set Conexao = Server.CreateObject("ADODB.Connection")

        Conexao.connectiontimeout = 120
        Conexao.Open buscarConexao()
        Conexao.CommandTimeout = 360 

    End Sub

   Private Sub Class_Terminate()
     Conexao.Close
     Set Conexao = nothing
   End Sub

   function ExecuteCMD (pstrSQL)
      	Set ExecuteCMD = conexao.execute (pstrSQL)
   end function 

   function QuerySQL(pstrSQL)

      dim objRs 
	  
      set objRs                  = server.CreateObject("ADODB.recordset")
  	  objRs.CursorLocation       = 3
	  
	  set objRs.ActiveConnection = conexao

      'response.Write strSQL
	  objRs.open pstrSQL

	  set objRs.ActiveConnection = nothing 

      Set QuerySQL         = objRs 
	  Set objRs            = nothing 
	  
   end function 

	Public Function ExecuteCMDINS(pstrSQL)

		dim intId, RecordsAffected 
	   ' response.write pstrSQL
		conexao.execute pstrSQL, RecordsAffected	

		If RecordsAffected=1 Then
			intId         = Conexao.execute("SELECT SCOPE_IDENTITY()")(0).Value 
			ExecuteCMDINS = cdbl(0 & intId)		   
		Else
			Response.Write "Erro na gravaчуo do registro." 
		End if
			
		err.clear
	
	end function  

end class
%>