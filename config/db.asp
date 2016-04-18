<%
function buscarConexao()

    dim strServer, strDataBase, strUser, strPassword

    strServer   = "localhost\SQLEXPRESS"    'Ip do Servidor
    strDataBase = "Prova"     'Nome do Banco de Dados
    strUser     = "sa"                   'Login do Banco
    strPassword = "131313"                  'Senha do Banco
    
    buscarConexao = "Provider=SQLOLEDB.1; SERVER=" & strServer & "; DATABASE=" & strDataBase & "; UID=" &  strUser & "; PWD=" & strPassword & ";PageTimeOut=80;"
    
end function
%>