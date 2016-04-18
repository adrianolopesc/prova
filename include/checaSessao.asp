<%
if session("IdUsuario") = "" then

    session.Abandon()
    %>
    <script>
        top.location = '<%=application("NmDiretorioBaseSistema") & "/default.asp" %>';
    </script>
    <%
    response.end
end if
%>