<%
option explicit
%>
<!-- #include file = "../include/configuracoes.asp" -->
<!-- #include file = "../include/checaSessao.asp" -->
<!-- #include file = "../include/topo.asp" -->
<!-- #include file = "../include/funcoesasp/string.asp" -->
<!-- #include file = "../include/classesEntidades/cSiteMenu.asp" -->
<%
dim objTbSiteMenu
    
Set objTbSiteMenu = new clsTbSiteMenu

%>
<script type="text/javascript">
    $(document).ready(function () {

        $("#liAlterarMenu").addClass("active");

        $('#btIncluir').bind('click', function () {
            self.location = "edicao.asp";
        });

        $(".btEditar").bind('click', function () {

            var idSiteMenu = $(this).attr("idSiteMenu");
            self.location = "edicao.asp?idSiteMenu=" + idSiteMenu;

        });

    });
</script>
<div class="page-header">
    <h2>Listagem de Menus do Site</h2>
</div>

<table class="table table-hover">
    <thead>
        <tr>
            <th style="width:50%">Menu Site
            </th>
            <th style="width:20%">Ordem
            </th>
            <th style="width:20%">Ativo
            </th>
            <th style="width:10%">Ações
            </th>
        </tr>
    </thead>
    <tbody>
        <%
        Set objRs = objTbSiteMenu.Listar()
        
        while not objRs.eof
            %>
            <tr>
                <td><%=objRs("DsSiteMenu") %>
                </td>
                <td><%=objRs("NrOrdem") %>
                </td>
                <td>
                    <% if cbool(objRs("FgAtivo")) then
                        response.write "Sim"
                    else
                        response.write "Não"
                    end if %>
                </td>
                <td class="pull-right">
                    <input type="button" class="btEditar btn btn-primary" name="btEditar" value="Editar" idSiteMenu="<%=objRs("idSiteMenu") %>" />
                </td>
            </tr>
            <%
            objRs.moveNext
        wend

        Set objRs = nothing
        %>
    </tbody>
</table>
    
<%
Set objTbSiteMenu = nothing
 %>
<br />

<div class="row">
    <div class="col-sm-11">
        <div  class="pull-right" >
            <input type="button" id="btIncluir" name="btIncluir" value="Incluir" class="btn btn-primary" />
        </div>
    </div>
</div>


<!-- #include file = "../include/base.asp" -->