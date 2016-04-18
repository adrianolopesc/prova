<%
option explicit
%>
<!-- #include file = "../include/configuracoes.asp" -->
<!-- #include file = "../include/checaSessao.asp" -->
<!-- #include file = "../include/topo.asp" -->
<!-- #include file = "../include/classesEntidades/cSiteMenu.asp" -->
<!-- #include file = "../include/funcoesasp/numero.asp" -->
<!-- #include file = "../include/funcoesasp/string.asp" -->
<!-- #include file = "../include/funcoesasp/data.asp" -->
<%
dim IdSiteMenu
dim objTbSiteMenu
dim DsSiteMenu
dim DsConteudo
dim NrOrdem
dim FgAtivo
dim Retorno
dim erro

IdSiteMenu = request.QueryString("IdSiteMenu")
Retorno = request.QueryString("Retorno")
erro = request.QueryString("erro")

FgAtivo = true

if IdSiteMenu <> "" then

    Set objTbSiteMenu = new clsTbSiteMenu

    objTbSiteMenu.AdicionarFiltro "IdSiteMenu", IdSiteMenu
    Set objRs = objTbSiteMenu.Listar()
    if not objRs.eof then
        DsSiteMenu = objRs("DsSiteMenu")
        DsConteudo = objRs("DsConteudo")
        NrOrdem = objRs("NrOrdem")
        FgAtivo = cbool(objRs("FgAtivo"))
    end if
    Set objRs = nothing

    Set objTbSiteMenu = nothing

end if

%>
<script src="../ckeditor/ckeditor.js"></script>
<script src="../include/jQuery/jquery.alphanumeric.js"></script>
<script src="../include/jQuery/jquery.validate.min.js"></script>
<script type="text/javascript">
    $(document).ready(function () {

        $("#liAlterarMenu").addClass("active");

        $('#btVoltar').bind('click', function () {
            self.location = "principal.asp";
        });
        /*
        CKEDITOR.replace('editor1', {
            toolbar: [
	            { name: 'basicstyles', groups: ['basicstyles', 'cleanup'], items: ['Bold', 'Italic', 'Underline', 'Strike', 'Subscript', 'Superscript', '-', 'RemoveFormat'] },
	            { name: 'paragraph', groups: ['list', 'indent', 'blocks', 'align', 'bidi'], items: ['NumberedList', 'BulletedList', '-', 'Outdent', 'Indent', '-', 'Blockquote', 'CreateDiv', '-', 'JustifyLeft', 'JustifyCenter', 'JustifyRight', 'JustifyBlock', '-', 'BidiLtr', 'BidiRtl'] },
	            { name: 'links', items: ['Link', 'Unlink', 'Anchor'] },
	            { name: 'insert', items: ['Image', 'Table', 'HorizontalRule', 'SpecialChar'] },
	            '/',
	            { name: 'styles', items: ['Styles', 'Format', 'Font', 'FontSize'] },
	            { name: 'colors', items: ['TextColor', 'BGColor'] },
	            { name: 'tools', items: ['Maximize', 'ShowBlocks'] },
	            { name: 'others', items: ['-'] }
            ]
        });
        */

        $("#nrOrdem").numeric();

        $("#frm").validate({
            rules: {
                dsSiteMenu: {
                    required: true
                },
                dsConteudo: {
                    required: true
                },
                nrOrdem: {
                    required: true
                }
            },
            messages: {
                dsSiteMenu: {
                    required: "Informe o menu"
                },
                dsConteudo: {
                    required: "Informe o conteúdo"
                },
                nrOrdem: {
                    required: "Informe a ordem"
                }
            },
            highlight: function (element) {
                $(element).closest('.form-group').addClass('has-error');
            },
            unhighlight: function (element) {
                $(element).closest('.form-group').removeClass('has-error');
            },
            errorElement: 'span',
            errorClass: 'help-block',
            errorPlacement: function (error, element) {
                if (element.parent('.input-group').length) {
                    error.insertAfter(element.parent());
                } else {
                    error.insertAfter(element);
                }
            }
        });



    });
</script>
<div class="page-header">
    <h2>Edição de Menu do Site</h2>
</div>
<%
if Retorno <> "" then 
    if erro = "0" then%>
        <div class="alert alert-success">
            <strong>Sucesso!</strong> <%=retorno %>
        </div>
    <%
    else
    %>
        <div class="alert alert-danger">
            <strong>Erro!</strong> <%=retorno %>
        </div>
    <%
    end if
    %>
    <br />
<%
end if %>
<form role="form" id="frm" name="frm" method="post" action="grava.asp">
    <input type="hidden" id="IdSiteMenu" name="IdSiteMenu" value="<%=IdSiteMenu %>">
    <div class="form-group col-lg-4" >
        <label for="dsSiteMenu">Menu</label>
        <input type="text" class="form-control input-sm" id="dsSiteMenu" name="dsSiteMenu" value="<%=dsSiteMenu %>">
    </div>
    <div class="form-group col-lg-12">
        <label for="dsConteudo">Conteúdo</label>
        <textarea class="ckeditor" cols="80" id="dsConteudo" name="dsConteudo" rows="50"><%=dsConteudo %>
        </textarea>
    </div>
    <div class="form-group  col-lg-2">
        <label for="nrOrdem">Ordem</label>
        <input type="text" class="form-control col-lg-1" id="nrOrdem" name="nrOrdem" value="<%=nrOrdem %>">
    </div>
    <div class="form-group col-lg-12">
        <div class="checkbox">
            <label><input type="checkbox" name="fgAtivo" id="fgAtivo" value="1" <%if FgAtivo then response.write " checked" %>>Exibir no site</label>
        </div>
    </div>
    <div class="row">
        <div class="col-sm-12">
            <div class="pull-right">
                <button type="submit" class="btn btn-primary">Salvar</button>
                <input type="button" id="btVoltar" name="btVoltar" value="Voltar" class="btn btn-warning" />
            </div>
        </div>
    </div>
</form>



<!-- #include file = "../include/base.asp" -->