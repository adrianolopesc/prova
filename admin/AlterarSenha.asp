<%
option explicit
%>
<!-- #include file = "../include/configuracoes.asp" -->
<!-- #include file = "../include/checaSessao.asp" -->
<!-- #include file = "../include/topo.asp" -->
<!-- #include file = "../include/funcoesASP/criptografia.asp" -->
<!-- #include file = "../include/funcoesASP/string.asp" -->
<!-- #include file = "../include/classesEntidades/cUsuario.asp" -->

<%
dim Retorno
dim erro
dim idUsuario

Retorno = request.QueryString("Retorno")
erro = request.QueryString("erro")
idUsuario = request.QueryString("idUsuario")
%>
<script src="../include/jQuery/jquery.validate.min.js"></script>

<script type="text/javascript">

    $(document).ready(function () {

        $("#liAlterarSenha").addClass("active");

        $("#frm").validate({
            rules: 
                {
                cdSenhaAtual: {
                    required: true,
                    remote: "validaCdSenha.asp"
                },
                cdSenha: {
                    required: true,
                    minlength: 6,
                    remote: "validaCdSenhaConfirma.asp"
                },
                cdSenhaConfirma: {
                    required: true,
                    equalTo: "#cdSenha"
                }
            },
            messages: {
                cdSenhaAtual: {
                    required: "Informe a senha atual",
                    remote: "Senha não confere com a senha atual"
                },
                cdSenha: {
                    required: "Informe a senha",
                    minlength: "Mínimo de 6 caracteres",
                    remote: "A nova senha não pode ser igual a senha atual"
                },
                cdSenhaConfirma: {
                    required: "Informe a confirmação da senha",
                    equalTo: "Confirmação diferente do campo senha"
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
    <h2>Alteração de Senha</h2>
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

<form role="form" id="frm" name="frm" method="post" action="gravaAlterarSenha.asp">

    <div class="form-group col-lg-4" >
        <label for="dsSiteMenu">Informe a senha atual</label>
        <input type="password" name="cdSenhaAtual" id="cdSenhaAtual" value="" class="form-control " maxlength="50" />
    </div>
    <div class="clearfix"></div>
    <div class="form-group col-lg-4" >
        <label for="dsSiteMenu">Nova senha </label>
        <input type="password" name="cdSenha" id="cdSenha" value="" class="form-control " maxlength="50" />
    </div>
    <div class="clearfix"></div>
    <div class="form-group col-lg-4" >
        <label for="dsSiteMenu">Confirme a nova senha  </label>
        <input type="password" name="cdSenhaConfirma" id="cdSenhaConfirma" value="" class="form-control " maxlength="50" />
    </div> 
    <div class="row">
        <div class="col-sm-12">
            <div class="pull-right">
                <button type="submit" class="btn btn-primary">Salvar</button>
            </div>
        </div>
    </div>
</form>

<!-- #include file = "../include/base.asp" -->