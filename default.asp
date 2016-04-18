
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Prova</title>

    <!-- Bootstrap core CSS -->
    <link href="bootstrap/css/bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>
    <!-- #include file = "include/configuracoes.asp" -->
    <!-- #include file = "include/classesEntidades/cSiteMenu.asp" -->
    <%
    dim idSiteMenu 
    dim DsSiteMenu
    dim DsConteudo

    idSiteMenu = request.QueryString("idSiteMenu")


        %>
  <body>

    <nav class="navbar navbar-inverse navbar-fixed-top">
      <div class="container">
        <div class="navbar-header">
          <a class="navbar-brand" href="<%=request.serverVariables("script_name")%>">Prova FC Moreno</a>
        </div>
        <div id="navbar" class="collapse navbar-collapse">
          <ul class="nav navbar-nav">
            <%
            Set objTbSiteMenu = new clsTbSiteMenu
            Set objRs = objTbSiteMenu.ListarMenusAtivos()
            while not objRs.eof 

                if idSiteMenu = "" then
                    idSiteMenu = objRs("idSiteMenu")
                end if

                if "" & idSiteMenu = "" & objRs("idSiteMenu") then
                    DsSiteMenu = objRs("DsSiteMenu")
                    DsConteudo = objRs("DsConteudo")
                end if
                %>
                <li <%if "" & idSiteMenu = "" & objRs("idSiteMenu") then response.write " class=""active""" %>><a href="<%=request.serverVariables("script_name")%>?IdSiteMenu=<%=objRs("IdSiteMenu") %> "><%=objRs("DsSiteMenu") %></a></li>
                <%

                objRs.moveNext
            wend
            Set objRs = nothing

            Set objTbSiteMenu = nothing
             %>
          </ul>
          <form class="navbar-form navbar-right" method="post" action="validaLogin.asp"> 
            <div class="form-group">
              <input type="text" placeholder="Email" class="form-control" name="login" id="login">
            </div>
            <div class="form-group">
              <input type="password" placeholder="Password" class="form-control" name="senha" id="senha">
            </div>
            <button type="submit" class="btn btn-success">Entrar</button>
          </form>
        </div><!--/.navbar-collapse -->
      </div>
    </nav>

    <!-- Main jumbotron for a primary marketing message or call to action -->
    <div class="jumbotron">
      <div class="container">
        <h1><%=DsSiteMenu %></h1>
        
      </div>
    </div>

    <div class="container">
      <!-- Example row of columns -->
      <div class="row">
        <div class="col-md-12">
          <p><%=DsConteudo %></p>
        </div>
      </div>
      <hr>

      <footer>
        <p>&copy; 2015 Company, Inc.</p>
      </footer>
    </div> <!-- /container -->

    
    <script src="include/jQuery/jquery.min.js"></script>
    <script src="bootstrap/js/bootstrap.min.js"></script>
  </body>
</html>
