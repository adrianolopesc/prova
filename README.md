# Prova

Este é um projeto de um site que o administrador edita os menus e os conteúdos dos menus através de um sistema com acesso restrito através de um login e senha.

# Instruções para Instalação

1- Copiar os arquivos desse projeto para a pasta no servidor com IIS (com ASP instalado) que ficará hospedado o site.
2- Criar um site no IIS apontando para esta pasta.
3- Restaurar o banco de dados (arquivo prova.bak na pasta banco) no servidor de SQL Server.
4- Configurar o arquivo config/db.asp com os parametros de configuração do servidor de banco.
5- Acessar o site via navegador http://servidor/site
6- O usuário se senha padrão para acesso ao sistema é:
    - Login: admin
    - Senha: adm123


# Sobre a Ferramenta

Foi feita utilizando bootstrap e JQuery.
No diretorio do sistema constam as seguintes pastas:
- admin - Contém os arquivos de administração do sistema
- bootstrap - Arquivos do bootstrap
- ckeditor - Arquivos da biblioteca de edição em HTML ckeditor
- config - Arquivo de configuração de banco
- Include - Arquivos include geral, como topo e parte inferior de layout dos arquivos de administração
- include/classes - Arquivos de classe utilizados pelo sistema em ASP
- include/classesEntidades - Arquivos de classe utilizados pelo sistema em ASP feitos para cada tabela do banco de dados
- include/funcoesASP - Funcões diversas em ASP utilizadas no sistema
- include/jQuery - Arquivos do jquery
