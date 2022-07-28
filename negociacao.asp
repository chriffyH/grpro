<!DOCTYPE html>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
if request.querystring("act") <> "" then 
    Dim acoes
    str = request.querystring("act")
    acoes = split(str,",")
    acao = acoes(0)
end if
'Conexao com o banco
Set conexao = Server.CreateObject("ADODB.Connection")
%>
<!--#include file="acesso.asp"-->
<!--#include file="../Connections/conectar.asp" -->
<!--#include file="../Connections/rede.asp" -->
<html>

<!--#include file="estrutura_head.asp"-->

<link href="css/plugins/steps/jquery.steps.css" rel="stylesheet">
<link href="css/plugins/select2/select2.min.css" rel="stylesheet">

<link href="css/plugins/steps/jquery.steps.css" rel="stylesheet">
<link href="css/plugins/select2/select2.min.css" rel="stylesheet">

<link href="css/plugins/clockpicker/clockpicker.css" rel="stylesheet">

<style>
    .md-skin .wrapper-content {
        padding: 15px 0px 15px 0px !important;
    }

    .seta::after {
        display: inline-block;
        margin-left: 0.255em;
        vertical-align: 0.255em;
        content: "";
        border-top: 0.3em solid;
        border-right: 0.3em solid transparent;
        border-bottom: 0;
        border-left: 0.3em solid transparent;
    }


    #submit_produtos {
        float: right;
        right: 15px;
    }
    
    .flexivel {
        position: relative;
        display: flex !important;
        justify-content: space-between !important;
        flex-wrap: nowrap !important;
    }

    .btt_custom {
        width: 100%;
        display: block;
        height: 34px;
        padding: 6px 2px;
        font-size: 14px;
        line-height: 1.42857143;
        background-image: none;
    }
    .btt_custom_2 {
        margin-top: 30px;
    }

    .overflow {
        overflow: hidden;
    }

    .estilo_link {
        display: block;
    }   

    .tratamento_td {
        position: relative;
        display: flex !important;
        justify-content: space-between !important;
        flex-wrap: nowrap !important;
    }

    .tratamento_td form {
        position: relative;
    }

    .tratamento div.div_link_tratamento_td {
        position: relative;
    }

    .div_link_tratamento_td a {
        padding: 0 !important;
        margin: auto !important;
    }

    #load_page {
        display: none;
        position: absolute;
        width: 100%;
        height: 100%;
        z-index: 10000000;
        background-color: rgba(255, 255, 255, .7);

        text-align: center;
    }

    #load {
        position: fixed;
        z-index: 100000000000;
        font-size: 60px;
        color: var(--cor-principal);
        margin-top: 200px;
    }   

    .col-lg-3{
        width: 29% !important;
    }
</style>

<%
    IF Request.Form("MM_Insert") = "formreuniao" THEN 
        id_neg = Request.Form("id_prospector")

        dia = day(Request.Form("agenda_data"))
        IF dia < 10 THEN dia = "0"&dia 
        mes = month(Request.Form("agenda_data"))
        IF mes < 10 THEN mes = "0"&mes 
        agenda_data = year(Request.Form("agenda_data"))&"/"&mes&"/"&dia

        agenda_hora = Request.Form("agenda_hora") & ":00"

        descricao = Replace(Request.Form("descricao"), VbCrLf, "<br>")

        conexao.execute("INSERT INTO neg_interacao (id_negociacao, id_usuario, descricao, agenda_data, agenda_hora, tipo) VALUES ("&id_neg&", "&Request.Cookies("grpro")("idusuario")&", '"&descricao&"', '"&agenda_data&"', '"&agenda_hora&"', 'Reunião')")

        SET rsnome_cliente = conexao.execute("SELECT empresa, responsavel, id_usuario FROM neg_prospecter WHERE del = 0 AND idnp = "&id_neg)

        datas = agenda_data & " " & agenda_hora

        conexao.execute("INSERT INTO agenda_compromissos (id_usuario, id_setor, id_np, titulo, descricao, data_inicio, data_termino, tipo) VALUES ( "&rsnome_cliente("id_usuario")&", "&Request.Cookies("grpro")("setor")&", "&id_neg&", 'Reunião com "&rsnome_cliente("empresa")&" - "&rsnome_cliente("responsavel")&"', '"&descricao&"', '"&datas&"', '"&datas&"', 'Pessoal')")

        Response.Redirect("?msg=agendamento")
    END IF 
%>


<body>

    <div id="wrapper">

        <div id="status-pesquisa-negociacao" style="display: none;position:fixed;background: rgba(255,255,255,0.8);width: 100%;height: 100%;z-index: 9999999;">
            <div class="ibox-content sk-loading">
                <div class="sk-spinner sk-spinner-double-bounce" style="position:fixed !important; top: 50%;left: 50%;">
                    <div class="sk-double-bounce1"></div>
                    <div class="sk-double-bounce2"></div>
                </div>  
            </div>      
        </div>

    <div id="load_page">
        <span id="load" class="loading toggle"></span>
    </div>

    <nav class="navbar-default navbar-static-side" role="navigation">
        <!--#include file="include_left_side.asp"-->
    </nav>

     <div id="page-wrapper" class="gray-bg">
        <div class="row border-bottom">
            <!--#include file="include_topo.asp"-->
        </nav>
        </div>
        <div class="row wrapper border-bottom white-bg page-heading" style="display: flex;">
            <div class="col-lg-9">
                <h2>Negociações</h2>
                <ol class="breadcrumb">
                    <li>
                        <a href="default.asp">Home</a>
                    </li>
                    <li class="active">
                        <strong>Gerenciamento de Negociações</strong>
                    </li>
                </ol>
            </div>
            <div class="col-lg-3" style="float: right; margin-top: 2%; ">
                <h2></h2>

                <div class="btn-group show" >
                    <button data-toggle="dropdown"  class="btn btn-primary seta" aria-expanded="false" style="float: right;">Escolha o Funil</button>
                    <ul class="dropdown-menu" x-placement="bottom-start" style="position: absolute; top: 32px; left: 0px; will-change: top, left; margin-left: 44%; ">
                        <li><a class="dropdown-item" onclick="filtro_neg('nome')" >Nome / CPF - CNPJ</a></li>
                        <li><a class="dropdown-item" onclick="filtro_neg('estado')">Estado</a></li>
                        <li><a class="dropdown-item" onclick="filtro_neg('responsavel')">Responsável</a></li>
                        <li><a class="dropdown-item" onclick="filtro_neg('perfil')">Perfil</a></li>
                        <li><a class="dropdown-item" onclick="filtro_neg('interesse')">Interesse</a></li>
                        <li><a class="dropdown-item" onclick="filtro_neg('aquisicao')">Aquisição</a></li>
                        <li class="dropdown-divider"></li>
                    </ul>
                </div>

                <script type="text/javascript">
                    function filtro_neg(tipo) {

                        ajaxGo({ url:"Ajax/ajax_pesquisa_negociacao_retorno.asp?tipo="+tipo, elem_return: document.getElementById("filtro-neg-tipo") });
                        document.getElementById("filtro-avanc").style.display = "block";
                    }
                </script>

            </div>

        </div>

        </nav>
        

        <div class="wrapper wrapper-content  animated fadeInRight resp271">

            <style type="text/css">
                .desativado-filtro{
                    display: none !important;
                }
            </style>

            <!-- Filtro -->

            <!-- desativado-filtro -->
            <div class="ibox-content m-b-sm border-bottom " style="display: none; padding: 0% 1%" id="filtro-avanc">
         
                    <div class="row">

                        <div class="col-sm-3" style="">

                            <div class="col-lg-1 tooltip-demo" style="display: flex; width: auto; margin-top: 9px; padding-left: 0px; ">

                                <%
                                    IF Request.Cookies("grpro")("visual_tarefas") = "Kanban" OR Request.Cookies("grpro")("visual_tarefas") = "" THEN 

                                        css_kanban_item = "border-left: 1px #ccc solid; padding: 5px 10px; cursor: pointer; background-color: var(--cor-secundaria); color: #FFF"
                                        css_lista_item = "border-right: 1px #ccc solid; border-left: 1px #ccc solid; padding: 5px 10px; cursor: pointer; color: var(--cor-secundaria)"
                                        css_calendario_item = "border-right: 1px #ccc solid; padding: 5px 10px; cursor: pointer; color: var(--cor-secundaria)"

                                    ELSEIF Request.Cookies("grpro")("visual_tarefas") = "Lista" THEN 

                                        css_kanban_item = "border-left: 1px #ccc solid; padding: 5px 10px; cursor: pointer; color: var(--cor-secundaria)"
                                        css_lista_item = "border-right: 1px #ccc solid; border-left: 1px #ccc solid; padding: 5px 10px; cursor: pointer; background-color: var(--cor-secundaria); color: #FFF"
                                        css_calendario_item = "border-right: 1px #ccc solid; padding: 5px 10px; cursor: pointer; color: var(--cor-secundaria)"

                                    END IF
                                %>

                                <div data-toggle="tooltip" data-placement="top" title="Modo Kanban" style="<%=css_kanban_item%>"  onClick="document.location.reload(true);" class="kanban">
                                    <i class="fa fa-th kanban_i"></i>
                                </div>

                                <div data-toggle="tooltip" data-placement="top" title="Modo Lista" style="<%=css_lista_item%>"  onClick="visualizacao_negociacoes('l', <%=Request.cookies("grpro")("idusuario")%>)" class="lista">
                                    <i class="fa fa-list lista_i" style="transform: rotate(90deg);"></i>
                                </div>

                                <a href="neg_relatorio.asp" style="background-color: #476371;border-color: #476371;margin-left: 7%;padding: 3% 14%" class="btn btn-primary center-block" name="buscar" type="submit" id="buscar" value="">Busca Detalhada</a>

                            </div>
                        </div>

                        <div class="col-sm-9" id="filtro-neg-tipo" style="padding-left: 2%">

                    
                        </div>
                        
                    </div>

            </div>

            <!-- Etapas -->
            <div class="wrapper1" style="margin-bottom: 6px;">
                <div class="div1"></div>
            </div>
            <div class="wrapper2" style="margin-bottom: 5%">
                <div class="div2">
                    
                    <div id="retorna_negociacoes" class="row" style="display: flex; min-height: auto;">

                        <%
                                    
                            usuario = "np.idnp is not null"  

                            perfil = "np.id_perfil is not null"  
                            IF Request.Form("perfil") <> "" and Request.Form("perfil") <> "todos" THEN 
                                perfil = "np.id_perfil = "&Request.Form("perfil")
                            END IF 

                            aquisicao = "np.id_aquisicao is not null"  
                            IF Request.Form("aquisicao") <> "" and Request.Form("aquisicao") <> "todos" THEN 
                                aquisicao = "np.id_aquisicao = "&Request.Form("aquisicao")
                            END IF 

                            produto = "np.id_produto is not null"  
                            IF Request.Form("interesse") <> "" and Request.Form("interesse") <> "todos" THEN 
                                produto = "np.id_produto = "&Request.Form("interesse")
                            END IF 


                            
                            SET rsetapas = conexao.execute("SELECT * FROM neg_funil_etapas WHERE status = 'Ativo' ORDER BY ordem ASC")
                            contador = 1
                            i = 1

                            IF Request.Cookies("grpro")("perm_103") = "1" THEN
                                visualizar_tipo = "np.idnp is not null"
                            ELSE 
                                visualizar_tipo = " EXISTS(SELECT id_usuario FROM neg_prospecter_responsaveis pr where id_negp = np.idnp and id_usuario = "&Request.Cookies("grpro")("idusuario")&")"
                            END IF 

                            total_tamanho = 0
                            IF NOT rsetapas.EOF THEN 
                                WHILE NOT rsetapas.EOF
                                    total_tamanho = total_tamanho + 324.5
                                ' Pega mes e ano atual
                                dat3 = date()
                                mes = month(dat3)
                                IF mes < 10 THEN
                                    mes = 0&mes
                                END IF
                                dat3 = year(dat3)&"-"&mes

                                if rsetapas("conversao") = "Sim" or rsetapas("nao_conversao") = "Sim"  then
                                    SET rstotal = conexao.execute("SELECT count(DISTINCT np.idnp) as total FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp WHERE "&visualizar_tipo&" and np.situacao = "&rsetapas("idnfe")&" AND "&usuario&" AND "&perfil&" AND "&aquisicao&" AND "&produto&" AND month(data) = "&month(date)&" AND np.del = 0 AND year(data) = "&year(date)&" ORDER BY nf.idnf DESC LIMIT 10")

                                    SET rstotaladesao = conexao.execute("SELECT SUM(DISTINCT valor_adesao) as totaladesao FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp WHERE "&visualizar_tipo&" AND np.situacao = "&rsetapas("idnfe")&" AND "&usuario&" AND "&perfil&" AND "&aquisicao&" AND "&produto&" AND month(data) = "&month(date)&" AND np.del = 0 AND year(data) = "&year(date)&" ORDER BY nf.idnf DESC")
                                ELSE

                                    SET rstotal = conexao.execute("SELECT Count(idnp) as total FROM neg_prospecter np WHERE "&visualizar_tipo&" AND np.del = 0 AND situacao = "&rsetapas("idnfe")&" AND "&usuario&" AND "&perfil&" AND "&aquisicao&" AND "&produto&" LIMIT 10")

                                    SET rstotaladesao = conexao.execute("SELECT SUM(valor_adesao) as totaladesao FROM neg_prospecter np WHERE "&visualizar_tipo&" AND np.del = 0 AND situacao = "&rsetapas("idnfe")&" AND "&usuario&" AND "&perfil&" AND "&aquisicao&" AND "&produto)
                                end if


                                IF IsNull(rstotaladesao("totaladesao")) THEN 
                                    totaladesao = "0,00"
                                ELSE 
                                    totaladesao = FormatNumber(rstotaladesao("totaladesao"))
                                END IF 
                        %>

                        <div class="col-lg-3">
                            <div class="ibox">
                                <div class="ibox-content" style="height: 568px; background-color: #e8e8ea; padding: 0px 6px 20px 6px;">
                                    
                                    <div style="overflow: auto;"></div>

                                    <%
                                        etapas = pontinhos(rsetapas("etapa"), 25) 
                                    %>
                                    
                                    <h3 style="margin-left: 3%; font-size: 17px; margin-bottom: 0;"><strong><%=etapas%></strong></h3> 

                                    <span style="margin-left: 3%; " id="count<%=rsetapas("idnfe")%>"><%=rstotal("total")%></span>&nbsp; contatos

                                    <label class="label label-success" style="background-color: #5cb85c; margin-top: 1%; margin-right: 3%; float: right; ">R$ <span id="count-adesao<%=rsetapas("idnfe")%>"><%=totaladesao%></span></label>

                                    <span style="position: absolute;right: 13%;top: 7%;cursor: pointer;">
                                        
                                    </span>
                                    <div style="clear:both"></div>
                                
                                    <%        
                                        if rsetapas("conversao") = "Sim" or rsetapas("nao_conversao") = "Sim" then

                                            SET rsnegoc = conexao.execute("SELECT DISTINCT np.idnp, np.empresa, np.responsavel, np.cidade, np.uf, data_validacao FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp WHERE "&visualizar_tipo&" and np.situacao = "&rsetapas("idnfe")&" and "&usuario&" and "&perfil&" and "&aquisicao&" and "&produto&" AND np.del = 0 and month(data) = "&month(date)&" and year(data) = "&year(date)&" ORDER BY nf.idnf DESC limit 10")
                                        ELSE
                                            SET rsnegoc = conexao.execute("SELECT DISTINCT np.idnp, np.empresa, np.responsavel, np.cidade, np.uf, data_validacao FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp WHERE "&visualizar_tipo&" AND np.del = 0 and np.situacao = "&rsetapas("idnfe")&" and "&usuario&" and "&perfil&" and "&aquisicao&" and "&produto&" ORDER BY nf.idnf DESC limit 10")
                                        end if

                                        SET rsid_entrada = conexao.execute("SELECT idnfe FROM neg_funil_etapas WHERE entrada = 'Sim'")


                                        %>

                                    <div class="barra" style="height: 500px; overflow: auto;overflow-x:hidden;">
                                        <style type="text/css">

                                            .info-element:hover .icones {
                                                display: block !important;
                                            }
                                            
                                        </style>
                                        <div id="retorno_neg<%=rsetapas("idnfe")%>">

                                            <ul style="" class="sortable-list connectList agile-list" id="etapa<%=rsetapas("idnfe")%>">

                                                <%
                                                    IF NOT rsnegoc.EOF THEN 
                                                
                                                        WHILE NOT rsnegoc.EOF 
                                                            habilita_acao = false
                                                            IF Request.Cookies("grpro")("perm_103") = "1" THEN 
                                                                habilita_acao = true
                                                            ELSE

                                                                SET verifica_participacao = conexao.execute("SELECT idprospecter FROM neg_prospecter_responsaveis where id_negp = "&rsnegoc("idnp")&" and id_usuario = "&Request.cookies("grpro")("idusuario")&" limit 1")

                                                                IF NOT verifica_participacao.EOF THEN habilita_acao = true

                                                            END IF
                                                %>
                                                    
                                                            <li class="info-element" id="<%IF NOT habilita_acao THEN response.write("nao-")%><%=rsnegoc("idnp")%>" style="position: relative; padding: 3%; padding-bottom: 0px !important">  
                                
                                                                <%IF habilita_acao THEN%>
                                                                    <div class="icones" style="text-align: center; position: absolute;  width: 100%; padding: 10px;display: block;bottom: 0px;z-index: -1; background-color: rgba(255, 255, 255, 0.7);">

                                                                        <div style="width: 100%; margin-left: auto;text-align: right; display: flex;">
                                                                            <i title="Adicionar Anotação" data-toggle="tooltip" data-placement="top" onclick="adicionar_anotacao(<%=rsnegoc("idnp")%>,'<%=rsnegoc("empresa")%> - <%=rsnegoc("responsavel")%>')" class="fas fa-folder-plus" style="font-size: 15px;color: #003366;cursor: pointer;margin-right: 3%;margin-left: auto !important;"></i><br>

                                                                            <i title="Adicionar Agendamento" data-toggle="tooltip" data-placement="top" onclick="adicionar_agendamento(<%=rsnegoc("idnp")%>,'<%=rsnegoc("empresa")%> - <%=rsnegoc("responsavel")%>')" class="fas fa-calendar" style="font-size: 15px;color: #003366;cursor: pointer;margin-right: 3%"></i><br>

                                                                            <i style="font-size: 15px; margin-right: 10px; color: #003366;cursor: pointer;" onclick="excluir(<%=rsnegoc("idnp")%>)" class="fa fa-trash pull-right" data-placement="top" title="Excluir" aria-hidden="true"></i><br><span class="sr-only"></span>
                                                                        </div>
                                                                    </div>
                                                                <%END IF%>

                                                                <div <%IF habilita_acao THEN%>onclick="window.open('neg_interacao_novo.asp?id=<%=rsnegoc("idnp")%>', '_blank') "<%END IF%>>
                                                                    <div style="height: 22px">

                                                                        <%
                                                                            SET count_rsparticipantes = conexao.execute("SELECT COUNT(pr.idprospecter) as total FROM usuarios u  INNER JOIN neg_prospecter_responsaveis pr ON pr.id_usuario = u.idusuario WHERE pr.id_negp = "&rsnegoc("idnp")&" AND u.status = 'Ativo' ")
                                                                            flag_part = false
                                                                            IF CLNG(count_rsparticipantes("total")) > 3 THEN 
                                                                                flag_part = true
                                                                        %>

                                                                                <div title="Demais responsáveis" class="tooltip-demo" style="width: 30px; height: 30px; position: relative; display: block; float: right; margin-right: 8px; text-align: center;  background-color: #f5f5f5; border-radius: 20px;">
                                                                                    <span style="line-height: 30px;font-size: 10px;font-weight: 700;color: #cccccc;" >+ <%=CLNG(count_rsparticipantes("total")) - 2%></span>
                                                                                </div>
                                                                        <%
                                                                            END IF

                                                                            SET rsusuario = conexao.execute("SELECT u.usuario, u.logo, s.setor FROM usuarios u INNER JOIN setores s ON u.id_setor = s.idsetor INNER JOIN neg_prospecter_responsaveis pr ON pr.id_usuario = u.idusuario WHERE pr.id_negp = "&rsnegoc("idnp")&" AND u.status = 'Ativo' ")

                                                                            IF NOT rsusuario.EOF THEN 

                                                                                contador = 1
                                                                                IF flag_part THEN
                                                                                    contador_aux = 2
                                                                                ELSE
                                                                                    contador_aux = 3
                                                                                END IF

                                                                                WHILE NOT rsusuario.EOF and contador <= contador_aux
                                                                                    IF IsNull(rsusuario("logo")) OR rsusuario("logo") = "" THEN
                                                                                        logo_usuario = "semfoto.png"
                                                                                    ELSE 
                                                                                        logo_usuario = rsusuario("logo")
                                                                                    END IF
                                                                        %>
                                                                                    <a class="tooltip-demo" href="usuarios.asp" style="width: 30px; height: 30px; position: relative; display: block; float: right; margin-right: 8px;">
                                                                                        <img style="width: 100%; height: 100%;"  data-toggle="tooltip" data-placement="bottom" title="<%=rsusuario("usuario")%> (<%=rsusuario("setor")%>)" alt="image" class="img-circle" src="../imagens/logo_usuario/<%=logo_usuario%>">
                                                                                    </a>
                                                                        <%
                                                                                    contador = contador + 1
                                                                                    rsusuario.movenext
                                                                                WEND
                                                                            END IF
                                                                        %>

                                                                        <div style="display: flex;">
                                                                            <%
                                                                                SET rsdata_etapa = conexao.execute("SELECT CAST(data AS DATE) as data FROM neg_funil WHERE id_negp = "&rsnegoc("idnp")&" and id_etapa_funil = "&rsetapas("idnfe")& " ORDER BY idnf DESC")

                                                                                IF NOT rsdata_etapa.EOF THEN
                                                                            %>
                                                                                    <div class="tooltip-demo">
                                                                                        <span data-toggle="tooltip" data-placement="top" title="Entrada na etapa" style="margin: 0px 10px 0 0" class="label label-primary pull-left"><%=rsdata_etapa("data")%></span>
                                                                                    </div> 
                                                                            <%
                                                                                END IF
                                                                            
                                                                                SET rs_last_historico = conexao.execute("SELECT data FROM neg_interacao WHERE id_negociacao = "&rsnegoc("idnp")&" ORDER BY idni DESC LIMIT 1")

                                                                                IF NOT rs_last_historico.EOF THEN
                                                                                    ultima_interacao = rs_last_historico("data")
                                                                            %>

                                                                                    <script>
                                                                                        var data1 = moment('<%=now%>', "DD/MM/YYYY hh:mm:ss");
                                                                                        var data2 = moment('<%=ultima_interacao%>', "DD/MM/YYYY hh:mm:ss");
                                                                                        var resultado<%=i%> = data1.diff(data2, 'days');

                                                                                        $(document).ready(function () {
                                                                                            document.getElementById('val<%=i%>').innerHTML = resultado<%=i%> + " dias";
                                                                                        });
                                                                                    </script>

                                                                                    <div class="tooltip-demo">
                                                                                        <span style="margin: 0px 20px 0 0" class="label label-danger pull-left" id="val<%=i%>" data-toggle="tooltip" data-placement="top" title="Última interação."></span>
                                                                                    </div> 
                                                                            <%
                                                                                ELSE 
                                                                            %>
                                                                                    <span style="margin: 0px 20px 0 0" class="label label-primary pull-left" id="val<%=i%>">N/A</span>
                                                                            <%
                                                                                END IF 
                                                                            %>
                                                                        </div>

                                                                    </div>

                                                                    <%
                                                                        n_empresa = pontinhos(rsnegoc("empresa"), 25)
                                                                    %>

                                                                    <strong><%=n_empresa%></strong> 

                                                                    <%
                                                                        resp = pontinhos(rsnegoc("responsavel"), 35)
                                                                    %>

                                                                    <p style="margin-bottom: 0px;"><small style="color: var(--cor-principal); "><%=resp%></small></p>

                                                                    <p style="margin-bottom: 3px;">(<%=rsnegoc("cidade")%> - <%=rsnegoc("uf")%>)</p> 

                                                                    <%

                                                                        SET rsagenda = conexao.execute("SELECT agenda_data, agenda_hora FROM neg_interacao WHERE tipo = 'Reunião' and id_negociacao = "&rsnegoc("idnp")&" and status = 'Aberto'")

                                                                        IF NOT rsagenda.EOF THEN 
                                                                    %>
                                                                            <div class="tooltip-demo" style=";clear: both;">
                                                                                <i data-toggle="tooltip" data-placement="top" title="Agendado: <%=rsagenda("agenda_data")%> " class="fa fa-calendar" style="color: var(--cor-secundaria) !important"></i>
                                                                            </div> 
                                                                    <%
                                                                        ELSE    
                                                                    %>
                                                                            <div class="tooltip-demo" style=";clear: both; height: 12px; width: 12px;">
                                                                               
                                                                            </div>

                                                                    <%  END IF  %>

                                                                    <%    
                                                                        IF NOT IsNull(rsnegoc("data_validacao")) THEN 
                                                                    %>
                                                                            <div class="tooltip-demo" style="height: 10px; clear: both;">
                                                                                <i data-toggle="tooltip" data-placement="top" title="Aprovado" class="fa fa-check-circle" ></i>
                                                                            </div> 
                                                                    <%
                                                                        
                                                                        END IF                                                                 
                              
                                                                    %>

                                                                    <%
                                                                        SET rsdata_entrada = conexao.execute("SELECT CAST(data AS DATE) as data FROM neg_funil WHERE id_negp = "&rsnegoc("idnp")&" and id_etapa_funil = "&rsid_entrada("idnfe"))

                                                                        IF NOT rsdata_entrada.EOF THEN
                                                                    %>
                                                                            <div>
                                                                                <p data-toggle="tooltip" data-placement="top" title="Data de cadastro" style="margin-top: -17px; text-align: right;"><small><i style="margin-right: 1%" class="fas fa-pennant" aria-hidden="true"></i><%=rsdata_entrada("data")%></small></p>
                                                                            </div> 
                                                                    <%
                                                                        END IF
                                                                    %>

                                                                    
                                                                </div>

                                                        </li>
                                                          
                                                <%
                                                            i = i + 1
                                                            rsnegoc.movenext
                                                        WEND
                                                    
                                                        IF Cint(rstotal("total")) > 2 THEN 
                                                %>
                                                            <div class="col-lg-12" style="text-align: center;padding: 2%">
                                                                <a onClick="neg_plus(<%=rsetapas("idnfe")%>, '<%=rsetapas("conversao")%>', '<%=rsetapas("nao_conversao")%>', '', '')" target="_blank">
                                                                    <div class="vertical-timeline-icon navy-bg" style="position: initial;margin: auto;cursor: pointer;">
                                                                        <i class="fa fa-plus" aria-hidden="true"></i>
                                                                    </div>
                                                                </a>
                                                            </div>
                                                <%      
                                                        END IF 
                                                    END IF
                                                %>
                                            </ul>
                                        </div>
                                    </div>

                                </div>
                            </div>
                        </div>

                        <%
                            i = i + 1
                            contador = contador + 1 
                            rsetapas.movenext 
                                WEND
                            ELSE 
                        %>
                            <p>Não há etapas.</p>
                        <%
                            END IF 

                            rsetapas.Close
                            SET rsetapas = nothing
                        %>

                    </div>
                   
                </div>
            </div>

            <style type="text/css">
                .wrapper1, .wrapper2 { width: 100%; overflow-x: scroll; overflow-y: hidden; }
                .wrapper1 { height: 20px; }
                .wrapper2 {}
                .div1 { height: 20px; }
                .div2 { overflow: none; }
            </style>

            <script type="text/javascript">
                $(function () {
                    $('.wrapper1').on('scroll', function (e) {
                        $('.wrapper2').scrollLeft($('.wrapper1').scrollLeft());
                    }); 
                    $('.wrapper2').on('scroll', function (e) {
                        $('.wrapper1').scrollLeft($('.wrapper2').scrollLeft());
                    });
                });
                $(window).on('load', function (e) {
                    $('.div1').width(<%=total_tamanho%>);
                    $('.div2').width($('#retorna_negociacoes').width());
                });

                console.log("Tamanho total: "+$('#retorna_negociacoes').width())
            </script>

            <script type="text/javascript">
                
                function adicionar_anotacao(x, y)
                {
                    console.log(x);
                    document.getElementById("idprospector").value = x;
                    document.getElementById("anotacao_desc").innerHTML = y;
                    document.getElementById("descricao_anotacao_funil").value = "";
                    $("#modalanota").modal("show");

                }

                function adicionar_agendamento(x, y)
                {
                    console.log(x);
                    document.getElementById("id_prospector_agenda").value = x;
                    document.getElementById("agendamento_desc").innerHTML = y;
                    $("#modagenda").modal("show");

                }

            </script>

            <!-- Modal Anotação -->
            <div id="modalanota" class="modal fade" role="dialog">
                <div class="modal-dialog">

                    <!-- Modal content-->
                    <div class="modal-content">
                        <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal">&times;</button>
                            <h4 class="modal-title">Adicionar Anotação</h4>
                            <p id="anotacao_desc">  </p>
                        </div>
                        <div class="modal-body">
                            <form ACTION="Ajax/ajax_cria_anotacao.asp" METHOD="POST">

                                <input type="hidden" name="idprospector" id="idprospector">

                                <label>Tipo de interação: *</label>
                                <select class="form-control" required="" name="id_tipo">

                                    <option value="">Selecione o tipo de interação</option>
                                    <%
                                        SET tipos = conexao.execute("SELECT * FROM neg_interacao_tipo where del = 0")
                                        WHILE NOT tipos.EOF
                                    %>
                                            <option value="<%=tipos("id")%>"><%=tipos("nome")%></option>
                                    <%
                                            tipos.movenext
                                        WEND
                                    %>

                                </select>

                                <label>Descrição: *</label>
                                <textarea style="height: 100px;" class="form-control" rows="10" id="descricao_anotacao_funil" name="descricao" required></textarea>

                                <br>
                                <div style="display: flex; justify-content: space-between;">

                                    <div class="col-md-4" style="padding-left: 0px">
                                        <input type="submit" formtarget="_blank" onclick="fechar_modal_anotacao()" class="btn btn-primary" value="Inserir">
                                    </div>

                                    <input hidden name="MM_Insert" value="formanotacao">

                                </div>
                            </form>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                        </div>
                    </div>

                </div>
            </div>

            <!-- Modal -->
            <div id="modagenda" class="modal fade" role="dialog">
                <div class="modal-dialog">

                    <!-- Modal content-->
                    <div class="modal-content">
                        <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal">&times;</button>
                            <h4 class="modal-title">Adicionar Agendamento</h4>
                            <p id="agendamento_desc">  </p>
                        </div>
                        <div class="modal-body">                            
                            <form ACTION="<%=MM_editAction%>" METHOD="POST">
                                <input type="hidden" name="id_prospector" id="id_prospector_agenda">

                                <label>Descrição: *</label>
                                <textarea class="form-control" rows="10" name="descricao" required></textarea>

                                <br>
                                <div style="display: flex; justify-content: space-between;">
                                    

                                    <div class="col-md-4" style="padding-left: 0px">
                                        <div class="form-group" id="data_agenda">
                                            <div class="input-group date">
                                                <span class="input-group-addon"><i class="fa fa-calendar"></i></span>
                                                <input autocomplete="OFF" value="<%=data_atual%>" name="agenda_data" type="text" class="form-control datasemhora" required value="<%=date%>" onBlur="verificacao_data_maior(this.value)">
                                            </div>
                                        </div>
                                    </div>


                                    <div class="col-md-4" style="padding-left: 0px;display: flex;">

                                        <div class="col-sm-12">
                                            <div class="input-group clockpicker" data-autoclose="true">
                                                <input autocomplete="off" name="agenda_hora" type="text" class="form-control" placeholder="00:00" value="<%=h_termino%>" value="<%=time%>">
                                                <span class="input-group-addon">
                                                    <span class="fa fa-clock-o"></span>
                                                </span>
                                            </div>
                                        </div>

                                    <!--  <span class="input-group-addon" style="padding-top: 3%;height: 34px;width: 41px"><i class="fa fa-clock"></i></span>
                                        <input type="text" class="form-control hora" value="<=time%>" name="agenda_hora" required> -->
                                    </div>

                                    <div class="col-md-4" style="padding-left: 0px">
                                        <input type="submit" class="btn btn-primary" value="Inserir">
                                    </div>

                                    <input hidden name="MM_Insert" value="formreuniao">
                                </div>
                            </form>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                        </div>
                    </div>

                </div>
            </div>


            <style type="text/css">
                .info-element:hover{
                    background: #ffff !important;
                    border-color: #1c84c6;
                }

                .info-element:hover  .opcoes{
                    display: block !important;
                }

                .info-element:hover  .agendamento{
                    display: block !important;
                }
            </style>

            <div class="row">
                <p id="resultao-list">
                    
                </p>
                <div id="funil_troca_resposta">
                    
                </div>
            </div>

            <!-- <div class="row">

                <div class="col-lg-12">
                    <div class="ibox">
                        <div class="ibox-content" style="height: 500px; overflow: auto">
                        
                            <h3 class="pull-left">Aguardando Aprovação</h3> 

                        <div style="clear:both"></div>
                        
                            <
                                SET rsnegoc = conexao.execute("SELECT np.idnp, np.empresa, np.responsavel, np.cidade, np.uf FROM neg_prospecter np WHERE np.situacao = 4 and "&usuario&" and data_validacao is null order by np.idnp desc")

                                IF NOT rsnegoc.EOF THEN 
                            %>
                        
                                    <ul class="sortable-list connectList agile-list" id="todo" >
                                    <
                                            WHILE NOT rsnegoc.EOF 
                                            cor_task = "info"
                                        
                                    %>
                                                <a href="neg_interacao_novo.asp?id=<=rsnegoc("idnp")%>" style="color: inherit !important;">
                                                    <li class="<%=cor_task%>-element">
                                                        <=rsnegoc("empresa")%> <=rsnegoc("cidade")%> - <=rsnegoc("uf")%> <br>
                                                        <small style="color: var(--cor-principal)"><=rsnegoc("responsavel")%></small>
                                                    </li>
                                                </a>
                                            
                                    <
                                            i = i + 1
                                            rsnegoc.movenext 
                                            WEND                                            
                                    %>                        
                                    </ul>
                            <
                                END IF
                            %>
                        </div>
                    </div>
                </div>

            </div> -->


        </div>
        <!--#include file="estrutura_rodape.asp"-->

        </div>
        </div>

    <div id="id_nada">
        
    </div>

    <!-- Modal -->
    <div id="modal_plus" class="modal fade" role="dialog">
        <div class="modal-dialog" style="width: 90%;">

            <!-- Modal content-->
            <div class="modal-content" style="padding-bottom: 70px;">
            <div class="modal-header">
                <h4 class="modal-title">Lista de <%=Application("nome_negociacoes")%></h4>
            </div>
            <div id="retorna_plus" class="modal-body">
                
            </div>
            <div class="modal-footer" style=" margin-top: 0px;">
                <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
            </div>
            </div>

        </div>
    </div>

    <!-- jquery UI -->
    <script src="js/plugins/jquery-ui/jquery-ui.min.js"></script>

    <!-- Custom and plugin javascript -->
    <script src="js/plugins/pace/pace.min.js"></script>

    <%
        aux = ""
        SET rsetapas2 = conexao.execute("SELECT idnfe FROM neg_funil_etapas")
        i = 1
        WHILE not rsetapas2.EOF

            if i <> 1 then
                aux = aux&", "
            end if

            aux = aux&"#etapa"&rsetapas2("idnfe")

            i = i +1
            rsetapas2.movenext

        WEND
    %>

    <script src="js/datatable2.js"></script>
    <script src="js/datatable2_css.js"></script>
    <script src="js/datatable2_dat.js"></script>

    <!-- Teste arrastar -->
    <script>
        function link_neg(id) {
            window.open("neg_interacao_novo.asp?id="+id)
        }
        
        function neg_plus(idnfe, conversao, nao_conversao, opcao, pesquisa) {

            console.log("Ajax/visualizacao_negociacoes_plus.asp?idnfe="+idnfe+"&conversao="+conversao+"&nao_conversao="+nao_conversao+"&opcao="+opcao+"&pesquisa="+pesquisa);

            ajaxGo({ url:"Ajax/visualizacao_negociacoes_plus.asp?idnfe="+idnfe+"&conversao="+conversao+"&nao_conversao="+nao_conversao+"&opcao="+opcao+"&pesquisa="+pesquisa, elem_return: document.getElementById("retorna_plus") });

            $("#modal_plus").modal("show");
        }

        $(document).ready(function(){

            $("<%=aux%>").sortable({
                connectWith: ".connectList",
                update: function( event, ui ) {

                    
                        text = "";
                        <%
                            SET rsetapas2 = conexao.execute("SELECT * FROM neg_funil_etapas WHERE status <> 'Inativo'")
                            WHILE not rsetapas2.EOF

                        %>
                            var etapa<%=rsetapas2("idnfe")%> = $( "#etapa<%=rsetapas2("idnfe")%>" ).sortable( "toArray" );
                            text = text + "etapa<%=rsetapas2("idnfe")%>" + window.JSON.stringify(etapa<%=rsetapas2("idnfe")%>)+"<br/>";
                            <!-- document.getElementById("resultao-list").innerHTML = text; -->
                        <%
                                rsetapas2.movenext
                            WEND
                        %>
                        
                        if (ui.sender != null)
                        {
                            
                            ajaxGo({ url:"Ajax/neg_muda_funil.asp?id_neg="+ui.item[0].id+"&id_etapa="+event.target.id+"&id_etapa_ult="+ui.sender[0].id, elem_return: document.getElementById("funil_troca_resposta") });
                        }
                }
            }).disableSelection();

        });
    </script>


    

    <!-- Função de arrastar teste -->
    <script>
        $(document).ready(function(){
            

            $(<%=aux%>).sortable({
                connectWith: ".connectList",
                update: function( event, ui ) {
                var text = "";
                <%
                    SET rsetapas2 = conexao.execute("SELECT * FROM neg_funil_etapas WHERE status <> 'Inativo' ")
                    WHILE not rsetapas2.EOF

                %>
                    
                    var etapa<%=rsetapas2("idnfe")%> = $( "#etapa<%=rsetapas2("idnfe")%>" ).sortable( "toArray" );
                    text += text + "<%=rsetapas2("etapa")%>" + window.JSON.stringify(etapa<%=rsetapas2("idnfe")%>)+"|";
                <%
                        rsetapas2.movenext
                    WEND
                %>
                    console.log(text);
                }
            }).disableSelection();

        });

    </script>

   


   <style>
        #form_inserir .content {
            height: 700px;
        }
    </style>

    

    <% if Request.Querystring("msg") = "insert" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 8000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Negociação inserida com sucesso!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "excluir" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 8000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Negociação excluida com sucesso!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "anotacao" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 8000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Anotação inserida com sucesso!")
            })
        </script>
    <%end if%>

    <script src="js/plugins/steps/jquery.steps.min.js"></script>

    <!-- Jquery Validate -->
    <script src="js/plugins/validate/jquery.validate.min.js"></script>

    <!-- Select2 -->
    <script src="js/plugins/select2/select2.full.min.js"></script>

    <style type="text/css">
        #retorna_negociacoes .ibox{
            width: 300px !important;
        }
    </style>


    <script>
        function visualizacao_negociacoes(tipo, idu) {
            $('#load_page').css({display:"block"});

            ajaxGo({ url:"Ajax/visualizacao_negociacoes.asp?tipo="+tipo+"&idu="+idu, elem_return: document.getElementById("retorna_negociacoes") });

            if (tipo == "k") {
                $('.kanban').css({backgroundColor:"var(--cor-secundaria)"});
                $('.kanban_i').css({color:"#FFF"});

                $('.lista').css({backgroundColor:"#FFF"});
                $('.lista_i').css({color:"var(--cor-secundaria)"});

            } else if (tipo == "l") {
                $('.kanban').css({backgroundColor:"#FFF"});
                $('.kanban_i').css({color:"var(--cor-secundaria)"});

                $('.lista').css({backgroundColor:"var(--cor-secundaria)"});
                $('.lista_i').css({color:"#FFF"});
            }
        }

        // $(document).ready(function() {
        //     $(".select_multiplo").select2()({
        //         placeholder: "Selecione",
        //         allowClear: true
        //     });
        // });

        $('.dinheiro').mask('#.##0,00', {reverse: true});
        $('.datasemhora').mask('00/00/0000', {placeholder: "00/00/0000"});
        $('.telefone_ddd').mask('(99) 99999-9999', {placeholder: "(00) 00000-0000"});

    </script>
    
    <script language="JavaScript" type="text/javascript" src="ajax/ajax_conteudo.js"></script>

    <!-- <script type="text/javascript">
        function abrirmodalnegociacao(){
            $('#inserir_prospector').modal('show');
            return false;
        }
    </script> -->


    <script>

        

        function pesquisa_neg(pesquisa, id_ret) {
            ajaxGo({ url:"Ajax/filtra_neg.asp?pesquisa="+pesquisa, elem_return: document.getElementById("retorno_neg"+id_ret) });
        }

        function altera_mascara_tel(ddd_pais) {
            ajaxGo({ url:"Ajax/atualiza_mascara_tel.asp?ddd_pais="+ddd_pais, elem_return: document.getElementById("retorna_input_tel") });
        }

        <%
            SET rsetapas = conexao.execute("SELECT * FROM neg_funil_etapas WHERE etapa NOT LIKE '%sem con'")
                IF NOT rsetapas.EOF THEN 
                    contador = 1
                    WHILE NOT rsetapas.EOF  

                        SET rsnegoc = conexao.execute("SELECT np.idnp, np.empresa, np.responsavel FROM neg_prospecter np WHERE np.del = 0 AND np.situacao = "&rsetapas("idnfe")&" and "&usuario&" and "&perfil&" and "&aquisicao&" and "&produto&" order by np.idnp desc")

                        IF NOT rsnegoc.EOF THEN 
        %>
                            $('.pesquisa_neg<%=contador%>').typeahead({
                            source: [
                                    <% 
                                        While not rsnegoc.EOF
                                    %>
                                            "<%=rsnegoc("idnp")%> - <%=Replace(Replace(Replace(Replace(rsnegoc("empresa"), "&", "e"), "'", ""), chr(13),""), chr(10),"")%>",
                                    <%
                                        rsnegoc.movenext 
                                        Wend

                                        rsnegoc.Close
                                        SET rsnegoc = nothing
                                    %>
                                ""]
                            }); 
        <%      
                        END IF
                    contador = contador + 1
                    rsetapas.movenext 
                    WEND
                END IF 
        %>

    </script>

    <script>
        var mem = $('#data_2 .input-group.date').datepicker({
            todayBtn: "linked",
            keyboardNavigation: false,
            forceParse: false,
            calendarWeeks: true,
            autoclose: true,
            format: 'dd/mm/yyyy',
        });
    </script>


    <script>
        var mem = $('#data_3 .input-group.date').datepicker({
            todayBtn: "linked",
            keyboardNavigation: false,
            forceParse: false,
            calendarWeeks: true,
            autoclose: true,
            format: 'dd/mm/yyyy',
        });
    </script>
     

     <!-- Clock picker -->
    <script src="js/plugins/clockpicker/clockpicker.js"></script>

    <script type="text/javascript">
        $('.clockpicker').clockpicker();
    </script>

    <style type="text/css">
        .clockpicker-popover, .popover{
            z-index: 99999 !important
        }

    </style>

    <script>
    $('.cpf').mask('#.##0,00', {reverse: true});
    </script>

    <script type="text/javascript">

        function filtrar(x)
        {
            filtro = document.getElementById("fill-"+x);
            filtro.classList.toggle("esconde");

            icone = document.getElementById("ic_filtro"+x);

            icone.classList.toggle("fa-search");
            icone.classList.toggle("fa-close");
            icone.classList.toggle("desativado");
            icone.classList.toggle("ativado");

        }
    </script>

    <script type="text/javascript">
        function excluir(idnp)
        {
            if (confirm("Você realmente deseja exluir esse prospecter?"))
            {
                //window.location.href = "neg_interacao_novo.asp?act=excluirp&id="+x
                ajaxGo({ url: "Ajax/ajax_deleta_neg_prospecter.asp?id="+idnp, elem_return: document.getElementById("id_nada")})
                document.getElementById(idnp).style.display = 'none';
                toastr.options.progressBar = true;
                toastr.options.timeOut = 8000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Negociação excluida com sucesso!")

                // pesquisa_negociacao(pesquisa, opcao);
            }
        }
    </script>

    <script type="text/javascript">
        function pesquisa_negociacao(pesquisa, opcao)
        {
            if (pesquisa.match(/./g) != null)
            {
                ocorrencias = pesquisa.match(/./g).length;
            }
            else
            {
                ocorrencias = 0;
            }
            for (i = 0; i < ocorrencias; i++)
            {
                pesquisa = pesquisa.replace('.', '');
                pesquisa = pesquisa.replace('-', '');
                pesquisa = pesquisa.replace('/', '');
            }

            let input = document.getElementById("pesquisa_nome_negociacao_input_checked_meu_funil")
            ver_todos = true;

            if (document.getElementById("pesquisa_nome_negociacao_input_checked_meu_funil"))
            {
                if(input.checked) ver_todos = false;
            }
            else
            {
                ver_todos = false;
            }

            if(pesquisa != "")
            {
                document.getElementById("status-pesquisa-negociacao").style.display="block";

                ajaxGo({ url:"Ajax/pesquisa_geral_funil.asp?pesq="+pesquisa+"&opcao="+opcao+"&meufunil="+ver_todos, elem_return: document.getElementById("retorna_negociacoes") });
            }
            else
            {
                <%IF Request.Cookies("grpro")("perm_103") = "1" THEN%>

                    document.getElementById("status-pesquisa-negociacao").style.display="block";

                    ajaxGo({ url:"Ajax/pesquisa_geral_funil.asp?pesq="+pesquisa+"&opcao="+opcao+"&meufunil="+ver_todos, elem_return: document.getElementById("retorna_negociacoes") });

                <%ELSE%>

                    swal('ATENÇÃO', 'Campo de pesquisa deve estar preenchido!', 'warning');

                    if (document.getElementById("pesquisa_nome_negociacao_input_checked_meu_funil"))
                    {
                        input.checked = true;
                    }
                    else
                    {
                        ver_todos = false;
                    }
                    
                <%END IF%>
            }

            
        }
    </script>

    <script type="text/javascript">
        function realizar_busca(x)
        {
            pesq = document.getElementById("pesq"+x).value;
           
            // alert("Ajax/filtrar_negociacao.asp?id="+x+"&empresa="+empresa+"&responsavel="+responsavel+"&cidade="+cidade);

            ajaxGo({ url:"Ajax/filtrar_negociacao_arrastar.asp?id="+x+"&pesq="+pesq, elem_return: document.getElementById("etapa"+x) });

        }

        
    </script>

    <script type="text/javascript">
        
        ver_todos = false;

        function mudar_meu_funil(input)
        {
            pesquisa = document.getElementById('pesquisa_nome_negociacao_input').value;
            opcao = "nome";

            if (pesquisa.match(/./g) != null)
            {
                ocorrencias = pesquisa.match(/./g).length;
            }
            else
            {
                ocorrencias = 0;
            }
            for (i = 0; i < ocorrencias; i++)
            {
                pesquisa = pesquisa.replace('.', '');
                pesquisa = pesquisa.replace('-', '');
                pesquisa = pesquisa.replace('/', '');
            }

            ver_todos = true;
            if(input.checked) ver_todos = false;

            if(pesquisa != "")
            {
                document.getElementById("status-pesquisa-negociacao").style.display="block";

                ajaxGo({ url:"Ajax/pesquisa_geral_funil.asp?pesq="+pesquisa+"&opcao="+opcao+"&meufunil="+ver_todos, elem_return: document.getElementById("retorna_negociacoes") });
            }
            else
            {
                <%IF Request.Cookies("grpro")("perm_103") = "1" THEN%>

                    document.getElementById("status-pesquisa-negociacao").style.display="block";

                    ajaxGo({ url:"Ajax/pesquisa_geral_funil.asp?pesq="+pesquisa+"&opcao="+opcao+"&meufunil="+ver_todos, elem_return: document.getElementById("retorna_negociacoes") });

                <%ELSE%>

                    swal('ATENÇÃO', 'Campo de pesquisa deve estar preenchido!', 'warning');
                    input.checked = true;
                    
                <%END IF%>
            }
        }

    </script>

    <style type="text/css">
        .esconde{
            display: none !important;
        }

        .desativado{
            color: #476371;
        }

        .ativado{
            color: #ed5565;
        }
    </style>
    <script type="text/javascript">
    <%
        SET rslic = conexao.execute("SELECT l.idlicenciado, l.licenciado, l.fantasia, l.cnpj, cidade FROM licenciados l WHERE l.status <> 'Inativo'")
    %>

    $('.lista_lic_neg').typeahead({
        maxItem: 0,
        source: [
            <% 
                IF NOT rslic.EOF THEN
                    While not rslic.EOF
            %>
                        "<%=rslic("cnpj")%> - <%=rslic("licenciado")%> | <%=rslic("cidade")%> | <%=rslic("fantasia")%>",
            <%
                    rslic.movenext 
                    Wend
                    rslic.movefirst
                END IF
            %>
        ""],
        limit: 10
    });

    </script>

    <script type="text/javascript">

        function fechar_modal_anotacao() {
            $("#modalanota").modal("hide");

            toastr.options.progressBar = true;
            toastr.options.timeOut = 8000;
            toastr.options.extendedTimeOut = 6000;
            toastr.success("Anotação inserida com sucesso!")
        }
        
    </script>

    <style type="text/css">

        .barra::-webkit-scrollbar{
            width: 6px;
            background: #aaa;
        }
        

        .info-element:hover .icones {
            z-index: 1 !important;
        }
        
    </style>

</body>


</html>
