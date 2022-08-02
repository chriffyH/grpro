<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
	
<%
'Conexao com o banco
Set conexao = Server.CreateObject("ADODB.Connection")
%>

<!--#include file="../../Connections/conectar.asp"-->
<!--#include file="../../Connections/rede.asp" -->
<!--#include file="../funcoes.asp" -->

<%                     
	meufunil = Request.querystring("meufunil")
    IF meufunil = "" THEN meufunil = "false"

	pesquisa = Request.querystring("pesq")
    opcao = Request.querystring("opcao")
%>
    <script type="text/javascript">

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

    </script>
<%

    IF opcao = "nome" THEN
	   pesquisa_sql = "and (np.responsavel like '%"&pesquisa&"%' or np.empresa like '%"&pesquisa&"%' or np.cidade like '%"&pesquisa&"%' or replace(replace(replace(np.cpf, '.', ''), '/', ''), '-', '') LIKE '%"&pesquisa&"%')"

    ELSEIF opcao = "estado" THEN

        pesquisa_sql = "AND (np.uf like '%"&pesquisa&"%')"

    ELSEIF opcao = "responsavel" THEN

        IF pesquisa <> "" THEN
            pesquisa_sql = " AND EXISTS(SELECT id_usuario FROM neg_prospecter_responsaveis pr where id_negp = np.idnp and id_usuario = "&pesquisa&")"
        ELSE 
            pesquisa_sql = ""
        END IF

    ELSEIF opcao = "perfil" THEN

        IF pesquisa <> "" THEN
            pesquisa_sql = "AND np.id_perfil = "&pesquisa&" "
        ELSE
            pesquisa_sql = ""
        END IF

    ELSEIF opcao = "interesse" THEN

        IF pesquisa <> "" THEN
            pesquisa_sql = "AND np.id_produto = "&pesquisa&" "
        ELSE
            pesquisa_sql = ""
        END IF

    ELSEIF opcao = "aquisicao" THEN

        pesquisa_sql = "AND np.id_aquisicao = "&pesquisa&" "

    ELSE

        pesquisa_sql = ""

    END IF  

    aquisicao = "np.id_aquisicao is not null"  
    IF Request.Form("aquisicao") <> "" and Request.Form("aquisicao") <> "todos" THEN 
        aquisicao = "np.id_aquisicao = "&Request.Form("aquisicao")
    END IF 

    
    SET rsetapas = conexao.execute("SELECT * FROM neg_funil_etapas WHERE status = 'Ativo' ORDER BY ordem ASC")
    contador = 1
    i = 1

    IF Request.Cookies("grpro")("perm_103") = "1" or meufunil = "true" THEN
        visualizar_tipo = "np.idnp is not null"
    ELSE 
        visualizar_tipo = " EXISTS(SELECT id_usuario FROM neg_prospecter_responsaveis pr where id_negp = np.idnp and id_usuario = "&Request.Cookies("grpro")("idusuario")&")"
    END IF 

    ' IF Request.Cookies("grpro")("perm_103") = "1" and meufunil = "false" THEN
    '     visualizar_tipo = " EXISTS(SELECT id_usuario FROM neg_prospecter_responsaveis pr where id_negp = np.idnp and id_usuario = "&Request.Cookies("grpro")("idusuario")&")"
    ' END IF

    IF NOT rsetapas.EOF THEN 
        WHILE NOT rsetapas.EOF

        ' Pega mes e ano atual
        dat3 = date()
        mes = month(dat3)
        IF mes < 10 THEN
            mes = 0&mes
        END IF
        dat3 = year(dat3)&"-"&mes

        IF (rsetapas("conversao") = "Sim" OR rsetapas("nao_conversao") = "Sim") AND opcao <> "nome" THEN
            SET rstotal = conexao.execute("SELECT count(DISTINCT np.idnp) as total FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp  WHERE "&visualizar_tipo&" AND np.del = 0 and np.situacao = "&rsetapas("idnfe")&"  and "&aquisicao&" and month(nf.data) = "&month(date)&" and year(nf.data) = "&year(date)&" "&pesquisa_sql&" ORDER BY nf.idnf DESC LIMIT 10")

            SET rstotaladesao = conexao.execute("SELECT SUM(DISTINCT valor_adesao) as totaladesao FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp INNER JOIN usuarios u ON u.idusuario = nf.id_usuario WHERE "&visualizar_tipo&" AND np.del = 0 and np.situacao = "&rsetapas("idnfe")&" and "&aquisicao&" and month(nf.data) = "&month(date)&" and year(nf.data) = "&year(date)&" "&pesquisa_sql&" ORDER BY nf.idnf DESC")
        ELSE
            SET rstotal = conexao.execute("SELECT Count(DISTINCT idnp) as total FROM neg_prospecter np INNER JOIN neg_prospecter_responsaveis npr ON npr.id_negp = np.idnp INNER JOIN usuarios u ON u.idusuario = npr.id_usuario WHERE "&visualizar_tipo&" and situacao = "&rsetapas("idnfe")&" AND np.del = 0  and "&aquisicao&" "&pesquisa_sql&" LIMIT 10")

            SET rstotaladesao = conexao.execute("SELECT SUM(valor_adesao) as totaladesao FROM neg_prospecter np INNER JOIN neg_prospecter_responsaveis npr ON npr.id_negp = np.idnp INNER JOIN usuarios u ON u.idusuario = npr.id_usuario WHERE "&visualizar_tipo&" and situacao = "&rsetapas("idnfe")&" AND np.del = 0 and "&aquisicao&" "&pesquisa_sql)
        END IF

        IF IsNull(rstotaladesao("totaladesao")) THEN 
            totaladesao = "0,00"
        ELSE 
            totaladesao = FormatNumber(rstotaladesao("totaladesao"))
        END IF 
%>

<div class="col-lg-3">
    <div class="ibox">
        <div class="ibox-content" style="height: 568px; background-color: #e8e8ea; padding: 0px 6px 20px 6px;">
           <!--  <i onclick="filtrar('<=rsetapas("idnfe")%>')" id="ic_filtro<=rsetapas("idnfe")%>" style="float: right;margin-left: 3%;cursor: pointer;" class="fa fa-search desativado" aria-hidden="true"></i> -->
            
            <div style="overflow: auto;"></div>

            <%
                etapas = pontinhos(rsetapas("etapa"), 25) 
            %>
            
            <h3 style="margin-left: 3%; font-size: 17px; margin-bottom: 0;"><strong><%=etapas%></strong></h3> 

            <span style="margin-left: 3%;" id="count<%=rsetapas("idnfe")%>"><%=rstotal("total")%></span>&nbsp; contatos

            <label class="label label-success" style="background-color: #5cb85c; margin-top: 1%; margin-right: 3%; float: right; ">R$ <span id="count-adesao<%=rsetapas("idnfe")%>"><%=totaladesao%></span></label>

            <span style="position: absolute;right: 13%;top: 7%;cursor: pointer;">
                
            </span>
            <div style="clear:both"></div>
        
            <%        
                IF (rsetapas("conversao") = "Sim" OR rsetapas("nao_conversao") = "Sim") AND opcao <> "nome" THEN

                    SET rsnegoc = conexao.execute("SELECT DISTINCT np.idnp, np.cpf, np.empresa, np.responsavel, np.cidade, np.uf, np.data_validacao FROM neg_prospecter np INNER JOIN neg_funil nf ON np.idnp = nf.id_negp WHERE "&visualizar_tipo&" and np.situacao = "&rsetapas("idnfe")&" and "&aquisicao&" AND np.del = 0 and month(nf.data) = "&month(date)&" and year(nf.data) = "&year(date)&" "&pesquisa_sql&" ORDER BY np.idnp DESC limit 10")
                ELSE
                    SET rsnegoc = conexao.execute("SELECT DISTINCT np.idnp, np.cpf, np.empresa, np.responsavel, np.cidade, np.uf, np.data_validacao FROM neg_prospecter np WHERE "&visualizar_tipo&" and np.situacao = "&rsetapas("idnfe")&" and "&aquisicao&" "&pesquisa_sql&" AND np.del = 0 ORDER BY np.idnp DESC limit 10")
                END IF

                SET rsid_entrada = conexao.execute("SELECT idnfe FROM neg_funil_etapas WHERE entrada = 'Sim'")

            %>

            <div class="barra" style="height: 500px; overflow: auto;overflow-x:hidden;">
                <style type="text/css">

                    .info-element:hover .icones {
                        z-index: 1 !important;
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
                            
                                    <li class="info-element" id="<%IF NOT habilita_acao THEN response.write("nao-")%><%=rsnegoc("idnp")%>" style="position: relative; padding: 3%; padding-bottom: 0px !important;">  
                                
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
                                        <a onClick="neg_plus(<%=rsetapas("idnfe")%>, '<%=rsetapas("conversao")%>', '<%=rsetapas("nao_conversao")%>', '<%=opcao%>', '<%=pesquisa%>')" target="_blank">
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

<div class="row">
    <p id="resultao-list">
        
    </p>
    <div id="funil_troca_resposta">
        
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


<script src="https://cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.10.19/js/dataTables.bootstrap4.min.js"></script>

<script>

    
    $(document).ready(function(){

        $("<%=aux%>").sortable({
            connectWith: ".connectList",
            update: function( event, ui ) {

                    var id_auxiliar = ui.item[0].id;

                    if(id_auxiliar.match(/nao/))
                    {
                        swal('ATENÇÃO', 'Ação cancelada, você não tem permissão para isso!', 'warning');
                        pesquisa_negociacao('<%=pesquisa%>', '<%=opcao%>')
                        exit();
                    }
                
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
                        
                        ajaxGo({ url:"Ajax/neg_muda_funil.asp?opcao=<%=opcao%>&pesquisa=<%=pesquisa%>&id_neg="+ui.item[0].id+"&id_etapa="+event.target.id+"&id_etapa_ult="+ui.sender[0].id, elem_return: document.getElementById("funil_troca_resposta") });
                    }
            }
        }).disableSelection();
    });

</script>

<script type="text/javascript">
    document.getElementById("status-pesquisa-negociacao").style.display="none";

    $('.tooltip-demo').tooltip({
        selector: "[data-toggle=tooltip]",
        container: "body"
    });

    


</script>