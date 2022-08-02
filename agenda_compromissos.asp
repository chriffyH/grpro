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

<!-- Select2 -->
<script src="js/plugins/select2/select2.full.min.js"></script>

<style>
    .md-skin .wrapper-content {
        padding: 15px 0px 15px 0px;
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

    .popover {
        z-index: 100000 !important;
    }

</style>

<%
    IF Request.Form("MM_insert") = "forminserir" THEN 
        dia = day(Request.Form("inicio"))
        IF dia < 10 THEN dia = "0"&dia 
        mes = month(Request.Form("inicio"))
        IF mes < 10 THEN mes = "0"&mes 
        inicio = year(Request.Form("inicio"))&"/"&mes&"/"&dia

        inicio_completo = inicio & " " & request.form("h_inicio")&":00"

        dia = day(Request.Form("termino"))
        IF dia < 10 THEN dia = "0"&dia 
        mes = month(Request.Form("termino"))
        IF mes < 10 THEN mes = "0"&mes 
        termino = year(Request.Form("termino"))&"/"&mes&"/"&dia

        termino_completo = termino & " " & request.form("h_termino")&":00"

        descricao = Replace(Request.Form("descricao"), VbCrLf, "<br>")
        descricao = Replace(descricao, Chr(13) & chr(10), "<br>")
        descricao = replace(descricao,"'","&apos;")
        descricao = replace(descricao,"""","&quot;") 

        titulo = Request.Form("titulo")
        titulo = replace(titulo,"'","&apos;")
        titulo = replace(titulo,"""","&quot;") 

        conexao.execute("INSERT INTO agenda_compromissos (id_usuario, id_setor, titulo, descricao, data_inicio, data_termino, tipo) VALUES ( "&Request.Cookies("grpro")("idusuario")&", "&Request.Cookies("grpro")("setor")&", '"&titulo&"', '"&descricao&"', '"&inicio_completo&"', '"&termino_completo&"', '"&Request.Form("tipo")&"')")

        SET rs_idac = conexao.execute("SELECT idac FROM agenda_compromissos ORDER BY idac DESC LIMIT 1;")

        id_idac = rs_idac("idac")
        rs_usuarios = request.Form("usuarios_inserir")

        IF rs_usuarios <> "" THEN
            separar = split(rs_usuarios, ",")
            FOR x = lbound(separar) TO ubound(separar)
                conexao.execute("INSERT INTO agenda_marcacoes (id_ac, id_usuario) VALUES ("&id_idac&", "&separar(x)&")")
            NEXT
        END IF 

        Response.Redirect("agenda_compromissos.asp?msg=insert")
    END IF 

    IF Request.Form("MM_Edit") = "formeditar" THEN 
        dia = day(Request.Form("inicio"))
        IF dia < 10 THEN dia = "0"&dia 
        mes = month(Request.Form("inicio"))
        IF mes < 10 THEN mes = "0"&mes 
        inicio = year(Request.Form("inicio"))&"/"&mes&"/"&dia

        inicio_completo = inicio & " " & request.form("h_inicio")&":00"

        dia = day(Request.Form("termino"))
        IF dia < 10 THEN dia = "0"&dia 
        mes = month(Request.Form("termino"))
        IF mes < 10 THEN mes = "0"&mes 
        termino = year(Request.Form("termino"))&"/"&mes&"/"&dia

        termino_completo = termino & " " & request.form("h_termino")&":00"

        descricao = Replace(Request.Form("descricao"), VbCrLf, "<br>")
        descricao = Replace(descricao, Chr(13) & chr(10), "<br>")
        descricao = replace(descricao,"'","&apos;")
        descricao = replace(descricao,"""","&quot;") 

        conexao.execute("UPDATE agenda_compromissos SET titulo = '"&Request.Form("titulo")&"', descricao = '"&Request.Form("descricao")&"', data_inicio = '"&inicio_completo&"', data_termino = '"&termino_completo&"', tipo = '"&Request.Form("tipo")&"' WHERE idac = "&Request.Form("idac"))

        Response.Redirect("agenda_compromissos.asp?msg=edit")
    END IF 

    IF Request.QueryString("act") = "excluir" THEN 
        conexao.execute("UPDATE agenda_compromissos SET status = 'Inativo' WHERE idac = "&Request.QueryString("idac"))

        Response.Redirect("agenda_compromissos.asp?msg=deletado")
    END IF  
%>

<body>
	<%
		IF request.form("config_calendario") = "Sim" THEN

			ani = request.form("calendario_aniversario")
			doc = request.form("calendario_documentos")
			flag_ani = 0
			flag_doc = 0
			IF ani = "on" THEN flag_ani = 1
			IF doc = "on" THEN flag_doc = 1
			conexao.execute("UPDATE usuarios SET calendario_aniversario = "&flag_ani&", calendario_documento = "&flag_doc&" WHERE idusuario = "&Request.Cookies("grpro")("idusuario"))
			response.redirect("agenda_compromissos.asp?msg=config-calendario")

		END IF

		SET config_calendario = conexao.execute("SELECT calendario_aniversario, calendario_documento FROM usuarios WHERE idusuario = "&Request.Cookies("grpro")("idusuario"))
	%>
    <div id="wrapper">

    <nav class="navbar-default navbar-static-side" role="navigation">
        <!--#include file="include_left_side.asp"-->
    </nav>

        <div id="page-wrapper" class="gray-bg">
        <div class="row border-bottom">
            <!--#include file="include_topo.asp"-->
        </nav>
        </div>
            <div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-8">
                    <h2>Compromissos</h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="default.asp">Home</a>
                        </li>
                        <li class="active">
                            <strong>Gerenciamento de Compromissos</strong>
                        </li>
                    </ol>
                </div>
                <div class="col-lg-4" style="text-align: right;">
                    <h2></h2>

                    <a id="btn_tour_6" class="btn btn-success" data-toggle="modal" data-target="#mod_config_calendario">Configurações</a>
                    <a class="btn btn-success" data-toggle="modal" data-target="#modinserir">Novo compromisso</a>

                </div>
            </div>

        <div class="wrapper wrapper-content animated fadeInRight ecommerce">

            <div class="row">
                <div class="col-lg-12">
                    <div class="ibox">
                        <div class="ibox-content" style="padding: 0">

                            <div class="ibox">
                                <div class="ibox-content" style="padding: 15px">

                                    <div class="ibox-content" style="padding: 0">

                                         <div id="legenda_calendario" style="border: 1px solid #f1f1f1;margin-bottom: 17px;">
                                            
                                            <h5 style="border-bottom: 1px solid #f1f1f1;padding-bottom: 5px;padding: 4px 10px;margin-bottom: 0px;">Legenda</h5>
                                            <ul style="padding: 8px;list-style: none;">

                                                <li>
                                                    <i style="color: purple" class="fa fa-circle" aria-hidden="true"></i> - Data de aniversário/inauguração dos <%=Application("nome_cliente")%>.
                                                </li>

                                                <li>
                                                    <i style="color: #003366" class="fa fa-circle" aria-hidden="true"></i> - Data de vencimento dos documentos.
                                                </li>

                                                <li>
                                                    <i style="color: green" class="fa fa-circle" aria-hidden="true"></i> - Compromisso pessoal.
                                                </li>

                                                <li>
                                                    <i style="color: #3788d8" class="fa fa-circle" aria-hidden="true"></i> - Compromisso Global.
                                                </li>

                                                <li>
                                                    <i style="color: #f8ac59" class="fa fa-circle" aria-hidden="true"></i> - Compromisso do Setor.
                                                </li>

                                                <li>
                                                    <i style="color: #000" class="fa fa-circle" aria-hidden="true"></i> - Reunião de prospecter.
                                                </li>

                                            </ul>


                                        </div>

                                        <div id="calendar"></div>

                                       

                                    </div>



                                </div>
                            </div>

                        </div>
                    </div>
                </div>
            </div>


        </div>
        <!--#include file="estrutura_rodape.asp"-->

        </div>
        </div>

    <!-- Modal -->
    <div id="modinserir" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Novo Compromisso</h4>
                </div>
                <div class="modal-body">
                    <form ACTION="<%=MM_editAction%>" METHOD="POST" name="forminserir" id="forminserir">
                        <strong>Título:</strong> <br>
                        <input name="titulo" type="text" class="form-control" maxlength="35" required> 

                        <strong>Descrição:</strong> <br>
                        <textarea class="form-control" name="descricao"></textarea>

                        <div class="col-lg-6" style="padding-left: 0">
                            <div class="col-lg-6" style="padding-left: 0;">
                                <strong>Início:</strong> <br>

                                <div class="form-group" id="data_inicio">     
                                    <input autocomplete="OFF" name="inicio" type="date" class="form-control" required value="<%=transforma_data_ua(date)%>" onChange="verificacao_data_maior_agenda(this.value, 'inicio')">
                                </div>
                                

                            </div>

                            <div class="col-lg-6" style="padding-right: 0;">
                                <strong>Horário de início:</strong>
                                <div class="input-group clockpicker" data-autoclose="true">
                                    <input autocomplete="off" name="h_inicio" type="text" class="form-control" placeholder="08:00" onChange="add_horario(this.value)">
                                    <span class="input-group-addon">
                                        <span class="fa fa-clock-o"></span>
                                    </span>
                                </div>
                            </div>
                        </div>

                        <div class="col-lg-6" style="padding-right: 0;">

                            <div class="col-lg-6" style="padding-left: 0;">
                                <strong>Término:</strong> <br>
                                <div class="form-group" id="data_termino"> 
                                    <input autocomplete="OFF" value="<%=transforma_data_ua(date)%>" name="termino" type="date" class="form-control" required onChange="verificacao_data_maior_agenda(this.value, 'termino')">
                                </div>
                            </div>

                            <div class="col-lg-6" style="padding-right: 0;">
                                <strong>Horário de término:</strong>
                                <div class="input-group clockpicker" data-autoclose="true">
                                    <input autocomplete="off" name="h_termino" type="text" class="form-control" placeholder="08:00">
                                    <span class="input-group-addon">
                                        <span class="fa fa-clock-o"></span>
                                    </span>
                                </div>
                            </div>
                            
                        </div>

                        
                        <div class="col-lg-12" style="padding: 0px;">
                            <div class="col-lg-6" style="padding-left: 0px;">
                                <strong>Marcar:</strong> <br>
                                <select name="usuarios_inserir" class="form-control select_multiplo" multiple="multiple">
                                    <%
                                        SET rs_marcar_user = conexao.execute("SELECT u.idusuario, u.usuario, s.setor FROM usuarios u INNER JOIN setores s ON u.id_setor = s.idsetor WHERE u.status = 'Ativo' and u.idusuario <> "&Request.Cookies("grpro")("idusuario"))

                                        WHILE NOT rs_marcar_user.EOF
                                    %>
                                            <option value="<%=rs_marcar_user("idusuario")%>">
                                                <%=rs_marcar_user("usuario")%> (<%=rs_marcar_user("setor")%>) 
                                            </option>
                                    <%
                                            rs_marcar_user.movenext 
                                        WEND
                                    %>
                                </select>
                            </div>

                            <div class="col-lg-6" style="padding-right: 0;">
                                <strong>Tipo:</strong> <br>
                                <select name="tipo" class="form-control" required>
                                    <option value="">Selecione</option>
                                    <option value="Pessoal">Pessoal</option>
                                    <option value="Setor">Setor</option>
                                    <option value="Global">Global</option>
                                </select>
                            </div>

                        </div>

                        <input class="btn btn-primary center-block btt_custom_2" name="salvar" type="submit" id="salvar" value="Salvar" style="position: relative; clear: both; top: 10px;">
                        <input type="hidden" name="MM_insert" value="forminserir">
                    </form> 
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                </div>
            </div>

        </div>
    </div>

    <div class="modal inmodal" id="modcompromisso" tabindex="-1" role="dialog" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content animated bounceInRight">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">&times;</span><span class="sr-only">Fechar</span></button>
                    <i class="fa fa-calendar modal-icon"></i>
                    <h4 class="modal-title">Detalhes do seu compromisso</h4>
                </div>
                <div id="retorna_compromisso">

                </div>
            </div>
        </div>
    </div>

    <script>
        $(document).ready(function() {
            $(".select_multiplo").select2()({
                placeholder: "Selecione",
                allowClear: true
            });
        });
    </script>

    <link href="css/plugins/clockpicker/clockpicker.css" rel="stylesheet">

    <script>

        function verificacao_data_maior_agenda(data, name) {

            <%
                dia_a = day(date)
                IF dia_a < 10 THEN dia_a = "0"&dia_a 
                mes_a = month(date)
                IF mes_a < 10 THEN mes_a = "0"&mes_a 
                data_a = year(date)&"-"&mes_a&"-"&dia_a
            %>

            dia = data.substring(0,2);
            mes = data.substring(3,5);
            ano = data.substring(6,10);

            var data_input = ano + "-" + mes + "-" + dia

            if(moment('<%=data_a%>').isAfter(data_input)) {
                swal({
                    title: "Sistema",
                    text: "A data inserida deve ser maior que a de hoje."
                });

                $("input[name="+name+"]").val("");
            }
            else {
                if (name === "inicio") {
                    $("input[name=termino]").val(data);
                }
            }
        }

        function add_horario(horario) {

            hora = horario.substring(0,2);
            minuto = horario.substring(3,5);

            hora_termino = parseInt(hora) + 1

            if (hora_termino < 10) {
                hora_termino = String(hora_termino)
                hora_termino = 0 + hora_termino
            }

            var hora_input = hora_termino + ":" + minuto

            $("input[name=h_termino]").val(hora_input);
        }

        function editar_compromisso(idac) {
            ajaxGo({ url:"Ajax/agenda_editar.asp?act=del&idac="+idac, elem_return: document.getElementById("retorna_compromisso") });
        }

        function deleta_marcacao(idac, idam) {
            ajaxGo({ url:"Ajax/agenda_marcacoes.asp?act=del&idam="+idam+"&idac="+idac, elem_return: document.getElementById("retorna_marcacoes") });
        }

        function marcar_usuario(idac) {
            
            usuarios = $("select[name=usuarios]").val();
            usuarios = usuarios.toString().replace(/,/g, '-');
            ajaxGo({ url:"Ajax/agenda_marcacoes.asp?act=inserir&usuarios="+usuarios+"&idac="+idac, elem_return: document.getElementById("retorna_marcacoes") });
        }

        $(document).ready(function() {
            var initialLocaleCode = 'pt-br';
            var localeSelectorEl = document.getElementById('locale-selector');
            var calendarEl = document.getElementById('calendar');

            var calendar = new FullCalendar.Calendar(calendarEl, {

            eventClick: function(info) {

                id = info.event.id;
                id = id.split("-");
                console.log(id);
                if (id.length == 1)
                {
                    ajaxGo({ url:"Ajax/agenda_compromisso.asp?idac="+info.event.id, elem_return: document.getElementById("retorna_compromisso") });
                    $("#modcompromisso").modal("show");
                }
                else if (id.length == 2)
                {
                    window.open('dashboard_lic_1.asp?act=editar&id='+id[1],'_blank')
                }
                else if (id.length == 3)
                {
                    window.open('documentos.asp?id='+id[2],'_blank')
                }
            },

            plugins: [ 'interaction', 'dayGrid', 'timeGrid', 'list' ],
            header: {
                left: 'prev,next today',
                center: 'title',
                right: 'dayGridMonth,timeGridWeek,timeGridDay,listMonth'
            },
            <%
                dia = day(date)
                IF dia < 10 THEN dia = "0"&dia 
                mes = month(date)
                IF mes < 10 THEN mes = "0"&mes 
                data_atual = year(date)&"-"&mes&"-"&dia
            %>
            defaultDate: '<%=data_atual%>',
            locale: initialLocaleCode,
            buttonIcons: false, // show the prev/next text
            weekNumbers: true,
            navLinks: true, // can click day/week names to navigate views
            editable: false,
            eventLimit: true, // allow "more" link when too many events
            events: 
                [

                    <%
                        ' INAUGURAÇÃO 
                        IF config_calendario("calendario_aniversario") THEN
                        	SET rs_inauguracao = conexao.execute("SELECT * FROM licenciados WHERE status = 'Ativo' AND inauguracao IS NOT NULL AND inauguracao <> '0000-00-00' ")

	                        WHILE NOT rs_inauguracao.EOF

	                            dia = day(rs_inauguracao("inauguracao"))
	                            IF dia < 10 THEN dia = "0"&dia 
	                            mes = month(rs_inauguracao("inauguracao"))
	                            IF mes < 10 THEN mes = "0"&mes 
	                            rs_data_inauguracao = year(date)&"-"&mes&"-"&dia

	                            rs_data_inauguracao = transforma_data_ua(rs_data_inauguracao)

	                            fantasia = rs_inauguracao("fantasia")
	                            IF fantasia = "" or isnull(fantasia) THEN
	                                fantasia = rs_inauguracao("licenciado")
	                            END IF

                            	IF rs_data_inauguracao <> "--" THEN 
                    %>
	                                {
	                                    title: '<%=fantasia%> (Inauguração)',
	                                    // url: '',
	                                    start: '<%=rs_data_inauguracao%>T00:00:00',
	                                    end: '<%=rs_data_inauguracao%>T23:00:00',
	                                    id: 'lic-<%=rs_inauguracao("idlicenciado")%>',
	                                    color: 'purple' // override!
	                                },

                    <%
                            	END IF
                            	rs_inauguracao.movenext
                        	WEND
                        END IF

                       	' DOCUMENTOS
                        IF config_calendario("calendario_documento") THEN
                        	SET rsdocumentos = conexao.execute("SELECT * FROM documentos WHERE del = 0")

                        	IF NOT rsdocumentos.EOF THEN
                            	WHILE NOT rsdocumentos.EOF

	                                dia = day(rsdocumentos("data_val"))
	                                IF dia < 10 THEN dia = "0"&dia 
	                                mes = month(rsdocumentos("data_val"))
	                                IF mes < 10 THEN mes = "0"&mes 
	                                ano = year(rsdocumentos("data_val"))
	                                vencimento = ano&"-"&mes&"-"&dia

	                                arquivo = rsdocumentos("arquivo")
	                                arquivo_titulo = rsdocumentos("titulo")
	                                IF vencimento <> "--" THEN
                    %>
	                                    {
	                                        title: '<%=arquivo%> (<%=arquivo_titulo%>)',
	                                        // url: '',
	                                        start: '<%=vencimento%>T00:00:00',
	                                        end: '<%=vencimento%>T23:00:00',
	                                        id: 'doc-venc-<%=rsdocumentos("id_licenciado")%>',
	                                        color: '#003366' // override!
	                                    },
                    <%
                                	END IF
                                	rsdocumentos.movenext
                            	WEND
                        	END IF
                        END IF 
                        
                        ' COMPROMISSOS
                        SET rscompromissos = conexao.execute("SELECT idac, titulo, descricao, data_inicio, data_termino, tipo, id_np FROM agenda_compromissos ac1 WHERE ac1.status = 'Ativo' AND (((tipo = 'Pessoal' and id_usuario = "&Request.Cookies("grpro")("idusuario")&") OR (tipo = 'Setor' AND id_setor = "&Request.Cookies("grpro")("setor")&") OR (tipo = 'Global')) OR EXISTS (SELECT id_usuario FROM agenda_marcacoes am WHERE "&Request.Cookies("grpro")("idusuario")&" = am.id_usuario and id_ac = idac))")

                      

                        IF NOT rscompromissos.EOF THEN
                            WHILE NOT rscompromissos.EOF

                                dia = day(rscompromissos("data_inicio"))
                                IF dia < 10 THEN dia = "0"&dia 
                                mes = month(rscompromissos("data_inicio"))
                                IF mes < 10 THEN mes = "0"&mes 
                                data_inicio = year(rscompromissos("data_inicio"))&"-"&mes&"-"&dia

                                hora = Hour(rscompromissos("data_inicio"))
                                IF hora < 10 THEN hora = "0"&hora

                                minuto = Minute(rscompromissos("data_inicio"))
                                IF minuto < 10 THEN minuto = "0"&minuto

                                hora_inicio = hora & ":" & minuto & ":00"

                                dia = day(rscompromissos("data_termino"))
                                IF dia < 10 THEN dia = "0"&dia 
                                mes = month(rscompromissos("data_termino"))
                                IF mes < 10 THEN mes = "0"&mes 
                                data_termino = year(rscompromissos("data_termino"))&"-"&mes&"-"&dia

                                hora = Hour(rscompromissos("data_termino"))
                                IF hora < 10 THEN hora = "0"&hora

                                minuto = Minute(rscompromissos("data_termino"))
                                IF minuto < 10 THEN minuto = "0"&minuto

                                hora_termino = hora & ":" & minuto & ":00"

                                SELECT CASE rscompromissos("tipo") 
                                    CASE "Pessoal" cor_evento = "green"
                                    CASE "Global" cor_evento = "#3788d8"
                                    CASE "Setor" cor_evento = "#f8ac59"
                                END SELECT

                                IF NOT isnull(rscompromissos("id_np")) THEN cor_evento = "#000"

                                IF data_inicio <> "--" THEN 

                                    descricao = Replace(rscompromissos("descricao"), VbCrLf, " ")
                                    descricao = Replace(descricao, "\", "")
                                    descricao = Replace(descricao, """", "")
                                    descricao = Replace(descricao, "'", "")
                                    descricao = Replace(descricao, chr(13), "")
                                    descricao = Replace(descricao, chr(12), "")
                                    descricao = Replace(descricao, chr(11), "")

                    %>
                                    {
                                        title: '<%=rscompromissos("titulo")%> (<%=descricao%>)',
                                        // url: '',
                                        start: '<%=data_inicio%>T<%=hora_inicio%>',
                                        end: '<%=data_termino%>T<%=hora_termino%>',
                                        id: '<%=rscompromissos("idac")%>',
                                        color: '<%=cor_evento%>' // override!
                                    },
                    <%
                                END IF 
                            data_prevista_termino = ""
                            rscompromissos.movenext
                            WEND
                            rscompromissos.movefirst
                        END IF 
                    %>
                ],
                timeFormat: 'H(:mm)' // uppercase H for 24-hour clock
            });

            calendar.render();
        });

        $('.datasemhora').mask('00/00/0000', {placeholder: "00/00/0000"});

    </script>




    <% if Request.Querystring("msg") = "insert" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 8000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Novo compromisso adicionado!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "deletado" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 8000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Compromisso inativado!")
            })
        </script>
    <%end if%>

    <!-- Clock picker -->
    <script src="js/plugins/clockpicker/clockpicker.js"></script>

    <script>
        $('.clockpicker').clockpicker();
    </script>

    <link href='js/packages/core/main.css' rel='stylesheet' />
    <link href='js/packages/daygrid/main.css' rel='stylesheet' />
    <link href='js/packages/timegrid/main.css' rel='stylesheet' />
    <link href='js/packages/list/main.css' rel='stylesheet' />
    <script src='js/packages/core/main.js'></script>
    <script src='js/packages/core/locales-all.js'></script>
    <script src='js/packages/interaction/main.js'></script>
    <script src='js/packages/daygrid/main.js'></script>
    <script src='js/packages/timegrid/main.js'></script>
    <script src='js/packages/list/main.js'></script>

    <script type="text/javascript">

        var mem = $('#data_termino .input-group.date').datepicker({
        todayBtn: "linked",
        keyboardNavigation: false,
        forceParse: false,
        calendarWeeks: true,
        autoclose: true,
        format: 'dd/mm/yyyy',
    });

        var mem = $('#data_inicio .input-group.date').datepicker({
        todayBtn: "linked",
        keyboardNavigation: false,
        forceParse: false,
        calendarWeeks: true,
        autoclose: true,
        format: 'dd/mm/yyyy',
    });


    </script>

    <script type="text/javascript">
       

            var mem = $('#data_anotacao .input-group.date').datepicker({
            todayBtn: "linked",
            keyboardNavigation: false,
            forceParse: false,
            calendarWeeks: true,
            autoclose: true,
            format: 'dd/mm/yyyy',
        });
    </script>

    <div id="mod_config_calendario" class="modal fade in" role="dialog" >
	    <div class="modal-dialog">

	        <!-- Modal content-->
	        <div class="modal-content" style="width: 50%;margin: auto;">
	            <div class="modal-header">
	                <button type="button" class="close" data-dismiss="modal">×</button>
	                <h4 class="modal-title">Configurações do calendário</h4>
	            </div>
	            <div class="modal-body">
	            <form action="agenda_compromissos.asp" method="POST">

	                <div class="skin-settings" style="">

		                <div class="setings-item" style="min-height: 70px;padding: 0px">
		                    <h5><strong>Aparecer aniversário/inauguração de <%=Application("nome_cliente")%></strong></h5>
		                    <div class="switch" style="float: left;">
		                      <div class="onoffswitch">
		                          <input type="checkbox" name="calendario_aniversario" class="onoffswitch-checkbox" id="calendario_aniversario" <%IF config_calendario("calendario_aniversario") THEN response.write("checked")%>>
		                          <label class="onoffswitch-label" for="calendario_aniversario">
		                              <span class="onoffswitch-inner"></span>
		                              <span class="onoffswitch-switch"></span>
		                              <input type="hidden" id="menu_fase" value="Não">
		                          </label>
		                      </div>
		                    </div>  
		                </div>

		                <div class="setings-item" style="min-height: 70px; padding: 0px;">
		                    <h5><strong>Aparecer data de validade de documentos</strong></h5>
		                    <div class="switch" style="float: left;">
		                      <div class="onoffswitch">
		                          <input type="checkbox" name="calendario_documentos" class="onoffswitch-checkbox" id="calendario_documentos" <%IF config_calendario("calendario_documento") THEN response.write("checked")%>>
		                          <label class="onoffswitch-label" for="calendario_documentos">
		                              <span class="onoffswitch-inner"></span>
		                              <span class="onoffswitch-switch"></span>
		                              <input type="hidden" name="email_tarefa" id="email_tarefa" value="Sim">
		                          </label>
		                      </div>
		                    </div> 
		                </div>

	                </div>

	                <input type="submit" value="Salvar" class="btn btn-primary" style="display: block;">
	                <input type="hidden" value="Sim" name="config_calendario">

	            </form>
	            </div>
	            <div class="modal-footer">
	                <button type="button" class="btn btn-default" data-dismiss="modal">Cancelar</button>
	            </div>
	        </div>

	    </div>
	</div>

	<script type="text/javascript">
		<%IF request.querystring("msg") = "config-calendario" THEN%>
			toastr.success('Configurações Salvas.','Sucesso');
		<%END IF%>

	</script>

</body>




</html>