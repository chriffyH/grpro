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
<!--#include file="wpp.asp" -->

<%
    Response.Cookies("grpro")("tamanho_contatos_wpp") = 10
    Response.Cookies("grpro")("tamanho_mensagens_wpp") = 10
%>

<html>


<link href="css/plugins/awesome-bootstrap-checkbox/awesome-bootstrap-checkbox.css" rel="stylesheet">

 

<!--#include file="estrutura_head.asp"-->
<link href="css/plugins/clockpicker/clockpicker.css" rel="stylesheet">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/ekko-lightbox/5.3.0/ekko-lightbox.css" />

<style>
    .md-skin .wrapper-content {
        padding: 15px 15px 15px 15px;
    }

    .desativado{
        display: none;
    }

    #mensagens-chat #mensagens{
        height: 446px !important;
    }

    #mensagens-internas .mensagem_enviada div div, #mensagens-internas .mensagem_enviada div p 
    {
        background: #dcf8c6 !important;
        border-bottom-right-radius: 0px !important;
        text-align: right;
    }

    #mensagens-internas{
        padding-left: 0px !important;
    }

    #mensagens-internas .mensagem_enviada div
    {
        margin-left: auto !important;
        width: 72% !important;
        margin-left: auto !important; 
    }

    #mensagens-internas .mensagem_recebida div div
    {
        border-bottom-left-radius: 0px !important;
        width: 50% !important;
    }

    #mensagens-internas .mensagem_recebida div p, #mensagens-internas .mensagem_enviada div p
    {
        word-wrap: break-word !important;
    }

    #mensagens-internasp .mensagem_enviada div div, #mensagens-internasp .mensagem_enviada div p 
    {
        background: #dcf8c6 !important;
        border-bottom-right-radius: 0px !important;
        text-align: right;
    }

    #mensagens-internasp{
        padding-left: 0px !important;
    }

    #mensagens-internasp .mensagem_enviada div
    {
        margin-left: auto !important;
        width: 72% !important;
        margin-left: auto !important; 
    }

    #mensagens-internasp .mensagem_recebida div div
    {
        border-bottom-left-radius: 0px !important;
        width: 50% !important;
    }

    #mensagens-internasp .mensagem_recebida div p, #mensagens-internasp .mensagem_enviada div p
    {
        word-wrap: break-word !important;
    }

    #submit_produtos {
        float: right;
        right: 15px;
    }

    .pace-active{
        display: none !important;
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

</style>

<%
    
    id_neg = request.querystring("id")

    IF request.form("act") = "editar_historico" THEN

        id_historico = request.form("idinteracao")

        descricao = Replace(Request.Form("descricao"), VbCrLf, "<br>")
        descricao = Replace(descricao, "'", "&quot;")
        descricao = Replace(descricao, "'", "&#34;")

        conexao.execute("UPDATE neg_interacao SET descricao = '"&descricao&"', id_tipo = "&request.form("id_tipo")&" WHERE idni =  "&id_historico)

        SET rs_historico = conexao.execute("SELECT id_negociacao FROM neg_interacao WHERE idni =  "&id_historico)

        acao_log = "Atualizou dados do histórico, interação "&id_historico
        tabela_log = "neg_interacao"
        obs_log = "Codigo da negociação alterada = "&rs_historico("id_negociacao")
        id_modulo_log = 9
        tipo_log = "update"
        Call registrar_log(acao_log, tabela_log, obs_log, id_modulo_log, tipo_log)

        response.redirect("neg_interacao_novo.asp?id="&id_neg&"&msg=edit-histo")

    END IF

    IF request.querystring("act") = "remover-histo" THEN

        id_historico = request.querystring("idhist")

        'conexao.execute("DELETE FROM neg_interacao WHERE idni = "&id_historico)
        conexao.execute("UPDATE neg_interacao SET del = 1 WHERE idni = "&id_historico)

        acao_log = "Deletou o histórico "&id_historico
        tabela_log = "neg_interacao"
        obs_log = "Codigo do histórico deletado = "&id_historico
        id_modulo_log = 9
        tipo_log = "delete"
        Call registrar_log(acao_log, tabela_log, obs_log, id_modulo_log, tipo_log)

        response.redirect("neg_interacao_novo.asp?id="&id_neg&"&msg=remove-histo")

    END IF

    IF isnull(id_neg) or id_neg = "" THEN 
        response.redirect("negociacao.asp")
    END IF

    SET rsemp = conexao.execute("SELECT np.*, nprod.produto, na.aquisicao, nfe.etapa, nfe.conversao, np.id_campanha FROM neg_prospecter np INNER JOIN neg_produtos nprod ON np.id_produto = nprod.idnp INNER JOIN neg_aquisicao na ON np.id_aquisicao = na.idna INNER JOIN neg_funil_etapas nfe ON np.situacao = nfe.idnfe WHERE np.idnp = "&id_neg)

    SET rsresponsavel = conexao.execute("SELECT p.id_usuario, u.usuario FROM neg_prospecter_responsaveis p INNER JOIN usuarios u ON u.idusuario = p.id_usuario WHERE p.id_negp = "&request.querystring("id"))
            
    if request.querystring("act") = "excluirp" then

        'conexao.execute("DELETE FROM neg_prospecter WHERE idnp = "&request.querystring("id"))
        conexao.execute("UPDATE neg_prospecter SET del = 1 WHERE idnp = "&request.querystring("id"))

        SET rs_negociacao = conexao.execute("SELECT empresa FROM neg_prospecter WHERE idnp = "&request.querystring("id"))

        acao_log = "Deletou a negociação "&rs_negociacao("empresa")
        tabela_log = "neg_prospecter"
        obs_log = "Codigo da negociação deletada = "&request.querystring("id")
        id_modulo_log = 9
        tipo_log = "delete"
        Call registrar_log(acao_log, tabela_log, obs_log, id_modulo_log, tipo_log)

        response.redirect("negociacao.asp?msg=excluir")

    end if
%>

<body>

    <div id="wrapper">

        <div id="status-envio-wpp" style="display: none;position:fixed;background: rgba(255,255,255,0.8);width: 100%;height: 100%;z-index: 9999999;">
            <div class="ibox-content sk-loading">
                <div class="sk-spinner sk-spinner-double-bounce" style="position:fixed !important; top: 50%;left: 50%;">
                    <div class="sk-double-bounce1"></div>
                    <div class="sk-double-bounce2"></div>
                </div>  
            </div>      
        </div>

    <nav class="navbar-default navbar-static-side" role="navigation">
        <!--#include file="include_left_side.asp"-->
    </nav>

        <div id="page-wrapper" class="gray-bg">
        <div class="row border-bottom">
            <!--#include file="include_topo.asp"-->
        </nav>
        </div>
            <div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-9">
                    <h2>Negócios</h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="default.asp">Home</a>
                        </li>
                        <li class="active">
                            <strong>Gerenciamento de Negócios</strong>
                        </li>
                    </ol>
                </div>
                <div class="col-lg-3" style="text-align: right;">
                    <h2></h2>
                    <%

                        SET rsvincular = conexao.execute("SELECT id_cliente FROM neg_prospecter WHERE del = 0 AND idnp = "&Request.querystring("id"))
                        IF rsvincular.EOF ThEN
                            response.redirect("negociacao.asp")
                        END IF

                        IF rsemp("conversao") = "Sim" AND rsvincular("id_cliente") = "0" THEN
                    %>
                            <a href="#" data-toggle="modal" data-target="#vincular_sistema" class="btn btn-success">Vincular ao Sistema</a>
                    <%
                        END IF
                    %>
                    <a href="negociacao.asp" class="btn btn-success">Voltar</a>
                </div>
            </div>

        <%
            IF Request.Form("MM_Insert") = "questionarios" THEN 
                SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo'")

                IF NOT rsquestionarios.EOF THEN 

                    WHILE NOT rsquestionarios.EOF

                        SET neg_campos = conexao.execute("SELECT * FROM neg_quest_campos where status = 'Ativo' and id_nq = "&rsquestionarios("idnq")&" order by idcampo desc")

                        while not neg_campos.EOF
                            campo = Request.Form("campo"&neg_campos("idcampo"))

                            ' Inserção Resposta questionario
                            conexao.execute("INSERT INTO neg_quest_respostas (id_prospecter, id_campo, resposta) VALUES ('"&Request.Form("idnp")&"', '"&neg_campos("idcampo")&"', '"&campo&"')")

                            neg_campos.movenext
                        WEND

                    rsquestionarios.movenext 
                    WEND
                END IF 

                Response.Redirect("neg_interacao_novo.asp?id="&Request.Form("idnp")&"&msg=questok")
                
            END IF 

            IF Request.Form("MM_Insert") = "formanotacao" THEN 

                dia = day(date)
                IF dia < 10 THEN dia = "0"&dia 
                mes = month(date)
                IF mes < 10 THEN mes = "0"&mes 
                agenda_data = year(date)&"/"&mes&"/"&dia

                agenda_hora_p = Split(now(), " ")
                agenda_hora = agenda_hora_p(1)

                id_usuario = Request.Form("id_usuario_")

                
                descricao = Replace(Request.Form("descricao"), VbCrLf, "<br>")
                descricao = Replace(descricao, "'", "&quot;")
                descricao = Replace(descricao, "'", "&#34;")

                conexao.execute("INSERT INTO neg_interacao (id_negociacao, id_usuario, descricao, agenda_data, agenda_hora, tipo, id_tipo) VALUES ("&id_neg&", "&Request.Cookies("grpro")("idusuario")&", '"&descricao&"', '"&agenda_data&"', '"&agenda_hora&"', 'Anotação', "&request.form("id_tipo")&")")

                Response.Redirect("neg_interacao_novo.asp?id="&id_neg&"&msg=histok")
            END IF 

            IF Request.Form("MM_Insert") = "formreuniao" THEN 
                dia = day(Request.Form("agenda_data"))
                IF dia < 10 THEN dia = "0"&dia 
                mes = month(Request.Form("agenda_data"))
                IF mes < 10 THEN mes = "0"&mes 
                agenda_data = year(Request.Form("agenda_data"))&"/"&mes&"/"&dia

                agenda_hora = Request.Form("agenda_hora") & ":00"

                descricao = Replace(Request.Form("descricao"), VbCrLf, "<br>")
                descricao = Replace(descricao, "'", "")

                tipo_reuniao = request.Form("tipo_reuniao")

                SET rsnome_responsavel = conexao.execute("SELECT npr.id_usuario, np.empresa, np.responsavel FROM neg_prospecter_responsaveis npr INNER JOIN neg_prospecter np ON np.idnp = npr.id_negp WHERE id_negp = "&id_neg)

                datas = agenda_data & " " & agenda_hora

                WHILE NOT rsnome_responsavel.EOF

                    conexao.execute("INSERT INTO neg_interacao (id_negociacao, id_usuario, descricao, agenda_data, agenda_hora, tipo) VALUES ("&id_neg&", "&rsnome_responsavel("id_usuario")&", '"&descricao&"', '"&agenda_data&"', '"&agenda_hora&"', '"&tipo_reuniao&"')")

                    conexao.execute("INSERT INTO agenda_compromissos (id_usuario, id_setor, id_np, titulo, descricao, data_inicio, data_termino, tipo) VALUES ( "&rsnome_responsavel("id_usuario")&", "&Request.Cookies("grpro")("setor")&", "&id_neg&", '"&tipo_reuniao&" com "&rsnome_responsavel("empresa")&" - "&rsnome_responsavel("responsavel")&"', '"&descricao&"', '"&datas&"', '"&datas&"', 'Pessoal')")

                    rsnome_responsavel.movenext
                WEND

                Response.Redirect("neg_interacao_novo.asp?id="&id_neg&"&msg=histok")
            END IF 

            IF Request.Form("MM_Insert") = "formfunil" THEN

                id_usuario_funil = Request.Form("id_usuario_funil")

                conexao.execute("INSERT INTO neg_funil (id_negp, id_etapa_funil, id_usuario) VALUES ("&id_neg&", "&Request.Form("funil_etapa")&", "&Request.Form("id_usuario_funil")&" )")

                conexao.execute("UPDATE neg_prospecter SET situacao = "&Request.Form("funil_etapa")&" WHERE idnp = "&id_neg)

                SET rsnomeetapa = conexao.execute("SELECT etapa FROM neg_funil_etapas WHERE idnfe = "&Request.Form("funil_etapa"))

                conexao.execute("INSERT INTO neg_interacao (id_negociacao, id_usuario, descricao, tipo) VALUES ("&id_neg&", "&Request.Cookies("grpro")("idusuario")&", 'Entrou no funil "&rsnomeetapa("etapa")&"', 'Funil')")

                Response.Redirect("neg_interacao_novo.asp?id="&id_neg&"&msg=funilok")
            END IF 

            IF Request.Form("MM_Update") = "negociacao" THEN 

                valor_estoque = Replace(Request.Form("valor_estoque"), ".", "")
                valor_estoque = Replace(valor_estoque, ",", ".")

                valor_investimento = Replace(Request.Form("valor_investimento"), ".", "")
                valor_investimento = Replace(valor_investimento, ",", ".")

                valor_adesao = Replace(Request.Form("valor_adesao"), ".", "")
                valor_adesao = Replace(valor_adesao, ",", ".")

                valor_recorrencia = Replace(Request.Form("valor_recorrencia"), ".", "")
                valor_recorrencia = Replace(valor_recorrencia, ",", ".")

                cidade = Request.Form("cidade")
                cidade = Replace(cidade, "'", "")

                quant_usuarios = split(request.Form("id_usuario"), ",")

                cpf = replace(request.Form("cpf_cnpj_P"), "'", "")

                SET rsatualizacao = conexao.execute("SELECT empresa, responsavel, email, telefone, cidade, uf, cpf FROM neg_prospecter WHERE idnp = "&id_neg)

                SET rs_usuario = conexao.execute("SELECT usuario FROM usuarios WHERE idusuario = "&Request.Cookies("grpro")("idusuario"))

                mudanca = "O colaborador(a) "&rs_usuario("usuario")
                mudanca_flag = false

                IF rsatualizacao("cpf") <> cpf THEN
                    mudanca_flag = true

                    IF rsatualizacao("cpf") = "" OR IsNull(rsatualizacao("cpf")) THEN
                        cpf_aux = "00.000.000/0000-00"
                    ELSE 
                        cpf_aux = rsatualizacao("cpf")
                    END IF 

                    mudanca = mudanca&"<br> Alterou o CPF / CNPJ "&cpf_aux&" para "&cpf
                END IF 

                IF rsatualizacao("empresa") <> Request.Form("empresa") THEN
                    mudanca_flag = true
                    mudanca = mudanca&"<br> Alterou o nome empresa "&rsatualizacao("empresa")&" para "&Request.Form("empresa")
                END IF 

                IF rsatualizacao("responsavel") <> Request.Form("responsavel") THEN
                    mudanca_flag = true
                    mudanca = mudanca&"<br> Alterou o nome do responsável "&rsatualizacao("responsavel")&" para "&Request.Form("responsavel")
                END IF 

                IF rsatualizacao("email") <> Request.Form("email") THEN
                    mudanca_flag = true
                    mudanca = mudanca&"<br> Alterou o email "&rsatualizacao("email")&" para "&Request.Form("email")
                END IF 

                IF rsatualizacao("telefone") <> Request.Form("telefone") THEN
                    mudanca_flag = true
                    mudanca = mudanca&"<br> Alterou o telefone "&rsatualizacao("telefone")&" para "&Request.Form("telefone")
                END IF 

                IF rsatualizacao("cidade") <> cidade THEN
                    mudanca_flag = true
                    mudanca = mudanca&"<br> Alterou o nome da cidade "&rsatualizacao("cidade")&" para "&cidade
                END IF 

                IF rsatualizacao("uf") <> Request.Form("uf") THEN
                    mudanca_flag = true
                    mudanca = mudanca&"<br> Alterou o uf "&rsatualizacao("uf")&" para "&Request.Form("uf")
                END IF 

                IF mudanca_flag = true THEN

                    dia = day(date)
                    IF dia < 10 THEN dia = "0"&dia 
                    mes = month(date)
                    IF mes < 10 THEN mes = "0"&mes 
                    agenda_data = year(date)&"/"&mes&"/"&dia

                    agenda_hora_p = Split(now(), " ")
                    agenda_hora = agenda_hora_p(1)

                    conexao.execute("INSERT INTO neg_interacao (id_negociacao, id_usuario, descricao, tipo, agenda_data, agenda_hora) VALUES ("&id_neg&", "&Request.Cookies("grpro")("idusuario")&", '"&mudanca&"', 'Anotação', '"&agenda_data&"', '"&agenda_hora&"')")

                END IF 

                conexao.execute("UPDATE neg_prospecter SET empresa = '"&Request.Form("empresa")&"', responsavel = '"&Request.Form("responsavel")&"', email = '"&Request.Form("email")&"', telefone = '"&Request.Form("telefone")&"', cidade = '"&cidade&"', uf = '"&Request.Form("uf")&"', valor_adesao = '"&valor_adesao&"', valor_recorrencia = '"&valor_recorrencia&"', id_produto = "&Request.Form("id_produto")&", id_aquisicao = "&Request.Form("id_aquisicao")&", habitantes = '"&Request.Form("habitantes")&"', area_vendas = '"&Request.Form("area_vendas")&"', montagem = '"&Request.Form("montagem")&"', valor_estoque = '"&valor_estoque&"', valor_investimento = '"&valor_investimento&"', ultimo_faturamento = '"&Request.Form("ultimo_faturamento")&"', zoneamento = '"&Request.Form("zoneamento")&"', id_perfil = '"&Request.Form("perfil")&"', cpf = '"&cpf&"' WHERE idnp = "&id_neg)
               
                conexao.execute("DELETE FROM neg_prospecter_responsaveis WHERE id_negp = "&request.querystring("id"))
                
                FOR i = 0 TO UBOUND(quant_usuarios)
                    conexao.execute("INSERT INTO neg_prospecter_responsaveis (id_negp, id_usuario) VALUES ("&request.querystring("id")&", "&quant_usuarios(i)&")")
                NEXT

                ' Inserindo respostas
                IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                    SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil INNER JOIN neg_questionarios ON id_questionario = idnq WHERE id_perfil = "&rsemp("id_perfil"))
                else

                    SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                end if

                IF NOT rsquestionarios.EOF THEN
                    aux = 1000
                    WHILE NOT rsquestionarios.EOF
                        set rs_upcampo = conexao.execute("SELECT * FROM neg_quest_campos where id_nq = "&rsquestionarios("idnq")&" and status <> 'inativo' order by ordem asc")

                        
                        while not rs_upcampo.EOF
                            set rsresposta = conexao.execute("SELECT * FROM neg_quest_respostas where id_campo = "&rs_upcampo("idcampo")&" and id_prospecter = "&id_neg)
                            campo = Request.Form("campo"&rs_upcampo("idcampo"))
                            if not rsresposta.eof then
                                idresposta = rsresposta("idresposta")
                                
                                ' Inserção Resposta questionario
                                conexao.execute("UPDATE neg_quest_respostas SET resposta = '"&campo&"' where idresposta = "&idresposta)
                            else
                                conexao.execute("INSERT INTO neg_quest_respostas (id_prospecter, id_campo, resposta) VALUES ('"&id_neg&"', '"&rs_upcampo("idcampo")&"', '"&campo&"')")
                            end if

                            aux = aux + 1
                            rs_upcampo.movenext
                        WEND

                    aux = aux + 1
                    rsquestionarios.movenext 
                    WEND 
                    rsquestionarios.movefirst
                END IF 

                Response.Redirect("neg_interacao_novo.asp?id="&id_neg&"&msg=editarok")
            END IF 

            IF Request.QueryString("act") = "marcar" THEN 
                idni = Request.QueryString("idni")

                SET rsverifica = conexao.execute("SELECT status FROM neg_interacao WHERE idni = "&idni)

                IF rsverifica("status") = "Aberto" THEN 
                    conexao.execute("UPDATE neg_interacao SET status = 'Concluida' WHERE idni = "&idni)
                ELSE 
                    conexao.execute("UPDATE neg_interacao SET status = 'Aberto' WHERE idni = "&idni)
                END IF 

                Response.Redirect("neg_interacao_novo.asp?id="&Request.QueryString("id"))
            END IF 

            IF Request.QueryString("act") = "marcar_c" THEN 
                idni = Request.QueryString("idni")

                SET rsverifica = conexao.execute("SELECT tipo, status, agenda_hora, agenda_data FROM neg_interacao WHERE idni = "&idni)

                WHILE NOT rsverifica.EOF

                    rshora = split(rsverifica("agenda_hora"), " ")

                    IF rsverifica("status") = "Aberto" AND rsverifica("tipo") = "Reunião" THEN 
                        conexao.execute("UPDATE neg_interacao SET status = 'Cancelada' WHERE id_negociacao = "&request.querystring("id")&" AND agenda_data = '"&transforma_data_ua(rsverifica("agenda_data"))&"' AND agenda_hora = '"&rshora(1)&"' ")
                    END IF 

                    rsverifica.movenext
                WEND

                Response.Redirect("neg_interacao_novo.asp?id="&Request.QueryString("id"))
            END IF 

            IF Request.QueryString("act") = "excluir" THEN 

                conexao.execute("UPDATE neg_documentos SET status = 'Inativo' WHERE idnd = "&Request.QueryString("idnd"))

                SET arquivo = conexao.execute("SELECT arquivo, id_np FROM neg_documentos where idnd = "&Request.QueryString("idnd"))

                varanexo = Server.MapPath("../imagens/negociacoes/"&arquivo("id_np")&"/"&arquivo("arquivo"))

                Set FSO = Server.CreateObject("Scripting.FileSystemObject")

                If Fso.FileExists(varanexo) Then

                    Set anexo = FSO.GetFile(varanexo)

                    anexo.delete

                end if

                Response.Redirect("neg_interacao_novo.asp?id="&Request.QueryString("id"))
            END IF 

            IF Request.Form("aprovar") = "sim" THEN  

                dia = day(date)
                IF dia < 10 THEN dia = "0"&dia 
                mes = month(date)
                IF mes < 10 THEN mes = "0"&mes 
                data_atual = year(date)&"/"&mes&"/"&dia

                observacao = Replace(Request.Form("observacao"), VbCrLf, " <br> ")
                
                conexao.execute("UPDATE neg_prospecter SET data_validacao = '"&data_atual&"', id_usuario_validacao = "&Request.Cookies("grpro")("idusuario")&", obs_validacao = '"&observacao&"' WHERE idnp = "&Request.Form("id"))

                'Processo'
                IF Request.Form("processo") <> "" THEN                 
                    SET rstarefas = conexao.execute("SELECT id_tarefa FROM processo_tarefas WHERE id_processo = "&Request.Form("processo"))

                    IF NOT rstarefas.EOF THEN 
                        tarefas = "("
                        WHILE NOT rstarefas.EOF 
                            tarefas = tarefas & rstarefas("id_tarefa") & ","
                        rstarefas.movenext 
                        WEND 
                        rstarefas.movefirst
                        total = Len(tarefas)
                        tarefas = Left(tarefas, total - 1)
                        tarefas = tarefas & ")"

                        idtarefa = tarefas 
                    END IF 

                    IF idtarefa <> "" OR (Len(tarefas) > 3) THEN 

                        dia = day(Request.Form("data_prevista"))
                        IF dia < 10 THEN dia = "0"&dia 
                        mes = month(Request.Form("data_prevista"))
                        IF mes < 10 THEN mes = "0"&mes 
                        data_prevista = year(Request.Form("data_prevista"))&"/"&mes&"/"&dia

                        descricao = Replace(Request.Form("descricao_p"), VbCrLf, "<br>")

                        tarefas = MID(tarefas, 2, total)
                        tarefas = Left(tarefas, total - 2)
                        separar = split(tarefas, ",")
                        for x = lbound(separar) to ubound(separar)

                            SET rstarefa = conexao.execute("SELECT * FROM tarefas WHERE idtarefa = "&separar(x))

                            SET rsverificaresp = conexao.execute("SELECT id_usuario FROM licenciado_equipe WHERE id_licenciado = "&Request.QueryString("id")&" and id_setor = "&rstarefa("id_setor"))
                            IF rsverificaresp.EOF THEN 
                                responsavel = 0
                            ELSE 
                                responsavel = rsverificaresp("id_usuario")
                            END IF 

                            conexao.execute("INSERT INTO tarefa_op (id_tarefa, id_usuario, id_criador, id_setor, id_licenciado, descricao, sequencial, data_prevista_termino, gut) VALUES ("&separar(x)&", "&responsavel&", "&Request.Cookies("grpro")("idusuario")&", "&rstarefa("id_setor")&", 1, '"&descricao&"', '"&rstarefa("sequencial")&"', '"&data_prevista&"', "&rstarefa("gut")&")")

                            SET rsidtarefa = conexao.execute("SELECT idtarefaop FROM tarefa_op WHERE id_tarefa = "&separar(x)&" ORDER BY idtarefaop DESC LIMIT 1")

                            IF rsverificaresp.EOF THEN
                                conexao.execute("UPDATE tarefa_op SET id_usuario = null WHERE idtarefaop = "&rsidtarefa("idtarefaop"))
                            END IF 

                            SET rsatividades = conexao.execute("SELECT * FROM tarefa_atividades WHERE id_tarefa = "&separar(x))

                            IF rstarefa("sequencial") = "Sim" THEN 
                                atv_status = "Aberta"
                                WHILE NOT rsatividades.EOF 
                                    conexao.execute("INSERT INTO atividade_op (id_tarefa, atividade, prazo, descricao, prioridade, status) VALUES ("&rsidtarefa("idtarefaop")&", '"&rsatividades("atividade")&"', "&rsatividades("prazo")&", '"&rsatividades("descricao")&"', '"&rsatividades("prioridade")&"', '"&atv_status&"')")
                                atv_status = "Aguardando"
                                rsatividades.movenext 
                                WEND 
                            ELSE 
                                WHILE NOT rsatividades.EOF 
                                    conexao.execute("INSERT INTO atividade_op (id_tarefa, atividade, prazo, descricao, prioridade, ordem, status) VALUES ("&rsidtarefa("idtarefaop")&", '"&rsatividades("atividade")&"', "&rsatividades("prazo")&", '"&rsatividades("descricao")&"', '"&rsatividades("prioridade")&"', "&rsatividades("ordem")&", 'Aberta')")
                                rsatividades.movenext 
                                WEND 
                            END IF 
                            rstarefa.movenext

                            IF request.Form("projeto") <> "" THEN 
                                conexao.execute("INSERT INTO projeto_tarefas (id_projeto, id_tarefaop) VALUES ("&Request.Form("projeto")&", "&rsidtarefa("idtarefaop")&")")
                            END IF 
                        next
                    END IF
                END IF 

                Response.Redirect("neg_interacao_novo.asp?id="&Request.Form("id"))
            END IF 


            
        %>

        <div class="wrapper wrapper-content animated fadeInRight ecommerce">

            <div class="row" style="margin-bottom: 15px;">

                <div id="div_1" class="col-lg-4" style="padding-left: 0;">
                    <div class="col-sm-12">
                        <div class="ibox-content m-b-sm border-bottom" style="min-height: 336px;background: transparent;padding: 0%; margin-bottom: 0px;">
                            <div class="row">

                                <div class="col-sm-12" style="padding: 5% 0%;background: #fff">
                                    <div class="col-sm-12" style="display: flex;margin-bottom: 6%">

                                        <div class="col-md-5" style="padding: 0px;display: flex;">
                                        
                                            <%
                                                SET rsprimeiraetapa = conexao.execute("SELECT idnfe FROM neg_funil_etapas WHERE entrada = 'Sim'")

                                                IF NOT IsNull(rsprimeiraetapa("idnfe")) THEN

                                                    SET rscadastro = conexao.execute("SELECT CAST(data AS DATE) as data FROM neg_funil WHERE id_negp = "&Request.QueryString("id")&" and id_etapa_funil = "&rsprimeiraetapa("idnfe")&" ORDER BY idnf ASC LIMIT 1")

                                                    IF NOT rscadastro.EOF THEN
                                                        dia = day(rscadastro("data"))
                                                        IF dia < 10 THEN dia = "0"&dia 
                                                        mes = month(rscadastro("data"))
                                                        IF mes < 10 THEN mes = "0"&mes 
                                                        data_cadastro = dia&"/"&mes&"/"&year(rscadastro("data"))
                                                    ELSE 
                                                        data_cadastro = "Não encontrado"
                                                    END IF 

                                                END IF
                                            %>

                                            <div class="tooltip-demo" >
                                                <label data-toggle="tooltip" data-placement="top" title="Data de Cadastro" class="label label-primary"><%=data_cadastro%></label>
                                            </div>

                                            <%
                                                SET rsinteracao = conexao.execute("SELECT CAST(data AS DATE) as dataa FROM neg_interacao WHERE id_negociacao = "&Request.QueryString("id")&" ORDER BY idni DESC LIMIT 1")

                                                IF NOT rsinteracao.EOF THEN
                                            %>

                                                    <div class="tooltip-demo" style="margin-left: 3%">
                                                        <label data-toggle="tooltip" data-placement="top" title="Data da Última Interação" class="label label-primary"><%=rsinteracao("dataa")%></label>   
                                                    </div>  

                                            <%
                                                END IF

                                                IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> 0 THEN
                                                    SET neg_perfil = conexao.execute("SELECT perfil FROM neg_perfil where idperfil = "&rsemp("id_perfil"))
                                            %>
                                                    <div class="tooltip-demo" style="margin-left: 3%">
                                                        <label style="background-color: #f7ac59 !important" data-toggle="tooltip" data-placement="top" title="Tipo de Perfil" class="label label-primary"><%=neg_perfil("perfil")%></label>   
                                                    </div>
                                            <%
                                                END IF
                                            %>
                                        </div>

                                        <div class="col-md-7" style="padding: 0px;text-align: right;">

                                            <%
                                                SET count_rsparticipantes = conexao.execute("SELECT COUNT(pr.idprospecter) as total FROM usuarios u  INNER JOIN neg_prospecter_responsaveis pr ON pr.id_usuario = u.idusuario WHERE pr.id_negp = "&request.querystring("id")&" AND u.status = 'Ativo' ")
                                                flag_part = false
                                                IF CLNG(count_rsparticipantes("total")) > 3 THEN 
                                                    flag_part = true
                                            %>

                                                    <div data-toggle="modal" data-target="#visualizar_responsaveis" title="Demais responsáveis" class="tooltip-demo" style="cursor: pointer; width: 30px; height: 30px; position: relative; display: block; float: right; margin-right: 8px; text-align: center;  background-color: #f5f5f5; border-radius: 20px;">
                                                        <span style="line-height: 30px;font-size: 10px;font-weight: 700;color: #cccccc;" >+ <%=CLNG(count_rsparticipantes("total")) - 2%></span>
                                                    </div>
                                            <%
                                                END IF
                                                'SET rsusuario = conexao.execute("SELECT u.idusuario, u.usuario, u.logo, u.status, s.setor FROM usuarios u INNER JOIN setores s ON u.id_setor = s.idsetor WHERE idusuario = "&rsnegoc("id_usuario"))

                                                SET rsusuario = conexao.execute("SELECT u.usuario, u.logo, s.setor FROM usuarios u INNER JOIN setores s ON u.id_setor = s.idsetor INNER JOIN neg_prospecter_responsaveis pr ON pr.id_usuario = u.idusuario WHERE pr.id_negp = "&request.querystring("id")&" AND u.status = 'Ativo' ")

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
                                                        <a class="tooltip-demo" href="usuarios.asp" style="cursor: auto; width: 30px; height: 30px; position: relative; display: block; float: right; margin-right: 8px;">
                                                            <img style="width: 100%; height: 100%;"  data-toggle="tooltip" data-placement="bottom" title="<%=rsusuario("usuario")%> (<%=rsusuario("setor")%>)" alt="image" class="img-circle" src="../imagens/logo_usuario/<%=logo_usuario%>">
                                                        </a>
                                            <%
                                                        contador = contador + 1
                                                        rsusuario.movenext
                                                    WEND
                                                        rsusuario.movefirst
                                                END IF
                                            %>

                                        </div>
                                    </div>
                                    <div class="col-sm-12">
                                        <h3><strong><%=rsemp("responsavel")%></strong></h3>
                                        <%
                                            nome_empresa = rsemp("empresa")
                                            IF len(nome_empresa) > 40 THEN 
                                                nome_empresa = pontinhos(nome_empresa,30)
                                            END IF
                                        %>
                                        <p style="margin-bottom: 3%"><i class="fas fa-building" style="margin-right: 2%;color: #23c6c8"></i> <%=nome_empresa%></p>

                                        <div class="col-sm-12" style="padding: 0px">
                                            <div class="form-group" style="margin-bottom: 0px" >
                                                <p style="margin-bottom: 3%"><i class="fad fa-phone-alt" style="margin-right: 2%;color: #23c6c8"></i> <%=rsemp("telefone")%></p>
                                            </div>
                                        </div>

                                        <%
                                            cidade = rsemp("cidade")&" - "&rsemp("uf")
                                            IF len(cidade) > 40 THEN 
                                                cidade = pontinhos(cidade,30)
                                            END IF
                                        %>

                                        <div class="col-sm-12" style="padding: 0px">
                                            <div class="form-group" style="margin-bottom: 3%">
                                                <p style="margin-bottom: 0px"><i style="margin-right: 2%;color: #23c6c8" class="fad fa-map-marker-alt"></i> <%=cidade%></p>
                                            </div>
                                        </div>

                                        <%
                                            email = rsemp("email")
                                            IF len(email) > 40 THEN 
                                                email = pontinhos(email,30)
                                            END IF
                                        %>

                                        <div class="col-sm-12" style="padding: 0px">
                                            <div class="form-group" style="margin-bottom: 0px">
                                                <p style="word-break: break-word;margin-bottom: 3%"><i style="margin-right: 2%;color: #23c6c8" class="fad fa-envelope"></i><%=email%></p>
                                            </div>
                                        </div>                                 
                                    </div>

                                    <div class="col-sm-12" style="float: initial; display: flex; justify-content: space-between; flex-wrap: wrap;border-top: 1px solid #f1f1f1;margin-top: 2%">

                                        <%
                                        
                                            IF Request.Cookies("grpro")("perm_53") = "1" THEN
                                        %>
                                                <div class="tooltip-demo" style="width: 100%;text-align: right;padding-top: 3%">

                                                    <script type="text/javascript">
                                                        function excluir(x)
                                                        {
                                                            if (confirm("Você realmente deseja exluir esse prospecter?"))
                                                            {
                                                                window.location.href = "?act=excluirp&id="+x
                                                            }
                                                        }
                                                    </script>

                                                    <span style="cursor: pointer;" data-toggle="modal" data-target="#editar_p" class="label label-success float-right">Editar <i class="fa fa-pencil" data-placement="top" title="Editar" aria-hidden="true"></i> </span>

                                                    <span style="cursor: pointer;margin-left: 5%" onclick="excluir(<%=request.querystring("id")%>)"  class="label label-danger float-right">Deletar <i class="fa fa-trash" data-placement="top"  aria-hidden="true"></i> </span> 
                                                </div>
                                        <%
                                            END IF
                                        %>

                                        <div class="col-sm-12" style="padding: 0px">
                                            <div class="form-group" style="margin-bottom: 0px">
                                                <label class="control-label" for="product_name" style="color: #1cc09f;font-weight: 800;margin-bottom: 0px">Interesse</label>
                                                <p style="margin-bottom: 0px"><a href="produto_detalhes.asp?id=<%=rsemp("id_produto")%>" target="_blank"><%=rsemp("produto")%></a></p>
                                            </div>
                                        </div>


                                        <div class="col-sm-12" style="padding: 0px; margin-top: 2%">
                                            <div class="form-group" style="margin-bottom: 0px">
                                                <label class="control-label" for="product_name" style="color: #1cc09f;font-weight: 800;margin-bottom: 0px">Aquisição</label>
                                                <%
                                                    dia = day(date)
                                                    IF dia < 10 THEN dia = "0"&dia 
                                                    mes = month(date)
                                                    IF mes < 10 THEN mes = "0"&mes 
                                                    data_cadastro = dia&"/"&mes&"/"&year(date)
                                                %>
                                                <p style="margin-bottom: 0px"><%=rsemp("aquisicao")%></p>
                                            </div>
                                        </div>

                                        <div class="col-sm-12" style="padding: 0px; margin-top: 2%;border-bottom: 1px solid #f1f1f1">
                                            <div class="form-group" style="margin-bottom: 0px">
                                                <label class="control-label" for="product_name" style="color: #1cc09f;font-weight: 800;margin-bottom: 0px">Campanha</label>
                                                <% 
                                                    
                                                    if rsemp("id_campanha") <> "" and rsemp("id_campanha") <> 0 then
                                                    SET campanha = conexao.execute("SELECT nome_campanha FROM neg_campanha WHERE idcampanha = "&rsemp("id_campanha"))
                                                %>
                                                    <p><%=campanha("nome_campanha")%></p>
                                                <%
                                                    Else
                                                %>
                                                    <p>Nenhuma</p>
                                                <%
                                                    End if
                                                %>
                                            </div>
                                        </div>

                                        <div class="col-sm-12" style="padding: 0px;margin-top: 3%">
                                            <div class="form-group" style="margin-bottom: 1%">
                                               
                                                <%
                                                    IF NOT IsNull(rsemp("valor_adesao")) THEN 
                                                        valor_adesao = FormatNumber(rsemp("valor_adesao"))
                                                    ELSE    
                                                        valor_adesao = "0,00"
                                                    END IF 
                                                %>
                                                <p style="margin-bottom: 0px">  <i class="fad fa-hand-holding-usd" style="color: #23c6c8"></i> <span class="control-label" for="product_name" style="font-weight: 700">Valor Adesão:</span> R$ <%=valor_adesao%></p>
                                            </div>
                                        </div>

                                        <div class="col-sm-12" style="padding: 0px; margin-top: 0%">
                                            <div class="form-group">
                                               
                                                <%
                                                    IF NOT IsNull(rsemp("valor_recorrencia")) THEN 
                                                        valor_recorrencia = FormatNumber(rsemp("valor_recorrencia"))
                                                    ELSE    
                                                        valor_recorrencia = "0,00"
                                                    END IF 
                                                %>
                                                <p><i style="color: #23c6c8" class="fad fa-envelope-open-dollar"></i> <span style="font-weight: 700" class="control-label" for="product_name">Valor de Recorrência:</span> R$ <%=valor_recorrencia%></p>
                                            </div>
                                        </div>
                                    </div>

                                    <%
                                        IF Request.Cookies("grpro")("modulo_wpp") = "Ativo" Then
                                    %>
                                            <div style="padding-left: 5%">
                                                <%
                                                        IF Request.Cookies("grpro")("perm_68") = "1" THEN
                                                            IF IsNull(rsemp("data_validacao")) THEN
                                                %>
                                                                <div style="margin: auto;" class="btn btn-warning" onClick="aprovar()">
                                                                    <%
                                                                        total = Len(Application("nome_cliente"))
                                                                        IF Application("nome_cliente") <> "" THEN
                                                                            nome_nome_cliente = Left(Application("nome_cliente"), total - 1)
                                                                        ELSE
                                                                            nome_nome_cliente = "Cliente"
                                                                        end if
                                                                    %>
                                                                    APROVAR <%=Ucase(nome_nome_cliente)%>
                                                                </div>
                                                <%
                                                            END IF
                                                        END IF
                                                
                                                        SET rsaprovado = conexao.execute("SELECT np.data_validacao, np.obs_validacao, u.usuario FROM neg_prospecter np INNER JOIN usuarios u ON np.id_usuario_validacao = u.idusuario WHERE np.del = 0 AND idnp = "&Request.QueryString("id"))

                                                        IF NOT rsaprovado.EOF THEN
                                                %>
                                                            <div class="col-sm-12" style="padding: 0px;margin: auto;">
                                                                <div class="form-group">
                                                                    <h4 style="margin-top: 3%;font-weight: 400">Cliente Aprovado em <%=rsaprovado("data_validacao")%> </h4>
                                                                    <div style="height: 100%; overflow: auto;">
                                                                        <p style="margin-bottom: 0;">por <%=rsaprovado("usuario")%></p>
                                                                        <p style="height: auto; overflow: auto;margin-top: 2%"><%=rsaprovado("obs_validacao")%></p>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                <%
                                                        END IF
                                                %>
                                            </div>
                                    <%
                                        END IF
                                    %>

                                </div>

                                <%
                                    IF Request.Cookies("grpro")("modulo_wpp") = "Ativo" Then
                                %>
                                        <div class="col-sm-12" style="padding: 0px;margin-top: 15px;background: #fff">
                                            <div  class="ibox-content m-b-sm border-bottom" style="margin-bottom: 0px;padding: 0px">
                                                <div class="row" style="padding: 0 20px;">
                                                    <div class="col-sm-12" style="background-color: #fff; padding-top: 15px; height: auto;">
                                                        <div class="form-group">
                                                            
                                                            <h4><i class="fal fa-users-class" style="color: #23c6c8"></i> <strong>Reuniões</strong></h4>
                                                            <div style="height: auto; overflow: auto; display: flex;">

                                                                <%
                                                                    SET historico_count = conexao.execute("SELECT COUNT(ni.idni) as total FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and hour(agenda_hora) < 24 and tipo = 'Reunião' and ni.status = 'Aberto' ORDER BY idni DESC")
                                                                %>

                                                                <p style="padding-right: 10px">Agendadas <span class="label label-info" style="background-color: #676a6c"><%=historico_count("total")%></span></p>

                                                                <%
                                                                    SET historico_count = conexao.execute("SELECT COUNT(ni.idni) as total FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and hour(agenda_hora) < 24 and tipo = 'Reunião' and ni.status = 'Concluida' ORDER BY idni DESC")
                                                                %>

                                                                <p style="padding-right: 10px">Concluidas <span class="label label-info"><%=historico_count("total")%></span></p>

                                                                <%
                                                                    SET historico_count = conexao.execute("SELECT COUNT(ni.idni) as total FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and hour(agenda_hora) < 24 and tipo = 'Reunião' and ni.status = 'Cancelada' ORDER BY idni DESC")
                                                                %>

                                                                <p>Canceladas <span style="background-color: #ed5565" class="label label-info"><%=historico_count("total")%></span></p>

                                                            </div>

                                                         
                                                            <div style="padding: 4px;" class="btn btn-warning" data-toggle="modal" data-target="#modagenda">Criar uma reunião</div>
                                                          

                                                        </div>
                                                    </div>
                                                    
                                                </div>
                                            </div>
                                        </div>
                                <%
                                    END IF
                                %>
                                
                            </div>
                        </div>
                    </div>
                </div>
                
                <%
                    SET rsdispositivos = conexao.execute("SELECT wi.id_pedbot, wi.nome FROM wpp_instancias wi INNER JOIN wpp_perm_user wp ON wp.id_instancia = wi.idinstancia WHERE wp.id_usuario = "&Request.Cookies("grpro")("idusuario")&" order by wi.idinstancia desc limit 1")

                    IF NOT rsdispositivos.EOF THEN
                        iddispositivo = rsdispositivos("id_pedbot")
                    else
                        iddispositivo = 0
                    END IF

                    IF Request.Cookies("grpro")("modulo_wpp") = "Ativo" and iddispositivo <> 0 Then
                %>
                            <div class="col-sm-8" style="padding: 0px;background: #fff;padding: 1%; height: 702px;">
                                
                                <div class="col-md-12" id="lista_instancias" style="padding: 0px;display: flex;">
                                    <div style="text-align: center;">
                                        <p style="text-align: center;margin: auto;padding: 2%">Buscando dispositivo <img style="width: 30px" src="https://painel.adaptweb.com.br/administrador/api_wpp/gif_wpp.gif"></p>
                                    </div>
                                </div>

                                <%
                                    numero = rsemp("telefone")
                                    numero = "55"&Mid(numero,2,2)&Mid(numero,7,4)&Mid(numero,12,4)
                                %>

                                <script type="text/javascript">
                                    ajaxGo({ url:"api_wpp/ajax_lista_instancia_neg.php?id=<%=Request.Cookies("grpro")("idusuario")%>&numero=<%=numero%>", elem_return: document.getElementById("lista_instancias") });
                                </script>

                                <div id="mensagens-chat" style="width: 100%;height: 252px;background: #f1f1f1">

                                </div>
                                <%
                                    
                                    SET rsdispositivos = conexao.execute("SELECT wi.id_pedbot FROM wpp_instancias wi INNER JOIN wpp_perm_user wp ON wp.id_instancia = wi.idinstancia WHERE wp.id_usuario = "&Request.Cookies("grpro")("idusuario")&" order by wi.idinstancia desc limit 1")

                                    IF NOT rsdispositivos.EOF THEN
                                        iddispositivo = rsdispositivos("id_pedbot")
                                    else
                                        iddispositivo = 0
                                    END IF

                                %>         

                                <script type="text/javascript">
                                    function enviar_mensagem(numero, idinstancia, nome, id_usuario)
                                    {

                                        if (document.getElementById("envio-msg").value != "")
                                        {
                                            document.getElementById("status-envio-wpp").style.display = "block";
                                            msg = document.getElementById("envio-msg").value;
                                            msg = msg.replace("\r\n", "\\n");

                                            enviar_mensagem_wpp(idinstancia, msg, numero, nome);

                                            document.getElementById("envio-msg").value = "";
                                            document.querySelector(".emojionearea-editor").innerHTML = "";
                                        }
                                    }
                                

                                    function enviar_doc(numero, instancia)
                                    {
                                        ajaxGo({ url:"Ajax/lista_galeria_envio.php?&instacia="+instancia+"&id=<%=Request.Cookies("grpro")("idusuario")%>&numero="+numero, elem_return: document.getElementById("galeria") });
                                        $("#anexo_wpp").modal("show");
                                    }
                                </script>

                                <div class="modal inmodal" id="anexo_wpp" tabindex="-1" role="dialog" aria-hidden="true">
                                    <div class="modal-dialog" style="width: 80% !important">
                                        <div id="div_mod" class="modal-content">

                                            <div class="modal-header" style="padding: 2%;border: 0;padding-bottom: 0px !important">

                                                <button onclick="habiliar_upload()" type="button" class="btn btn-success btn-xs" style="float: right;"><i onclick="" style="cursor: pointer" class="fas fa-plus" aria-hidden="true"></i> Novo Arquivo</button>

                                                <h3 style="text-align: left;">Seus Anexos:</h3>

                                               <div id="upload_new_wpp" class="desativado" style="padding: 2%;border: 1px solid #f1f1f1;border-radius: 0px;">

                                                    
                                                    <form method="post" enctype="multipart/form-data" action="upload_fotos_galeria_wpp_imediato.php" style="text-align: left;">
                                            
                                                    <div id="custom-queue"></div> 
                                                    <br>

                                                    <input type="file" name="upload[]" multiple="" required="" >
                                                    <input type="hidden" name="id" value="<%=Request.Cookies("grpro")("idusuario")%>">
                                                    
                                                    <button class="btn btn-w-m btn-success" onclick="att_galeria()" formtarget="_blank" type="submit" style="margin-top: 2%">Enviar</button>
                                            
                                                  </form>
                                                </div>

                                            </div>

                                            <div class="modal-body" id="galeria" style="background: #fff !important;padding: 0% 2% 2% 2%">
                                            </div>

                                            <div class="modal-footer">
                                                <button type="button" class="btn btn-white" data-dismiss="modal">Fechar</button>
                                            </div>

                                        </div>
                                    </div>
                                </div>

                                <script type="text/javascript">
                                    function habiliar_upload()
                                    {
                                        document.getElementById("upload_new_wpp").classList.toggle("desativado");
                                    }
                                </script>

                                <script type="text/javascript">
                                    
                                    function busca_conversa(numero, iddispo)
                                    {

                                        <%
                                            nome_empresa = replace(Replace(rsemp("empresa"),"""",""),"&", "")
                                        %>

                                        ajaxGo({ url:"Ajax/wpp_mensagens_neg.asp?nome=<%=nome_empresa%>&numero="+numero+"&dispo="+iddispo, elem_return: document.getElementById("mensagens-chat") });
                                       
                                        busca_msg(numero, iddispo);
                                       
                                    }



                                    function busca_msg(numero, iddispo)
                                    {

                                        ajaxGo({ url:"Ajax/wpp_mensagens_internas.asp?numero="+numero+"&dispo="+iddispo, elem_return: document.getElementById("mensagens-internas") });

                                        // console.log(numero);
                                        
                                       clearInterval(busca_wpp);

                                       busca_wpp = setInterval(function(){ 

                                            busca_msg(numero, iddispo);
                                        }, 4000);
                                    }

                                    busca_wpp = setInterval(function(){ busca_conversa('<%=numero%>', <%=iddispositivo%>) }, 100);
                                </script>

                            </div>
                <%
                    ELSE
                %>
                        <div class="col-sm-8" style="padding: 0px;">
                            <div  class="ibox-content m-b-sm border-bottom" style="height: auto">
                                <div class="row">
                                    <div class="col-sm-12" style="background-color: #fff; padding-top: 15px; height: auto;">
                                        <div class="form-group">
                                            <h4><strong>Área de Aprovação</strong></h4>   

                                            <%
                                                IF Request.Cookies("grpro")("perm_68") = "1" THEN
                                                    IF IsNull(rsemp("data_validacao")) THEN
                                            %>
                                                        <div class="btn btn-warning" onClick="aprovar()">
                                                            <%
                                                                total = Len(Application("nome_cliente"))
                                                                IF Application("nome_cliente") <> "" THEN
                                                                    nome_nome_cliente = Left(Application("nome_cliente"), total - 1)
                                                                ELSE
                                                                    nome_nome_cliente = "Cliente"
                                                                end if
                                                            %>
                                                            APROVAR <%=Ucase(nome_nome_cliente)%>
                                                        </div>
                                            <%
                                                    END IF
                                                END IF
                                            %>

                                            <%
                                                SET rsaprovado = conexao.execute("SELECT np.data_validacao, np.obs_validacao, u.usuario FROM neg_prospecter np INNER JOIN usuarios u ON np.id_usuario_validacao = u.idusuario WHERE np.del = 0 AND idnp = "&Request.QueryString("id"))

                                                IF NOT rsaprovado.EOF THEN
                                            %>
                                                    <div class="col-sm-12" style="padding: 0px">
                                                        <div class="form-group">
                                                            <h4 style="margin-top: 3%;font-weight: 400">Cliente Aprovado em <%=rsaprovado("data_validacao")%> </h4>
                                                            <div style="height: 100%; overflow: auto;">
                                                                <p style="margin-bottom: 0;">por <%=rsaprovado("usuario")%></p>
                                                                <p style="height: auto; overflow: auto;margin-top: 2%"><%=rsaprovado("obs_validacao")%></p>
                                                            </div>
                                                        </div>
                                                    </div>
                                            <%
                                                END IF
                                            %>

                                        </div>
                                    </div>
                                    <br>
                                </div>
                            </div> 
                        </div>

                        <div class="col-sm-8" style="padding: 0px;">
                            <div  class="ibox-content m-b-sm border-bottom">
                                <div class="row" style="padding: 0 20px;">
                                    <div class="col-sm-12" style="background-color: #fff; padding-top: 15px; height: auto;">
                                        <div class="form-group">
                                            <%
                                                SET rsagenda = conexao.execute("SELECT idni, agenda_data, status FROM neg_interacao WHERE id_negociacao = "&id_neg&" and status = 'Aberto' and tipo = 'Reunião' ORDER BY idni ASC LIMIT 1")
                                                IF NOT rsagenda.EOF THEN
                                            %>
                                                    <div class="pull-right btn btn-warning" style="margin-right: 15px;" onClick="marcar_neg(<%=rsagenda("idni")%>)">
                                                        Realizada
                                                    </div>
                                            <%
                                                END IF
                                            %>
                                            
                                            <h4><strong>Agendamento</strong></h4>
                                            <div style="height: auto; overflow: auto;">
                                                <%
                                                    SET rsagendamento = conexao.execute("SELECT agenda_data, descricao FROM neg_interacao WHERE id_negociacao = "&Request.QueryString("id")&" and status = 'Aberto' and tipo = 'Reunião' ORDER BY idni DESC LIMIT 1")

                                                    IF NOT rsagendamento.EOF THEN
                                                %>
                                                        <p><%=rsagendamento("agenda_data")%></p>
                                                        <p><%=rsagendamento("descricao")%></p>
                                                <%
                                                    ELSE 
                                                %>
                                                        <p>Nenhum encontrado.</p>
                                                <%
                                                    END IF
                                                %>
                                            </div>

                                            <% IF rsagendamento.EOF THEN %>
                                                <div style="width: 100%; margin-top: 20px; display: block;" class="btn btn-primary" data-toggle="modal" data-target="#modagenda">Agendar</div>
                                            <% END IF %>

                                        </div>
                                    </div>
                                    
                                </div>
                            </div>
                        </div>
                <%
                    END IF
                %>


            </div>

            <!-- Etapas -->
            <div class="row">
                <div class="col-lg-12" style="padding: 0px">
                    <div class="ibox-content m-b-sm border-bottom" style="overflow: auto; padding: 15px; margin-bottom: 15px;">

                        <div class="col-lg-12" style="text-align: right;">
                            <%
                                IF Request.Cookies("grpro")("perm_54") = "1" THEN
                            %>

                                    <div class="col-sm-12" style="margin-top: 0%; padding-right: 0px;">
                                        <div class="form-group" style="margin-bottom: 0px">
                                            <a class="btn btn-warning" data-toggle="modal" data-target="#mod_funil" style="cursor: pointer;">Mudar funil</a>
                                        </div>
                                    </div>

                            <%
                                ELSE
                            %>
                                    <div class="col-sm-12" style="margin-top: 0%">
                                        <div class="form-group" style="margin-bottom: 0px; padding-right: 0px;">
                                            <button disabled class="btn btn-warning" data-toggle="modal" data-target="#mod_funil" style="cursor: pointer;background: #7c7c80;border:#7c7c80">Mudar funil</button>
                                        </div>
                                    </div>
                            <%
                                END IF
                            %>
                        </div>

                        <div class="col-lg-12" style="padding-left: 0px;">
                            <%
                                SET rs_historico_funil = conexao.execute("SELECT nfe.processo, nf.id_etapa_funil, nf.data, nfe.etapa, u.usuario FROM neg_funil nf INNER JOIN neg_funil_etapas nfe ON nf.id_etapa_funil = nfe.idnfe INNER JOIN usuarios u ON nf.id_usuario = u.idusuario WHERE id_negp = "&id_neg&" ORDER BY nf.data ASC")

                                IF NOT rs_historico_funil.EOF THEN
                                    WHILE NOT rs_historico_funil.EOF

                                        dia = day(rs_historico_funil("data"))
                                        IF dia < 10 THEN dia = "0"&dia 
                                        mes = month(rs_historico_funil("data"))
                                        IF mes < 10 THEN mes = "0"&mes 
                                        data_historico_funil = dia&"/"&mes&"/"&year(rs_historico_funil("data"))

                                        hora_h_funil = Hour(rs_historico_funil("data"))
                                        IF hora_h_funil < 10 THEN hora_h_funil = "0"&hora_h_funil

                                        minuto_h_funil = Minute(rs_historico_funil("data"))
                                        IF minuto_h_funil < 10 THEN minuto_h_funil = "0"&minuto_h_funil

                                        hora_historico_funil = hora_h_funil& ":" &minuto_h_funil
                            %>
                                        <div class="col-lg-4" style="padding-left: 0px;">
                                            <div class="col-lg-1">
                                                <div class="vertical-timeline-icon navy-bg">
                                                    <i class="fa fa-arrows"></i>
                                                </div>
                                            </div>

                                            <div class="col-lg-11" style="margin-left: 40px;">
                                                <h3 style="margin-top: 0; margin-bottom: 2px;"><%=rs_historico_funil("etapa")%></h3>
                                                <p style="font-size: 11px; margin-bottom: 2px;"><%=data_historico_funil%> às <%=hora_historico_funil%></p>
                                                <p style="margin-bottom: 2px;" ><%=rs_historico_funil("usuario")%></p>
                                            </div>
                                            
                                        </div>
                            <%
                                        rs_historico_funil.movenext 
                                    WEND
                                END IF 
                            %>
                        </div>
                        
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-lg-12" style="padding: 0px; display: flex; margin-bottom: 40px;">
                    
                    <div class="ibox-content m-b-sm border-bottom col-lg-9" style="min-height: 545px; margin-right: 15px;">
                        <div class="col-lg-12" style="display: flex; padding: 0px;">
                            <div class="col-lg-8">
                                <h3>INTERAÇÃO E INFORMAÇÃO</h3>
                            </div>

                            <%

                                SET etapa_conversao = conexao.execute("SELECT nf.data, nfe.conversao FROM neg_funil nf INNER JOIN neg_funil_etapas nfe ON nf.id_etapa_funil = nfe.idnfe WHERE nfe.status = 'Ativo' AND id_negp = "&request.querystring("id")&" ORDER BY nf.data DESC limit 1")

                                'response.write("SELECT nf.data, nfe.conversao FROM neg_funil nf INNER JOIN neg_funil_etapas nfe ON nf.id_etapa_funil = nfe.idnfe WHERE nfe.status = 'Ativo' AND id_negp = "&request.querystring("id")&" ORDER BY nf.data DESC limit 1")

                                SET data_follow = conexao.execute("SELECT data FROM neg_interacao WHERE id_negociacao = "&request.querystring("id")&" ORDER BY idni ASC LIMIT 1")

                                'response.write("SELECT data FROM neg_interacao WHERE id_negociacao = "&request.querystring("id")&" ORDER BY idni ASC LIMIT 1")

                                IF etapa_conversao("conversao") <> "Sim" AND NOT data_follow.EOF THEN

                                    ' verificando quantidade de dias que está em negociação
                                    quat_dias = datediff("y", data_follow("data"), now())

                                    texto = "Tempo em negociação ("&quat_dias&") dias"

                                ELSEIF etapa_conversao("conversao") <> "Nao" AND NOT data_follow.EOF THEN

                                    ' verificando quantidade de dias que está em etapa de conversao
                                    quat_dias = datediff("y", data_follow("data"), etapa_conversao("data"))
                                
                                    texto = "Ciclo de venda ("&quat_dias&") dias"

                                END IF
                            
                                IF quat_dias > 0 THEN
                            %>
                                    <div class="col-lg-4" style="padding: 0px; text-align: right;">
                                        <button style="padding: 2px 12px;" class="btn btn-warning"><%=texto%></button>
                                    </div>
                            <%
                                END IF
                            %>
                            
                        </div>
                        
                        <br>

                        <div class="tabs-container">
                            <ul class="nav nav-tabs" role="tablist">

                                <%
                                    SET historico_count = conexao.execute("SELECT COUNT(ni.idni) as total FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and (hour(agenda_hora) < 24 or agenda_hora is null) AND ni.del = 0 and tipo = 'Anotação' ORDER BY idni DESC")
                                %>

                                <li class="active"><a class="nav-link" data-toggle="tab" href="#tabhi">Histórico 
                                    <span class="label label-info" style="background-color: #1cc09f"><%=historico_count("total")%></span>
                                </a></li>

                                <%
                                    SET historico_reuniao = conexao.execute("SELECT COUNT(ni.idni) as total FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and hour(agenda_hora) < 24 AND ni.del = 0 AND tipo = 'Reunião' ORDER BY idni DESC")

                                    SET historico_lembrete = conexao.execute("SELECT COUNT(ni.idni) as total FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and hour(agenda_hora) < 24 AND ni.del = 0 AND tipo = 'Lembrete' ORDER BY idni DESC ")
                                %>

                                <li>
                                    <a class="nav-link" data-toggle="tab" href="#tabreuni">Reuniões / Lembretes
                                        <!-- <span class="label label-info" style="background-color: #013366"><=historico_reuniao("total")%></span>
                                        <span class="label label-info" style="background-color: #f8ac59"><=historico_lembrete("total")%></span> -->
                                    </a>
                                </li>

                                <li class=""><a class="nav-link" data-toggle="tab" href="#tabarq">Arquivos</a></li>
                                

                                <%
                                    IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then
                                        SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
                                    else
                                        SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                                    end if

                                    IF NOT rsquestionarios.EOF THEN
                                %>
                                        <li class=""><a class="nav-link" data-toggle="tab" href="#tabquest">Questionário</a></li>
                                <%
                                    END IF
                                
                                    SET metricas = conexao.execute("SELECT * FROM neg_prospecter_metrica where id_negp = "&request.querystring("id"))
                                    IF NOT metricas.EOF THEN
                                %>
                                        <li class=""><a class="nav-link" data-toggle="tab" href="#tabutm">Métricas UTM</a></li>
                                <%
                                    END IF
                                %>

                            </ul>
                            <div class="tab-content">
                                
                                <div role="tabpanel" id="tabutm" class="tab-pane">
                                    <div class="panel-body">

                                       <table class="table">
                                           <thead>
                                               <tr>
                                                   <th>Variável</th>
                                                   <th>Valor</th>
                                               </tr>
                                           </thead>

                                           <tbody>
                                                <%
                                                WHILE NOT metricas.EOF
                                                %>

                                                    <tr>
                                                        <td><%=metricas("campo")%></td>
                                                        <td><%=metricas("valor")%></td>
                                                    </tr>
                                                <%
                                                    metricas.movenext
                                                WEND
                                               %>
                                           </tbody>
                                       </table>

                                    </div>
                                </div>


                                <div role="tabpanel" id="tabhi" class="tab-pane active">
                                    <div class="panel-body">

                                        <div class="col-lg-12" style="text-align: center;padding: 2%">
                                            <div class="vertical-timeline-icon navy-bg" style="position: initial;margin: auto;cursor: pointer;background-color: #1cc09f">
                                                <i data-toggle="modal" data-target="#modalanota" class="fa fa-plus" aria-hidden="true"></i>
                                            </div>
                                        </div>

                                        <%

                                            SET rshistorico = conexao.execute("SELECT nit.nome, ni.idni, ni.descricao, ni.agenda_data, ni.agenda_hora, ni.data, ni.status, ni.tipo, u.usuario, u.logo, ni.id_tipo FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and (hour(agenda_hora) < 24 or agenda_hora is null) AND ni.del = 0 and tipo = 'Anotação' ORDER BY idni DESC")

                                           

                                            IF NOT rshistorico.EOF THEN 
                                                WHILE NOT rshistorico.EOF

                                                dia = day(rshistorico("data"))
                                                IF dia < 10 THEN dia = "0"&dia 
                                                mes = month(rshistorico("data"))
                                                IF mes < 10 THEN mes = "0"&mes 
                                                data_historico = dia&"/"&mes&"/"&year(rshistorico("data"))

                                                hora_historico = Hour(rshistorico("data"))& ":" &Minute(rshistorico("data"))

                                                dia = day(rshistorico("agenda_data"))
                                                IF dia < 10 THEN dia = "0"&dia 
                                                mes = month(rshistorico("agenda_data"))
                                                IF mes < 10 THEN mes = "0"&mes 
                                                data_agenda = dia&"/"&mes&"/"&year(rshistorico("agenda_data"))

                                                IF IsNull(rshistorico("agenda_hora")) THEN
                                                    hora_h_funil = 0
                                                    minuto_h_funil = 0
                                                ELSE
                                                
                                                    hora_h_funil = Hour(rshistorico("agenda_hora"))
                                                    IF hora_h_funil < 10 THEN hora_h_funil = "0"&hora_h_funil


                                                    minuto_h_funil = Minute(rshistorico("agenda_hora"))
                                                    IF minuto_h_funil < 10 THEN minuto_h_funil = "0"&minuto_h_funil
                                                END IF
                                                

                                                hora_agenda = hora_h_funil& ":" &minuto_h_funil
                                        %>
                                                    <div class="timeline-item">
                                                        <div class="row">
                                                            <div class="col-lg-3 date">
                                                                <i class="fa fa-circle" style="color: #1cc09f" aria-hidden="true"></i>

                                                                <%
                                                                    if isnull(rshistorico("logo")) then 
                                                                        foto = "semfoto.png"
                                                                    else
                                                                        foto = rshistorico("logo") 
                                                                    end if
                                                                %>
                                                                <div style="text-align: center;">
                                                                    <img style="width: 40px;height: 40px;border-radius: 50%;object-fit: cover" alt="image" src="../imagens/logo_usuario/<%=foto%>">

                                                                    <p style="font-size: 10px;margin-top: 2%"><%=pontinhos(rshistorico("usuario"),12)%></p>
                                                                </div>
                                                                
                                                            </div>

                                                            <% 
                                                                descricao = rshistorico("descricao") 
                                                                idhistorico = rshistorico("idni")
                                                            %>

                                                            <div class="col-lg-10 content">
                                                                <p class="m-b-xs"><strong>Histórico</strong></p>
                                                                <p style="word-break: break-all;"><%=descricao%></p>

                                                                <!-- icone para editar e excluir historico -->

                                                                
                                                                <label class="label label-primary pull-left" style="margin-right: 8px;"><%=rshistorico("nome")%></label>
                                                                

                                                                <label class="label label-warning pull-left" style="position: relative; right: 0px;"><%=data_agenda%> &nbsp; <%=hora_agenda%></label>


                                                                <div style="text-align: right;">
                                                                    <span style="cursor: pointer;" onclick="editar_historico(<%=idhistorico%>, '<%=replace(descricao,"""","")%>', <%=rshistorico("id_tipo")%> )" data-toggle="modal" data-target="#modeditarHistorico" class="label label-success float-right">Editar <i class="fa fa-pencil" data-placement="top" title="Editar" aria-hidden="true"></i> </span>

                                                                    <span style="cursor: pointer;margin-left: 1%;" onclick="deleta_historico(<%=idhistorico%>)"  class="label label-danger float-right">Deletar <i class="fa fa-trash" data-placement="top"  aria-hidden="true"></i> </span>
                                                                </div>

                                                            </div>
                                                        </div>
                                                    </div>

                                        <%
                                                rshistorico.movenext 
                                                WEND 
                                            ELSE
                                        %>
                                                Não há registros.
                                        <%
                                            END IF 
                                        %>
                                    </div>
                                </div>

                                <div role="tabpanel" id="tabreuni" class="tab-pane">
                                    <div class="panel-body">
                                 
                                        <div class="col-lg-12" style="text-align: center;padding: 2%">
                                            <div class="vertical-timeline-icon navy-bg" style="position: initial;margin: auto;cursor: pointer;background-color: #f8ac59">
                                                <i data-toggle="modal" data-target="#modagenda" class="fa fa-plus" aria-hidden="true"></i>
                                            </div>
                                        </div>

                                        <%
                                            SET rshistorico = conexao.execute("SELECT nit.nome, ni.idni, ni.descricao, ni.agenda_data, ni.agenda_hora, ni.data, ni.status, ni.tipo, u.usuario, u.logo, ni.id_tipo FROM neg_interacao ni INNER JOIN usuarios u ON ni.id_usuario = u.idusuario INNER JOIN neg_interacao_tipo nit ON nit.id = ni.id_tipo WHERE id_negociacao = "&id_neg&" and hour(agenda_hora) < 24 AND ni.del = 0 AND (ni.tipo = 'Reunião' OR ni.tipo = 'Lembrete') ORDER BY ni.idni DESC")

                                            IF NOT rshistorico.EOF THEN 
                                                WHILE NOT rshistorico.EOF

                                                dia = day(rshistorico("data"))
                                                IF dia < 10 THEN dia = "0"&dia 
                                                mes = month(rshistorico("data"))
                                                IF mes < 10 THEN mes = "0"&mes 
                                                data_historico = dia&"/"&mes&"/"&year(rshistorico("data"))

                                                hora_historico = Hour(rshistorico("data"))& ":" &Minute(rshistorico("data"))

                                                dia = day(rshistorico("agenda_data"))
                                                IF dia < 10 THEN dia = "0"&dia 
                                                mes = month(rshistorico("agenda_data"))
                                                IF mes < 10 THEN mes = "0"&mes 
                                                data_agenda = dia&"/"&mes&"/"&year(rshistorico("agenda_data"))

                                                IF IsNull(rshistorico("agenda_hora")) THEN
                                                    hora_h_funil = 0
                                                    minuto_h_funil = 0
                                                ELSE
                                                
                                                    hora_h_funil = Hour(rshistorico("agenda_hora"))
                                                    IF hora_h_funil < 10 THEN hora_h_funil = "0"&hora_h_funil


                                                    minuto_h_funil = Minute(rshistorico("agenda_hora"))
                                                    IF minuto_h_funil < 10 THEN minuto_h_funil = "0"&minuto_h_funil
                                                END IF
                                                

                                                hora_agenda = hora_h_funil& ":" &minuto_h_funil
                                        %>

                                                    <%
                                                        IF rshistorico("status") = "Cancelada" THEN
                                                            cor_reuniao = "#ed5565"
                                                        END IF

                                                        IF rshistorico("status") = "Concluida" THEN
                                                            cor_reuniao = "#1cc09f"
                                                        END IF

                                                        IF rshistorico("status") = "Aberto" THEN
                                                            cor_reuniao = "#676a6c"
                                                        END IF
                                                    %>

                                                    <div class="timeline-item">
                                                        <div class="row">
                                                            <div class="col-lg-3 date">
                                                                <i class="fa fa-circle" style="color: <%=cor_reuniao%>" aria-hidden="true"></i>

                                                                <%
                                                                    if isnull(rshistorico("logo")) then 
                                                                        foto = "semfoto.png"
                                                                    else
                                                                        foto = rshistorico("logo") 
                                                                    end if
                                                                %>
                                                                <div style="text-align: center;">
                                                                    <img style="width: 40px;height: 40px;border-radius: 50%;object-fit: cover" alt="image" src="../imagens/logo_usuario/<%=foto%>">

                                                                    <p style="font-size: 10px;margin-top: 2%"><%=pontinhos(rshistorico("usuario"),12)%></p>
                                                                </div>
                                                                
                                                            </div>

                                                            <% 
                                                                descricao = rshistorico("descricao") 
                                                                idhistorico = rshistorico("idni")
                                                            %>

                                                            <div class="col-lg-10 content">
                                                                <div style="display: flex;">
                                                                    <p class="m-b-xs" style="margin-right: 20px;"><strong>Histórico</strong></p>
                                                                    

                                                                    <!-- icone para editar e excluir historico -->

                                                                    <%
                                                                        IF rshistorico("tipo") = "Reunião" THEN
                                                                            rstipo = "Reunião"
                                                                            cor = "#013366"
                                                                        ELSE
                                                                            rstipo = "Lembrete"
                                                                            cor = "#f8ac59"
                                                                        END IF
                                                                    %>

                                                                    <label class="label label-primary pull-left" style="margin-right: 8px; background-color: <%=cor%>; "><%=rstipo%></label>

                                                                    <label class="label label-primary pull-left" style="margin-right: 8px;background-color: <%=cor_reuniao%>"><%=rshistorico("status")%></label>
                                                                    

                                                                    <label class="label label-warning pull-left" style="position: relative; right: 0px;"><%=data_agenda%> &nbsp; <%=hora_agenda%></label>
                                                                </div>

                                                                
                                                                <p style="word-break: break-all;"><%=descricao%></p>


                                                                <div style="text-align: right; margin-bottom: 10px;">

                                                                    
                                                                    <%IF rshistorico("status") = "Aberto" THEN%>

                                                                        <%
                                                                            dia = day(rshistorico("agenda_data"))
                                                                            IF dia < 10 THEN dia = "0"&dia 
                                                                            mes = month(rshistorico("agenda_data"))
                                                                            IF mes < 10 THEN mes = "0"&mes 
                                                                            data_inicio_formatada = year(rshistorico("agenda_data"))&mes&dia

                                                                            hora = hour(rshistorico("agenda_hora")) + 3
                                                                            if hora < 10 THEN hora = 0&hora
                                                                            minuto = minute(rshistorico("agenda_hora"))
                                                                            if minuto < 10 THEN minuto = 0&minuto
                                                                            hora_inicio = hora & minuto

                                                                            data_termino_formatada = data_inicio_formatada

                                                                            hora = hora + 1
                                                                            IF hora >= 24 THEN 
                                                                                hora = 3
                                                                                data_termino_semformato = year(rshistorico("agenda_data"))&"-"&mes&"-"&dia
                                                                                data_termino_formatada = DateAdd("d",1,data_termino_semformato)
                                                                                data_termino_formatada = year(data_termino_formatada)&month(data_termino_formatada)&day(data_termino_formatada)

                                                                          
                                                                            END IF
                                                                            hora_termino = hora & minuto

                                                                            ' descricao = "Reunião com "&rsemp("responsavel")&" sobre o produto "&rsemp("produto")&" - "&rsemp("telefone")

                                                                        %>

                                                                        <a target="_blank" href="https://calendar.google.com/calendar/r/eventedit?dates=<%=data_inicio_formatada%>T<%=hora_inicio%>00Z%2F<%=data_termino_formatada%>T<%=hora_termino%>00Z&details=<%=Replace(descricao, " ", "+")%>&text=<%=Replace(rsemp("responsavel"), " ", "+")%>&trp=true&sf=true"><span style="cursor: pointer;margin-right: 1%;background-color: #1cc09f !important"  class="label label-success float-right">Google agenda <i class="fab fa-google-plus-g" data-placement="top" ></i> </span></a>

                                                                        <span style="cursor: pointer;margin-right: 1%;background-color: #1cc09f !important" onClick="marcar_neg(<%=idhistorico%>)" class="label label-success float-right">Concluir <i class="fa fa-check" data-placement="top" ></i> </span>

                                                                        <span style="cursor: pointer;margin-right: 1%;" onClick="marcar_neg_c(<%=idhistorico%>)" class="label label-danger float-right">Cancelar <i class="fa fa-times" data-placement="top" ></i> </span>
                                                                    <%END IF%>

                                                                    <span style="cursor: pointer;" onclick="editar_historico(<%=idhistorico%>, '<%=descricao%>', <%=rshistorico("id_tipo")%> )" data-toggle="modal" data-target="#modeditarHistorico" class="label label-success float-right">Editar <i class="fa fa-pencil" data-placement="top" title="Editar" aria-hidden="true"></i> </span>

                                                                    <span style="cursor: pointer;margin-left: 1%;" onclick="deleta_historico(<%=idhistorico%>)"  class="label label-danger float-right">Deletar <i class="fa fa-trash" data-placement="top"  aria-hidden="true"></i> </span>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>

                                        <%
                                                rshistorico.movenext 
                                                WEND 
                                            ELSE
                                        %>
                                                Não há registros.
                                        <%
                                            END IF 
                                        %>
                                    </div>
                                </div>

                                <style>
                                    .gallery {
                                    -webkit-column-count: 3;
                                    -moz-column-count: 3;
                                    column-count: 3;
                                    -webkit-column-width: 33%;
                                    -moz-column-width: 33%;
                                    column-width: 33%; }
                                    .gallery .pics {
                                    -webkit-transition: all 350ms ease;
                                    transition: all 350ms ease; }
                                    .gallery .animation {
                                    -webkit-transform: scale(1);
                                    -ms-transform: scale(1);
                                    transform: scale(1); }

                                    @media (max-width: 450px) {
                                    .gallery {
                                    -webkit-column-count: 1;
                                    -moz-column-count: 1;
                                    column-count: 1;
                                    -webkit-column-width: 100%;
                                    -moz-column-width: 100%;
                                    column-width: 100%;
                                    }
                                    }

                                    @media (max-width: 400px) {
                                    .btn.filter {
                                    padding-left: 1.1rem;
                                    padding-right: 1.1rem;
                                    }
                                    }
                                </style>
                                
                                <div id="tabarq" class="tab-pane">
                                    <div class="panel-body" style="padding-top: 70px;">
                                        <%
                                            data = year(date())&"-"&month(date())&"-"&day(date())
                                        %>

                                        <%
                                            SET verificaarq = conexao.execute("SELECT titulo FROM neg_documentos where status = 'Ativo' and id_np = "&request.querystring("id")&" limit 1")
                                            IF NOT verificaarq.EOF THEN
                                        %>
                                            <div style="position: absolute; top: 15px !important; right: 270px; padding: 0;">
                                                <div class="btn btn-success" onclick="doc_upload()"><i class="fa fa-upload" style="margin-right: 4%"></i>Inserir Arquivos</div>
                                            </div>

                                            <div style="position: absolute; top: 15px !important; right: 15px; padding: 0;">
                                                <div class="btn btn-success" onclick="doc_download_compact(<%=request.querystring("id")%>)"><i class="fa fa-download" style="margin-right: 4%"></i>Baixar Arquivos Compactados</div>
                                            </div>

                                            <script type="text/javascript">

                                                function doc_download_compact(id)
                                                {
                                                    ajaxGo({ url:"compactar_arquivos_neg.php?id="+id, elem_return: document.getElementById("btn-compact") });
                                                    $("#modcompact").modal("show");
                                                }

                                            </script>

                                        <%
                                            ELSE
                                        %>
                                            <div style="position: absolute; top: 15px !important; right: 30px; padding: 0;">
                                                <div class="btn btn-success" onclick="doc_upload()"><i class="fa fa-upload" style="margin-right: 4%"></i>Inserir Arquivos</div>
                                            </div>
                                        <%
                                            END IF
                                        %>

                                        <div class="row">
                                            <div class="col-lg-12" style="display: flex; justify-content: space-between">

                                                <div class="ibox " style="width: 100%;">
                                                    <div class="ibox-title">
                                                        <h5>Lista de Arquivos</h5>
                                                        <div class="ibox-tools">
                                                            <a class="collapse-link">
                                                                <i class="fa fa-chevron-up"></i>
                                                            </a>
                                                        </div>
                                                    </div>
                                                    <div class="ibox-content">
                                                        <div class="table-responsive">
                                                            <table class="table table-striped">
                                                                <%
                                                                    SET rsarq = conexao.execute("SELECT nd.*, nc.categoria FROM neg_documentos nd inner join neg_doc_categoria nc on nd.id_categoria = nc.idndc WHERE id_np = "&Request.QueryString("id")&" and nd.status = 'Ativo' ")

                                                                    IF NOT rsarq.EOF THEN 
                                                                %>
                                                                <thead>
                                                                    <tr>
                                                                        <th></th>

                                                                        <th>Título </th>
                                                                        <th>Arquivo</th>
                                                                        <th>Categoria</th>
                                                                        <th class="text-center">Ações</th>
                                                                    </tr>
                                                                </thead>
                                                                <tbody>

                                                                    <%
                                                                            i = 1
                                                                            WHILE NOT rsarq.EOF
                                                                    %>
                                                                    <tr>

                                                                        <%
                                                                            IF rsarq("tipo") = "jpg" or rsarq("tipo") = "jpeg" or rsarq("tipo") = "png" then
                                                                        %>
                                                                            <td><i class="fas fa-images"></i></td>
                                                                        <%
                                                                            Else
                                                                        %>
                                                                            <td><i class="far fa-file-alt"></i></td>
                                                                        <%
                                                                            end if
                                                                        %>

                                                                        

                                                                        <td>
                                                                            <% IF IsNull(rsarq("titulo")) THEN %>
                                                                                <span style="color: #ed5565;">Defina um título para esse arquivo...</span>
                                                                            <% ELSE %>
                                                                                <%=rsarq("titulo")%>
                                                                            <% END IF %>
                                                                        </td>
                                                                        <td><%=rsarq("arquivo")%></td>

                                                                        <td>
                                                                            <%=rsarq("categoria")%>
                                                                        </td>

                                                                        <td style="width: 15px;">
                                                                            <div style="display: flex; justify-content: space-between;">
                                                                                <div class="tooltip-demo"> 
                                                                                    <i class="fa fa-eye" data-toggle="tooltip" data-placement="top" title="Visualizar arquivo" style="cursor: pointer;" onClick="ver_arq('<%=rsarq("arquivo")%>', <%=Request.QueryString("id")%>, '<%=rsarq("tipo")%>', <%=rsarq("idnd")%>)"></i>
                                                                                </div>

                                                                                <div class="tooltip-demo"> 
                                                                                    <i class="fa fa-minus-circle" data-toggle="tooltip" data-placement="top" title="Deletar arquivo" style="cursor: pointer; color: red;" onClick="deletar_arq(<%=rsarq("idnd")%>, <%=Request.QueryString("id")%>)"></i>
                                                                                </div>
                                                                            </div>
                                                                        </td>
                                                                    </tr>
                                                                    <%
                                                                            i = i + 1
                                                                            rsarq.movenext
                                                                            WEND
                                                                        ELSE
                                                                    %>
                                                                        Nenhum arquivo foi encontrado.
                                                                    <%
                                                                        END IF 
                                                                    %>



                                                                </tbody>
                                                            </table>
                                                        </div>

                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div id="tabquest" class="tab-pane">
                                    <div class="panel-body" >
                                        <div class="row">
                                            <div class="col-lg-12">
                                                

                                                <h3>Respostas Questionário</h3>
                                                

                                            
                                                <div class="tabs-container">
                                                    <ul class="nav nav-tabs" role="tablist">
                                                        <%

                                                        IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                                                            SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
                                                        else

                                                            SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                                                        end if

                                                                IF NOT rsquestionarios.EOF THEN
                                                                    aa = 1
                                                                    ativo = "active"
                                                                    WHILE NOT rsquestionarios.EOF
                                                        %>
                                                                        <li class="<%=ativo%>">
                                                                            <a class="nav-link" data-toggle="tab" href="#tab-<%=aa%>">
                                                                                <%=rsquestionarios("titulo")%>
                                                                            </a>
                                                                        </li>
                                                        <%
                                                                    aa = aa + 1
                                                                    ativo = ""
                                                                    rsquestionarios.movenext 
                                                                    WEND 
                                                                    rsquestionarios.movefirst
                                                                END IF 
                                                        %>
                                                    </ul>
                                                    <div class="tab-content">
                                                        <%
                                                            IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                                                                SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
                                                            else

                                                                SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                                                            end if

                                                            

                                                            IF NOT rsquestionarios.EOF THEN
                                                                aa = 1
                                                                ativo = "active"
                                                                WHILE NOT rsquestionarios.EOF
                                                        %>
                                                                    <div role="tabpanel" id="tab-<%=aa%>" class="tab-pane <%=ativo%>">
                                                                        <div class="panel-body">    
                                                                            <% 
                                                                                    set rscampo = conexao.execute("SELECT * FROM neg_quest_campos where status = 'Ativo' and id_nq = "&rsquestionarios("idnq")&" order by ordem asc")    
                                                                            %>
                                                                                    <div role="tabpanel" id="#quest<%=aa%>" class="tab-pane <%=ativo%>">
                                                                                        <div class="panel-body" style="border: none;">
                                                                            <%
                                                                                            contador = 1
                                                                                            while not rscampo.EOF 

                                                                                                IF rscampo("obrigatorio") = "Sim" THEN 
                                                                                                    obrigatorio = ""
                                                                                                ELSE 
                                                                                                    obrigatorio = ""
                                                                                                END IF 

                                                                                                mask = ""

                                                                                                if rscampo("mascara") = "telefone" then
                                                                                                    mask="MascaraTelefone"
                                                                                                elseif rscampo("mascara") = "data" then
                                                                                                    mask="MascaraData"
                                                                                                elseif rscampo("mascara") = "valor" then
                                                                                                    mask = "moeda"
                                                                                                elseif rscampo("mascara") = "cpf" then
                                                                                                    mask = "MascaraCPF"
                                                                                                elseif rscampo("mascara") = "cnpj" then 
                                                                                                    mask = "MascaraCNPJ"
                                                                                                elseif rscampo("mascara") = "link" then 
                                                                                                    mask = "link"
                                                                                                elseif rscampo("mascara") = "inteiro" then 
                                                                                                    mask = "inteiro"
                                                                                                end if  

                                                                                                set rsresposta = conexao.execute("SELECT * FROM neg_quest_respostas where id_campo = "&rscampo("idcampo")&" and id_prospecter = "&request.querystring("id"))      
                                                                                        
                                                                                                if not rsresposta.eof then
                                                                                    
                                                                                            
                                                                                                    if rscampo("tipo") = "text" then
                                                                                                        if mask <> "link" THEN 
                                                                            %>
                                                                                                            <label><%=rscampo("nome")%></label>
                                                                                                            <div style="display: flex;">
                                                                                                                <%IF mask = "moeda" THEN%>
                                                                                                                    <input type="" name="" disabled="" placeholder="R$" style="width: 50px;border-right: 0px" class="form-control">
                                                                                                                <%END IF%>
                                                                                                                <input disabled <%if rscampo("mascara") <> "valor" then%>onkeypress="<%=mask%>(this)"<%end if%> <%if rscampo("mascara") = "valor" then%>onkeypress="return(moeda(this,'.',',',event))"<%end if%>  value="<%=rsresposta("resposta")%>" type="text" class="form-control <%=rscampo("mascara")%>" <%=obrigatorio%> name="campo<%=rscampo("idcampo")%>">
                                                                                                            </div>
                                                                            <%                      
                                                                                                        else 
                                                                                                            link_acesso = rsresposta("resposta")
                                                                            %>
                                                                                                            <label><%=rscampo("nome")%></label> <br>
                                                                                                            <a target="_blank" href="<%=link_acesso%>" class="btn btn-primary"><%=link_acesso%></a> 
                                                                            <%
                                                                                                        end if 
                                                                                                    elseif rscampo("tipo") = "check" then
                                                                            %>
                                                                                                        <input disabled type="checkbox" <%if rsresposta("resposta") = "on" then response.write("checked")%> class="<%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                        <label><%=rscampo("nome")%></label>
                                                                            <%
                                                                                                    else
                                                                            %>
                                                                                                        <label><%=rscampo("nome")%></label>
                                                                                                        <textarea disabled <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>"><%=rsresposta("resposta")%></textarea>
                                                                            <%
                                                                                                    end if
                                                                                    
                                                                                                else
                                                                            
                                                                                                    if rscampo("tipo") = "text" then
                                                                                                        if mask <> "link" THEN 
                                                                            %>
                                                                                                            <label><%=rscampo("nome")%></label>
                                                                                                            <div style="display: flex;">
                                                                                                                <%IF mask = "moeda" THEN%>
                                                                                                                    <input type="" name="" disabled="" placeholder="R$" style="width: 50px;border-right: 0px" class="form-control">
                                                                                                                <%END IF%>
                                                                                                                <input disabled <%if rscampo("mascara") <> "valor" then%>onkeypress="<%=mask%>(this)"<%end if%> <%if rscampo("mascara") = "valor" then%>onkeypress="return(moeda(this,'.',',',event))"<%end if%> type="text" <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                            </div>
                                                                            <%
                                                                                                        else 
                                                                            %>
                                                                                                            <label><%=rscampo("nome")%></label> <br>
                                                                                                            <a class="btn btn-primary disabled">Vazio</a> 
                                                                            <%
                                                                                                        end if 
                                                                                                    elseif rscampo("tipo") = "check" then
                                                                            %>
                                                                                                        <input disabled type="checkbox" class="<%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                        <label><%=rscampo("nome")%></label>
                                                                            <%
                                                                                                    elseif rscampo("tipo") = "text-area" then
                                                                            %>
                                                                                                        <label><%=rscampo("nome")%></label>
                                                                                                        <textarea disabled <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>"></textarea>
                                                                            <%
                                                                                                    end if
                                                                                                end if
                                                                            %>
                                                                                                <br>
                                                                            <%
                                                                                            contador = contador + 1
                                                                                            rscampo.movenext
                                                                                            WEND
                                                                            %>      
                                                                                        </div>
                                                                                    </div>
                                                                        </div>
                                                                    </div>
                                                        <%
                                                                aa = aa + 1
                                                                ativo = ""
                                                                rsquestionarios.movenext 
                                                                WEND 
                                                                rsquestionarios.movefirst
                                                            END IF 
                                                        %>
                                                    </div>
                                                </div>
                               
                                        </div>
                                        </div>
                                        <div style="display: block; position: relative; clear: both;">&nbsp;</div>
                                    </div>
                                </div>

                                <div role="tabpanel" id="tabcriar" class="tab-pane">
                                    <div class="panel-body">
                                        
                                    </div>
                                </div>
                            </div>
                        </div>

                    </div>

                    <%
                        'SET rs_processo = conexao.execute("SELECT nfe.idnfe, nfe.processo, nfe.processo, nf.id_etapa_funil, nfe.etapa FROM neg_funil nf INNER JOIN neg_funil_etapas nfe ON nf.id_etapa_funil = nfe.idnfe WHERE id_negp = "&id_neg&" ORDER BY idnfe DESC LIMIT 1")

                        SET rs_processo = conexao.execute("SELECT nfe.processo, nf.id_etapa_funil, nf.data, nfe.etapa, u.usuario FROM neg_funil nf INNER JOIN neg_funil_etapas nfe ON nf.id_etapa_funil = nfe.idnfe INNER JOIN usuarios u ON  nf.id_usuario = u.idusuario WHERE id_negp = "&id_neg&" ORDER BY nf.data DESC LIMIT 1")
                    %>
                        <div class="ibox-content m-b-sm border-bottom col-lg-3">
                    <%
                            IF NOT rs_processo.EOF AND rs_processo("processo") <> "" THEN  
                                'WHILE NOT rs_processo.EOF  
                    %>
                                <h3><%=rs_processo("etapa")%></h3>
                                <br>
                                <p><%=rs_processo("processo")%></p>
                    <%
                            ELSE 
                    %>
                                <h3><%=rs_processo("etapa")%></h3>
                                <br>
                                <p>Nenhuma descrição para essa Etapa.</p>
                    <%
                                    'rs_processo.movenext
                                'WEND
                            END IF 
                    %>
                        </div>
                    
                </div>
            </div>

        </div>
        <!--#include file="estrutura_rodape.asp"-->

        </div>
        </div>

    <!-- Modal Editar Histórico-->
    <div id="modeditarHistorico" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Editar Histórico</h4>
                </div>
                <div class="modal-body">
                    <form ACTION="?id=<%=request.querystring("id")%>" METHOD="POST">

                        <label>Tipo de interação: *</label>
                        <select class="form-control" required="" id="interacao_editar" name="id_tipo">

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
                        <textarea style="height: 100px;margin-bottom: 0px" class="form-control" id="descricao_editar" name="descricao" required></textarea>

                        <br>
                        <div style="display: flex; justify-content: space-between;">

                            <div class="col-md-4" style="padding-left: 0px">
                                <input type="submit" class="btn btn-primary" value="Atualizar">
                            </div>

                            <input type="hidden" name="act" value="editar_historico">
                            <input type="hidden" name="idinteracao" value="" id="idinteracao">

                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal Anotação -->
    <div id="modalanota" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Adicionar Anotação</h4>
                </div>
                <div class="modal-body">
                    <form ACTION="<%=MM_editAction%>" METHOD="POST">

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
                        <textarea style="height: 100px;" class="form-control" rows="10" name="descricao" required></textarea>
                        <br>

                        <div style="display: flex; justify-content: space-between;">

                            <!-- <div class="col-md-4" style="padding-left: 0px">
                                <div class="form-group" id="data_anotacao">
                                    <div class="input-group date">
                                        <span class="input-group-addon"><i class="fa fa-calendar"></i></span>
                                        <input autocomplete="OFF" value="<%=data_atual%>" name="agenda_data" type="text" class="form-control datasemhora" required value="<=date%>" onBlur="verificacao_data_maior(this.value)">
                                    </div>
                                </div>
                            </div>


                            <div class="col-md-4" style="padding-left: 0px;display: flex;">
                                <span class="input-group-addon" style="padding-top: 3%;height: 34px;width: 41px"><i class="fa fa-clock"></i></span>
                                <input type="text" class="form-control hora" value="<=time%>" name="agenda_hora" required>
                            </div> -->

                            <div class="col-md-4" style="padding-left: 0px">
                                <input type="submit" class="btn btn-primary" value="Inserir">
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
                    <h4 class="modal-title">Agendar uma reunião</h4>
                </div>
                <div class="modal-body">
                    <form ACTION="<%=MM_editAction%>" METHOD="POST">
                    
                        <div class="col-lg-12 form-control" style="border: none !important; display: flex; padding: 0px;margin-bottom: 0px">
                            <div class="col-lg-3" style="padding-left: 0;">
                                <input name="tipo_reuniao" style="float: left;margin-right: 3%" type="radio" value="Reunião" checked> Reunião
                            </div>
                            <div class="col-lg-3" style="padding-left: 0;">
                                <input name="tipo_reuniao" style="float: left;margin-right: 3%" type="radio" value="Lembrete"> Lembrete 
                            </div>
                        </div>
            
                        <label>Descrição:</label>
                        <textarea class="form-control" style="height: 100px" name="descricao" required></textarea>
                        <div> 
                            <div style="display: flex;">
                                <div class="col-md-6" style="padding-left: 0px;float: initial;">
                                    <div class="form-group" id="data_agenda">
                                        <div class="input-group date">
                                            <span class="input-group-addon"><i class="fa fa-calendar"></i></span>
                                            <input autocomplete="OFF" value="<%=data_atual%>" name="agenda_data" type="text" class="form-control datasemhora" required value="<%=date%>" onBlur="verificacao_data_maior(this.value)">
                                        </div>
                                    </div>
                                </div>


                                <div class="col-md-6" style="padding-left: 0px;display: flex;float: initial;padding-right: 0px">

                                    <div class="col-sm-12" style="padding-right: 0px">
                                        <div class="input-group clockpicker" data-autoclose="true">
                                            <input autocomplete="off" name="agenda_hora" type="text" class="form-control" placeholder="00:00" value="<%=h_termino%>" value="<%=time%>">
                                            <span class="input-group-addon">
                                                <span class="fa fa-clock-o"></span>
                                            </span>
                                        </div>
                                    </div>

                                </div>
                            </div>

                            <div class="col-md-12" style="padding-left: 0px;float: initial;">
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

    <!-- Modal -->
    <div id="mod_funil" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Alterar Funil</h4>
                </div>
                <div class="modal-body">
                    <form ACTION="<%=MM_editAction%>" METHOD="POST" name="forminserir" id="forminserir">
                        <strong>Funil:</strong> <br>
                        <select class="form-control" name="funil_etapa">
                            <%
                                SET rsfunil = conexao.execute("SELECT idnfe, etapa FROM neg_funil_etapas WHERE status = 'Ativo'")

                                WHILE NOT rsfunil.EOF
                            %>
                                <option value="<%=rsfunil("idnfe")%>"><%=rsfunil("etapa")%></option>
                            <%
                                rsfunil.movenext 
                                WEND
                            %>
                        </select>
                        <br>

                        <input type="hidden" name="id_usuario_funil" value="<%=Request.Cookies("grpro")("idusuario")%>">

                        <input class="btn btn-primary center-block btt_custom_2" name="salvar" type="submit" id="salvar" value="Salvar">
                        <input type="hidden" name="MM_Insert" value="formfunil">
                    </form> 
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal -->
    <div id="moduploadetapas" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-body">
                    <div class="panel-body">
                        <form id="form_inserir" action="../upload_doc_neg.php" method="post" enctype="multipart/form-data">

                            <h1 style="margin-top: 0; margin-bottom: 45px;">Faça o upload de documentos</h1>
                            <fieldset>
                                <%
                                    SET rsdocumentos = conexao.execute("SELECT idndc, categoria FROM neg_doc_categoria WHERE status = 'Ativo' and obrigatorio = 'Sim'")

                                    i_doc = 1
                                    WHILE NOT rsdocumentos.EOF 
                                %>
                                        <div class="col-lg-6">
                                            <label><%=rsdocumentos("categoria")%>:</label>
                                            <input onChange="verifica_arquivo(<%=i_doc%>)" class="form-control" name="doc<%=rsdocumentos("idndc")%>" id="file<%=i_doc%>" type="file" required>
                                        </div>

                                        <div class="col-lg-6">
                                            <label>Título:</label>
                                            <input type="text" class="form-control" name="titulo<%=rsdocumentos("idndc")%>" required>
                                        </div>                                            
                                        
                                        <div style="overflow: auto; clear: both; "></div>

                                        <br>  
                                <%
                                    i_doc = i_doc + 1
                                    rsdocumentos.movenext 
                                    WEND
                                %>   
                            </fieldset>

                            <input name="idnp" type="hidden" value="<%=Request.QueryString("id")%>">
                            
                            <input class="btn btn-primary pull-right" type="submit" value="Upload">
                        </form>
                    </div>    
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Pular</button>
                </div>
            </div>

        </div>
    </div>

    <div id="editar_p" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Editar Negociação</h4>
                </div>
                <form <%=MM_editAction%> METHOD="POST" name="formedit" id="formedit">
                    <div class="modal-body">
                        <div class="tabs-container">
                            <ul class="nav nav-tabs">
                                <li class="active"><a class="nav-link active" data-toggle="tab" href="#dados_p">Dados Primários</a></li>

                                <%
                                    IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                                        SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
                                    else

                                        SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                                    end if

                                    IF NOT rsquestionarios.EOF THEN
                                %>
                                        <li><a class="nav-link" data-toggle="tab" href="#dados_sec">Dados Questionário</a></li>
                                <%
                                    END IF
                                %>
                                
                            </ul>
                            
                            <div class="tab-content">
                                <div id="dados_p" class="tab-pane active">
                                    <div class="panel-body">

                                        <label>Perfil</label>
                                        <%
                                            SET neg_perfil = conexao.execute("SELECT * FROM neg_perfil where status = 'Ativo'")
                                        %>
                                        <select class="form-control" name="perfil">
                                            <option value="">Nenhum</option>
                                            <%
                                                WHILE NOT neg_perfil.EOF
                                            %>
                                                <option <%if neg_perfil("idperfil") = rsemp("id_perfil") then%>selected<%end if%> value="<%=neg_perfil("idperfil")%>"><%=neg_perfil("perfil")%></option>
                                            <%
                                                    neg_perfil.movenext
                                                WEND
                                            %>
                                        </select>
                                        <br>

                                        <%
                                            SET neg_cpf = conexao.execute("SELECT cpf FROM neg_prospecter WHERE del = 0 AND idnp = "&id_neg)

                                            IF Len(neg_cpf("cpf")) = "18" THEN 
                                                mascara_cnpj_cpf = "cnpj"
                                                j_check = "checked"
                                                aux_tipo = "CNPJ"
                                                mascara = "MascaraCNPJ"
                                                aux_mascara = "cnpj"
                                                max_aux = "18"
                                            ELSEIF Len(neg_cpf("cpf")) = "14" THEN
                                                mascara_cnpj_cpf = "cpf"
                                                f_check = "checked"
                                                aux_tipo = "CPF"
                                                mascara = "MascaraCPF"
                                                aux_mascara = "cpf"
                                                max_aux = "14"
                                            ELSE
                                                mascara_cnpj_cpf = "cnpj"
                                                j_check = "checked"
                                                aux_tipo = "CNPJ"
                                                mascara = "MascaraCNPJ"
                                                aux_mascara = "cnpj"
                                                max_aux = "18"
                                            END IF 
                                        %>

                                        <strong>Tipo:</strong>
                                        <div class="form-control" style="border: none !important; display: flex; justify-content: space-between;padding: 0px;margin-bottom: 0px">
                                            <div class="col-lg-6" style="padding-left: 0;">
                                                <input name="tipo_prospecter" style="float: left;margin-right: 3%" type="radio" value="j" onClick="mascara_cpf_cnpj_edit('juridica', <%=id_neg%>);" <%=j_check%>> Jurídica
                                            </div>
                                            <div class="col-lg-6" style="padding-left: 0;">
                                                <input name="tipo_prospecter" style="float: left;margin-right: 3%" type="radio" value="f" onClick="mascara_cpf_cnpj_edit('fisica', <%=id_neg%>);" <%=f_check%>> Física 
                                            </div>
                                        </div>
                   
                                        <div id="retorno_cpfcnpj_prospecter">
                                            <strong><%=aux_tipo%>:</strong>
                                            <div id="exibe_cnpj_prospecter">
                                                <!-- <small><%=aux_tipo%> Válido.</small> -->
                                                <input type="hidden" id="cnpj_status" value="nao"></input>
                                                <input type="hidden" id="tipo" value="cnpj">
                                                <input id="<%=aux_mascara%>" style="border: 1px green solid; color: green;" onblur="verifica_cnpj_edit_prospecter(this.value, <%=id_neg%>)" name="cpf_cnpj_P" onKeyPress="<%=mascara%>(formedit.<%=aux_mascara%>);" value="<%=(neg_cpf.Fields.Item("cpf").Value)%>" type="text" class="form-control <%=mascara_cnpj_cpf%>" maxlength="<%=max_aux%>">
                                            </div>  
                                        </div>

                                        <label>Empresa: *</label>
                                        <input required type="text" class="form-control" name="empresa" value="<%=rsemp("empresa")%>">
                                        <br>

                                        <label>Responsável: *</label>
                                        <input required type="text" class="form-control" name="responsavel" value="<%=rsemp("responsavel")%>">
                                        <br>

                                        <label>E-mail:</label>
                                        <input type="text" class="form-control" name="email" value="<%=rsemp("email")%>">
                                        <br>

                                        <label>Telefone:</label>
                                        <input type="text" class="form-control telefone_ddd" name="telefone" value="<%=rsemp("telefone")%>">
                                        <br>

                                        <div class="col-md-12" style="display: flex;padding: 0px">
                                            <div class="col-md-8" style="padding: 0px;padding-right: 2%">
                                                <label>Cidade: *</label>
                                                <input required type="text" class="form-control " name="cidade" value="<%=rsemp("cidade")%>">
                                                <br>
                                            </div>
                                            <div class="col-md-4" style="padding: 0px">
                                                <label>UF: *</label>
                                                <input required type="text" class="form-control " name="uf" value="<%=rsemp("uf")%>">
                                                <br>
                                            </div>
                                        </div>

                                        <label>Interesse: *</label>
                                        <select required type="text" class="form-control" name="id_produto">
                                            <option value="">Selecione...</option>
                                            <%
                                                SET rsprod = conexao.execute("SELECT idnp, produto FROM neg_produtos")
                                                WHILE NOT rsprod.EOF
                                            %>
                                                    <option <% IF rsprod("idnp") = rsemp("id_produto") THEN Response.Write("SELECTED")%> value="<%=rsprod("idnp")%>"><%=rsprod("produto")%></option>
                                            <%
                                                    rsprod.movenext
                                                WEND 
                                            %>
                                        </select>
                                        <br>

                                        <div class="col-md-12" style="display: flex;padding: 0px">
                                            <div class="col-md-6" style="padding: 0px;padding-right: 2%">
                                                <label>Valor de Adesão: *</label>
                                                <input required type="text" class="form-control dinheiro" name="valor_adesao" value="<%=valor_adesao%>" placeholder="0,00">
                                                <br>
                                            </div>
                                            <div class="col-md-6" style="padding: 0px">
                                                <label>Valor de Recorrência: *</label>
                                                <input required type="text" class="form-control dinheiro" name="valor_recorrencia" value="<%=valor_recorrencia%>" placeholder="0,00">
                                                <br>
                                            </div>
                                        </div>

                                        <label>Canal de Aquisição: *</label>
                                        <select required type="text" class="form-control" name="id_aquisicao">
                                            <option value="">Selecione...</option>
                                            <%
                                                SET rsaqui = conexao.execute("SELECT idna, aquisicao FROM neg_aquisicao")
                                                WHILE NOT rsaqui.EOF
                                            %>
                                                <option <% IF rsaqui("idna") = rsemp("id_aquisicao") THEN Response.Write("SELECTED")%> value="<%=rsaqui("idna")%>"><%=rsaqui("aquisicao")%></option>
                                            <%
                                                rsaqui.movenext
                                                WEND 
                                            %>
                                        </select>
                                        <br>

                                        <label>Usuário Responsável: *</label>
                                        <select type="text" class="select2_demo_2 form-control" multiple="multiple" name="id_usuario" required="">
                                            <%
                                                SET rsusuarios = conexao.execute("SELECT u.idusuario, u.usuario, s.setor, s.idsetor FROM usuarios u INNER JOIN setores s ON u.id_setor = s.idsetor WHERE u.status = 'Ativo' ORDER BY setor ASC")

                                                WHILE NOT rsusuarios.EOF

                                                    SET rsverifica = conexao.execute("SELECT * FROM permissoes WHERE id_setor = "&rsusuarios("idsetor")&" and id_modulo = 9")

                                                    if not rsverifica.EOF then
                                                        selected = ""

                                                        rsresponsavel.filter = "id_usuario = "&rsusuarios("idusuario")
                                                        IF NOT rsresponsavel.EOF THEN selected = "selected"
                                            %>
                                                        <option <%=selected%> value="<%=rsusuarios("idusuario")%>"><%=rsusuarios("usuario")%></option>
                                            <%
                                                    end if
                                                    rsusuarios.movenext
                                                WEND 
                                            %>
                                        </select>

                                    </div>            
                                </div>

                                <div id="dados_sec" class="tab-pane">
                                    <div class="panel-body">
                                        <div class="tabs-container">
                                            <ul class="nav nav-tabs" role="tablist">
                                                <%


                                                    IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                                                        SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
                                                    else

                                                        SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                                                    end if

                                                        IF NOT rsquestionarios.EOF THEN
                                                            aa = 1000
                                                            ativo = "active"
                                                            WHILE NOT rsquestionarios.EOF
                                                %>
                                                                <li class="<%=ativo%>">
                                                                    <a class="nav-link" data-toggle="tab" href="#tab-<%=aa%>">
                                                                        <%=rsquestionarios("titulo")%>
                                                                    </a>
                                                                </li>
                                                <%
                                                            aa = aa + 1
                                                            ativo = ""
                                                            rsquestionarios.movenext 
                                                            WEND 
                                                            rsquestionarios.movefirst
                                                        END IF 
                                                %>
                                            </ul>
                                            <div class="tab-content">
                                                <%
                                                    IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                                                            SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
                                                        else

                                                            SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
                                                        end if

                                                    IF NOT rsquestionarios.EOF THEN
                                                        aa = 1000
                                                        aaa = 1000
                                                        ativo = "active"
                                                        WHILE NOT rsquestionarios.EOF
                                                %>
                                                            <div role="tabpanel" id="tab-<%=aaa%>" class="tab-pane <%=ativo%>">
                                                                <div class="panel-body">    
                                                                    <% 
                                                                            set rscampo = conexao.execute("SELECT * FROM neg_quest_campos where status = 'Ativo' and id_nq = "&rsquestionarios("idnq")&" order by ordem asc")    
                                                                    %>
                                                                            <div role="tabpanel" id="#quest<%=aaa%>" class="tab-pane <%=ativo%>">
                                                                                <div class="panel-body" style="border: none;">
                                                                    <%
                                                                                    contador = 1
                                                                                    while not rscampo.EOF 

                                                                                        IF rscampo("obrigatorio") = "Sim" THEN 
                                                                                            obrigatorio = ""
                                                                                        ELSE 
                                                                                            obrigatorio = ""
                                                                                        END IF 

                                                                                        mask = ""

                                                                                        if rscampo("mascara") = "telefone" then
                                                                                            mask="MascaraTelefone"
                                                                                        elseif rscampo("mascara") = "data" then
                                                                                            mask="MascaraData"
                                                                                        elseif rscampo("mascara") = "valor" then
                                                                                            mask = "moeda"
                                                                                        elseif rscampo("mascara") = "cpf" then
                                                                                            mask = "MascaraCPF"
                                                                                        elseif rscampo("mascara") = "cnpj" then 
                                                                                            mask = "MascaraCNPJ"
                                                                                        elseif rscampo("mascara") = "link" then 
                                                                                            mask = "link"
                                                                                        elseif rscampo("mascara") = "inteiro" then 
                                                                                            mask = "inteiro"
                                                                                        end if  

                                                                                        set rsresposta = conexao.execute("SELECT * FROM neg_quest_respostas where id_campo = "&rscampo("idcampo")&" and id_prospecter = "&request.querystring("id"))      
                                                                                
                                                                                        if not rsresposta.eof then
                                                                            
                                                                                    
                                                                                            if rscampo("tipo") = "text" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>

                                                                                                <div style="display: flex;">
                                                                                                    <%if mask = "moeda" then%>
                                                                                                        <input type="" name="" disabled="" placeholder="R$" style="width: 50px;border-right: 0px" class="form-control">
                                                                                                    <%end if%>

                                                                                                    <input <%if rscampo("mascara") <> "valor" then%>onkeypress="<%=mask%>(this)"<%end if%> <%if rscampo("mascara") = "valor" then%>onkeypress="return(moeda(this,'.',',',event))"<%end if%>  value="<%=rsresposta("resposta")%>" type="text" class="form-control <%=rscampo("mascara")%>" <%=obrigatorio%> name="campo<%=rscampo("idcampo")%>">
                                                                                                </div>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "check" then
                                                                    %>
                                                                                                <input  type="checkbox" <%if rsresposta("resposta") = "on" then response.write("checked")%> class="<%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                <label><%=rscampo("nome")%></label>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "text-area" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <textarea <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>"><%=rsresposta("resposta")%></textarea>

                                                                    <%
                                                                                            elseif rscampo("tipo") = "selecao" THEN
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <select class="form-control" name="campo<%=rscampo("idcampo")%>">
                                                                                                    <%
                                                                                                        resposta_value = rsresposta("resposta")
                                                                                                    %>
                                                                                                    <option value="">Selecione</option>
                                                                                                    <%
                                                                                                      valores = split(rscampo("campos_selecao"),"|")
                                                                                                      FOR i = 0 to uBound(valores)
                                                                                                    %>
                                                                                                        <option <%IF resposta_value = valores(i) THEN response.write("selected")%> value="<%=valores(i)%>"><%=valores(i)%></option>
                                                                                                    <%
                                                                                                      Next
                                                                                                    %>


                                                                                                </select> 
                                                                    <%
                                                                                            end if
                                                                            
                                                                                        else
                                                                    
                                                                                            if rscampo("tipo") = "text" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <div style="display: flex;">
                                                                                                    <%if mask = "moeda" then%>
                                                                                                        <input type="" name="" disabled="" placeholder="R$" style="width: 50px;border-right: 0px" class="form-control">
                                                                                                    <%end if%>

                                                                                                    <input <%if rscampo("mascara") <> "valor" then%>onkeypress="<%=mask%>(this)"<%end if%> <%if rscampo("mascara") = "valor" then%>onkeypress="return(moeda(this,'.',',',event))"<%end if%> type="text" class="form-control <%=rscampo("mascara")%>" <%=obrigatorio%> name="campo<%=rscampo("idcampo")%>">
                                                                                                </div>
                                                                    <%  
                                                                                            elseif rscampo("tipo") = "check" then
                                                                    %>
                                                                                                <input type="checkbox" class="<%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                <label><%=rscampo("nome")%></label>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "text-area" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <textarea <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>"></textarea>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "selecao" THEN
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <select class="form-control" name="campo<%=rscampo("idcampo")%>">
                                                                                                    <option value="">Selecione</option>
                                                                                                    <%
                                                                                                      valores = split(rscampo("campos_selecao"),"|")
                                                                                                      FOR i = 0 to uBound(valores)
                                                                                                    %>
                                                                                                        <option value="<%=valores(i)%>"><%=valores(i)%></option>
                                                                                                    <%
                                                                                                      Next
                                                                                                    %>


                                                                                                </select>  
                                                                    <%
                                                                                            end if
                                                                                        end if
                                                                    %>
                                                                                        <br>
                                                                    <%
                                                                                    aa = aa + 1
                                                                                    contador = contador + 1
                                                                                    rscampo.movenext
                                                                                    WEND
                                                                    %>      
                                                                                </div>
                                                                            </div>
                                                                </div>
                                                            </div>
                                                <%
                                                        aa = aa + 1
                                                        aaa = aaa + 1
                                                        ativo = ""
                                                        rsquestionarios.movenext 
                                                        WEND 
                                                        rsquestionarios.movefirst
                                                    END IF 
                                                %>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                
                            </div>
                            
                        </div>
                    </div>
                    <div class="modal-footer">
                        <input class="btn btn-primary pull-right"  type="submit" value="Salvar">
                        <input hidden name="MM_Update" value="negociacao">
                    </div>
                </form>
            </div>
        </div>
    </div>

    <script type="text/javascript">
        $(document).on('hide.bs.modal','#mod_quest', function () {
            window.location.href = "neg_interacao_novo.asp?id=<%=request.querystring("id")%>"
        });

        $(document).on('hide.bs.modal','#moduploadetapas', function () {
            window.location.href = "neg_interacao_novo.asp?act=questionarios&id=<%=request.querystring("id")%>"
        });

        
    </script>

        <%
            SET rsperfil = conexao.execute("SELECT id_perfil FROM neg_prospecter WHERE del = 0 AND idnp = "&request.querystring("id")&" and  id_perfil is not null  and id_perfil <> 0")

            if not rsperfil.EOF then
                SET rsquestionarios = conexao.execute("SELECT DISTINCT q.idnq, q.titulo FROM neg_quest_perfil as qp INNER JOIN neg_questionarios as q ON q.idnq = qp.id_questionario INNER JOIN neg_perfil as p ON p.idperfil = qp.id_perfil INNER JOIN neg_quest_campos nc ON nc.id_nq = q.idnq WHERE q.status = 'Ativo' and qp.id_perfil = "&rsperfil("id_perfil")&" ORDER BY q.ordem ASC")
            else
                SET rsquestionarios = conexao.execute("SELECT DISTINCT q.idnq, q.titulo FROM neg_quest_perfil as qp INNER JOIN neg_questionarios as q ON q.idnq = qp.id_questionario INNER JOIN neg_perfil as p ON p.idperfil = qp.id_perfil INNER JOIN neg_quest_campos nc ON nc.id_nq = q.idnq WHERE q.status = 'Ativo' ORDER BY q.ordem ASC")
            end if

            IF NOT rsquestionarios.EOF THEN
        %>

                <div id="mod_quest"  class="modal fade" role="dialog">
                    <div class="modal-dialog">

                        <!-- Modal content-->
                        <div class="modal-content">
                            
                            <div class="modal-body">

                                <div class="panel-body">
                                    <form id="form_inserir" action="neg_interacao_novo.asp?id=<%=request.querystring("id")%>" method="post">
                                        <h1 style="margin-top: 0; margin-bottom: 45px;font-size: 28px">Responda aos questionários para finalizar o cadastro</h1>
                                        <div class="tabs-container">
                                            <ul class="nav nav-tabs" role="tablist">
                                                <%
                                                            aa = 2000
                                                            ativo = "active"
                                                            WHILE NOT rsquestionarios.EOF
                                                %>
                                                                <li class="<%=ativo%>">
                                                                    <a class="nav-link" data-toggle="tab" href="#tab-<%=aa%>">
                                                                        <%=rsquestionarios("titulo")%>
                                                                    </a>
                                                                </li>
                                                <%
                                                            aa = aa + 1
                                                            ativo = ""
                                                            rsquestionarios.movenext 
                                                            WEND 
                                                            rsquestionarios.movefirst
                                                %>
                                            </ul>
                                            <div class="tab-content">
                                                <%
                                                        aa = 2000
                                                        aaa = 2000
                                                        ativo = "active"
                                                        WHILE NOT rsquestionarios.EOF
                                                %>
                                                            <div role="tabpanel" id="tab-<%=aaa%>" class="tab-pane <%=ativo%>">
                                                                <div class="panel-body">    
                                                                    <% 
                                                                            set rscampo = conexao.execute("SELECT * FROM neg_quest_campos where status = 'Ativo' and id_nq = "&rsquestionarios("idnq")&" order by ordem asc")    
                                                                    %>
                                                                            <div role="tabpanel" id="#quest<%=aaa%>" class="tab-pane <%=ativo%>">
                                                                                <div class="panel-body" style="border: none;">
                                                                    <%
                                                                                    contador = 1
                                                                                    while not rscampo.EOF 

                                                                                        IF rscampo("obrigatorio") = "Sim" THEN 
                                                                                            obrigatorio = ""
                                                                                        ELSE 
                                                                                            obrigatorio = ""
                                                                                        END IF 

                                                                                        mask = ""

                                                                                        if rscampo("mascara") = "telefone" then
                                                                                            mask="MascaraTelefone"
                                                                                        elseif rscampo("mascara") = "data" then
                                                                                            mask="MascaraData"
                                                                                        elseif rscampo("mascara") = "valor" then
                                                                                            mask = "moeda"
                                                                                        elseif rscampo("mascara") = "cpf" then
                                                                                            mask = "MascaraCPF"
                                                                                        elseif rscampo("mascara") = "cnpj" then 
                                                                                            mask = "MascaraCNPJ"
                                                                                        elseif rscampo("mascara") = "link" then 
                                                                                            mask = "link"
                                                                                        elseif rscampo("mascara") = "inteiro" then 
                                                                                            mask = "inteiro"
                                                                                        end if  

                                                                                        set rsresposta = conexao.execute("SELECT * FROM neg_quest_respostas where id_campo = "&rscampo("idcampo")&" and id_prospecter = "&request.querystring("id"))      
                                                                                
                                                                                        if not rsresposta.eof then
                                                                            
                                                                                    
                                                                                            if rscampo("tipo") = "text" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>

                                                                                                <div style="display: flex;">
                                                                                                    <%if mask = "moeda" then%>
                                                                                                        <input type="" name="" disabled="" placeholder="R$" style="width: 50px;border-right: 0px" class="form-control">
                                                                                                    <%end if%>

                                                                                                    <input <%if rscampo("mascara") <> "valor" then%>onkeypress="<%=mask%>(this)"<%end if%> <%if rscampo("mascara") = "valor" then%>onkeypress="return(moeda(this,'.',',',event))"<%end if%>  value="<%=rsresposta("resposta")%>" type="text" class="form-control <%=rscampo("mascara")%>" <%=obrigatorio%> name="campo<%=rscampo("idcampo")%>">
                                                                                                </div>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "check" then
                                                                    %>
                                                                                                <input  type="checkbox" <%if rsresposta("resposta") = "on" then response.write("checked")%> class="<%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                <label><%=rscampo("nome")%></label>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "text-area" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <textarea <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>"><%=rsresposta("resposta")%></textarea>

                                                                    <%
                                                                                            elseif rscampo("tipo") = "selecao" THEN
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <select class="form-control" name="campo<%=rscampo("idcampo")%>">
                                                                                                    <option value="">Selecione</option>
                                                                                                    <%
                                                                                                      valores = split(rscampo("campos_selecao"),"|")
                                                                                                      FOR i = 0 to uBound(valores)
                                                                                                    %>
                                                                                                        <option <%IF rsresposta("resposta") = valores(i) THEN response.write("selected")%> value="<%=valores(i)%>"><%=valores(i)%></option>
                                                                                                    <%
                                                                                                      Next
                                                                                                    %>


                                                                                                </select> 
                                                                    <%
                                                                                            end if
                                                                            
                                                                                        else
                                                                    
                                                                                            if rscampo("tipo") = "text" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <div style="display: flex;">
                                                                                                    <%if mask = "moeda" then%>
                                                                                                        <input type="" name="" disabled="" placeholder="R$" style="width: 50px;border-right: 0px" class="form-control">
                                                                                                    <%end if%>

                                                                                                    <input <%if rscampo("mascara") <> "valor" then%>onkeypress="<%=mask%>(this)"<%end if%> <%if rscampo("mascara") = "valor" then%>onkeypress="return(moeda(this,'.',',',event))"<%end if%> type="text" class="form-control <%=rscampo("mascara")%>" <%=obrigatorio%> name="campo<%=rscampo("idcampo")%>">
                                                                                                </div>
                                                                    <%  
                                                                                            elseif rscampo("tipo") = "check" then
                                                                    %>
                                                                                                <input type="checkbox" class="<%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>">
                                                                                                <label><%=rscampo("nome")%></label>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "text-area" then
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <textarea <%=obrigatorio%> class="form-control <%=rscampo("mascara")%>" name="campo<%=rscampo("idcampo")%>"></textarea>
                                                                    <%
                                                                                            elseif rscampo("tipo") = "selecao" THEN
                                                                    %>
                                                                                                <label><%=rscampo("nome")%></label>
                                                                                                <select class="form-control" name="campo<%=rscampo("idcampo")%>">
                                                                                                    <option value="">Selecione</option>
                                                                                                    <%
                                                                                                      valores = split(rscampo("campos_selecao"),"|")
                                                                                                      FOR i = 0 to uBound(valores)
                                                                                                    %>
                                                                                                        <option value="<%=valores(i)%>"><%=valores(i)%></option>
                                                                                                    <%
                                                                                                      Next
                                                                                                    %>


                                                                                                </select>  
                                                                    <%
                                                                                            end if
                                                                                        end if
                                                                    %>
                                                                                        <br>
                                                                    <%
                                                                                    aa = aa + 1
                                                                                    contador = contador + 1
                                                                                    rscampo.movenext
                                                                                    WEND
                                                                    %>      
                                                                                </div>
                                                                            </div>
                                                                </div>
                                                            </div>
                                                <%
                                                        aa = aa + 1
                                                        aaa = aaa + 1
                                                        ativo = ""
                                                        rsquestionarios.movenext 
                                                        WEND 
                                                        rsquestionarios.movefirst
                                                %>
                                            </div>
                                        </div>

                                        <input type="hidden" name="idnp" value="<%=Request.QueryString("id")%>">
                                        
                                        <input class="btn btn-primary pull-right" onclick="verificar_campos()"  type="submit" style="margin-top: 2%" value="Salvar">
                                        <input hidden name="MM_Insert" value="questionarios">

                                        <script type="text/javascript">
                                            
                                            function verificar_campos()
                                            {
                                                // alert("Verifique se todos os campos do(s) questionário(s) foram preenchidos!");
                                            }

                                        </script>
                                    </form>
                                </div>    

                            </div>
                            <div class="modal-footer">
                                <button type="button" class="btn btn-default" data-dismiss="modal">Pular</button>
                            </div>
                        </div>
                    </div>
                </div>

        <%
            END IF
        %>

    <style>
        .espacamento {
            margin: 15px 0;
        }
    </style>

    <!-- Modal Compactação de arquivos -->
    <div id="modcompact" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-body">
                    

                    <div id="btn-compact">
                    
                    </div>

                </div>
               
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal -->
    <div id="modupload" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-body">
                    <form name="formarquivo" method="post" enctype="multipart/form-data" action="neg_documentos_upload.php">
                        <h1 style="margin-top: 0; margin-bottom: 45px;">Faça o upload de documentos</h1>

                        <%
                            i_doc = i_doc + 1
                        %>
                        <span class="label label-danger">Número maxímo de arquivos enviados por vez: 10</span><br><br>
                        <strong>Selecione o  Arquivo:</strong>
                        <input onChange="verifica_arquivo(<%=i_doc%>);valida_tamanho(<%=i_doc%>)" class="form-control" name="upload[]" id="file<%=i_doc%>" type="file" multiple style="margin: 0px auto 0 0;" required>

                        <script type="text/javascript">
                            
                            function valida_tamanho(x)
                            {
                                if ($("#file"+x)[0].files.length > 10) {
                                    alert("Não é permitido fazer upload multíplo de mais de 10 arquivos");
                                    $("#file"+x).val("");
                                } 
                            }

                        </script>
                       
                        
                        <div class="espacamento">
                            <strong>Categoria:</strong>
                            <select name="id_categoria" class="form-control" required>
                                <option value="">Selecione uma categoria</option>
                                <% 
                                    SET rscat = conexao.execute("SELECT idndc, categoria FROM neg_doc_categoria WHERE status = 'Ativo'")
                                    WHILE NOT rscat.EOF
                                %>
                                    <option value="<%=rscat("idndc")%>"><%=rscat("categoria")%></option>
                                <% 
                                    rscat.movenext()
                                    WEND
                                    rscat.movefirst
                                %>
                            </select>
                        </div>

                        <div class="espacamento">
                            <strong>Título:</strong>
                            <input type="text" class="form-control" name="titulo" required>
                        </div>

                        <input name="idnp" type="hidden" value="<%=Request.QueryString("id")%>">

                        <input class="espacamento btn btn-success center-block" name="Salvar" type="submit" id="Salvar" value="Enviar">

                        <p>* Enviar apenas documentos em <strong>PDF</strong>, <strong>EXCEL</strong> ou <strong>JPG/PNG</strong></p>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal -->
    <div id="modvisu" class="modal fade" role="dialog">
        <div class="modal-dialog" style="width: 70%;">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Visualização de Arquivo</h4>
                </div>
                <div id="visualizar_arq" class="modal-body">
                    
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal -->
    <div id="modaprovar" class="modal fade" role="dialog">
        <div class="modal-dialog">

            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Aprovação de Prospecter</h4>
                </div>
                <div class="modal-body">
                    <form method="POST" action="neg_interacao_novo.asp?id=<%=request.querystring("id")%>">
                        <label>Observação</label>
                        <textarea class="form-control" name="observacao" rows="7"></textarea>
                        <br>

                        <hr>

                        <strong>Disparar processo:</strong>
                        <select name="processo" class="form-control">
                            <option value="">Não disparar</option>
                            <% 
                                SET rsprocessos = conexao.execute("SELECT idprocesso, processo FROM processos WHERE status = 'Ativo' ORDER BY processo ASC")

                                WHILE NOT rsprocessos.EOF 
                            %>
                                    <option value="<%=rsprocessos("idprocesso")%>"><%=rsprocessos("processo")%></option>
                            <% 
                                rsprocessos.movenext()
                                WEND
                            %>
                        </select> 
                        <br>

                        <strong>Vincular à um projeto:</strong> <br>
                        <select name="projeto" class="form-control"> 
                            <option value="">Selecione</option>
                            <%
                                SET rsprojetos = conexao.execute("SELECT idprojeto, titulo FROM projetos WHERE status = 'Ativo'")

                                WHILE NOT rsprojetos.EOF 
                            %>
                                    <option value="<%=rsprojetos("idprojeto")%>"><%=rsprojetos("titulo")%></option>
                            <%
                                rsprojetos.movenext 
                                WEND 
                            %>
                        </select> 
                        <br>

                        <strong>Data Prevista:</strong>
                        <input name="data_prevista" type="text" class="form-control datasemhora" maxlength="35">
                        <br>

                        <strong>Descrição:</strong>
                        <textarea class="form-control" rows="5" name="descricao_p"></textarea>
                        <br><br>

                        <input type="submit" class="btn btn-primary" value="Salvar">

                        <input type="hidden" name="aprovar" value="sim">
                        <input type="hidden" name="id" value="<%=Request.QueryString("id")%>">
                    </form>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal -->
    <div id="vincular_sistema" class="modal fade" role="dialog">
        <div class="modal-dialog">
            <!-- Modal content-->
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h4 class="modal-title">Vincular <%=Application("nome_cliente")%></h4>
                </div>
                <%
                    SET rsparametros = conexao.execute("SELECT dominio, nome_bandeira FROM parametros")
                %>
                <div class="modal-body">

                    <form ACTION="<%=MM_editAction%>" METHOD="POST" name="formvincular" id="formvincular" class="wizard-big mod_lic">
                        <h1>Dados Cadastrais</h1>
                        <fieldset>
                            <div class="row">
                                <div class="col-lg-6">
                                    <strong>Nome / Razão Social</strong>
                                    <input autocomplete="OFF" name="licenciado" class="form-control" required type="text" id="licenciado" maxlength="150" value="<%=rsemp("empresa")%>">
                                    
                                    <strong>Apelido / Fantasia:</strong>
                                    <input name="fantasia" class="form-control" type="text" id="fantasia" maxlength="45" required value="<%=rsemp("empresa")%>">

                                    <strong>Tipo de Cadastro:</strong>
                                    <div class="form-control" style="border: none !important; display: flex; justify-content: space-between;padding: 0px;margin-bottom: 0px">
                                        <div class="col-lg-6" style="padding-left: 0;">
                                            <input name="tipo_licenciado" style="float: left;margin-right: 3%" type="radio" value="j" onClick="mascara_cpf_cnpj_crm('juridica');" checked> Jurídica
                                        </div>
                                        <div class="col-lg-6" style="padding-left: 0;">
                                            <input name="tipo_licenciado" style="float: left;margin-right: 3%" type="radio" value="f" onClick="mascara_cpf_cnpj_crm('fisica');"> Física 
                                        </div>
                                    </div>

                                    <div id="retorno_cpfcnpj_crm">
                                        <strong>CPF / CNPJ:</strong>
                                        <div id="exibe_cnpj_crm">
                                            <small>-</small>
                                            <input type="hidden" id="cnpj_status_crm" value="nao">
                                            <input type="hidden" id="tipo_crm" value="cnpj">
                                            <input onBlur="verifica_cnpj_crm(this.value)" name="cpf_cnpj" placeholder="00.000.000/0000-00" class="form-control" onKeyPress="MascaraCNPJ(formvincular.cnpj);" maxlength="18" type="text" id="cnpj" required>
                                        </div>  
                                    </div>

                                    <strong>RG / Inscrição Estadual:</strong>
                                    <input name="ie" class="form-control" type="text" id="ie" maxlength="15">

                                    <strong>Telefone:</strong>
                                    <input name="telefone" placeholder="(00) 0000-0000" class="form-control" onKeyPress="MascaraTelefone(formvincular.telefone);" maxlength="14" type="text" id="telefone" value="<%=rsemp("telefone")%>">

                                    <strong>Celular:</strong>
                                    <input name="celular" placeholder="(00) 9 0000-0000" class="form-control" onKeyPress="MascaraCelular(formvincular.celular);" maxlength="16" type="text" id="celular" value="<%=rsemp("telefone")%>">
                                    
                                    <strong><%=rsparametros("nome_bandeira")%>:</strong>
                                    <select required name="idband" id="idband" class="form-control">
                                        <option value="">Selecione</option>
                                        <%
                                            set rsbandeiras = conexao.execute("SELECT * FROM bandeiras WHERE status_band = 'Ativo'")

                                            WHILE NOT rsbandeiras.EOF
                                        %>
                                            <option value="<%=rsbandeiras("idbandeira")%>"><%=rsbandeiras("nome")%></option>
                                        <%
                                            rsbandeiras.movenext
                                            WEND
                                        %>
                                    </select>
                                </div>

                                <div class="col-lg-6">  
                                    <% 
                                        variavel = rsdominio("dominio") 
                                        valor = "adapt"
                                        IF (instr(variavel, valor) > 0) THEN
                                    %>
                                            <strong>Quantidade de E-mails:</strong>
                                            <input type="number" name="qtd_emails" class="form-control">
                                    <%  END IF %>

                                    <strong>E-mail:</strong>
                                    <div class="tooltip-demo pull-right" style="margin-right: 240px;">
                                        <i title="Separe cada e-mail por um ponto e vírgula. EX: email@email.com.br; email2@email2.com.br" data-toggle="tooltip" data-placement="right" class="fa fa-question-circle"></i>
                                    </div> 
                                    <input name="email" class="form-control" type="text" id="email" size="40" maxlength="250" required value="<%=rsemp("email")%>">

                                    <%
                                        SELECT CASE rsemp("uf")
                                            CASE "AC" ac = "selected"
                                            CASE "AL" al = "selected"
                                            CASE "AM" am = "selected"
                                            CASE "AP" ap = "selected"
                                            CASE "BA" ba = "selected"
                                            CASE "CE" c = "selected"
                                            CASE "DF" df = "selected"
                                            CASE "GO" go = "selected"
                                            CASE "MA" ma = "selected"
                                            CASE "MG" mg = "selected"
                                            CASE "MT" mt = "selected"
                                            CASE "MS" ms = "selected"
                                            CASE "PA" pa = "selected"
                                            CASE "PB" pb = "selected"
                                            CASE "PE" pe = "selected"
                                            CASE "PI" pi = "selected"
                                            CASE "PR" pr = "selected"
                                            CASE "RJ" rj = "selected"
                                            CASE "RN" rn = "selected"
                                            CASE "RR" rr = "selected"
                                            CASE "RO" ro = "selected"
                                            CASE "RS" rs = "selected"
                                            CASE "SC" sc = "selected"
                                            CASE "SE" se = "selected"
                                            CASE "SP" sp = "selected"
                                            CASE "TO" toc = "selected"
                                        END SELECT 
                                    %>

                                    <strong>Estado:</strong>
                                    <select name="estado" class="form-control" id="estado" required>
                                        <option value="">Selecione</option>
                                        <option value="AC" <%=ac%>>Acre</option>
                                        <option value="AL" <%=al%>>Alagoas</option>
                                        <option value="AM" <%=am%>>Amazonas</option>
                                        <option value="AP" <%=ap%>>Amapa</option>
                                        <option value="BA" <%=ba%>>Bahia</option>
                                        <option value="CE" <%=ce%>>Ce&aacute;ra</option>
                                        <option value="DF" <%=df%>>Distrito Federal</option>
                                        <option value="ES" <%=es%>>Espirito Santo</option>
                                        <option value="GO" <%=go%>>Goi&aacute;s</option>
                                        <option value="MA" <%=ma%>>Maranh&atilde;o</option>
                                        <option value="MG" <%=mg%>>Minas Gerais</option>
                                        <option value="MT" <%=mt%>>Mato Grosso</option>
                                        <option value="MS" <%=ms%>>Mato Grosso do Sul</option>
                                        <option value="PA" <%=pa%>>Par&aacute;</option>
                                        <option value="PB" <%=pb%>>Paraiba</option>
                                        <option value="PE" <%=pe%>>Pernambuco</option>
                                        <option value="PI" <%=pi%>>Piau&iacute;</option>
                                        <option value="PR" <%=pr%>>Paran&aacute;</option>
                                        <option value="RJ" <%=rj%>>Rio de Janeiro</option>
                                        <option value="RN" <%=rn%>>Rio Grande do Norte</option>
                                        <option value="RR" <%=rr%>>Roraima</option>
                                        <option value="RO" <%=ro%>>Rondônia</option>
                                        <option value="RS" <%=rs%>>Rio Grande do Sul</option>
                                        <option value="SC" <%=sc%>>Santa Catarina</option>
                                        <option value="SE" <%=se%>>Sergipe</option>
                                        <option value="SP" <%=sp%>>S&atilde;o Paulo</option>
                                        <option value="TO" <%=toc%>>Tocantis</option>
                                    </select>

                                    <strong>Cidade:</strong>
                                    <br>
                                    <small>-</small>
                                    <input autocomplete="OFF" name="cidade" class="form-control typeahead_1" type="text" id="cidade" maxlength="30" required value="<%=rsemp("cidade")%>">

                                    <strong>Endereço:</strong>
                                    <input name="endereco" class="form-control" type="text" id="endereco" maxlength="150" required>

                                    <strong>Número:</strong>
                                    <input name="numero" class="form-control" type="text" required>

                                    <strong>Bairro:</strong>
                                    <input name="bairro" class="form-control" type="text" id="bairro" maxlength="30" required>

                                    <strong>Complemento:</strong>
                                    <input name="complemento" class="form-control" type="text" id="complemento" maxlength="30">

                                    <strong>Cep:</strong>
                                    <input name="cep" placeholder="00000-000" class="form-control" onKeyPress="MascaraCep(formvincular.cep);" maxlength="9" type="text" id="cep" required>

                                    
                                </div>
                            </div>

                        </fieldset>

                        <!-- <
                            SET rsbanco = conexao.execute("SELECT idbanco, banco FROM bancos WHERE principal = 1")

                            IF NOT rsbanco.EOF then
                        %>

                        <h1>Dados Bancários</h1>
                        <fieldset>
                            <div class="row">
                                <div class="col-sm-12"  style="text-align: left !important;">
                                    <strong>Banco:</strong>
                                    <select required name="idbanco" id="idbanco" class="form-control">
                                    < 
                                        While not rsbanco.EOF
                                    %>
                                        <option value="<=rsbanco("idbanco")%>"><=rsbanco("banco")%></option>
                                    <
                                        rsbanco.movenext
                                        Wend
                                    %>
                                    </select>
                                    <strong>Dia melhor Vencimento:</strong>
                                    <input required name="dia_fat" class="form-control" type="number" id="dia_fat" maxlength="2" min="1" max="31">
                                </div>
                            </div>
                        </fieldset>

                        <
                            END IF
                        %> -->

                        <h1>Dados de Acesso</h1>
                        <fieldset>
                            <%
                                dia = day(date)
                                IF dia < 10 THEN dia = "0"&dia 
                                mes = month(date)
                                IF mes < 10 THEN mes = "0"&mes 
                                data_atual = dia&"/"&mes&"/"&year(date)
                            %>
                            <strong>Data Início:</strong>
                            <input name="datainicio" class="form-control" placeholder="00/00/0000" type="text" id="datainicio" onKeyPress="MascaraData(formvincular.datainicio);" value="<%=data_atual%>" maxlength="10" onBlur="validadata(this, 0);" required>
                            
                            <!-- <strong>Login:</strong>
                            <input autocomplete="OFF" name="login" class="form-control" type="text" id="login" size="20" maxlength="15" onBlur="mandalogin(this.value, 'licenciados', 0);" required>

                            <div style="min-height: 15px;"></div>
                            <div id="exibe_login_lic_crm">
                                <input class="form-control" name="avisologin" class="form-control btn btn-success cor_branca" type="text" disabled="disabled" id="avisologin" value="Login aceito. Continue">
                                <input type="hidden" name="avisologin2" value="Login aceito. Continue">
                            </div>

                            <strong>Senha:</strong>
                            <input name="senha" class="form-control" type="password" id="senha" size="15" maxlength="10" required> -->

                            <input type="hidden" name="MM_insert" value="forminserir_cliente">
                            <input type="hidden" name="vincular_crm" value="ok">
                            <input type="hidden" name="id_neg" value="<%=request.querystring("id")%>">
                            <!-- <input name="registrar" class="btn btn-primary pull-right" type="submit" id="registrar" value="Salvar"> -->
                        </fieldset>

                        <h1>Responsáveis</h1>
                        <fieldset>
                            <%
                                SET rssetores = conexao.execute("SELECT idsetor, setor FROM setores WHERE equipe = 'Sim'")
                                
                                WHILE NOT rssetores.EOF
                            %>
                                    <br>
                                    <strong><%=rssetores("setor")%>:</strong>
                                    <select name="setor_<%=rssetores("setor")%>" class="form-control">
                                        <option value="">Selecione</option>
                                        <% 
                                            SET rsusuarios = conexao.execute("SELECT * FROM usuarios WHERE id_setor = "&rssetores("idsetor")&" and status = 'Ativo' ORDER BY usuario ASC")

                                            WHILE NOT rsusuarios.EOF 
                                        %>
                                                    <option value="<%=rsusuarios("idusuario")%>"><%=rsusuarios("usuario")%></option>
                                        <% 
                                            rsusuarios.movenext()
                                            WEND
                                        %>
                                    </select>   
                            <%
                                rssetores.movenext 
                                WEND
                            %>      
                        </fieldset>

                        <h1>Ações</h1>
                        <fieldset>
                            <strong>Disparar processo:</strong>
                            <select name="processo" class="form-control">
                                <option value="">Não disparar</option>
                                <% 
                                    SET rsprocessos = conexao.execute("SELECT idprocesso, processo FROM processos WHERE status = 'Ativo' ORDER BY processo ASC")

                                    WHILE NOT rsprocessos.EOF 
                                %>
                                        <option value="<%=rsprocessos("idprocesso")%>"><%=rsprocessos("processo")%></option>
                                <% 
                                    rsprocessos.movenext()
                                    WEND
                                %>
                            </select> 

                            <strong>Vincular à um projeto:</strong> <br>
                            <select name="projeto" class="form-control"> 
                                <option value="">Selecione</option>
                                <%
                                    SET rsprojetos = conexao.execute("SELECT idprojeto, titulo FROM projetos")

                                    WHILE NOT rsprojetos.EOF 
                                %>
                                        <option value="<%=rsprojetos("idprojeto")%>"><%=rsprojetos("titulo")%></option>
                                <%
                                    rsprojetos.movenext 
                                    WEND 
                                %>
                            </select> 

                            <strong>Data Prevista:</strong>
                            <input name="data_prevista" type="text" class="form-control datasemhora" maxlength="35">

                            <strong>Descrição:</strong>
                            <textarea class="form-control" rows="5" name="descricao_p"></textarea>
                        </fieldset>

                        <h1>Redes Sociais</h1>
                        <fieldset>

                            <strong>Site:</strong>
                            <input name="site" type="text" class="form-control" maxlength="35">

                            <strong>Instagram:</strong>
                            <input name="instagram" type="text" class="form-control" maxlength="35">

                            <strong>Facebook:</strong>
                            <input name="facebook" type="text" class="form-control" maxlength="35">
                        </fieldset>
                    </form>           

                </div>

                <div class="modal-footer">
                    <!-- <div onClick="verifica_cnpj_crm" id="div_verifica_cnpj_crm"></div> -->
                    <button type="button" class="btn btn-basic" data-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>

    <style type="text/css">
        <%
            SET usuario = conexao.execute("SELECT fundo_wp FROM usuarios WHERE idusuario = "&Request.Cookies("grpro")("idusuario"))
            foto = usuario("fundo_wp")
            IF isnull(foto) OR foto = "fundo_padrao.jpg" THEN
                img_f = "arquivos/wpp/fundo_padrao.jpg"
            ELSE
                img_f = "arquivos/wpp/"&Request.Cookies("grpro")("idusuario")&"/"&foto
            END IF
        %>

        #mensagens-chat #mensagens, #mensagensp
        {
            background-image: url(<%=img_f%>);
        }
    </style>

    <script>
        // Script do step
        $(document).ready(function(){
            $("#formvincular").steps({
                bodyTag: "fieldset",
                onStepChanging: function (event, currentIndex, newIndex)
                {
                    // Always allow going backward even if the current step contains invalid fields!
                    if (currentIndex > newIndex)
                    {
                        return true;
                    }

                    var form = $(this);

                    // Clean up if user went backward before
                    if (currentIndex < newIndex)
                    {
                        // To remove error styles
                        $(".body:eq(" + newIndex + ") label.error", form).remove();
                        $(".body:eq(" + newIndex + ") .error", form).removeClass("error");
                    }

                    // Start validation; Prevent going forward if false
                    return form.valid();
                },
                onStepChanged: function (event, currentIndex, priorIndex)
                {
                    // Suppress (skip) "Warning" step if the user is old enough.
                    if (currentIndex === 2)
                    {
                        $(this).steps("next");
                    }

                    // Suppress (skip) "Warning" step if the user is old enough and wants to the previous step.
                    if (currentIndex === 2 && priorIndex === 3)
                    {
                        $(this).steps("previous");
                    }
                },
                onFinishing: function (event, currentIndex)
                {
                    var form = $(this);

                    // Start validation; Prevent form submission if false
                    return form.valid();
                },
                onFinished: function (event, currentIndex)
                {
                    var form = $(this);

                    // Submit form input
                    form.submit();
                }
            })
            .validate({
                        errorPlacement: function (error, element)
                        {
                            element.before(error);
                        },
                        rules: {
                            confirm: {
                                equalTo: "#password"
                            }
                        }
                    });
        });

        function mandalogin(login, tabela, id){
            ajaxGo({ url:"retorna_login.asp?login="+login+"&tabela="+tabela+"&id="+id, elem_return: document.getElementById("exibe_login_lic_crm") });
        }

        function mascara_cpf_cnpj_crm(estado){
            ajaxGo({ url:"Ajax/mascara_cpf_cnpj_crm_licenciado.asp?estado="+estado, elem_return: document.getElementById("retorno_cpfcnpj_crm") });
        }

        function verifica_cnpj_crm(cnpj){
        if (typeof cnpj != "undefined") {
    
            var tipo = document.getElementById("tipo_crm").value;
            var nao = "nao"
        
                ajaxGo({ url:"Ajax/verifica_cnpj_crm.asp?cnpj="+cnpj+"&tipo="+tipo, elem_return: document.getElementById("exibe_cnpj_crm") });
            
        }
        
        }
    </script>


    <% if Request.Querystring("msg") = "histok" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 2000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Histórico inserido com sucesso!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "remove-histo" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 2000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Histórico deletado com sucesso!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "edit-histo" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 2000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Histórico atualizado com sucesso!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "funilok" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 2000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Funil alterado com sucesso!")
            })
        </script>
    <%end if%>

    <% if Request.Querystring("msg") = "editarok" then %>
        <script type="text/javascript">
            $(document).ready(function(){
                // console.log(toastr)
                toastr.options.progressBar = true;
                toastr.options.timeOut = 2000;
                toastr.options.extendedTimeOut = 6000;
                toastr.success("Dados editados com sucesso!")
            })
        </script>
    <%end if%>

    <link rel="stylesheet" type="text/css" href="css/emojionearea.min.css">
    <script src="js/emojionearea.js"></script> 

    <%
        modal_documento = false

        IF Request.QueryString("act") = "doc_upload" THEN 

            SET rsdocumentos = conexao.execute("SELECT categoria FROM neg_doc_categoria WHERE status = 'Ativo' and obrigatorio = 'Sim'")

            IF NOT rsdocumentos.EOF THEN 
                modal_documento = true
    %>
                <script>
                    $("#moduploadetapas").modal("show");
                </script>
    <%
            ELSE 

            IF rsemp("id_perfil") <> "" and rsemp("id_perfil") <> "0" then

                SET rsquestionarios = conexao.execute("SELECT * FROM neg_quest_perfil as p INNER JOIN neg_questionarios as q ON p.id_questionario = q.idnq WHERE q.status = 'Ativo'  and  id_perfil = "&rsemp("id_perfil")&" order by q.ordem asc")
            else

                SET rsquestionarios = conexao.execute("SELECT idnq, titulo FROM neg_questionarios WHERE status = 'Ativo' ORDER BY ordem ASC")
            end if

                IF NOT rsquestionarios.EOF THEN
    %>
                <script>
                    $("#mod_quest").modal("show");
                </script>
    <%
                END IF 
            END IF 
        END IF 
    %>

    <%
        SET verifica_campos_obrigatorio = conexao.execute("SELECT DISTINCT q.idnq, q.titulo FROM neg_quest_perfil as qp INNER JOIN neg_questionarios as q ON q.idnq = qp.id_questionario INNER JOIN neg_perfil as p ON p.idperfil = qp.id_perfil INNER JOIN neg_quest_campos nc ON nc.id_nq = q.idnq WHERE nc.status = 'Ativo' and nc.obrigatorio = 'Sim' and NOT EXISTS(SELECT * FROM neg_quest_respostas nr where nr.id_prospecter = "&rsemp("idnp")&" and nr.id_campo = nc.idcampo and nr.resposta <> '') and q.status = 'Ativo' and qp.id_perfil = "&rsemp("id_perfil")&" ORDER BY q.ordem ASC LIMIT 1")

        IF request.querystring("act") = "questionarios" or (NOT modal_documento and NOT verifica_campos_obrigatorio.EOF) THEN 
    %>
            <script>
                $("#mod_quest").modal("show");
            </script>
    <%
        END IF 
    %>

    <script>
        $('.dinheiro').mask('#.##0,00', {reverse: true});
        $('.cnpj').mask('00.000.000/0000-00', {placeholder: "00.000.000/0000-00"});
        $('.datasemhora').mask('00/00/0000', {placeholder: "00/00/0000"});
        $('.data_fat').mask('00', {placeholder: "12"});
        $('.hora').mask('00:00:00');
        $('.data_hora').mask('99/99/9999 00:00:00');
        $('.cep').mask('99999-999');
        $('.telefone_ddd').mask('(99) 99999-9999', {placeholder: "(00) 00000-0000"});
    </script>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/ekko-lightbox/5.3.0/ekko-lightbox.min.js"></script>

    <script>
        function verifica_arquivo(i) {
            var tamanhoArquivo = parseInt(document.getElementById("file"+i).files[0].size);
            if(tamanhoArquivo > 5242880)
            { 
                swal({
                    title: "SISTEMA",
                    text: "O arquivo que você está tentando subir excede o limite de upload (5 MB)."
                });
                $("#file"+i).val("");
            }
        }

        $(document).on('click', '[data-toggle="lightbox"]', function(event) {
            event.preventDefault();
            $(this).ekkoLightbox();
        });

        $(document).ready(function(){
            var alturaDiv1 = $('#div_1').height();

            $("#div_2").css("height", alturaDiv1 - 25);  
        });

        function doc_upload(){
            $("#modupload").modal("show");
        }

        $('.datasemhora').mask('00/00/0000', {placeholder: "00/00/0000"});
        $('.hora').mask('00:00', {placeholder: "00:00"});

        function marcar_neg(idni) {
            window.location.href = "neg_interacao_novo.asp?id=<%=Request.QueryString("id")%>&idni="+idni+"&act=marcar"; 
        }

        function marcar_neg_c(idni) {
            window.location.href = "neg_interacao_novo.asp?id=<%=Request.QueryString("id")%>&idni="+idni+"&act=marcar_c"; 
        }

        function aprovar() {
            $("#modaprovar").modal("show");
        }

        function ver_arq(arquivo, id, tipo, idta) {
            ajaxGo({ url:"Ajax/mostra_neg_doc.asp?arquivo="+arquivo+"&id="+id+"&tipo="+tipo+"&idta="+idta, elem_return: document.getElementById("visualizar_arq") });

            $("#modvisu").modal("show");
        }

        function deletar_arq(idnd, idnp) {
            swal({
                title: "Você tem certeza?",
                text: "Você não poderá mais recuperar este arquivo!",
                type: "warning",
                showCancelButton: true,
                confirmButtonColor: "#DD6B55",
                confirmButtonText: "Sim, desejo deletar o arquivo!",
                cancelButtonText: "Não, desejo manter o arquivo!",
                closeOnConfirm: false,
                closeOnCancel: false },
            function (isConfirm) {
                if (isConfirm) {
                    // swal("Deletado!", "Seu arquivo foi deletado!", "success");
                    window.location.assign("neg_interacao_novo.asp?act=excluir&idnd="+idnd+"&id="+idnp);
                } else {
                    swal("Cancelado", "Seu arquivo não foi deletado!", "error");
                }
            });
        };

        function editar_historico(idni, descricao, interacao)
        {
            document.getElementById("idinteracao").value=idni;
            document.getElementById("interacao_editar").value=interacao;

            if (descricao.match(/<br>/g) != null)
            {
                ocorrencias = descricao.match(/<br>/g).length;
            }
            else
            {
                ocorrencias = 0;
            }
            for (i = 0;i < ocorrencias; i++)
            {
                descricao = descricao.replace("<br>", "\n");
            }

            document.getElementById("descricao_editar").innerHTML=descricao;

        }

        function deleta_historico(idni)
        {
            swal({
                title: "Deletar Interação",
                text: "Deseja deletar essa interação?",
                type: "warning",
                showCancelButton: true,
                confirmButtonColor: "#DD6B55",
                confirmButtonText: "Sim",
                cancelButtonText: "Não",
                closeOnConfirm: false,
                closeOnCancel: false },
                function (isConfirm) {
                if (isConfirm) {
                    window.location.href = "neg_interacao_novo.asp?idhist="+idni+"&act=remover-histo&id=<%=request.querystring("id")%>"
                } else {
                    swal("Cancelado", "Ação cancelada!", "error");
                }
            });
        }

    </script>

    <script src="js/plugins/pdfjs/pdf.js"></script>

    <script src="assets/gallery/player.min.js"></script>
    <script src="assets/gallery/script.js"></script>

    <!-- Clock picker -->
    <script src="js/plugins/clockpicker/clockpicker.js"></script>

    <script type="text/javascript">
        $('.clockpicker').clockpicker();
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

        data_agenda

        var mem = $('#data_agenda .input-group.date').datepicker({
        todayBtn: "linked",
        keyboardNavigation: false,
        forceParse: false,
        calendarWeeks: true,
        autoclose: true,
        format: 'dd/mm/yyyy',
    });

    </script>

    <style type="text/css">
        .clockpicker-popover, .popover{
            z-index: 99999 !important
        }

    </style>

    <style type="text/css">
        
        #mensagens-internas .mensagem_enviada div div, #mensagens-internas .mensagem_enviada div p 
        {
            background: #dcf8c6 !important;
            border-bottom-right-radius: 0px !important;
            text-align: right;
        }

        #mensagens-internas .mensagem_enviada div
        {
            margin-left: auto !important;
        }

        #mensagens-internas .mensagem_recebida div div, #mensagens-internas .mensagem_recebida div p
        {
            border-bottom-left-radius: 0px !important;
        }

    </style>

    <script type="text/javascript">
        function mascara_cpf_cnpj_edit(estado, id){
            ajaxGo({ url:"Ajax/mascara_cpf_cnpj_prospecter_edit.asp?estado="+estado+"&id="+id, elem_return: document.getElementById("retorno_cpfcnpj_prospecter") });
            $("input[name='salvar']").prop("disabled", true);s
        }

        function verifica_cnpj_edit_prospecter(cnpj, id){

            $("input[name='salvar']").prop("disabled", false);
            var att_ver_cnpj = document.getElementById("cnpj_status").value;
            var tipo = document.getElementById("tipo").value;
            var nao = "nao"
            if (att_ver_cnpj == nao) {
                ajaxGo({ url:"Ajax/verifica_cnpj_edit_prospecter.asp?cnpj="+cnpj+"&tipo="+tipo+"&id="+id, elem_return: document.getElementById("exibe_cnpj_prospecter") });
            }
        }
    </script>

    <script type="text/javascript">
        function abrir_midia(link, tipo)
        {
            console.log(link);
            ajaxGo({ url:"Ajax/wpp_abrir_midia.asp?link="+link+"&tipo="+tipo, elem_return: document.getElementById("midia_msg") });
            $("#mod_midia").modal("show");
        }
    </script>

    <div id="mod_midia" class="modal fade" role="dialog">
        <div class="modal-dialog">
            <div class="modal-content" id="midia_msg">
            </div>
        </div>
    </div>


    <script type="text/javascript">
        
        function atualizar_viewer_arq(arquivo, id)
        {
            ajaxGo({ url:"Ajax/wpp_retorna_arquivo.php?id="+id+"&arquivo="+arquivo, elem_return: document.getElementById("visualizacao_arquivo_wpp") });
        }

    </script>

    <script type="text/javascript">
        
        function deletar_arq_wpp(arquivo, id, instancia, numero)
        {
            // alert(arquivo+" "+id);
            if(confirm("Deseja realmente excluir o arquivo "+arquivo+" ?"))
            {
                ajaxGo({ url:"Ajax/deletar_arquivo_wpp.php?id="+id+"&arquivo="+arquivo, elem_return: document.getElementById("aviso_deletar_arq_wpp") });
                att_galeria(instancia, numero);
            }
        }

    </script>

    


   

    <script type="text/javascript">
            
        function enviar_doc(numero, instancia)
        {

            ajaxGo({ url:"lista_galeria_envio.asp?instacia="+instancia+"&id=<%=Request.Cookies("grpro")("idusuario")%>&numero="+numero, elem_return: document.getElementById("galeria") });


            $("#anexo_wpp").modal("show");
        }

        function att_galeria(numero, instancia)
        {

            setTimeout(function(){ ajaxGo({ url:"lista_galeria_envio.asp?&instacia="+instancia+"&id=<%=Request.Cookies("grpro")("idusuario")%>&numero="+numero, elem_return: document.getElementById("galeria") }); }, 2000);
            
        }
    </script>

    <script type="text/javascript">
        function gravar_audio(numero, dispo)
        {
            console.log('recorder.asp?numero='+numero+"&dispo="+dispo);
            window.open('recorder.asp?numero='+numero+"&dispo="+dispo, '_blank', 'location=yes,height=450,width=400,scrollbars=yes,status=yes'); 
        }
    </script>

    <!-- Modal -->
    <div id="visualizar_responsaveis" class="modal fade" role="dialog">
        <div class="modal-dialog" style="width: 35%;">

            <!-- Modal content-->
            <div class="modal-content" >
            <div class="modal-header">
                <h4 class="modal-title">Usuários responsáveis pela negociação</h4>
            </div>
            <div class="col-lg-12 modal-body" style="border-bottom: 2px solid #f5f5f5; padding: 15px 15px 15px 15px;">
                <div style="display: flex !important; justify-content: center !important;">
                <%
                    IF NOT rsusuario.EOF THEN
                        WHILE NOT rsusuario.EOF
                            IF IsNull(rsusuario("logo")) OR rsusuario("logo") = "" THEN
                                logo_usuario = "semfoto.png"
                            ELSE 
                                logo_usuario = rsusuario("logo")
                            END IF
                %>
                            <div>
                                <a class="tooltip-demo" href="usuarios.asp" style="cursor: auto; width: 30px; height: 30px; position: relative; display: block; float: right; margin-right: 8px;">
                                    <img style="width: 100%; height: 100%;"  data-toggle="tooltip" data-placement="bottom" title="<%=rsusuario("usuario")%> (<%=rsusuario("setor")%>)" alt="image" class="img-circle" src="../imagens/logo_usuario/<%=logo_usuario%>">
                                </a>
                            </div>
                <%
                            rsusuario.movenext
                        WEND
                    END IF 
                %>
                </div>
            </div>
            <div class="modal-footer" style=" margin-top: 0px;">
                <button type="button" class="btn btn-default" data-dismiss="modal" style="margin-top: 10px;">Fechar</button>
            </div>
            </div>

        </div>
    </div>

    <div class="modal inmodal" id="ver_fundo" tabindex="-1" role="dialog" aria-hidden="true">
        <div class="modal-dialog">
            <div id="div_mod" class="modal-content" style="padding: 2%">

                <form action="upload_foto_fundo_wpp.php" method="post" enctype="multipart/form-data">
                                            
                    <input type="hidden" name="editar-foto" value="ok">

                    <input type="hidden" name="id" value="<%=Request.Cookies("grpro")("idusuario")%>">

                    <h2 style="text-align: center;margin-top: 0px">Personalizar Whatsapp</h2>

                    <div style="text-align: center;" id="mensagensp">

                        <div id="mensagens-internasp">
                            <div class="mensagem_recebida">
                                <div style="padding: 2%">
                                    <div style="margin-bottom: 0px;padding: 1.5% 3%;background: #fff;border-radius: 10px">
                                        
                                        <p style="margin-bottom: 0px;white-space: pre-line;text-align: left;">Hello how are you?</p>

                                        <p style="margin-bottom: 0px;text-align: right;color: #1cc09f;margin-top: 1%">08/03/2022 15:31:56</p>
                                    </div>
                                </div>
                            </div>

                            <div class="mensagem_enviada">
                                <div style="padding: 2%">
                                    <div style="margin-bottom: 0px;padding: 1.5% 3%;background: #fff;border-radius: 10px">
                                        
                                        <p style="margin-bottom: 0px;white-space: pre-line">Yes, and you?</p>

                                        <p style="margin-bottom: 0px;text-align: right;color: #1cc09f;margin-top: 1%">08/03/2022 15:45:56</p>
                                    </div>
                                </div>
                            </div>

                            <div class="mensagem_recebida">
                                <div style="padding: 2%">
                                    <div style="margin-bottom: 0px;padding: 1.5% 3%;background: #fff;border-radius: 10px">
                                        
                                        <p style="margin-bottom: 0px;white-space: pre-line;text-align: left;">all yes</p>

                                        <p style="margin-bottom: 0px;text-align: right;color: #1cc09f;margin-top: 1%">08/03/2022 15:48:35</p>
                                    </div>
                                </div>
                            </div>

                            <div class="mensagem_enviada">
                                <div style="padding: 2%">
                                    <div style="margin-bottom: 0px;padding: 1.5% 3%;background: #fff;border-radius: 10px">
                                        
                                        <p style="margin-bottom: 0px;white-space: pre-line">Have a good day</p>

                                        <p style="margin-bottom: 0px;text-align: right;color: #1cc09f;margin-top: 1%">08/03/2022 16:34:58</p>
                                    </div>
                                </div>
                            </div>

                            <div class="mensagem_enviada">
                                <div style="padding: 2%">
                                    <div style="margin-bottom: 0px;padding: 1.5% 3%;background: #fff;border-radius: 10px">
                                        
                                        <p style="margin-bottom: 0px;white-space: pre-line">good work!</p>

                                        <p style="margin-bottom: 0px;text-align: right;color: #1cc09f;margin-top: 1%">08/03/2022 16:54:23</p>
                                    </div>
                                </div>
                            </div>

                            <div class="mensagem_recebida">
                                <div style="padding: 2%">
                                    <div style="margin-bottom: 0px;padding: 1.5% 3%;background: #fff;border-radius: 10px">
                                        
                                        <p style="margin-bottom: 0px;white-space: pre-line;text-align: left;">thanks</p>

                                        <p style="margin-bottom: 0px;text-align: right;color: #1cc09f;margin-top: 1%">08/03/2022 17:00:34</p>
                                    </div>
                                </div>
                            </div>
                        </div>

                    </div>

                    <!-- <img id="preview" src="<%=img_f%>" style="height: 345px;width: 100%;" alt="" class="rounded-circle img-thumbnail"> -->

                    <div class="fallback mt-3" style="margin-top: 3%">
                        <input accept="image/*" id="img-input" name="file" type="file" />
                    </div>

                    <div class="mt-3" style="margin-top: 2%">
                        <button type="submit" class="btn btn-success">Salvar Imagem</button>
                    </div>

                </form>

            </div>
        </div>
    </div>

    <script type="text/javascript">

        $(document).ready(function() {
            $(".select2_demo_2").select2()({
                placeholder: "Selecione",
                allowClear: true
            });
        } );

        function readImage() {
            if (this.files && this.files[0]) {
                var file = new FileReader();
                file.onload = function(e) {
                    document.getElementById("mensagensp").style.background = "url('"+e.target.result+"')";
                };       
                file.readAsDataURL(this.files[0]);
            }
        }

        document.getElementById("img-input").addEventListener("change", readImage, false);
    </script>

     <script type="text/javascript">
        function fechar_resposta()
        {
            document.getElementById("bloco_resposta_msg").style.display = "none";
            ajaxGo({ url:"Ajax/vazio.asp", elem_return: document.getElementById("bloco_resposta_msg") });
        }

        function responder_wpp_msg(codigo_resposta)
        {
            if (codigo_resposta != "")
            {
                ajaxGo({ url:"Ajax/resposta_msg_wpp.asp?codigo="+codigo_resposta, elem_return: document.getElementById("bloco_resposta_msg") });
                document.getElementById("bloco_resposta_msg").style.display = "block";
            }
        }
    </script>

    <style type="text/css">
        .emojionearea-editor{
            min-height: 40px !important;
            max-height: 80px !important;
        }
    </style>

    <script type="text/javascript">
        function abrir_audio(link)
        {
            // document.getElementById("src_audi").src = link;
            // $("#mod_midia_audio").modal("show");
            window.open(link,'_blank')
        }
    </script>
</body>

</html>