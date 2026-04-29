// ============================================================
// SISTEMA DE ATAS DE REUNIÃO — Orgânico do Chico  v2
// ============================================================

const SPREADSHEET_ID  = "1lqtTSH238fzG5gJ_31VeJ61dbLK1i0UXEMioJBbVIyU";
const DRIVE_FOLDER_ID = "16WDblZKEWRpuApX82qC9E2LbGQ1mlmch";
const EVOLUTION_URL   = "https://SUA-EVOLUTION-API.com";
const EVOLUTION_KEY   = "SUA_API_KEY";
const EVOLUTION_INST  = "SUA_INSTANCIA";


// ============================================================
// doGet — serve dados para o formulário e o dashboard
// ============================================================
function doGet(e) {
  const action = e?.parameter?.action || "";

  if (action === "participantes") {
    const area = e.parameter.area || "";
    return _json(buscarParticipantesArea(area));
  }

  if (action === "dashboard") {
    return _json(dadosDashboard());
  }

  if (action === "configuracoes") {
    return _json(buscarConfiguracoes());
  }

  return _json({ status: "ok" });
}


// ============================================================
// doPost — recebe o formulário de ata
// ============================================================
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);

    if (dados.telefone && dados.resposta) {
      processarRespostaWhatsApp(dados.telefone, dados.resposta);
      return _ok("resposta processada");
    }

    const linkDrive = gerarAtaNoDrive(dados);
    registrarNaPlanilha(dados, linkDrive);
    enviarAtaWhatsApp(dados, linkDrive);
    criarProximaReuniaoCalendario(dados, linkDrive);

    return _ok(linkDrive);

  } catch (err) {
    return _erro(err.message);
  }
}


// ============================================================
// PARTICIPANTES — lê da aba "Participantes"
// ============================================================
function buscarParticipantesArea(area) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Participantes");
  if (!aba) return [];

  return aba.getDataRange().getValues()
    .slice(1)
    .filter(r => r[0] === area && r[1])
    .map(r => ({ nome: String(r[1]).trim(), telefone: String(r[2] || "").trim(), email: String(r[3] || "").trim() }));
}


// ============================================================
// CONFIGURAÇÕES — lê da aba "Configurações"
// ============================================================
function buscarConfiguracoes() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Configurações");
  if (!aba) return [];

  return aba.getDataRange().getValues()
    .slice(1)
    .map(r => ({
      area:       String(r[0]).trim(),
      frequencia: String(r[1]).trim(),
      dia:        String(r[2]).trim(),
      horario:    String(r[3]).trim(),
      ativo:      String(r[4]).trim() === "Sim",
    }));
}


// ============================================================
// DASHBOARD — dados de reuniões e ações do mês atual
// ============================================================
function dadosDashboard() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAtas = ss.getSheetByName("Atas");
  const abaAco  = ss.getSheetByName("Ações");
  const configs = buscarConfiguracoes();

  const hoje      = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const fimMes    = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);

  // Reuniões do mês
  const todasAtas = abaAtas ? abaAtas.getDataRange().getValues().slice(1) : [];
  const atasMes   = todasAtas.filter(r => {
    if (!r[2]) return false;
    const d = new Date(r[2] + "T12:00:00");
    return d >= inicioMes && d <= fimMes;
  });

  // Quantas semanas tem no mês
  const diasNoMes = fimMes.getDate();
  const semanas   = Math.round(diasNoMes / 7);

  // Monta dados por área
  const areas = configs.map(c => {
    const atasDaArea = atasMes.filter(r => r[1] === c.area);
    let esperadas = 0;
    if (c.frequencia === "Semanal")    esperadas = semanas;
    if (c.frequencia === "Quinzenal")  esperadas = Math.ceil(semanas / 2);
    if (c.frequencia === "Mensal")     esperadas = 1;

    const realizadas = atasDaArea.length;
    let status = "sem_periodicidade";
    if (esperadas > 0) {
      if (realizadas >= esperadas)     status = "em_dia";
      else if (realizadas > 0)         status = "parcial";
      else                             status = "atrasada";
    }

    return {
      area:       c.area,
      frequencia: c.frequencia,
      esperadas,
      realizadas,
      status,
      atas: atasDaArea.map(r => ({
        id:            String(r[0]),
        data:          String(r[2]),
        participantes: String(r[5]),
        qtdAcoes:      Number(r[6]) || 0,
        link:          String(r[7]),
      })).sort((a, b) => b.data.localeCompare(a.data)),
    };
  });

  // Resumo de ações
  const acoes     = abaAco ? abaAco.getDataRange().getValues().slice(1) : [];
  const pendentes = acoes.filter(r => r[7] === "Pendente").length;
  const concluidas= acoes.filter(r => r[7] === "Concluída").length;
  const atrasadas = acoes.filter(r => {
    if (r[7] !== "Pendente" || !r[6]) return false;
    const prazo = new Date(r[6] + "T12:00:00"); prazo.setHours(0,0,0,0);
    const h = new Date(); h.setHours(0,0,0,0);
    return prazo < h;
  }).length;

  return {
    mes:   hoje.toLocaleDateString("pt-BR", { month: "long", year: "numeric" }),
    areas,
    acoes: { pendentes, concluidas, atrasadas },
  };
}


// ============================================================
// GERA ATA EM PDF NO GOOGLE DRIVE
// ============================================================
function gerarAtaNoDrive(dados) {
  const pasta   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const dataFmt = _data(dados.data);
  const nomeArq = `Ata - ${dados.tipo} - ${dataFmt}`;

  const doc  = DocumentApp.create(nomeArq);
  const body = doc.getBody();

  const estiloTit = {};
  estiloTit[DocumentApp.Attribute.BOLD]                 = true;
  estiloTit[DocumentApp.Attribute.FONT_SIZE]            = 18;
  estiloTit[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

  body.appendParagraph("ATA DE REUNIÃO").setAttributes(estiloTit);
  body.appendParagraph(dados.tipo).setAttributes(estiloTit).setBold(false).setFontSize(14);
  body.appendParagraph(`Data: ${dataFmt}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor("#555555");
  body.appendHorizontalRule();

  _secao(body, "PARTICIPANTES");
  dados.participantes.forEach(p => body.appendListItem(p.nome));

  _secao(body, "ASSUNTOS DISCUTIDOS");
  dados.topicos.forEach((t, i) => body.appendParagraph(`${i + 1}. ${t}`));

  _secao(body, "AÇÕES DEFINIDAS");
  if (!dados.acoes.length) {
    body.appendParagraph("Nenhuma ação definida nesta reunião.");
  } else {
    dados.acoes.forEach((a, i) => {
      const prazo = a.prazo ? _data(a.prazo) : "Sem prazo";
      body.appendParagraph(`${i + 1}. ${a.descricao}`).setBold(true);
      body.appendParagraph(`   👤 ${a.responsavelNome}   |   📅 Prazo: ${prazo}`)
        .setForegroundColor("#666666").setFontSize(10);
    });
  }

  body.appendParagraph(`\nAta gerada em ${new Date().toLocaleString("pt-BR")}`)
    .setForegroundColor("#aaaaaa").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdf     = docFile.getAs("application/pdf");
  pdf.setName(nomeArq + ".pdf");
  const pdfFile = pasta.createFile(pdf);
  docFile.setTrashed(true);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return pdfFile.getUrl();
}


// ============================================================
// REGISTRA NA PLANILHA
// ============================================================
function registrarNaPlanilha(dados, linkDrive) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Aba Atas ─────────────────────────────────────────────
  let abaAtas = ss.getSheetByName("Atas");
  if (!abaAtas) {
    abaAtas = ss.insertSheet("Atas");
    abaAtas.appendRow(["ID", "Área", "Data", "Participantes", "Qtd Ações", "Link PDF"]);
    abaAtas.setFrozenRows(1);
    _estiloCabecalho(abaAtas);
  }

  const idAta = "ATA-" + Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyyMMdd-HHmmss");

  abaAtas.appendRow([
    idAta,
    dados.tipo,
    dados.data,
    dados.participantes.map(p => p.nome).join(", "),
    dados.acoes.filter(a => a.descricao.trim()).length,
    linkDrive,
  ]);

  // ── Aba Ações ─────────────────────────────────────────────
  let abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) {
    abaAcoes = ss.insertSheet("Ações");
    abaAcoes.appendRow(["ID Ata", "Área", "Data Reunião", "Ação", "Responsável", "Telefone", "Prazo", "Status", "Data Conclusão"]);
    abaAcoes.setFrozenRows(1);
    _estiloCabecalho(abaAcoes);
  }

  dados.acoes.forEach(a => {
    if (!a.descricao.trim()) return;
    abaAcoes.appendRow([idAta, dados.tipo, dados.data, a.descricao, a.responsavelNome, a.responsavelTelefone, a.prazo || "", "Pendente", ""]);
  });
}


// ============================================================
// CRIA PRÓXIMA REUNIÃO NO GOOGLE AGENDA
// ============================================================
function criarProximaReuniaoCalendario(dados, linkDrive) {
  const configs = buscarConfiguracoes();
  const config  = configs.find(c => c.area === dados.tipo);
  if (!config || config.frequencia === "Sem reunião periódica" || !config.ativo) return;

  const proxima = calcularProximaData(config);
  if (!proxima) return;

  const [hora, min] = (config.horario || "09:00").split(":").map(Number);
  const inicio = new Date(proxima); inicio.setHours(hora, min, 0, 0);
  const fim    = new Date(inicio);  fim.setHours(fim.getHours() + 1);

  const convidados = dados.participantes
    .filter(p => p.email)
    .map(p => p.email)
    .join(",");

  if (!convidados) return;

  CalendarApp.getDefaultCalendar().createEvent(
    `Reunião ${dados.tipo} — Orgânico do Chico`,
    inicio, fim,
    {
      description: `Reunião periódica de ${dados.tipo}.\n\nÚltima ata: ${linkDrive}`,
      guests:      convidados,
      sendInvites: true,
    }
  );
}

function calcularProximaData(config) {
  const map = { "Domingo": 0, "Segunda": 1, "Terça": 2, "Quarta": 3, "Quinta": 4, "Sexta": 5, "Sábado": 6 };
  const alvo = map[config.dia];
  if (alvo === undefined) return null;

  const d = new Date();
  d.setDate(d.getDate() + 1);
  while (d.getDay() !== alvo) d.setDate(d.getDate() + 1);
  return d;
}


// ============================================================
// WHATSAPP — envia ata para todos
// ============================================================
function enviarAtaWhatsApp(dados, linkDrive) {
  const dataFmt = _data(dados.data);

  dados.participantes.forEach(p => {
    if (!p.telefone) return;

    let msg = `📋 *ATA DE REUNIÃO*\n*${dados.tipo}* — ${dataFmt}\n\n`;
    msg += `*Participantes:*\n` + dados.participantes.map(x => `• ${x.nome}`).join("\n") + "\n\n";
    msg += `*Assuntos discutidos:*\n` + dados.topicos.map((t, i) => `${i+1}. ${t}`).join("\n") + "\n\n";

    if (dados.acoes.length > 0) {
      msg += `*Ações definidas:*\n`;
      dados.acoes.forEach((a, i) => {
        msg += `${i+1}. ${a.descricao}\n   👤 ${a.responsavelNome}  📅 ${a.prazo ? _data(a.prazo) : "sem prazo"}\n`;
      });
      msg += "\n";
    }

    msg += `📄 *Ata completa (PDF):*\n${linkDrive}`;

    _whatsapp(p.telefone, msg);
    Utilities.sleep(600);
  });
}


// ============================================================
// LEMBRETE DIÁRIO DE AÇÕES — rode todo dia às 8h
// ============================================================
function enviarLembretes() {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) return;

  const linhas = abaAcoes.getDataRange().getValues();
  const hoje   = new Date(); hoje.setHours(0, 0, 0, 0);

  for (let i = 1; i < linhas.length; i++) {
    const [, reuniao, , acao, responsavel, telefone, prazoRaw, status] = linhas[i];
    if (status === "Concluída" || !telefone || !prazoRaw) continue;

    const prazo = new Date(prazoRaw + "T12:00:00"); prazo.setHours(0, 0, 0, 0);
    const diff  = Math.round((prazo - hoje) / 864e5);

    let msg = "";
    if      (diff === 1) msg = `⏰ *Lembrete de prazo*\n\nOlá, ${responsavel}! Sua ação da reunião *${reuniao}* vence *amanhã*:\n\n📌 ${acao}\n\nJá concluiu?\n✅ Responda *SIM* para marcar como concluída\n⏳ Responda *NÃO* se ainda está em andamento`;
    else if (diff === 0) msg = `🔴 *Prazo hoje!*\n\nOlá, ${responsavel}! Esta ação vence *hoje*:\n\n📌 ${acao}\n_(Reunião: ${reuniao})_\n\nJá concluiu?\n✅ Responda *SIM*\n⏳ Responda *NÃO*`;
    else if (diff < 0)  msg = `🚨 *Ação atrasada ${Math.abs(diff)} dia(s)!*\n\nOlá, ${responsavel}!\n\n📌 ${acao}\n_(Reunião: ${reuniao})_\n\nJá concluiu?\n✅ Responda *SIM*\n⏳ Responda *NÃO*`;

    if (msg) { _whatsapp(telefone, msg); Utilities.sleep(600); }
  }
}


// ============================================================
// PROCESSA RESPOSTA SIM/NÃO DO WHATSAPP
// ============================================================
function processarRespostaWhatsApp(telefone, mensagem) {
  const resp = mensagem.trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (resp !== "SIM" && resp !== "NAO") return;

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) return;

  const linhas   = abaAcoes.getDataRange().getValues();
  const telBusca = String(telefone).replace(/\D/g, "");

  for (let i = 1; i < linhas.length; i++) {
    const telPlan = String(linhas[i][5]).replace(/\D/g, "");
    const status  = linhas[i][7];
    const match   = telPlan === telBusca || telPlan === telBusca.slice(-11) || telBusca === telPlan.slice(-11);

    if (match && status === "Pendente") {
      if (resp === "SIM") {
        abaAcoes.getRange(i+1, 8).setValue("Concluída");
        abaAcoes.getRange(i+1, 9).setValue(Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy"));
        _whatsapp(telefone, `✅ Ação marcada como *concluída*! Bom trabalho, ${linhas[i][4]}! 👏`);
      } else {
        _whatsapp(telefone, `⏳ Entendido! A ação continua como pendente. Qualquer atualização é só responder por aqui.`);
      }
      break;
    }
  }
}


// ============================================================
// CONFIGURA GATILHO DIÁRIO — rode UMA VEZ só
// ============================================================
function configurarGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "enviarLembretes") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("enviarLembretes").timeBased().everyDays(1).atHour(8).create();
  Logger.log("✅ Gatilho diário configurado para 8h!");
}


// ============================================================
// CRIA ABAS DE CONFIGURAÇÃO — rode UMA VEZ só
// ============================================================
function criarAbaConfiguracoes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Aba Configurações ────────────────────────────────────
  let abaCfg = ss.getSheetByName("Configurações");
  if (!abaCfg) abaCfg = ss.insertSheet("Configurações");
  else abaCfg.clearContents();

  abaCfg.appendRow(["Área", "Frequência", "Dia da Semana", "Horário", "Ativo"]);
  _estiloCabecalho(abaCfg);

  [
    ["Operação",          "Semanal",               "Terça",  "09:00", "Sim"],
    ["Atendimento",       "Semanal",               "Quarta", "14:00", "Sim"],
    ["Compras",           "Sem reunião periódica",  "",       "",      "Sim"],
    ["Financeiro/Fiscal", "Sem reunião periódica",  "",       "",      "Sim"],
    ["Entregadores",      "Mensal",                 "Segunda","08:00", "Sim"],
    ["Geral",             "Sem reunião periódica",  "",       "",      "Sim"],
  ].forEach(r => abaCfg.appendRow(r));

  // Dropdowns
  _dropdown(abaCfg, 2, 2, 6, ["Semanal","Quinzenal","Mensal","Sem reunião periódica"]);
  _dropdown(abaCfg, 2, 3, 6, ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]);
  _dropdown(abaCfg, 2, 5, 6, ["Sim","Não"]);
  abaCfg.autoResizeColumns(1, 5);

  // ── Aba Participantes ────────────────────────────────────
  let abaPart = ss.getSheetByName("Participantes");
  if (!abaPart) abaPart = ss.insertSheet("Participantes");
  else abaPart.clearContents();

  abaPart.appendRow(["Área", "Nome", "WhatsApp", "Email (opcional)"]);
  _estiloCabecalho(abaPart);

  // Exemplos — substitua pelos dados reais da equipe
  abaPart.appendRow(["Operação",    "Bernardo",    "31999990000", "bernardo@organico.com.br"]);
  abaPart.appendRow(["Operação",    "João Silva",  "31988880000", ""]);
  abaPart.appendRow(["Atendimento", "Maria Santos","31977770000", "maria@organico.com.br"]);
  abaPart.autoResizeColumns(1, 4);

  Logger.log("✅ Abas criadas! Agora edite a aba 'Participantes' com os dados reais da equipe.");
}


// ============================================================
// AUXILIARES
// ============================================================
function _whatsapp(telefone, mensagem) {
  const num    = String(telefone).replace(/\D/g, "");
  const numero = num.startsWith("55") ? num : "55" + num;
  const opts   = {
    method: "POST",
    headers: { "Content-Type": "application/json", "apikey": EVOLUTION_KEY },
    payload: JSON.stringify({ number: numero, text: mensagem }),
    muteHttpExceptions: true,
  };
  try { UrlFetchApp.fetch(`${EVOLUTION_URL}/message/sendText/${EVOLUTION_INST}`, opts); }
  catch (err) { console.log(`Erro WhatsApp ${telefone}: ${err.message}`); }
}

function _data(str) {
  if (!str) return "—";
  return new Date(str + "T12:00:00").toLocaleDateString("pt-BR");
}

function _secao(body, titulo) {
  body.appendParagraph(titulo).setHeading(DocumentApp.ParagraphHeading.HEADING2).setForegroundColor("#1a5c35");
}

function _estiloCabecalho(sheet) {
  const r = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  r.setBackground("#1a5c35"); r.setFontColor("#ffffff"); r.setFontWeight("bold");
}

function _dropdown(sheet, startRow, col, numRows, values) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule);
}

function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function _ok(info) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", info })).setMimeType(ContentService.MimeType.JSON);
}

function _erro(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: msg })).setMimeType(ContentService.MimeType.JSON);
}
