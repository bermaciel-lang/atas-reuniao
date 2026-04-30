// ============================================================
// SISTEMA DE ATAS DE REUNIÃO — Orgânico do Chico  v3
// ============================================================

const SPREADSHEET_ID      = "1lqtTSH238fzG5gJ_31VeJ61dbLK1i0UXEMioJBnVIyU";
const DRIVE_FOLDER_ID     = "16WDblZKEWRpuApX82qC9E2LbGQ1mlmch";
const EVOLUTION_URL       = "https://evolution-api-production-32ab.up.railway.app";
const EVOLUTION_KEY       = "organico123";
const EVOLUTION_INSTANCE  = "organico";


// ============================================================
// doGet
// ============================================================
function doGet(e) {
  const action = e?.parameter?.action || "";
  if (action === "participantes")       return _json(buscarParticipantesArea(e.parameter.area || ""));
  if (action === "padrao_area")         return _json(buscarNomesPadraoArea(e.parameter.area || ""));
  if (action === "todos_participantes") return _json(buscarTodosParticipantes());
  if (action === "dashboard")           return _json(dadosDashboard());
  if (action === "configuracoes")       return _json(buscarConfiguracoes());
  return _json({ status: "ok" });
}


// ============================================================
// doPost — recebe formulário de ata E webhook do Z-API
// ============================================================
function doPost(e) {
  try {
    _log("=== doPost chamado ===");
    const raw = e.postData.contents;
    _log("Raw: " + raw.substring(0, 500));
    
    const dados = JSON.parse(raw);
    _log("event: " + dados.event);

    if (dados.event === "messages.upsert" || dados.event === "messages.update") {
      const msg      = dados.data?.message || dados.data;
      const fromMe   = msg?.key?.fromMe || false;
      _log("fromMe: " + fromMe);
      if (fromMe) return _ok("ignorado");
      const telefone = (msg?.key?.remoteJid || "").replace("@s.whatsapp.net","").replace("@c.us","");
      const texto    = msg?.message?.conversation
                    || msg?.message?.extendedTextMessage?.text
                    || msg?.message?.buttonsResponseMessage?.selectedButtonId
                    || "";
      _log("telefone: " + telefone + " | texto: " + texto);
      if (telefone && texto) processarRespostaWhatsApp(telefone, texto);
      return _ok("webhook processado");
    }

    _log("evento nao reconhecido: " + JSON.stringify(dados).substring(0, 300));

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
    _log("ERRO: " + err.message);
    return _erro(err.message);
  }
}


// ============================================================
// PARTICIPANTES
// ============================================================
function buscarNomesPadraoArea(area) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Configurações");
  if (!aba) return [];
  const linha = aba.getDataRange().getValues().slice(1).find(r => String(r[0]).trim() === area);
  if (!linha) return [];
  return linha.slice(5, 15).map(v => String(v).trim()).filter(v => v);
}

function buscarParticipantesArea(area) {
  const nomes = buscarNomesPadraoArea(area);
  if (!nomes.length) return [];
  const todos = buscarTodosParticipantes();
  return nomes.map(nome => todos.find(p => p.nome === nome) || { nome, telefone: "", email: "" });
}

function buscarTodosParticipantes() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Participantes");
  if (!aba) return [];
  const vistos = new Set();
  return aba.getDataRange().getValues()
    .slice(1)
    .filter(r => r[1] && !vistos.has(String(r[1]).trim()) && vistos.add(String(r[1]).trim()))
    .map(r => ({ nome: String(r[1]).trim(), telefone: String(r[2]||"").trim(), email: String(r[3]||"").trim() }))
    .sort((a, b) => a.nome.localeCompare(b.nome));
}


// ============================================================
// CONFIGURAÇÕES
// ============================================================
function buscarConfiguracoes() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Configurações");
  if (!aba) return [];
  return aba.getDataRange().getValues().slice(1).map(r => ({
    area: String(r[0]).trim(), frequencia: String(r[1]).trim(),
    dia:  String(r[2]).trim(), horario: r[3] instanceof Date
  ? Utilities.formatDate(r[3], "America/Sao_Paulo", "HH:mm")
  : String(r[3]).trim(),
    ativo: String(r[4]).trim() === "Sim",
  }));
}


// ============================================================
// DASHBOARD
// ============================================================
function dadosDashboard() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAtas = ss.getSheetByName("Atas");
  const abaAco  = ss.getSheetByName("Ações");
  const configs = buscarConfiguracoes();
  const hoje    = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const fimMes    = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
  const semanas   = Math.round(fimMes.getDate() / 7);

  const todasAtas = abaAtas ? abaAtas.getDataRange().getValues().slice(1) : [];
  const atasMes   = todasAtas.filter(r => {
    if (!r[2]) return false;
    const d = new Date(r[2] + "T12:00:00");
    return d >= inicioMes && d <= fimMes;
  });

  const areas = configs.map(c => {
    const atasDaArea = atasMes.filter(r => r[1] === c.area);
    let esperadas = 0;
    if (c.frequencia === "Semanal")   esperadas = semanas;
    if (c.frequencia === "Quinzenal") esperadas = Math.ceil(semanas / 2);
    if (c.frequencia === "Mensal")    esperadas = 1;
    const realizadas = atasDaArea.length;
    let status = "sem_periodicidade";
    if (esperadas > 0) {
      if (realizadas >= esperadas)  status = "em_dia";
      else if (realizadas > 0)      status = "parcial";
      else                          status = "atrasada";
    }
    return {
      area: c.area, frequencia: c.frequencia, esperadas, realizadas, status,
      atas: atasDaArea.map(r => ({
        id: String(r[0]), data: String(r[2]),
        participantes: String(r[3]), qtdAcoes: Number(r[4])||0, link: String(r[5]),
      })).sort((a, b) => b.data.localeCompare(a.data)),
    };
  });

  const acoes      = abaAco ? abaAco.getDataRange().getValues().slice(1) : [];
  const pendentes  = acoes.filter(r => r[7] === "Pendente").length;
  const concluidas = acoes.filter(r => r[7] === "Concluída").length;
  const atrasadas  = acoes.filter(r => {
    if (r[7] !== "Pendente" || !r[6]) return false;
    const prazo = new Date(r[6] + "T12:00:00"); prazo.setHours(0,0,0,0);
    const h = new Date(); h.setHours(0,0,0,0);
    return prazo < h;
  }).length;

  return {
    mes: hoje.toLocaleDateString("pt-BR", { month:"long", year:"numeric" }),
    areas, acoes: { pendentes, concluidas, atrasadas },
  };
}


// ============================================================
// GERA ATA EM PDF NO GOOGLE DRIVE
// ============================================================
function gerarAtaNoDrive(dados) {
  const pasta   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const dataFmt = _data(dados.data);
  const nomeArq = `Ata - ${dados.tipo} - ${dataFmt}`;
  const doc     = DocumentApp.create(nomeArq);
  const body    = doc.getBody();

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
  (dados.topicos||[]).forEach((t, i) => body.appendParagraph(`${i + 1}. ${t}`));

  _secao(body, "AÇÕES DEFINIDAS");
  if (!dados.acoes || !dados.acoes.length) {
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
  const pdf = docFile.getAs("application/pdf");
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

  let abaAtas = ss.getSheetByName("Atas");
  if (!abaAtas) {
    abaAtas = ss.insertSheet("Atas");
    abaAtas.appendRow(["ID","Área","Data","Participantes","Qtd Ações","Link PDF"]);
    abaAtas.setFrozenRows(1);
    _estiloCabecalho(abaAtas);
  }
  const idAta = "ATA-" + Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyyMMdd-HHmmss");
  abaAtas.appendRow([
    idAta, dados.tipo, dados.data,
    dados.participantes.map(p => p.nome).join(", "),
    (dados.acoes||[]).filter(a => a.descricao && a.descricao.trim()).length,
    linkDrive,
  ]);

  let abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) {
    abaAcoes = ss.insertSheet("Ações");
    abaAcoes.appendRow(["ID Ata","Área","Data Reunião","Ação","Responsável","Telefone","Prazo","Status","Data Conclusão"]);
    abaAcoes.setFrozenRows(1);
    _estiloCabecalho(abaAcoes);
  }
  (dados.acoes||[]).forEach(a => {
    if (!a.descricao || !a.descricao.trim()) return;
    abaAcoes.appendRow([idAta, dados.tipo, dados.data, a.descricao, a.responsavelNome||"", a.responsavelTelefone||"", a.prazo||"", "Pendente", ""]);
  });
}


// ============================================================
// ENVIA ATA + AÇÕES POR WHATSAPP (mensagens separadas)
// ============================================================
function enviarAtaWhatsApp(dados, linkDrive) {
  const dataFmt = _data(dados.data);

  // Mensagem 1 — ata para todos
  let msgAta = `📋 *ATA DE REUNIÃO*\n*${dados.tipo}* — ${dataFmt}\n\n`;
  msgAta += `*Participantes:*\n` + dados.participantes.map(x => `• ${x.nome}`).join("\n") + "\n\n";
  msgAta += `*Assuntos discutidos:*\n` + (dados.topicos||[]).map((t,i) => `${i+1}. ${t}`).join("\n") + "\n\n";
  msgAta += `📄 *Ata completa (PDF):*\n${linkDrive}`;

  dados.participantes.forEach(p => {
    if (!p.telefone) return;
    _whatsapp(p.telefone, msgAta);
    Utilities.sleep(600);
  });

  // Mensagem 2 — ações agrupadas por responsável
  const acoesPorPessoa = {};
  (dados.acoes||[]).forEach(a => {
    if (!a.descricao || !a.responsavelTelefone) return;
    const tel = a.responsavelTelefone;
    if (!acoesPorPessoa[tel]) acoesPorPessoa[tel] = { nome: a.responsavelNome, acoes: [] };
    acoesPorPessoa[tel].acoes.push(a);
  });

  Object.entries(acoesPorPessoa).forEach(([tel, pessoa]) => {
    const primeiroNome = pessoa.nome.split(" ")[0];
    let msgAcoes = `📌 *Suas ações — ${dados.tipo}*\n\nOlá, ${primeiroNome}! Aqui estão as ações que ficaram com você:\n\n`;
    pessoa.acoes.forEach((a, i) => {
      msgAcoes += `${i+1}. ${a.descricao}\n   📅 Prazo: ${a.prazo ? _data(a.prazo) : "Sem prazo"}\n\n`;
    });
    msgAcoes += `Você receberá lembretes automáticos antes do prazo.\n✅ Responda *SIM* quando concluir | ⏳ Responda *NÃO* se ainda pendente`;
    _whatsapp(tel, msgAcoes);
    Utilities.sleep(600);
  });
}

function enviarLembretesAcoes() {
  // Não envia sábado (6) ou domingo (0)
  const diaSemana = new Date().getDay();
  if (diaSemana === 0 || diaSemana === 6) return;

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) return;
  const linhas = abaAcoes.getDataRange().getValues();
  const hoje   = new Date(); hoje.setHours(0, 0, 0, 0);

  for (let i = 1; i < linhas.length; i++) {
    const [, reuniao, , acao, responsavel, telefone, prazoRaw, status] = linhas[i];
    if (status === "Concluída" || !telefone || !prazoRaw) continue;

    const prazoDate = prazoRaw instanceof Date
      ? new Date(prazoRaw.getFullYear(), prazoRaw.getMonth(), prazoRaw.getDate())
      : new Date(String(prazoRaw).split("T")[0] + "T12:00:00");
    const diff = Math.round((prazoDate - hoje) / 864e5);

    let titulo = "";
    if      (diff === 1) titulo = `⏰ Lembrete — prazo amanhã`;
    else if (diff === 0) titulo = `🔴 Prazo hoje!`;
    else if (diff < 0)   titulo = `🚨 Ação atrasada ${Math.abs(diff)} dia(s)!`;
    else continue;

    const prazoFmt = Utilities.formatDate(prazoDate, "America/Sao_Paulo", "dd/MM/yyyy");
    const msg = `${titulo}\n\nOlá, ${responsavel}!\n\n📌 ${acao}\n_(Reunião: ${reuniao} | Prazo: ${prazoFmt})_\n\nJá concluiu essa ação?\n\n*Digite 1* para confirmar que concluiu ✅\n*Digite 2* se ainda está em andamento ⏳`;

    _whatsapp(telefone, msg);
    Utilities.sleep(800);
  }
}

function enviarLembreteReuniao() {
  // Não envia sábado (6) ou domingo (0)
  const agora     = new Date();
  const diaSemana = agora.getDay();
  if (diaSemana === 0 || diaSemana === 6) return;

  const configs = buscarConfiguracoes();
  const mapDia  = { "Domingo":0,"Segunda":1,"Terça":2,"Quarta":3,"Quinta":4,"Sexta":5,"Sábado":6 };
  const horaAgora = agora.getHours();

  configs.forEach(c => {
    if (!c.ativo || !c.dia || !c.horario) return;
    if (c.frequencia === "Sem reunião periódica") return;
    if (mapDia[c.dia] !== diaSemana) return;

    const [hReuniao] = c.horario.split(":").map(Number);
    const is7h      = horaAgora === 7;
    const is1hAntes = horaAgora === hReuniao - 1;

    if (!is7h && !is1hAntes) return;

    const tipoAviso = is1hAntes ? "em 1 hora" : "hoje";
    const emoji     = is1hAntes ? "⏰" : "📅";

    buscarParticipantesArea(c.area).forEach(p => {
      if (!p.telefone) return;
      _whatsapp(p.telefone,
        `${emoji} *Lembrete de reunião — ${tipoAviso}!*\n\nOlá, ${p.nome}!\n\nVocê tem uma reunião de *${c.area}* hoje às *${c.horario}h*.\n\nPrepare os tópicos que deseja discutir. 💪`
      );
      Utilities.sleep(600);
    });
  });
}

function configurarTodosGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["enviarLembretesAcoes","enviarLembreteReuniao","enviarLembretes"].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Lembretes de ações — todo dia às 8h Brasília
  ScriptApp.newTrigger("enviarLembretesAcoes")
    .timeBased().everyDays(1).atHour(8)
    .inTimezone("America/Sao_Paulo").create();

  // Lembrete de reunião — a cada hora (verifica se é 7h ou 1h antes)
  ScriptApp.newTrigger("enviarLembreteReuniao")
    .timeBased().everyHours(1)
    .inTimezone("America/Sao_Paulo").create();

  Logger.log("✅ Gatilhos configurados!");
  Logger.log("   • Lembretes de ações: dias úteis às 8h");
  Logger.log("   • Lembretes de reunião: dias úteis às 7h + 1h antes");
}

// ============================================================
// PROCESSA RESPOSTA (texto digitado OU botão clicado)
// ============================================================
function processarRespostaWhatsApp(telefone, mensagem) {
  const resp  = String(mensagem).trim();
  const isSim = resp === "1" || resp.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"") === "SIM";
  const isNao = resp === "2" || resp.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"") === "NAO";
  if (!isSim && !isNao) return;

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) return;
  const linhas   = abaAcoes.getDataRange().getValues();

  // Remove tudo exceto números e pega só os últimos 11 dígitos para comparar
    const telBusca = String(telefone).replace(/\D/g, "").slice(-8);


  _log("Buscando telefone: " + telBusca);

  for (let i = 1; i < linhas.length; i++) {
    const telPlan = String(linhas[i][5]).replace(/\D/g, "").slice(-8);
    const status   = linhas[i][7];
    
    _log(`Linha ${i}: telPlan=${telPlan} status=${status} match=${telPlan === telBusca}`);

    if (telPlan === telBusca && status === "Pendente") {
      if (isSim) {
        abaAcoes.getRange(i+1, 8).setValue("Concluída");
        abaAcoes.getRange(i+1, 9).setValue(
          Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy")
        );
        _whatsapp(telefone, `✅ Ação marcada como *concluída*! Bom trabalho, ${linhas[i][4]}! 👏`);
        _log("Ação concluída para " + linhas[i][4]);
      } else {
        _whatsapp(telefone, `⏳ Entendido! A ação continua como pendente.`);
      }
      break;
    }
  }
}


// ============================================================
// GOOGLE AGENDA
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
  const convidados = dados.participantes.filter(p => p.email).map(p => p.email).join(",");
  if (!convidados) return;
  CalendarApp.getDefaultCalendar().createEvent(
    `Reunião ${dados.tipo} — Orgânico do Chico`, inicio, fim,
    { description: `Reunião periódica de ${dados.tipo}.\n\nÚltima ata: ${linkDrive}`, guests: convidados, sendInvites: true }
  );
}

function calcularProximaData(config) {
  const map  = { "Domingo":0,"Segunda":1,"Terça":2,"Quarta":3,"Quinta":4,"Sexta":5,"Sábado":6 };
  const alvo = map[config.dia];
  if (alvo === undefined) return null;
  const d = new Date();
  d.setDate(d.getDate() + 1);
  while (d.getDay() !== alvo) d.setDate(d.getDate() + 1);
  return d;
}

// ============================================================
// AUXILIARES
// ============================================================

// Mensagem de texto simples — Evolution API
function _whatsapp(telefone, mensagem) {
  const num    = String(telefone).replace(/\D/g, "");
  const numero = num.startsWith("55") ? num : "55" + num;
  const opts   = {
    method: "POST",
    headers: { "Content-Type": "application/json", "apikey": EVOLUTION_KEY },
    payload: JSON.stringify({ number: numero, textMessage: { text: mensagem } }),
    muteHttpExceptions: true,
  };
  try {
    const resp = UrlFetchApp.fetch(
      `${EVOLUTION_URL}/message/sendText/${EVOLUTION_INSTANCE}`,
      opts
    );
    Logger.log("Evolution texto " + telefone + ": " + resp.getContentText());
  } catch (err) {
    Logger.log("Erro texto " + telefone + ": " + err.message);
  }
}

// Mensagem com botões — Evolution API
function _whatsappBotoes(telefone, mensagem, botoes) {
  _whatsapp(telefone, mensagem + "\n\n*Digite 1* para confirmar que concluiu ✅\n*Digite 2* se ainda está em andamento ⏳");
}

function _data(str) {
  if (!str) return "—";
  if (str instanceof Date) return str.toLocaleDateString("pt-BR");
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
  return ContentService.createTextOutput(JSON.stringify({ status:"ok", info })).setMimeType(ContentService.MimeType.JSON);
}

function _erro(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status:"erro", mensagem:msg })).setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// TESTES
// ============================================================
function testeWhatsApp() {
  _whatsapp("31994599539", "✅ Teste de texto simples — funcionando!");
}

function testeCobrancaAcoesReais() {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  const linhas   = abaAcoes.getDataRange().getValues();

  for (let i = 1; i < linhas.length; i++) {
    const [, reuniao, , acao, responsavel, telefone, prazoRaw, status] = linhas[i];
    if (!telefone) continue;

    // Formata a data independente do tipo
    let prazoFmt = "Sem prazo";
    if (prazoRaw) {
      const d = prazoRaw instanceof Date ? prazoRaw : new Date(prazoRaw);
      prazoFmt = Utilities.formatDate(d, "America/Sao_Paulo", "dd/MM/yyyy");
    }

    const msg = `⏰ *Lembrete de prazo — amanhã*\n\nOlá, ${responsavel}!\n\nSua ação da reunião de *${reuniao}* vence *amanhã*:\n\n📌 ${acao}\n   📅 Prazo: ${prazoFmt}\n\nJá concluiu essa ação?`;

    _whatsappBotoes(telefone, msg, [
      { id: `sim_${i}`, label: "✅ Sim, concluí!" },
      { id: `nao_${i}`, label: "⏳ Não, ainda pendente" }
    ]);
    Logger.log(`Enviado para ${responsavel} (${telefone})`);
    Utilities.sleep(800);
  }
}

function _log(msg) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    let aba   = ss.getSheetByName("Logs");
    if (!aba) { aba = ss.insertSheet("Logs"); aba.appendRow(["Data","Mensagem"]); }
    aba.appendRow([new Date().toLocaleString("pt-BR"), String(msg).substring(0, 500)]);
  } catch(e) {}
}

function testeDoPostManual() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        event: "messages.upsert",
        data: {
          message: {
            key: { remoteJid: "553194599539@s.whatsapp.net", fromMe: false },
            message: { conversation: "1" }
          }
        }
      })
    }
  };
  doPost(fakeEvent);
}