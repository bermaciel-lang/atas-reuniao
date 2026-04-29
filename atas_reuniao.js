// ============================================================
// SISTEMA DE ATAS DE REUNIÃO — Orgânico do Chico  v3
// ============================================================
// ⚙️  PREENCHA AS 4 VARIÁVEIS ABAIXO
// ============================================================

const SPREADSHEET_ID  = "1lqtTSH238fzG5gJ_31VeJ61dbLK1i0UXEMioJBnVIyU";
const DRIVE_FOLDER_ID = "16WDblZKEWRpuApX82qC9E2LbGQ1mlmch";

// Z-API — cole aqui depois de criar a instância em z-api.io
const ZAPI_INSTANCE   = "3F2630320F71D193AB6962108CBB360D";   // ex: 3DF5A2B1C4D...
const ZAPI_TOKEN      = "33A064CB1201CA3C5B2C6F6A";          // ex: F23K9ABC...


// ============================================================
// doGet — serve dados para o formulário e o dashboard
// ============================================================
function doGet(e) {
  const action = e?.parameter?.action || "";

  if (action === "participantes") {
    const area = e.parameter.area || "";
    return _json(buscarParticipantesArea(area));
  }
  if (action === "padrao_area") {
    const area = e.parameter.area || "";
    return _json(buscarNomesPadraoArea(area));
  }
  if (action === "todos_participantes") {
    return _json(buscarTodosParticipantes());
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
// PARTICIPANTES
// ============================================================
// Retorna os NOMES padrão de uma área (colunas F-O da aba Configurações)
function buscarNomesPadraoArea(area) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Configurações");
  if (!aba) return [];
  const linha = aba.getDataRange().getValues().slice(1).find(r => String(r[0]).trim() === area);
  if (!linha) return [];
  // Colunas F-O = índices 5 a 14
  return linha.slice(5, 15).map(v => String(v).trim()).filter(v => v);
}

// Retorna participantes de uma área com contato completo
// Lê os nomes padrão da Configurações e busca o contato na aba Participantes
function buscarParticipantesArea(area) {
  const nomes = buscarNomesPadraoArea(area);
  if (!nomes.length) return [];
  const todos = buscarTodosParticipantes();
  return nomes
    .map(nome => todos.find(p => p.nome === nome) || { nome, telefone: "", email: "" })
    .filter(p => p.nome);
}

// Retorna todos os participantes da aba Participantes (ordem alfabética, sem duplicatas)
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
// DASHBOARD
// ============================================================
function dadosDashboard() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAtas = ss.getSheetByName("Atas");
  const abaAco  = ss.getSheetByName("Ações");
  const configs = buscarConfiguracoes();

  const hoje      = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const fimMes    = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);

  const todasAtas = abaAtas ? abaAtas.getDataRange().getValues().slice(1) : [];
  const atasMes   = todasAtas.filter(r => {
    if (!r[2]) return false;
    const d = new Date(r[2] + "T12:00:00");
    return d >= inicioMes && d <= fimMes;
  });

  const semanas = Math.round(fimMes.getDate() / 7);

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
    mes: hoje.toLocaleDateString("pt-BR", { month:"long", year:"numeric" }),
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

  let abaAtas = ss.getSheetByName("Atas");
  if (!abaAtas) {
    abaAtas = ss.insertSheet("Atas");
    abaAtas.appendRow(["ID", "Área", "Data", "Participantes", "Qtd Ações", "Link PDF"]);
    abaAtas.setFrozenRows(1);
    _estiloCabecalho(abaAtas);
  }

  const idAta = "ATA-" + Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyyMMdd-HHmmss");
  abaAtas.appendRow([
    idAta, dados.tipo, dados.data,
    dados.participantes.map(p => p.nome).join(", "),
    dados.acoes.filter(a => a.descricao.trim()).length,
    linkDrive,
  ]);

  let abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) {
    abaAcoes = ss.insertSheet("Ações");
    abaAcoes.appendRow(["ID Ata","Área","Data Reunião","Ação","Responsável","Telefone","Prazo","Status","Data Conclusão"]);
    abaAcoes.setFrozenRows(1);
    _estiloCabecalho(abaAcoes);
  }

  dados.acoes.forEach(a => {
    if (!a.descricao.trim()) return;
    abaAcoes.appendRow([idAta, dados.tipo, dados.data, a.descricao, a.responsavelNome, a.responsavelTelefone, a.prazo||"", "Pendente", ""]);
  });
}


// ============================================================
// ENVIA ATA POR WHATSAPP
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
// LEMBRETE DIÁRIO DE AÇÕES — roda todo dia às 8h
// Avisa: 1 dia antes do prazo, no dia do prazo, e ações atrasadas
// ============================================================
function enviarLembretesAcoes() {
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
    if (diff === 1) {
      msg = `⏰ *Lembrete de prazo — amanhã*\n\nOlá, ${responsavel}!\n\nSua ação da reunião de *${reuniao}* vence *amanhã*:\n\n📌 ${acao}\n\nJá concluiu?\n✅ Responda *SIM* para marcar como concluída\n⏳ Responda *NÃO* se ainda está em andamento`;
    } else if (diff === 0) {
      msg = `🔴 *Prazo hoje!*\n\nOlá, ${responsavel}!\n\nEsta ação vence *hoje*:\n\n📌 ${acao}\n_(Reunião: ${reuniao})_\n\nJá concluiu?\n✅ Responda *SIM*\n⏳ Responda *NÃO*`;
    } else if (diff < 0) {
      msg = `🚨 *Ação atrasada ${Math.abs(diff)} dia(s)!*\n\nOlá, ${responsavel}!\n\n📌 ${acao}\n_(Reunião: ${reuniao})_\n\nJá concluiu?\n✅ Responda *SIM*\n⏳ Responda *NÃO*`;
    }

    if (msg) { _whatsapp(telefone, msg); Utilities.sleep(600); }
  }
}


// ============================================================
// LEMBRETE DE REUNIÃO — roda a cada hora
// Avisa 2h antes de cada reunião agendada
// ============================================================
function enviarLembreteReuniao() {
  const configs = buscarConfiguracoes();
  const agora   = new Date();
  const agoraH  = agora.getHours();
  const agoraM  = agora.getMinutes();
  const diaSemana = agora.getDay(); // 0=Dom, 1=Seg...

  const mapDia = { "Domingo":0,"Segunda":1,"Terça":2,"Quarta":3,"Quinta":4,"Sexta":5,"Sábado":6 };

  configs.forEach(c => {
    if (!c.ativo || !c.horario || !c.dia) return;
    if (c.frequencia === "Sem reunião periódica") return;
    if (mapDia[c.dia] !== diaSemana) return;

    const [hReuniao, mReuniao] = c.horario.split(":").map(Number);

    // Verifica se a reunião começa em 2h (±15 min de tolerância)
    const minutosReuniao = hReuniao * 60 + mReuniao;
    const minutosAgora   = agoraH * 60 + agoraM;
    const diff           = minutosReuniao - minutosAgora;

    if (diff < 105 || diff > 135) return; // fora da janela de 2h (±15min)

    // Busca participantes da área
    const participantes = buscarParticipantesArea(c.area);
    participantes.forEach(p => {
      if (!p.telefone) return;
      const msg = `📅 *Lembrete de reunião — em 2 horas*\n\nOlá, ${p.nome}!\n\nVocê tem uma reunião de *${c.area}* hoje às *${c.horario}h*.\n\nPrepare os tópicos que deseja discutir. 💪`;
      _whatsapp(p.telefone, msg);
      Utilities.sleep(600);
    });
  });
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

  const convidados = dados.participantes.filter(p => p.email).map(p => p.email).join(",");
  if (!convidados) return;

  CalendarApp.getDefaultCalendar().createEvent(
    `Reunião ${dados.tipo} — Orgânico do Chico`,
    inicio, fim,
    { description: `Reunião periódica de ${dados.tipo}.\n\nÚltima ata: ${linkDrive}`, guests: convidados, sendInvites: true }
  );
}

function calcularProximaData(config) {
  const map = { "Domingo":0,"Segunda":1,"Terça":2,"Quarta":3,"Quinta":4,"Sexta":5,"Sábado":6 };
  const alvo = map[config.dia];
  if (alvo === undefined) return null;
  const d = new Date();
  d.setDate(d.getDate() + 1);
  while (d.getDay() !== alvo) d.setDate(d.getDate() + 1);
  return d;
}


// ============================================================
// CONFIGURA TODOS OS GATILHOS — rode UMA VEZ só
// ============================================================
function configurarTodosGatilhos() {
  // Remove todos os gatilhos antigos
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "enviarLembretesAcoes" || fn === "enviarLembreteReuniao" || fn === "enviarLembretes") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Lembrete de ações — todo dia às 8h
  ScriptApp.newTrigger("enviarLembretesAcoes")
    .timeBased().everyDays(1).atHour(8).create();

  // Lembrete de reunião — a cada hora (verifica se alguma começa em 2h)
  ScriptApp.newTrigger("enviarLembreteReuniao")
    .timeBased().everyHours(1).create();

  Logger.log("✅ Gatilhos configurados!");
  Logger.log("   • Lembretes de ações: todo dia às 8h");
  Logger.log("   • Lembretes de reunião: verificação a cada hora (avisa 2h antes)");
}


// ============================================================
// CRIA ABAS DE CONFIGURAÇÃO — rode UMA VEZ só
// ============================================================
function criarAbaConfiguracoes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

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

  _dropdown(abaCfg, 2, 2, 6, ["Semanal","Quinzenal","Mensal","Sem reunião periódica"]);
  _dropdown(abaCfg, 2, 3, 6, ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]);
  _dropdown(abaCfg, 2, 5, 6, ["Sim","Não"]);
  abaCfg.autoResizeColumns(1, 5);

  let abaPart = ss.getSheetByName("Participantes");
  if (!abaPart) abaPart = ss.insertSheet("Participantes");
  else abaPart.clearContents();

  abaPart.appendRow(["Área", "Nome", "WhatsApp", "Email (opcional)"]);
  _estiloCabecalho(abaPart);
  abaPart.appendRow(["Operação",    "Bernardo",    "31999990000", "bernardo@organico.com.br"]);
  abaPart.appendRow(["Operação",    "João Silva",  "31988880000", ""]);
  abaPart.appendRow(["Atendimento", "Maria Santos","31977770000", "maria@organico.com.br"]);
  abaPart.autoResizeColumns(1, 4);

  Logger.log("✅ Abas criadas! Edite a aba Participantes com os dados reais da equipe.");
}


// ============================================================
// AUXILIARES
// ============================================================

// Envia mensagem via Z-API
function _whatsapp(telefone, mensagem) {
  const num    = String(telefone).replace(/\D/g, "");
  const numero = num.startsWith("55") ? num : "55" + num;

  const opts = {
    method: "POST",
    headers: { "Content-Type": "application/json", "client-token": ZAPI_TOKEN },
    payload: JSON.stringify({ phone: numero, message: mensagem }),
    muteHttpExceptions: true,
  };

  try {
    UrlFetchApp.fetch(
      `https://api.z-api.io/instances/${ZAPI_INSTANCE}/token/${ZAPI_TOKEN}/send-text`,
      opts
    );
  } catch (err) {
    console.log(`Erro WhatsApp ${telefone}: ${err.message}`);
  }
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
  return ContentService.createTextOutput(JSON.stringify({ status:"ok", info })).setMimeType(ContentService.MimeType.JSON);
}

function _erro(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status:"erro", mensagem:msg })).setMimeType(ContentService.MimeType.JSON);
}