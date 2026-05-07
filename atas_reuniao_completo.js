// ============================================================
// SISTEMA DE ATAS DE REUNIÃO — Orgânico do Chico  v4
// ============================================================

const SPREADSHEET_ID      = "1lqtTSH238fzG5gJ_31VeJ61dbLK1i0UXEMioJBnVIyU";
const DRIVE_FOLDER_ID     = "16WDblZKEWRpuApX82qC9E2LbGQ1mlmch";
const EVOLUTION_URL       = "https://evolution-api-production-32ab.up.railway.app";
const EVOLUTION_KEY       = "organico123";
const EVOLUTION_INSTANCE  = "organico";

// ============================================================
// Estrutura da aba Configurações:
// A=Título, B=Frequência, C=Data única (se DATA), D=Dia da semana,
// E=Horário, F=Lembrete (ex: 7H/1H), G=Ativo,
// H-Q=Participantes 1-10, R=Data de Alteração (automática)
// ============================================================

// ============================================================
// doGet
// ============================================================
function doGet(e) {
  const action = e?.parameter?.action || "";
  if (action === "participantes")       return _json(buscarParticipantesArea(e.parameter.area || ""));
  if (action === "padrao_area")         return _json(buscarNomesPadraoArea(e.parameter.area || ""));
  if (action === "todos_participantes") return _json(buscarTodosParticipantes());
  if (action === "dashboard")           return _json(dadosDashboard(e.parameter.mes || ""));
  if (action === "acoes_completas")     return _json(buscarTodasAcoes());
  if (action === "configuracoes")       return _json(buscarConfiguracoes());
  return _json({ status: "ok" });
}

// ============================================================
// doPost
// ============================================================
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);

    if (dados.event === "messages.upsert" || dados.event === "messages.update") {
      const msg    = dados.data || {};
      if (msg.status) return _ok("ignorado - status update");
      const fromMe = msg?.key?.fromMe || false;
      if (fromMe) return _ok("ignorado - fromMe");

      let telefone = (msg?.key?.remoteJid || "");
      if (telefone.includes("@s.whatsapp.net")) {
        telefone = telefone.replace("@s.whatsapp.net", "");
      } else if (telefone.includes("@c.us")) {
        telefone = telefone.replace("@c.us", "");
      } else if (telefone.includes("@lid")) {
        const pushName = dados.data?.pushName || "";
        telefone = buscarTelefonePorNome(pushName);
      }

      const texto = msg?.message?.conversation
                 || msg?.message?.extendedTextMessage?.text
                 || "";

      if (telefone && texto) processarRespostaWhatsApp(telefone, texto);
      return _ok("webhook processado");
    }

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
function buscarTelefonePorNome(nome) {
  if (!nome) return "";
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Participantes");
  if (!aba) return "";
  const nomeNorm = nome.trim().toLowerCase();
  const encontrado = aba.getDataRange().getValues().slice(1).find(r =>
    String(r[1]).trim().toLowerCase().includes(nomeNorm) ||
    nomeNorm.includes(String(r[1]).trim().toLowerCase())
  );
  return encontrado ? String(encontrado[2]).trim() : "";
}

function buscarNomesPadraoArea(titulo) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Configurações");
  if (!aba) return [];
  const linha = aba.getDataRange().getValues().slice(1)
    .find(r => String(r[0]).trim() === titulo);
  if (!linha) return [];
  // Participantes: cols H-Q = índices 7-16
  return linha.slice(7, 17).map(v => String(v).trim()).filter(v => v);
}

function buscarParticipantesArea(titulo) {
  const nomes = buscarNomesPadraoArea(titulo);
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
// CONFIGURAÇÕES — nova estrutura
// ============================================================
function buscarConfiguracoes() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Configurações");
  if (!aba) return [];
  return aba.getDataRange().getValues().slice(1)
    .filter(r => r[0] && String(r[0]).trim())
    .map(r => ({
      titulo:        String(r[0]).trim(),                          // A
      frequencia:    String(r[1]).trim(),                          // B
      dataUnica:     r[2] instanceof Date ? r[2]                   // C
                   : (String(r[2]).trim() || null),
      dia:           String(r[3]).trim(),                          // D
      horario:       r[4] instanceof Date                          // E
        ? Utilities.formatDate(r[4], "America/Sao_Paulo", "HH:mm")
        : String(r[4]).trim(),
      lembrete:      String(r[5]).trim(),                          // F ex: 7H/1H
      ativo:         String(r[6]).trim() === "Sim",                // G
      dataAlteracao: r[17] instanceof Date ? r[17] : null,         // R
    }));
}

// ============================================================
// onEdit — carimbo automático na col R quando col B muda
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Configurações") return;
  if (e.range.getColumn() !== 2) return; // col B
  const linha = e.range.getRow();
  if (linha < 2) return;
  // Col R = coluna 18
  sheet.getRange(linha, 18).setValue(
    Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy")
  );
}

// ============================================================
// DASHBOARD
// ============================================================
function dadosDashboard(mesFiltro) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAtas = ss.getSheetByName("Atas");
  const abaAco  = ss.getSheetByName("Ações");
  const configs = buscarConfiguracoes();
  const hoje    = new Date();

  let inicioMes, fimMes;
  if (mesFiltro) {
    const [ano, mes] = mesFiltro.split("-").map(Number);
    inicioMes = new Date(ano, mes-1, 1);
    fimMes    = new Date(ano, mes, 0);
  } else {
    inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    fimMes    = new Date(hoje.getFullYear(), hoje.getMonth()+1, 0);
  }

  const todasAtas = abaAtas ? abaAtas.getDataRange().getValues().slice(1) : [];
  const atasMes   = todasAtas.filter(r => {
    if (!r[2]) return false;
    const d = r[2] instanceof Date ? r[2] : new Date(r[2] + "T12:00:00");
    return d >= inicioMes && d <= fimMes;
  });

  const areas = configs
    .filter(c => c.titulo && c.titulo.trim())
    .map(c => {
      const atasDaArea = atasMes.filter(r => r[1] === c.titulo);
      let esperadas = 0;

      if (c.frequencia !== "Sem reunião periódica" && c.frequencia !== "DATA" && c.ativo) {
        const contarDe = (c.dataAlteracao && c.dataAlteracao >= inicioMes && c.dataAlteracao <= fimMes)
          ? c.dataAlteracao : inicioMes;
        const diasRestantes = Math.round((fimMes - contarDe) / 864e5) + 1;
        if (c.frequencia === "Semanal")   esperadas = Math.max(1, Math.floor(diasRestantes / 7));
        if (c.frequencia === "Quinzenal") esperadas = Math.max(1, Math.floor(diasRestantes / 14));
        if (c.frequencia === "Mensal")    esperadas = 1;
        if (c.frequencia === "ULT/MES")   esperadas = 1;
        if (c.frequencia === "ULT/3M") {
          const mes = inicioMes.getMonth();
          esperadas = [2,5,8,11].includes(mes) ? 1 : 0;
        }
      }

      if (c.frequencia === "DATA" && c.dataUnica) {
        const dataEvento = c.dataUnica instanceof Date
          ? c.dataUnica : _parsarData_(String(c.dataUnica));
        if (dataEvento && dataEvento >= inicioMes && dataEvento <= fimMes) esperadas = 1;
      }

      const realizadas = atasDaArea.length;
      let status = "sem_periodicidade";
      if (esperadas > 0) {
        if (realizadas >= esperadas)  status = "em_dia";
        else if (realizadas > 0)      status = "parcial";
        else                          status = "atrasada";
      }

      return {
        area: c.titulo, frequencia: c.frequencia, esperadas, realizadas, status, ativo: c.ativo,
        atas: atasDaArea.map(r => ({
          id: String(r[0]),
          data: r[2] instanceof Date ? Utilities.formatDate(r[2],"America/Sao_Paulo","yyyy-MM-dd") : String(r[2]),
          participantes: String(r[3]), qtdAcoes: Number(r[4])||0, link: String(r[5]),
        })).sort((a,b) => b.data.localeCompare(a.data)),
      };
    })
    .filter(c => c.esperadas > 0 || c.realizadas > 0);

  const acoes     = abaAco ? abaAco.getDataRange().getValues().slice(1) : [];
  const pendentes = acoes.filter(r => r[7] === "Pendente").length;
  const atrasadas = acoes.filter(r => {
    if (r[7] !== "Pendente" || !r[6]) return false;
    const prazo = r[6] instanceof Date ? r[6] : new Date(r[6]+"T12:00:00");
    prazo.setHours(0,0,0,0);
    const h = new Date(); h.setHours(0,0,0,0);
    return prazo < h;
  }).length;

  const mesTxt = mesFiltro
    ? new Date(mesFiltro+"-15").toLocaleDateString("pt-BR",{month:"long",year:"numeric"})
    : hoje.toLocaleDateString("pt-BR",{month:"long",year:"numeric"});

  return { mes: mesTxt, areas, acoes: { pendentes, atrasadas } };
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
  (dados.participantes || []).forEach(p => body.appendListItem(p.nome));
  body.appendParagraph("");

  if (dados.pautas && dados.pautas.length) {
    _secao(body, "ASSUNTOS DISCUTIDOS / AÇÕES");
    dados.pautas.forEach((p, pi) => {
      if (pi > 0) body.appendParagraph("");
      const tit = body.appendParagraph(`${pi + 1}. ${p.titulo || "Pauta " + (pi+1)}`);
      tit.setBold(true).setFontSize(12);
      if (p.discussao) body.appendParagraph(p.discussao).setFontSize(11).setForegroundColor("#333333");
      const acoesPauta = (p.acoes||[]).filter(a => a.descricao && a.descricao.trim());
      if (acoesPauta.length) {
        body.appendParagraph("Ações:").setBold(true).setFontSize(10).setForegroundColor("#1a5c35");
        acoesPauta.forEach((a, ai) => {
          const prazo = a.prazo ? _data(a.prazo) : "Sem prazo";
          body.appendParagraph(`   ${ai+1}. ${a.descricao}`).setFontSize(10).setBold(false);
          body.appendParagraph(`      👤 ${a.responsavelNome}   |   📅 ${prazo}`)
            .setForegroundColor("#888888").setFontSize(9);
        });
      } else {
        body.appendParagraph("Nenhuma ação definida.").setFontSize(10).setForegroundColor("#aaaaaa");
      }
    });
  } else {
    const topicos = dados.topicos || [];
    const acoes   = dados.acoes || [];
    if (!topicos.length && !acoes.length) body.appendParagraph("Nenhum assunto ou ação registrada.");
    topicos.forEach((t, i) => body.appendParagraph(`${i+1}. ${t}`).setBold(true).setFontSize(12).setForegroundColor("#1a5c35").setSpacingAfter(8));
    if (acoes.length) {
      body.appendParagraph("Ações:").setBold(true).setFontSize(10).setForegroundColor("#1a5c35").setSpacingBefore(8);
      acoes.forEach((a, i) => {
        if (!a.descricao || !a.descricao.trim()) return;
        const prazo = a.prazo ? _data(a.prazo) : "Sem prazo";
        body.appendParagraph(`${i+1}. ${a.descricao.trim()}`).setBold(true).setFontSize(10);
        body.appendParagraph(`   👤 ${a.responsavelNome||"Sem responsável"}   |   📅 Prazo: ${prazo}`)
          .setForegroundColor("#666666").setFontSize(9).setSpacingAfter(6);
      });
    }
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
// ENVIA ATA + AÇÕES POR WHATSAPP
// ============================================================
function enviarAtaWhatsApp(dados, linkDrive) {
  const dataFmt = _data(dados.data);
  let msgAta = `📋 *ATA DE REUNIÃO*\n*${dados.tipo}* — ${dataFmt}\n\n`;
  msgAta += `*Participantes:*\n` + (dados.participantes||[]).map(x => `• ${x.nome}`).join("\n");
  msgAta += `\n\n*ASSUNTOS DISCUTIDOS / AÇÕES*\n`;

  const pautas = dados.pautas || [];
  if (pautas.length) {
    pautas.forEach((p, pi) => {
      const titulo = p.titulo?.trim() || `Assunto ${pi+1}`;
      msgAta += `\n*${pi+1}. ${titulo}*\n`;
      if (p.discussao?.trim()) msgAta += `${p.discussao.trim()}\n`;
      const acoesPauta = (p.acoes||[]).filter(a => a.descricao?.trim());
      if (acoesPauta.length) {
        msgAta += `\n_Ações:_\n`;
        acoesPauta.forEach((a, ai) => {
          msgAta += `${ai+1}. ${a.descricao.trim()}\n`;
          msgAta += `   👤 ${a.responsavelNome||"Sem responsável"} | 📅 ${a.prazo ? _data(a.prazo) : "Sem prazo"}\n`;
        });
      } else {
        msgAta += `_Nenhuma ação definida para este assunto._\n`;
      }
      msgAta += `\n`;
    });
  } else {
    (dados.topicos||[]).forEach((t,i) => msgAta += `\n*${i+1}. ${t}*\n`);
    const acoes = dados.acoes||[];
    if (acoes.length) {
      msgAta += `\n_Ações:_\n`;
      acoes.forEach((a,i) => {
        if (!a.descricao?.trim()) return;
        msgAta += `${i+1}. ${a.descricao.trim()}\n`;
        msgAta += `   👤 ${a.responsavelNome||"Sem responsável"} | 📅 ${a.prazo ? _data(a.prazo) : "Sem prazo"}\n`;
      });
    }
  }
  msgAta += `\n📄 *Ata completa em PDF:*\n${linkDrive}`;

  (dados.participantes||[]).forEach(p => {
    if (!p.telefone) return;
    _whatsapp(p.telefone, msgAta);
    Utilities.sleep(600);
  });

  const acoesPorPessoa = {};
  if (pautas.length) {
    pautas.forEach((p, pi) => {
      const tituloPauta = p.titulo?.trim() || `Assunto ${pi+1}`;
      (p.acoes||[]).forEach(a => {
        if (!a.descricao?.trim() || !a.responsavelTelefone) return;
        if (!acoesPorPessoa[a.responsavelTelefone]) acoesPorPessoa[a.responsavelTelefone] = { nome: a.responsavelNome||"", acoes: [] };
        acoesPorPessoa[a.responsavelTelefone].acoes.push({ descricao: a.descricao.trim(), prazo: a.prazo, pauta: tituloPauta });
      });
    });
  } else {
    (dados.acoes||[]).forEach(a => {
      if (!a.descricao?.trim() || !a.responsavelTelefone) return;
      if (!acoesPorPessoa[a.responsavelTelefone]) acoesPorPessoa[a.responsavelTelefone] = { nome: a.responsavelNome||"", acoes: [] };
      acoesPorPessoa[a.responsavelTelefone].acoes.push({ descricao: a.descricao.trim(), prazo: a.prazo, pauta: "" });
    });
  }

  Object.entries(acoesPorPessoa).forEach(([tel, pessoa]) => {
    const primeiroNome = pessoa.nome ? pessoa.nome.split(" ")[0] : "Olá";
    let msgAcoes = `📌 *Suas ações — ${dados.tipo}*\n\nOlá, ${primeiroNome}! Aqui estão as ações que ficaram com você:\n\n`;
    pessoa.acoes.forEach((a, i) => {
      if (a.pauta) msgAcoes += `*Assunto:* ${a.pauta}\n`;
      msgAcoes += `${i+1}. ${a.descricao}\n   📅 Prazo: ${a.prazo ? _data(a.prazo) : "Sem prazo"}\n\n`;
    });
    msgAcoes += `Você receberá lembretes automáticos antes do prazo.\n*Digite 1* quando concluir ✅ | *Digite 0* se ainda pendente ⏳`;
    _whatsapp(tel, msgAcoes);
    Utilities.sleep(600);
  });
}

// ============================================================
// LEMBRETE DIÁRIO DE AÇÕES — roda todo dia às 8h
// ============================================================
function enviarLembretesAcoes() {
  const diaSemana = new Date().getDay();
  if (diaSemana === 0 || diaSemana === 6) return;

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) return;
  const linhas = abaAcoes.getDataRange().getValues();
  const hoje   = new Date(); hoje.setHours(0, 0, 0, 0);

  const porTelefone = {};
  for (let i = 1; i < linhas.length; i++) {
    const [, reuniao, , acao, responsavel, telefone, prazoRaw, status] = linhas[i];
    if (status === "Concluída" || !telefone) continue;

    const prazoDate = prazoRaw instanceof Date
      ? new Date(prazoRaw.getFullYear(), prazoRaw.getMonth(), prazoRaw.getDate())
      : prazoRaw ? new Date(String(prazoRaw).split("T")[0] + "T12:00:00") : null;

    const diff = prazoDate ? Math.round((prazoDate - hoje) / 864e5) : null;
    const prazoFmt = prazoDate
      ? Utilities.formatDate(prazoDate, "America/Sao_Paulo", "dd/MM/yyyy")
      : "Sem prazo";

    let urgencia = "📌";
    if (diff === 1)      urgencia = "⏰";
    else if (diff === 0) urgencia = "🔴";
    else if (diff !== null && diff < 0) urgencia = "🚨";

    const tel = String(telefone);
    if (!porTelefone[tel]) porTelefone[tel] = { responsavel, acoes: [], linhas: [] };
    porTelefone[tel].acoes.push({ acao, prazoFmt, urgencia, diff });
    porTelefone[tel].linhas.push(i);
  }

  Object.entries(porTelefone).forEach(([tel, pessoa]) => {
    const primeiroNome = pessoa.responsavel.split(" ")[0];
    let msg = `📋 *Lembretes de ações pendentes*\n\nOlá, ${primeiroNome}! Você tem ${pessoa.acoes.length} ação(ões) pendente(s):\n\n`;

    pessoa.acoes.forEach((a, idx) => {
      const atrasado = a.diff !== null && a.diff < 0 ? ` — 🚨 ATRASADA ${Math.abs(a.diff)} dia(s)` : "";
      msg += `*${idx+1}.* ${a.urgencia} ${a.acao}\n   📅 Prazo: ${a.prazoFmt}${atrasado}\n\n`;
    });

    if (pessoa.acoes.length === 1) {
      msg += `*Digite 1* para confirmar que concluiu ✅\n*Digite 0* se ainda está em andamento ⏳`;
    } else {
      msg += `Responda com o *número da ação* que concluiu (ex: *1*, *2*, *3*)\nOu *0* se nenhuma foi concluída ainda ⏳`;
    }

    _whatsapp(tel, msg);
    const agora = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    pessoa.linhas.forEach(i => abaAcoes.getRange(i+1, 10).setValue(agora));
    Utilities.sleep(800);
  });
}

// ============================================================
// LEMBRETE DE REUNIÃO — roda a cada hora
// Interpreta códigos: 7H, 1H, XH, 1D, 2D e combinações (ex: 7H/1H)
// ============================================================
function enviarLembreteReuniao() {
  const agora      = new Date();
  const diaSemana  = agora.getDay();
  if (diaSemana === 0 || diaSemana === 6) return; // só seg-sex

  const configs        = buscarConfiguracoes();
  const mapDia         = { "Domingo":0,"Segunda":1,"Terça":2,"Quarta":3,"Quinta":4,"Sexta":5,"Sábado":6 };
  const horaAgora      = parseInt(Utilities.formatDate(agora, "America/Sao_Paulo", "H"));
  const minAgora       = parseInt(Utilities.formatDate(agora, "America/Sao_Paulo", "m"));
  const minTotalAgora  = horaAgora * 60 + minAgora;
  const diaBrasilia    = parseInt(Utilities.formatDate(agora, "America/Sao_Paulo", "u")) % 7;

  configs.forEach(c => {
    if (!c.ativo || !c.horario || !c.lembrete) return;
    if (c.frequencia === "Sem reunião periódica") return;

    const [hR, mR]   = c.horario.split(":").map(Number);
    const minReuniao = hR * 60 + mR;
    const codigos    = c.lembrete.split("/").map(s => s.trim().toUpperCase());

    codigos.forEach(cod => {
      let deveEnviar = false;
      let textoAviso = "";

      if (cod.endsWith("D")) {
        // XD antes — avisa às 8h do dia
        const diasAntes = parseInt(cod.replace("D",""));
        if (!isNaN(diasAntes) && horaAgora === 8 && minAgora < 60) {
          const futuro = new Date(agora);
          futuro.setDate(futuro.getDate() + diasAntes);
          if (_ehHojeReuniao_(c, futuro, futuro.getDay(), mapDia)) {
            deveEnviar = true;
            textoAviso = diasAntes === 1 ? `*amanhã* às *${c.horario}h*` : `*em ${diasAntes} dias* (${c.horario}h)`;
          }
        }
      } else if (cod === "7H") {
        // Avisa às 7h do dia DA reunião
        if (horaAgora === 7 && minAgora < 60 && _ehHojeReuniao_(c, agora, diaBrasilia, mapDia)) {
          deveEnviar = true;
          textoAviso = `hoje às *${c.horario}h*`;
        }
      } else if (cod.endsWith("H")) {
        // XH antes — avisa X horas antes do evento (janela de 30min)
        const hAntes = parseInt(cod.replace("H",""));
        if (!isNaN(hAntes) && _ehHojeReuniao_(c, agora, diaBrasilia, mapDia)) {
          const diff = minReuniao - minTotalAgora;
          const minAntes = hAntes * 60;
          if (diff >= minAntes - 30 && diff <= minAntes + 30) {
            deveEnviar = true;
            textoAviso = hAntes === 1 ? `em *1 hora* (${c.horario}h)` : `em *${hAntes} horas* (${c.horario}h)`;
          }
        }
      }

      if (!deveEnviar) return;

      buscarParticipantesArea(c.titulo).forEach(p => {
        if (!p.telefone) return;
        _whatsapp(p.telefone,
          `📅 *Lembrete de reunião — ${textoAviso}*\n\nOlá, ${p.nome}!\n\nVocê tem a reunião *${c.titulo}* ${textoAviso}.\n\nPrepare os tópicos que deseja discutir. 💪`
        );
        Utilities.sleep(600);
      });
    });
  });
}

// ============================================================
// AUXILIAR — verifica se hoje é dia desta reunião
// ============================================================
function _ehHojeReuniao_(c, data, diaSemana, mapDia) {
  const hoje = new Date(data); hoje.setHours(0,0,0,0);

  if (c.frequencia === "DATA") {
    if (!c.dataUnica) return false;
    const dataEvento = c.dataUnica instanceof Date ? new Date(c.dataUnica) : _parsarData_(String(c.dataUnica));
    if (!dataEvento) return false;
    dataEvento.setHours(0,0,0,0);
    return dataEvento.getTime() === hoje.getTime();
  }

  if (c.frequencia === "Semanal") return mapDia[c.dia] === diaSemana;

  if (c.frequencia === "Quinzenal") {
    if (mapDia[c.dia] !== diaSemana) return false;
    if (!c.dataAlteracao) return true;
    const ref = new Date(c.dataAlteracao); ref.setHours(0,0,0,0);
    return Math.round((hoje - ref) / 864e5) % 14 === 0;
  }

  if (c.frequencia === "Mensal" || c.frequencia === "ULT/MES") {
    if (mapDia[c.dia] !== diaSemana) return false;
    return _ehUltimaOcorrenciaNoMes_(hoje);
  }

  if (c.frequencia === "ULT/3M") {
    if (mapDia[c.dia] !== diaSemana) return false;
    const mes = hoje.getMonth();
    if (![2,5,8,11].includes(mes)) return false;
    return _ehUltimaOcorrenciaNoMes_(hoje);
  }

  return false;
}

function _ehUltimaOcorrenciaNoMes_(data) {
  const proxSemana = new Date(data);
  proxSemana.setDate(proxSemana.getDate() + 7);
  return proxSemana.getMonth() !== data.getMonth();
}

function _parsarData_(str) {
  if (!str) return null;
  const partes = String(str).split("/");
  if (partes.length !== 3) return null;
  return new Date(parseInt(partes[2]), parseInt(partes[1])-1, parseInt(partes[0]));
}

// ============================================================
// PROCESSA RESPOSTA WHATSAPP
// ============================================================
function processarRespostaWhatsApp(telefone, mensagem) {
  const resp  = String(mensagem).trim();
  const num   = parseInt(resp);
  const isNao = resp === "0" || resp.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"") === "NAO";

  if (isNaN(num) && !isNao) return;
  if (num < 1 && !isNao) return;

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaAcoes = ss.getSheetByName("Ações");
  if (!abaAcoes) return;
  const linhas   = abaAcoes.getDataRange().getValues();
  const telBusca = String(telefone).replace(/\D/g, "").slice(-8);

  if (isNao) {
    _whatsapp(telefone, `⏳ Entendido! As ações continuam como pendentes.`);
    return;
  }

  const acoesPendentes = [];
  for (let i = 1; i < linhas.length; i++) {
    const telPlan = String(linhas[i][5]).replace(/\D/g, "").slice(-8);
    if (telPlan === telBusca && linhas[i][7] === "Pendente") acoesPendentes.push(i);
  }

  if (!acoesPendentes.length) return;

  let linhaAlvo;
  if (num <= acoesPendentes.length) {
    linhaAlvo = acoesPendentes[num - 1];
  } else {
    let melhorData = null;
    acoesPendentes.forEach(i => {
      const cob = linhas[i][9];
      if (cob) {
        const d = new Date(cob);
        if (!melhorData || d > melhorData) { melhorData = d; linhaAlvo = i; }
      }
    });
    if (!linhaAlvo) linhaAlvo = acoesPendentes[0];
  }

  const nomeAcao = linhas[linhaAlvo][3];
  const nomeResp = linhas[linhaAlvo][4];

  abaAcoes.getRange(linhaAlvo+1, 8).setValue("Concluída");
  abaAcoes.getRange(linhaAlvo+1, 9).setValue(Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy"));
  abaAcoes.getRange(linhaAlvo+1, 10).setValue("");

  _whatsapp(telefone, `✅ Ação marcada como *concluída*!\n\n📌 _${nomeAcao}_\n\nBom trabalho, ${nomeResp.split(" ")[0]}! 👏`);
}

// ============================================================
// GOOGLE AGENDA (desabilitado)
// ============================================================
function criarProximaReuniaoCalendario(dados, linkDrive) {
  // Convite de calendário desabilitado
}

// ============================================================
// GATILHOS — rode UMA VEZ só
// ============================================================
function configurarTodosGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["enviarLembretesAcoes","enviarLembreteReuniao","enviarLembretes"].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("enviarLembretesAcoes")
    .timeBased().everyDays(1).atHour(8)
    .inTimezone("America/Sao_Paulo").create();
  ScriptApp.newTrigger("enviarLembreteReuniao")
    .timeBased().everyHours(1)
    .inTimezone("America/Sao_Paulo").create();
  Logger.log("✅ Gatilhos configurados!");
  Logger.log("   • Lembretes de ações: dias úteis às 8h");
  Logger.log("   • Lembretes de reunião: a cada hora");
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
    payload: JSON.stringify({ number: numero, textMessage: { text: mensagem } }),
    muteHttpExceptions: true,
  };
  try {
    const resp = UrlFetchApp.fetch(`${EVOLUTION_URL}/message/sendText/${EVOLUTION_INSTANCE}`, opts);
    Logger.log("Evolution texto " + telefone + ": " + resp.getContentText());
  } catch (err) {
    Logger.log("Erro texto " + telefone + ": " + err.message);
  }
}

function buscarTodasAcoes() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName("Ações");
  if (!aba) return [];
  return aba.getDataRange().getValues().slice(1).filter(r => r[3]).map(r => ({
    area:          String(r[1]||""),
    dataReuniao:   r[2] instanceof Date ? Utilities.formatDate(r[2],"America/Sao_Paulo","yyyy-MM-dd") : String(r[2]||""),
    acao:          String(r[3]||""),
    responsavel:   String(r[4]||""),
    prazo:         r[6] instanceof Date ? Utilities.formatDate(r[6],"America/Sao_Paulo","yyyy-MM-dd") : String(r[6]||""),
    status:        String(r[7]||"Pendente"),
    dataConclusao: r[8] instanceof Date ? Utilities.formatDate(r[8],"America/Sao_Paulo","yyyy-MM-dd") : String(r[8]||""),
  }));
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

function _log(msg) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    let aba   = ss.getSheetByName("Logs");
    if (!aba) { aba = ss.insertSheet("Logs"); aba.appendRow(["Data","Mensagem"]); }
    aba.appendRow([new Date().toLocaleString("pt-BR"), String(msg).substring(0, 500)]);
  } catch(e) {}
}

// ============================================================
// TESTES
// ============================================================
function testeWhatsApp() {
  _whatsapp("31994599539", "✅ Teste de texto simples — funcionando!");
}

function testeLembreteReuniao() {
  Logger.log("Testando lembretes de reunião...");
  enviarLembreteReuniao();
  Logger.log("Concluído.");
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
