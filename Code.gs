
 * WFS SHUTTLE - V-0.4

// --- CONFIGURAÇÕES E IDs ---
const ID_PLANILHA_DEMITIDOS = "ID_planilha"; 
const ID_BASE_EFETIVOS = "ID_planilha"; 

const ABA_BASE = 'BASE_GERAL';
const ABA_HIST = 'HISTORICO';
const ABA_SOLIC = 'SOLICITACOES';
const ABA_CONFIG = 'CONFIG';

const MAPA_GERENTES = {
  "CANAL": "Gestor", "CARGA": "Gestor", "SEGURANÇA": "Gestor",
  "RAMPA NARROW": "Gestor", "RAMPA WIDE": "Gestor", "LIMPEZA": "Gestor",
  "ESTEIRA": "Gestor", "PASSAGEIROS": "Gestor"
};

const COL = {
  CARRO: 0, VAN: 1, MOT: 2, MATRICULA: 3, NOME: 4, FONE: 5, CEP: 6, RUA: 7, 
  NUM: 8, BAIRRO: 9, TIPO: 10, HORARIO: 11, FOCAL: 43, GERENTE: 44, 
  AREA: 45, CONTRATO: 46, DATA: 47, FERIAS_INI: 48, FERIAS_FIM: 49 
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('WFS SHUTTLE').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function salvarColagemInteligente(dadosRecebidos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(ABA_BASE);
    const dadosDb = ws.getDataRange().getValues();
    let alterados = 0;
    let setorEditado = ""; // Sensor de qual setor está sendo editado

    let mapaLinhas = {};
    for(let i=1; i<dadosDb.length; i++) {
        let mat = String(dadosDb[i][COL.MATRICULA]).trim();
        if(mat) mapaLinhas[mat] = i + 1; 
    }

    dadosRecebidos.forEach(item => {
        let linha = mapaLinhas[String(item.matricula)];
        if (linha) {
            let diasParaSalvar = item.dias;
            while (diasParaSalvar.length < 31) diasParaSalvar.push("");
            diasParaSalvar = diasParaSalvar.slice(0, 31);
            ws.getRange(linha, 13, 1, diasParaSalvar.length).setValues([diasParaSalvar]);
            alterados++;
            setorEditado = dadosDb[linha - 1][COL.AREA]; // Descobre a área
        }
    });

    // NOVIDADE: Se colou com sucesso, avisa o painel do Admin!
    if (alterados > 0 && setorEditado) registrarAtividadeSetor(setorEditado);

    return { success: true, msg: `${alterados} escalas atualizadas com sucesso!` };
  } catch (e) { return { success: false, msg: "Erro no servidor: " + e.message }; }
}

// --- FUNÇÕES PADRÃO ---
function setLastUpdate() {
  let hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM HH:mm");
  PropertiesService.getScriptProperties().setProperty('LAST_UPLOAD', hoje); return hoje;
}
function getLastUpdate() { return PropertiesService.getScriptProperties().getProperty('LAST_UPLOAD') || "Pendente"; }
function getListaSetores() { return Object.keys(MAPA_GERENTES).concat(["ADMINISTRADOR"]); }

function tentarLogin(setor, nome, senha) {
  // 1. Verificação do Administrador (Pede senha)
  if (setor === "ADMINISTRADOR") { 
      if (senha === "SENHA ADMIN") return { sucesso: true, perfil: 'ADMIN', setor: 'VISÃO GLOBAL' }; 
      return { erro: "Senha de Admin incorreta." }; 
  }
  
  // 2. Verificação dos Gestores/Focais (Pede apenas o nome)
  if (!nome || nome.length < 3) return { erro: "Digite seu nome completo." }; 
  
  return { sucesso: true, perfil: 'FOCAL', setor: setor };
}

function setPrazoSistema(ini, fim) {
  const ws = SpreadsheetApp.getActive().getSheetByName(ABA_CONFIG); 
  ws.getRange("A2").setValue(ini); ws.getRange("B2").setValue(fim); 
  // NOVIDADE: Zera a memória de atividades sempre que o Admin atualiza o prazo!
  PropertiesService.getScriptProperties().setProperty('SETORES_ATIVOS', "[]"); 
  return "Período Atualizado!";
}

// NOVIDADE: O Rastreador Silencioso
function registrarAtividadeSetor(setor) {
    if (!setor || setor === "ADMINISTRADOR" || setor === "VISÃO GLOBAL") return;
    try {
        let prop = PropertiesService.getScriptProperties();
        let ativos = JSON.parse(prop.getProperty('SETORES_ATIVOS') || "[]");
        if (!ativos.includes(setor)) {
            ativos.push(setor);
            prop.setProperty('SETORES_ATIVOS', JSON.stringify(ativos));
        }
    } catch(e) {}
}
function getPrazoSistema() {
  const ws = SpreadsheetApp.getActive().getSheetByName(ABA_CONFIG); const d = ws.getRange("A2:B2").getValues()[0];
  if(!d[0] || !d[1]) return { inicio: "", fim: "", isLocked: false };
  const h = new Date(); h.setHours(0,0,0,0); const i = new Date(d[0]); i.setHours(0,0,0,0); const f = new Date(d[1]); f.setHours(23,59,59,999);
  return { inicio: Utilities.formatDate(i, Session.getScriptTimeZone(), "yyyy-MM-dd"), fim: Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"), isLocked: (h < i || h > f) };
}
function buscarNomeAtivo(mat) {
  const strMat = String(mat).trim().toUpperCase(); const ss = SpreadsheetApp.getActive(); let wsBase = ss.getSheetByName(ABA_BASE);
  if(wsBase.getLastRow() > 1) { let f = wsBase.getRange(2, COL.MATRICULA+1, wsBase.getLastRow()-1, 2).getValues().find(r => String(r[0]).trim().toUpperCase() === strMat); if(f) return { status: 'EXISTE_BASE', nome: f[1] }; }
  try { const shAtivos = SpreadsheetApp.openById(ID_BASE_EFETIVOS).getSheetByName("ATIVOS"); if(shAtivos && shAtivos.getLastRow() > 1) { let f = shAtivos.getRange(2, 1, shAtivos.getLastRow()-1, 2).getValues().find(r => String(r[0]).trim().toUpperCase() === strMat); if(f) return { status: 'ENCONTRADO', nome: f[1] }; } } catch(e) {} return null;
}

function carregarDadosSetor(setorUser, perfilUser) {
  try {
    const ss = SpreadsheetApp.getActive(); 
    let ws = ss.getSheetByName(ABA_BASE);
    const data = ws.getDataRange().getValues(); 
    const systemInfo = getPrazoSistema(); 
    const lastUpdate = getLastUpdate();
    
    let solicitacoes = [];
    if(perfilUser === 'ADMIN') {
        let wsSol = ss.getSheetByName(ABA_SOLIC); 
        if(wsSol && wsSol.getLastRow() > 1) {
           let dSol = wsSol.getDataRange().getValues();
           for(let i=1; i<dSol.length; i++) {
               // Pesca as folgas da aba de solicitação
               let folgasSol = [];
               for(let f=12; f<=42; f++) { if(dSol[i][f] === 'F') folgasSol.push(f-11); }
               
               solicitacoes.push({ 
                   linha: i+1, matricula: dSol[i][COL.MATRICULA], nome: dSol[i][COL.NOME], area: dSol[i][COL.AREA], carro: dSol[i][COL.CARRO], status: 'PENDENTE', 
                   tipo: dSol[i][COL.TIPO], horario: formatarHora(dSol[i][COL.HORARIO]), fone: dSol[i][COL.FONE], bairro: dSol[i][COL.BAIRRO],
                   cep: dSol[i][COL.CEP], rua: dSol[i][COL.RUA], num: dSol[i][COL.NUM], contrato: dSol[i][COL.CONTRATO], folgas: folgasSol
               });
           }
        }
    }
    
    let demitidosIDs = [];
    try { const shDem = SpreadsheetApp.openById(ID_PLANILHA_DEMITIDOS).getSheetByName("DEMITIDOS"); if(shDem) demitidosIDs = shDem.getRange("A2:A").getValues().flat().map(String); } catch(e) {}
    
    let stats = { total: 0, novatos: 0, alertas: 0, semrota: 0, ferias: 0 };
    let filtrados = [], contratosUnicos = new Set(), listaFeriasKPI = [], hoje = new Date();

    for(let i=1; i<data.length; i++) {
      let row = data[i]; if(row.length < 46) continue;
      
      if ((perfilUser === 'ADMIN') || (String(row[COL.AREA]||"").trim().toUpperCase() === String(setorUser||"").trim().toUpperCase())) {
        let contrato = String(row[COL.CONTRATO]||"").trim(); if(contrato) contratosUnicos.add(contrato);
        let mat = String(row[COL.MATRICULA] || ""); if(!mat) continue;
        
        let isNovato = !isNaN(new Date(row[COL.DATA]).getTime()) && ((hoje - new Date(row[COL.DATA]))/(8.64e7) < 30);
        let isDem = demitidosIDs.includes(mat);
        
        let ic = String(row[COL.CARRO]||"").trim().toUpperCase();
        let iv = String(row[COL.VAN]||"").trim().toUpperCase();
        let icLimpo = ic.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
        let ivLimpo = iv.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
        
        let isSemRota = (
            icLimpo === "" || icLimpo === "A DEFINIR" || icLimpo === "UNDEFINED" || icLimpo === "NULL" || 
            icLimpo.includes("SEM VAGA") || icLimpo.includes("SEM ROTA") || icLimpo.includes("NAO TEM") || 
            ivLimpo.includes("SEM ROTA") || ivLimpo.includes("SEM VAGA") || ivLimpo.includes("NAO TEM")
        );
        
        let statusFinal = isDem ? "DEMITIDO" : (isSemRota ? "SEM ROTA" : "OK");

        let folgas = [], diasGrid = []; 
        for(let f=12; f<=42; f++) { 
            let v = row[f]; 
            if(v === 'F') { folgas.push(f-11); diasGrid.push("F"); } 
            else if(v === 'FE') { diasGrid.push("FE"); } 
            else { diasGrid.push(""); } 
        }
        
        stats.total++; if(isNovato) stats.novatos++; if(isDem) stats.alertas++; 
        if(isSemRota && !isDem) stats.semrota++;
        
        let fIni = row[COL.FERIAS_INI] ? (row[COL.FERIAS_INI] instanceof Date ? Utilities.formatDate(row[COL.FERIAS_INI], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[COL.FERIAS_INI]).substring(0,10)) : "";
        let fFim = row[COL.FERIAS_FIM] ? (row[COL.FERIAS_FIM] instanceof Date ? Utilities.formatDate(row[COL.FERIAS_FIM], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[COL.FERIAS_FIM]).substring(0,10)) : "";

        filtrados.push({ 
            linha: i+1, matricula: mat, nome: String(row[COL.NOME]||"").toUpperCase(), status: statusFinal, situacao: isNovato ? "NOVATO" : "FIXO", 
            carro: row[COL.CARRO], van: row[COL.VAN], mot: row[COL.MOT], fone: row[COL.FONE], cep: row[COL.CEP], rua: row[COL.RUA], num: row[COL.NUM], 
            bairro: row[COL.BAIRRO], tipo: row[COL.TIPO], horario: formatarHora(row[COL.HORARIO]), contrato: contrato, focal: row[COL.FOCAL], 
            area: row[COL.AREA], folgas: folgas, diasGrid: diasGrid,
            feriasIni: fIni, feriasFim: fFim
        });
      }
    }
    
    let wsHist = ss.getSheetByName(ABA_HIST);
    if(wsHist && wsHist.getLastRow() > 1) {
       let dH = wsHist.getDataRange().getValues();
       for(let h=1; h<dH.length; h++) {
           if(String(dH[h][COL.CARRO]).toUpperCase().includes("FÉRIAS")) {
               if (perfilUser === 'ADMIN' || String(dH[h][COL.AREA]).toUpperCase() === String(setorUser).toUpperCase()) {
                   let cH = String(dH[h][COL.CONTRATO]||"").trim(); if(cH) contratosUnicos.add(cH);
                   listaFeriasKPI.push({ area: dH[h][COL.AREA], contrato: cH }); 

                   let hIni = dH[h][COL.FERIAS_INI] ? (dH[h][COL.FERIAS_INI] instanceof Date ? Utilities.formatDate(dH[h][COL.FERIAS_INI], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dH[h][COL.FERIAS_INI]).substring(0,10)) : "";
                   let hFim = dH[h][COL.FERIAS_FIM] ? (dH[h][COL.FERIAS_FIM] instanceof Date ? Utilities.formatDate(dH[h][COL.FERIAS_FIM], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dH[h][COL.FERIAS_FIM]).substring(0,10)) : "";

                   filtrados.push({ 
                       linha: "HIST:"+(h+1), matricula: dH[h][COL.MATRICULA], nome: String(dH[h][COL.NOME]).toUpperCase(), status: "FÉRIAS", situacao: "EM FÉRIAS", 
                       carro: "EM FÉRIAS", van: dH[h][COL.VAN], mot: "", fone: dH[h][COL.FONE], cep: dH[h][COL.CEP], rua: dH[h][COL.RUA], num: dH[h][COL.NUM], 
                       bairro: dH[h][COL.BAIRRO], tipo: dH[h][COL.TIPO], horario: formatarHora(dH[h][COL.HORARIO]), contrato: cH, focal: dH[h][COL.FOCAL], 
                       area: dH[h][COL.AREA], folgas: [], diasGrid: new Array(31).fill(""),
                       feriasIni: hIni, feriasFim: hFim
                   });
               }
           }
       }
    }

    // --- NOVIDADE: LEITURA DO RASTREADOR SILENCIOSO ---
    let statusEnvioSetores = [];
    if (perfilUser === 'ADMIN') {
        let setoresAtivosCache = JSON.parse(PropertiesService.getScriptProperties().getProperty('SETORES_ATIVOS') || "[]");
        let mapa = {}; 
        for(let i=1; i<data.length; i++) {
            let s = String(data[i][COL.AREA]||"").trim().toUpperCase(); if(!s || s==="ADMINISTRADOR") continue;
            if(!mapa[s]) mapa[s] = {t:0}; mapa[s].t++;
        }
        for (let s in mapa) { 
            let editou = setoresAtivosCache.includes(s); // Verifica se o setor caiu no sensor
            statusEnvioSetores.push({ 
                setor: s, 
                total: mapa[s].t, 
                status: editou ? "ATUALIZADO" : "PENDENTE", 
                percentual: editou ? 100 : 0 
            }); 
        }
        statusEnvioSetores.sort((a, b) => a.status.localeCompare(b.status));
    }
    // ----------------------------------------------------
    
    filtrados.sort((a, b) => a.nome.localeCompare(b.nome));
    return { lista: filtrados, stats: stats, contratos: Array.from(contratosUnicos).sort(), system: systemInfo, solicitacoes: solicitacoes, historicoFerias: listaFeriasKPI, lastUpdate: lastUpdate, monitorEnvio: statusEnvioSetores };
  } catch (e) { return { erro: e.toString() }; }
}

// --- VERSÃO ATUALIZADA: RESTAURAÇÃO DE FÉRIAS ---
function salvarDados(d, nomeFocal, setorFocal, perfilUser) { 
  try { 
    const ss = SpreadsheetApp.getActive(); 
    const prazos = getPrazoSistema(); 
    let isSolicitacao = false; 
    
    if (prazos.isLocked && perfilUser !== 'ADMIN') { 
        if (d.isEdit) return "Sistema Fechado!"; 
        isSolicitacao = true; 
    } 
    
    let target = isSolicitacao ? ABA_SOLIC : ABA_BASE; 
    let ws = ss.getSheetByName(target); 

    // NOVIDADE: O sistema descobre se o gestor está puxando alguém de volta do Histórico
    let isRestaurandoHist = d.isEdit && String(d.linha).startsWith("HIST:");
    let histLinha = isRestaurandoHist ? Number(String(d.linha).replace("HIST:", "")) : null;
    
    let linha = (d.isEdit && !isSolicitacao && !isRestaurandoHist) ? Number(d.linha) : ws.getLastRow() + 1; 
    let dados = new Array(50).fill(""); 
    
    if(d.isEdit && target === ABA_BASE) { 
        let v;
        if (isRestaurandoHist) {
            v = ss.getSheetByName(ABA_HIST).getRange(histLinha, 1, 1, 48).getValues()[0];
        } else {
            v = ws.getRange(linha, 1, 1, 48).getValues()[0];
        }
        for(let k=0; k<v.length; k++) dados[k] = v[k]; 
    } else if (!d.isEdit && !isSolicitacao) { 
        if(ws.getDataRange().getValues().some(r => String(r[COL.MATRICULA]) == String(d.matricula))) return "Duplicado!"; 
    } 
    
    dados[COL.MATRICULA] = d.matricula; 
    dados[COL.NOME] = String(d.nome).toUpperCase(); 
    dados[COL.FONE] = d.telefone; 
    dados[COL.CEP] = d.cep; 
    dados[COL.RUA] = String(d.endereco).toUpperCase(); 
    dados[COL.NUM] = d.numero; 
    dados[COL.BAIRRO] = String(d.bairro).toUpperCase(); 
    dados[COL.TIPO] = d.tipo; 
    dados[COL.HORARIO] = d.horario; 
    dados[COL.CONTRATO] = String(d.contrato).toUpperCase(); 
    
    if (!d.isEdit && !dados[COL.CARRO]) dados[COL.CARRO] = "A DEFINIR"; 
    
    // Se voltar do histórico de férias, tira o nome "EM FÉRIAS" do carro
    if (isRestaurandoHist && String(dados[COL.CARRO]).includes("FÉRIAS")) dados[COL.CARRO] = "A DEFINIR";

    for(let f=12; f<=42; f++) dados[f] = ""; 
    
    if(d.isFerias && d.feriasIni && d.feriasFim && !isSolicitacao) { 
        let hoje = new Date(); hoje.setHours(0,0,0,0); 
        let ini = new Date(d.feriasIni+"T00:00:00"); 
        let fim = new Date(d.feriasFim+"T23:59:59"); 
        
        if (hoje >= ini && hoje <= fim) { 
            let wsHist = ss.getSheetByName(ABA_HIST); 
            dados[COL.CARRO] = "EM FÉRIAS"; 
            while(dados.length < 50) dados.push("");
            dados[50] = "FÉRIAS";
            
            if (wsHist.getMaxColumns() < 51) wsHist.insertColumnsAfter(wsHist.getMaxColumns(), 51 - wsHist.getMaxColumns());
            wsHist.appendRow(dados); 
            
            if(d.isEdit && !isRestaurandoHist) ws.deleteRow(Number(d.linha)); 
            if(isRestaurandoHist) wsHist.deleteRow(histLinha); // Se editou mas manteve de férias, atualiza no hist
            return "Enviado p/ Histórico (Férias)."; 
        } else { 
            dados[COL.FERIAS_INI] = d.feriasIni; 
            dados[COL.FERIAS_FIM] = d.feriasFim; 
            let hY = ini.getFullYear(), hM = ini.getMonth(); 
            for(let dia=1; dia<=31; dia++) { 
                let dt = new Date(hY, hM, dia); 
                if(dt >= ini && dt <= fim && 11+dia <= 42) dados[11+dia] = "FE"; 
            } 
        } 
    } else {
        // Se o gestor desmarcou a caixinha, limpamos as datas para não dar erro futuro
        dados[COL.FERIAS_INI] = ""; 
        dados[COL.FERIAS_FIM] = ""; 
    }
    
    if(d.folgas) d.folgas.forEach(dia => { let idx = 11 + parseInt(dia); if(idx <= 42 && dados[idx] !== "FE") dados[idx] = "F"; }); 
    
    dados[COL.AREA] = dados[COL.AREA] || (perfilUser === 'ADMIN' && d.setorManual ? d.setorManual : setorFocal); 
    dados[COL.FOCAL] = dados[COL.FOCAL] || nomeFocal; 
    
    if (MAPA_GERENTES[dados[COL.AREA]]) dados[COL.GERENTE] = MAPA_GERENTES[dados[COL.AREA]]; 
    if (!dados[COL.DATA]) dados[COL.DATA] = new Date(); 
    
    if(ws.getMaxColumns() < dados.length) ws.insertColumnsAfter(ws.getMaxColumns(), dados.length - ws.getMaxColumns()); 
    ws.getRange(linha, 1, 1, dados.length).setValues([dados]); 
    
    // NOVIDADE: Se voltou do histórico e NÃO está mais de férias ativas, apaga do Histórico de vez!
    if (isRestaurandoHist && (!d.isFerias || (d.isFerias && !(new Date() >= new Date(d.feriasIni+"T00:00:00") && new Date() <= new Date(d.feriasFim+"T23:59:59"))))) {
        ss.getSheetByName(ABA_HIST).deleteRow(histLinha);
    }
    
    // --- NOVIDADE: RASTREADOR SILENCIOSO DE ATIVIDADE ---
    // Se não for uma solicitação e quem estiver salvando não for o ADMIN, registra a atividade.
    if (!isSolicitacao && perfilUser !== 'ADMIN') {
        registrarAtividadeSetor(dados[COL.AREA]);
    }
    // ----------------------------------------------------
    
    return isSolicitacao ? "Solicitação enviada!" : "Salvo com Sucesso!"; 
  } catch(e) { 
    return "ERRO: " + e.toString(); 
  } 
}

function formatarHora(val) { if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm"); return String(val); }

function gerarExcelFormatado() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const originalSheet = ss.getSheetByName(ABA_BASE); 
  const tempSheet = originalSheet.copyTo(ss); tempSheet.setName("TEMP_" + new Date().getTime()); 
  let lastRow = tempSheet.getLastRow(); if (lastRow >= 2) { tempSheet.getRange(2, 1, lastRow - 1, 3).clearContent(); SpreadsheetApp.flush(); } 
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx&gid=" + tempSheet.getSheetId(); 
  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } }; 
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName("Escala_WFS.xlsx"); ss.deleteSheet(tempSheet); 
  return { base64: Utilities.base64Encode(blob.getBytes()), filename: "Escala_WFS.xlsx" }; 
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// FUNÇÃO DE BUSCA DE CEP

function buscarCEP(cep) { 
  if(!cep) return null; 
  cep = cep.replace(/\D/g, ''); 
  try { 
    let r = UrlFetchApp.fetch(`https://viacep.com.br/ws/${cep}/json/`); 
    let j = JSON.parse(r.getContentText()); 
    if(!j.erro) return { rua: j.logradouro, bairro: j.bairro }; 
  } catch(e) {} 
  try { 
    let r2 = UrlFetchApp.fetch(`https://brasilapi.com.br/api/cep/v1/${cep}`); 
    let j2 = JSON.parse(r2.getContentText()); 
    if(j2.street) return { rua: j2.street, bairro: j2.neighborhood }; 
  } catch(e){} 
  return null; 
}

function remover(ln, mat) { 
  try { 
    const ss = SpreadsheetApp.getActive(); 
    let wsBase = ss.getSheetByName(ABA_BASE); 
    let wsHist = ss.getSheetByName(ABA_HIST); 
    
    // NOVIDADE: O sistema agora sabe lidar com exclusões do Histórico (Férias)
    let isHist = String(ln).startsWith("HIST:");
    let targetSheet = isHist ? wsHist : wsBase;
    let numLinha = isHist ? Number(String(ln).replace("HIST:", "")) : Number(ln);
    
    if (isNaN(numLinha) || numLinha < 2) return "Erro: Linha inválida.";

    // --- MÁGICA DA SEGURANÇA ---
    let targetRow = numLinha;
    let currentMat = String(targetSheet.getRange(numLinha, COL.MATRICULA + 1).getValue()).trim();
    
    if (mat && currentMat !== String(mat).trim()) {
        let dataMat = targetSheet.getRange(1, COL.MATRICULA + 1, targetSheet.getLastRow(), 1).getValues();
        let found = false;
        for (let i = 1; i < dataMat.length; i++) {
            if (String(dataMat[i][0]).trim() === String(mat).trim()) {
                targetRow = i + 1;
                found = true;
                break;
            }
        }
        if (!found) return "Erro: Colaborador não encontrado na base para exclusão.";
    }

    if (isHist) {
        // Se a pessoa estava de férias, o sistema apenas troca o carimbo dela para EXCLUÍDO
        targetSheet.getRange(targetRow, 51).setValue("EXCLUÍDO");
    } else {
        // Se estava na Base normal, copia para o Histórico e apaga da Base
        let lastCol = wsBase.getLastColumn();
        let dadosOriginais = wsBase.getRange(targetRow, 1, 1, lastCol).getValues()[0]; 
        
        if(wsHist) {
            let novoDado = [];
            for(let i = 0; i < dadosOriginais.length; i++) novoDado.push(dadosOriginais[i]);
            while(novoDado.length < 50) novoDado.push(""); 
            novoDado[50] = "EXCLUÍDO"; 
            
            if (wsHist.getMaxColumns() < 51) wsHist.insertColumnsAfter(wsHist.getMaxColumns(), 51 - wsHist.getMaxColumns());
            wsHist.appendRow(novoDado); 
            
            let lr = wsHist.getLastRow();
            wsHist.getRange(lr, 1, 1, wsHist.getMaxColumns()).setBorder(true,true,true,true,true,true).setVerticalAlignment("middle");
            wsHist.getRange(lr, COL.DATA + 1).setNumberFormat("dd/MM/yyyy"); 
        }
        wsBase.deleteRow(targetRow); 
    }
    
    SpreadsheetApp.flush(); 
    return "Colaborador arquivado como EXCLUÍDO!"; 
  } catch(e) { 
    return "Erro no servidor: " + e.message; 
  } 
}

function receberDadosFretado(dadosExcel) {
  try {
    if (!dadosExcel || dadosExcel.length < 2) return "Erro: Arquivo vazio ou inválido.";
    
    const ss = SpreadsheetApp.getActive();
    const ws = ss.getSheetByName(ABA_BASE);
    let baseData = ws.getDataRange().getValues();
    
    let mapaMat = {};
    let mapaNome = {};
    for(let i = 1; i < baseData.length; i++) {
        let mat = String(baseData[i][COL.MATRICULA]).trim();
        let nome = String(baseData[i][COL.NOME] || "").trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "");
        if (mat) mapaMat[mat] = i;
        if (nome) mapaNome[nome] = i;
    }
    
    let headerRow = [];
    let startRow = -1;
    for(let i = 0; i < Math.min(20, dadosExcel.length); i++) {
        let txt = dadosExcel[i].join("").toUpperCase();
        if (txt.includes("NOME") && (txt.includes("BAIRRO") || txt.includes("HORA"))) {
            headerRow = dadosExcel[i].map(h => String(h).toUpperCase().trim());
            startRow = i + 1; 
            break;
        }
    }
    
    if (startRow === -1) return "Erro: Não encontrei a linha de cabeçalho. O Excel precisa ter as colunas 'NOME' e 'BAIRRO'.";
    
    let idx = {
        mat: headerRow.findIndex(h => h.includes("MATRICULA") || h.includes("RE ")),
        nome: headerRow.findIndex(h => h.includes("NOME DO COLABORADOR") || h === "NOME" || h === "COLABORADOR"),
        carro: headerRow.findIndex(h => h.includes("Nº CARRO") || h === "CARRO"),
        van: headerRow.findIndex(h => h.includes("COD") || h === "VAN"),
        mot: headerRow.findIndex(h => h.includes("MOTORI")),
        fone: headerRow.findIndex(h => h.includes("FONE")),
        rua: headerRow.findIndex(h => h.includes("ENDEREÇO") || h.includes("RUA")),
        num: headerRow.findIndex(h => h === "Nº" || h === "NUM" || h === "NUMERO"),
        bairro: headerRow.findIndex(h => h.includes("BAIRRO")),
        tipo: headerRow.findIndex(h => h.includes("TIPO")),
        hora: headerRow.findIndex(h => h.includes("HORA") || h.includes("HORÁRIO")),
        focal: headerRow.findIndex(h => h.includes("FOCAL")),
        gerente: headerRow.findIndex(h => h.includes("GERENTE") || h.includes("COORD")),
        area: headerRow.findIndex(h => h.includes("AREA") || h.includes("ÁREA")),
        contrato: headerRow.findIndex(h => h.includes("CONTRATO"))
    };

    if (idx.nome === -1) return "Erro: A coluna com o Nome do Colaborador não foi encontrada no Excel.";
    
    let idxDias = idx.hora > -1 ? idx.hora + 1 : -1;
    let atualizados = 0, novos = 0;
    
    function formatarHoraExcel(val) {
        if (!val && val !== 0) return "";
        if (!isNaN(val) && val < 1) { 
            let totalMin = Math.round(val * 24 * 60);
            let h = Math.floor(totalMin / 60); let m = totalMin % 60;
            return (h < 10 ? "0"+h : h) + ":" + (m < 10 ? "0"+m : m);
        }
        return String(val).trim();
    }

    for (let r = startRow; r < dadosExcel.length; r++) {
        let exc = dadosExcel[r];
        let nomeEx = String(exc[idx.nome] || "").trim().toUpperCase();
        let nomeBusca = nomeEx.normalize('NFD').replace(/[\u0300-\u036f]/g, ""); 
        let matEx = idx.mat > -1 ? String(exc[idx.mat] || "").trim() : "";
        
        if (!nomeEx || nomeEx === "NOME" || nomeEx.includes("NOME DO COLABORADOR")) continue;
        
        let sysRow;
        let isNew = false;
        
        if (matEx && mapaMat.hasOwnProperty(matEx)) {
            sysRow = baseData[mapaMat[matEx]];
            atualizados++;
        } else if (mapaNome.hasOwnProperty(nomeBusca)) {
            sysRow = baseData[mapaNome[nomeBusca]];
            atualizados++;
        } else {
            sysRow = new Array(50).fill(""); 
            isNew = true;
            novos++;
        }
        
        if (idx.carro > -1) sysRow[COL.CARRO] = String(exc[idx.carro] || "").toUpperCase().trim() || "A DEFINIR";
        if (idx.van > -1) sysRow[COL.VAN] = String(exc[idx.van] || "").toUpperCase().trim();
        if (idx.mot > -1) sysRow[COL.MOT] = String(exc[idx.mot] || "").toUpperCase().trim();
        sysRow[COL.NOME] = nomeEx;
        if (matEx) sysRow[COL.MATRICULA] = matEx;
        if (idx.fone > -1) sysRow[COL.FONE] = exc[idx.fone] || sysRow[COL.FONE];
        if (idx.rua > -1) sysRow[COL.RUA] = String(exc[idx.rua] || "").toUpperCase();
        if (idx.num > -1) sysRow[COL.NUM] = exc[idx.num] || "";
        if (idx.bairro > -1) sysRow[COL.BAIRRO] = String(exc[idx.bairro] || "").toUpperCase();
        if (idx.tipo > -1) sysRow[COL.TIPO] = String(exc[idx.tipo] || "").toUpperCase();
        if (idx.hora > -1) sysRow[COL.HORARIO] = formatarHoraExcel(exc[idx.hora]);
        
        if (idxDias > -1) {
            for(let d = 0; d < 31; d++) {
                let val = String(exc[idxDias + d] || "").trim().toUpperCase();
                if (['F','FOLGA','DSR','FERIAS','FE','FR','X','COMP','FC','X '].includes(val)) {
                    sysRow[12 + d] = "F";
                } else if (val !== "") {
                    sysRow[12 + d] = val; 
                } else {
                    sysRow[12 + d] = "";
                }
            }
        }
        
        if (idx.focal > -1) sysRow[COL.FOCAL] = exc[idx.focal] || sysRow[COL.FOCAL];
        if (idx.gerente > -1) sysRow[COL.GERENTE] = exc[idx.gerente] || sysRow[COL.GERENTE];
        if (idx.area > -1) sysRow[COL.AREA] = exc[idx.area] || sysRow[COL.AREA];
        if (idx.contrato > -1) sysRow[COL.CONTRATO] = exc[idx.contrato] || sysRow[COL.CONTRATO];
        
        if (isNew) {
            sysRow[COL.DATA] = new Date(); 
            baseData.push(sysRow);
        }
    }
    
    let dadosFinais = baseData.map(r => {
        let nova = [];
        for(let k=0; k<50; k++) nova.push(r[k] !== undefined ? r[k] : "");
        return nova;
    });
    
    ws.getRange(1, 1, dadosFinais.length, 50).setValues(dadosFinais);
    setLastUpdate();
    
    return `✅ Escala processada com sucesso!\n\n${atualizados} cadastros cruzados e atualizados.\n${novos} novatos inseridos na base.`;

  } catch(e) { return "ERRO GRAVE: " + e.message; }
}

function rotinaVerificarFerias() {
  try {
    const ss = SpreadsheetApp.getActive();
    const wsBase = ss.getSheetByName(ABA_BASE);
    const wsHist = ss.getSheetByName(ABA_HIST);
    
    if (!wsBase || !wsHist) return;

    const data = wsBase.getDataRange().getValues();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];
      let iniStr = row[COL.FERIAS_INI]; 
      
      if (iniStr) {
        let dataIni = new Date(iniStr + "T00:00:00");
        if (hoje >= dataIni) {
          row[COL.CARRO] = "EM FÉRIAS"; 
          
          let novoDado = [];
          for (let k = 0; k < row.length; k++) {
            novoDado.push(row[k]);
          }
          
          while (novoDado.length < 50) {
            novoDado.push("");
          }
          novoDado[50] = "FÉRIAS"; 
          
          if (wsHist.getMaxColumns() < 51) {
            wsHist.insertColumnsAfter(wsHist.getMaxColumns(), 51 - wsHist.getMaxColumns());
          }
          wsHist.appendRow(novoDado);
          
          let lr = wsHist.getLastRow();
          wsHist.getRange(lr, 1, 1, wsHist.getMaxColumns()).setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
          wsHist.getRange(lr, COL.DATA + 1).setNumberFormat("dd/MM/yyyy");
          
          wsBase.deleteRow(i + 1); 
        }
      }
    }
  } catch(e) {
    console.error("Erro no acionador de férias: " + e.message);
  }
}

function obterUrlApp() {
  return ScriptApp.getService().getUrl();
}

function limparGradeFolgas() {
  try {
    const ss = SpreadsheetApp.getActive();
    const ws = ss.getSheetByName(ABA_BASE);
    let data = ws.getDataRange().getValues();
    let alterados = false;

    for(let i=1; i<data.length; i++) {
      for(let c=12; c<=42; c++) {
        // Apaga tudo que foi digitado, MAS preserva as Férias ("FE") intactas!
        if(data[i][c] !== "FE" && data[i][c] !== "") { 
          data[i][c] = "";
          alterados = true;
        }
      }
    }

    if(alterados) {
      ws.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
    }
    return "Grade de folgas limpa com sucesso! (As Férias ativas foram preservadas).";
  } catch(e) { return "Erro no servidor: " + e.message; }
}

function removerLote(lista) {
  try {
    let sucessos = 0;
    // O loop passa por todos os selecionados e usa a sua função remover que já é super segura
    for (let i = 0; i < lista.length; i++) {
      remover(lista[i].linha, lista[i].mat);
      sucessos++;
    }
    return `${sucessos} colaboradores foram excluídos e arquivados com sucesso!`;
  } catch(e) { return "Erro ao excluir em lote: " + e.message; }
}

function processarSolicitacao(linhaOrigem, acao) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const wsSol = ss.getSheetByName(ABA_SOLIC);
        const wsBase = ss.getSheetByName(ABA_BASE);
        
        if (!wsSol) return "Erro: Aba SOLICITAÇÕES não encontrada.";
        
        let numLinha = Number(linhaOrigem);
        if (isNaN(numLinha) || numLinha < 2) return "Erro: Linha da solicitação inválida.";
        
        let lastCol = wsSol.getLastColumn();
        let dados = wsSol.getRange(numLinha, 1, 1, lastCol).getValues()[0];
        
        // Se APROVAR, joga o colaborador na base principal
        if (acao === 'APROVAR') {
            let mat = String(dados[COL.MATRICULA]).trim();
            let baseData = wsBase.getDataRange().getValues();
            let targetRow = wsBase.getLastRow() + 1; // Padrão: Adicionar no final
            
            // O robô é inteligente: se o colaborador já existir na base, ele sobrepõe a linha antiga.
            for (let i = 1; i < baseData.length; i++) {
                if (String(baseData[i][COL.MATRICULA]).trim() === mat) {
                    targetRow = i + 1;
                    break;
                }
            }
            
            if (wsBase.getMaxColumns() < dados.length) wsBase.insertColumnsAfter(wsBase.getMaxColumns(), dados.length - wsBase.getMaxColumns());
            wsBase.getRange(targetRow, 1, 1, dados.length).setValues([dados]);
        }
        
        // Seja aprovado ou rejeitado, removemos da aba de pendentes
        wsSol.deleteRow(numLinha);
        SpreadsheetApp.flush();
        
        return acao === 'APROVAR' ? "Solicitação Aprovada e salva na Base Principal!" : "Solicitação Rejeitada e descartada.";
    } catch(e) {
        return "Erro no servidor: " + e.message;
    }
}
