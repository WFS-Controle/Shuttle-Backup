/**
 * ============================================================
 * WFS MANAGER - VERSÃO FINAL (BASE-ZERO + AUTH + AUDITORIA)
 * ============================================================
 */

// --- CONFIGURAÇÕES E IDs ---
const ID_PLANILHA_DEMITIDOS = "19wHIhlwOh2gwWOgn8sC_gok5iYAf1W8DTFAfbNFXV8I"; 
const ID_BASE_EFETIVOS = "19wHIhlwOh2gwWOgn8sC_gok5iYAf1W8DTFAfbNFXV8I"; 

const ABA_BASE = 'BASE_GERAL';
const ABA_HIST = 'HISTORICO';
const ABA_CONFIG = 'CONFIG';

const MAPA_GERENTES = {
  "CANAL": "ROBERTA NADAI", "CARGA": "CLEBER DE OLIVEIRA", "SEGURANÇA": "JOSY MENDES",
  "RAMPA NARROW": "RENATO LEAL", "RAMPA WIDE": "KARYN ALVES", "LIMPEZA": "RENATA MENEZES",
  "ESTEIRA": "MATHEUS PIRES", "PASSAGEIROS": "GUILHERME IGLESIAS", "COI": "MACIEL PONTECIANO"
};

const COL = { CARRO: 0, VAN: 1, MOT: 2, MATRICULA: 3, NOME: 4, FONE: 5, CEP: 6, RUA: 7, NUM: 8, BAIRRO: 9, TIPO: 10, HORARIO: 11, FOCAL: 43, GERENTE: 44, AREA: 45, CONTRATO: 46, DATA: 47 };

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('WFS SHUTTLE').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =======================================================
// AUTENTICAÇÃO E PRIMEIRO ACESSO
// =======================================================
function tentarLogin(user, pass) {
  let u = String(user).trim(); let p = String(pass).trim();
  if (u === "admin" && p === "Orbital2026") return { sucesso: true, perfil: 'ADMIN', setor: 'VISÃO GLOBAL', nome: 'Administrador Geral' };

  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("USUARIOS");
  if (!ws) return { erro: "Nenhum usuário cadastrado. Faça o Primeiro Acesso." };

  let data = ws.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
      if(String(data[i][1]).trim() === u && String(data[i][2]).trim() === p) {
          let agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
          ws.getRange(i+1, 6).setValue(agora); 
          return { sucesso: true, perfil: data[i][4], setor: data[i][3], nome: data[i][0] };
      }
  }
  return { erro: "Usuário ou senha incorretos." };
}

function registrarPrimeiroAcesso(nome, user, pass, setor) {
  if (!nome || !user || !pass || !setor) return { erro: "Preencha todos os campos." };
  let u = String(user).trim();
  
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("USUARIOS");
  if (!ws) {
      ws = ss.insertSheet("USUARIOS");
      ws.appendRow(["NOME", "USUARIO", "SENHA", "SETOR", "PERFIL", "ULTIMO_ACESSO"]);
      ws.getRange("A1:F1").setBackground("#27272a").setFontColor("white").setFontWeight("bold");
  }
  
  let data = ws.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
      if(String(data[i][1]).trim() === u) return { erro: "Este nome de usuário já existe." };
  }
  
  let agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  ws.appendRow([nome.toUpperCase(), u, pass, setor, "FOCAL", agora]);
  return { sucesso: true, msg: "Cadastro realizado com sucesso! Pode entrar." };
}

function getListaSetores() { return Object.keys(MAPA_GERENTES); }

function setLastUpdate() {
  let hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM HH:mm");
  PropertiesService.getScriptProperties().setProperty('LAST_UPLOAD', hoje); return hoje;
}
function getLastUpdate() { return PropertiesService.getScriptProperties().getProperty('LAST_UPLOAD') || "Pendente"; }

function setPrazoSistema(ini, fim) {
  const ws = SpreadsheetApp.getActive().getSheetByName(ABA_CONFIG); 
  ws.getRange("A2").setValue(ini); ws.getRange("B2").setValue(fim); 
  return "Período Atualizado!";
}

function getPrazoSistema() {
  const ws = SpreadsheetApp.getActive().getSheetByName(ABA_CONFIG); const d = ws.getRange("A2:B2").getValues()[0];
  if(!d[0] || !d[1]) return { inicio: "", fim: "", isLocked: false };
  const h = new Date(); h.setHours(0,0,0,0); const i = new Date(d[0]); i.setHours(0,0,0,0); const f = new Date(d[1]); f.setHours(23,59,59,999);
  return { inicio: Utilities.formatDate(i, Session.getScriptTimeZone(), "yyyy-MM-dd"), fim: Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"), isLocked: (h < i || h > f) };
}

function formatarHora(val) { if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm"); return String(val); }

function carregarDadosSetor(setorUser, perfilUser) {
  try {
    const ss = SpreadsheetApp.getActive(); 
    let ws = ss.getSheetByName(ABA_BASE);
    const data = ws.getDataRange().getValues(); 
    const systemInfo = getPrazoSistema(); 
    const lastUpdate = String(getLastUpdate() || "Pendente");
    const hoje = new Date();
    
    let demitidosIDs = [];
    try { 
        const shDem = SpreadsheetApp.openById(ID_PLANILHA_DEMITIDOS).getSheetByName("DEMITIDOS"); 
        if(shDem) demitidosIDs = shDem.getRange("A2:A").getValues().flat().map(String); 
    } catch(e) {}

    function processarLinhas(matrizDados, prefixoLinha) {
        let lista = []; let statsLocais = { total: 0, novatos: 0, alertas: 0, semrota: 0 }; let contratosLocais = new Set();
        
        for(let i=1; i<matrizDados.length; i++) {
            let row = matrizDados[i]; if(row.length < 46) continue;
            
            if ((perfilUser === 'ADMIN') || (String(row[COL.AREA]||"").trim().toUpperCase() === String(setorUser||"").trim().toUpperCase())) {
                let contrato = String(row[COL.CONTRATO]||"").trim(); if(contrato) contratosLocais.add(contrato);
                let mat = String(row[COL.MATRICULA] || ""); if(!mat) continue;
                
                let isNovato = false;
                if(row[COL.DATA]) {
                    let objData = new Date(row[COL.DATA]);
                    if(!isNaN(objData.getTime())) { isNovato = ((hoje - objData)/(8.64e7) < 30); }
                }
                
                let isDem = demitidosIDs.includes(mat);
                let ic = String(row[COL.CARRO]||"").trim().toUpperCase(); let iv = String(row[COL.VAN]||"").trim().toUpperCase();
                let icLimpo = ic.normalize('NFD').replace(/[\u0300-\u036f]/g, ""); let ivLimpo = iv.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
                
                let isAdefinir = (icLimpo === "" || icLimpo === "A DEFINIR" || icLimpo === "UNDEFINED");
                let isSemRota = (icLimpo.includes("SEM VAGA") || icLimpo.includes("SEM ROTA") || icLimpo.includes("NAO TEM") || ivLimpo.includes("SEM ROTA") || ivLimpo.includes("SEM VAGA") || ivLimpo.includes("NAO TEM"));
                
                let statusFinal = "OK";
                if (isDem) statusFinal = "DEMITIDO"; else if (isSemRota) statusFinal = "SEM ROTA"; else if (isAdefinir) statusFinal = "A DEFINIR";

                let folgas = [], diasGrid = []; 
                for(let f=12; f<=42; f++) { let v = String(row[f]||"").trim(); if(v === 'F') { folgas.push(f-11); diasGrid.push("F"); } else { diasGrid.push(""); } }
                
                statsLocais.total++; if(isNovato) statsLocais.novatos++; if(isDem) statsLocais.alertas++; if((isSemRota || isAdefinir) && !isDem) statsLocais.semrota++;
                
                lista.push({ 
                    linha: prefixoLinha === "BASE" ? (i+1) : String(prefixoLinha + i), 
                    matricula: mat, 
                    nome: String(row[COL.NOME]||"").toUpperCase(), 
                    status: statusFinal, 
                    situacao: isNovato ? "NOVATO" : "FIXO", 
                    carro: String(row[COL.CARRO]||""), 
                    van: String(row[COL.VAN]||""), 
                    mot: String(row[COL.MOT]||""), 
                    fone: String(row[COL.FONE]||""), 
                    cep: String(row[COL.CEP]||""), 
                    rua: String(row[COL.RUA]||""), 
                    num: String(row[COL.NUM]||""), 
                    bairro: String(row[COL.BAIRRO]||""), 
                    tipo: String(row[COL.TIPO]||""), 
                    horario: String(typeof formatarHora === 'function' ? formatarHora(row[COL.HORARIO]) : row[COL.HORARIO] || ""), 
                    contrato: contrato, 
                    focal: String(row[COL.FOCAL]||""), 
                    area: String(row[COL.AREA]||""), 
                    folgas: folgas, 
                    diasGrid: diasGrid 
                });
            }
        }
        lista.sort((a, b) => a.nome.localeCompare(b.nome)); return { lista: lista, stats: statsLocais, contratos: contratosLocais };
    }

    let objBase = processarLinhas(data, "BASE");
    let wsBackup = ss.getSheetByName("BACKUP");
    let objBackup = wsBackup ? processarLinhas(wsBackup.getDataRange().getValues(), "BKP-") : { lista: [], stats: { total: 0, novatos: 0, alertas: 0, semrota: 0 }, contratos: new Set() };
    
    let statusEnvioSetores = [];
    if (perfilUser === 'ADMIN') {
        let expectedSectors = new Set(); 
        let currentSectors = {}; 
        let sectorFocalCount = {}; 
        let usersMap = {}; 
        
        // Limpa acentos e espaços invisíveis
        const normalizarTexto = (txt) => String(txt || "").normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
        
        let dictIdentidade = {}; // Dicionário Mágico: [Qualquer Nome/Login] -> [Nome Oficial Exato]
        
        // 1. LÊ A ABA USUARIOS PRIMEIRO PARA CRIAR O DICIONÁRIO
        try {
            let wsUsr = ss.getSheetByName("USUARIOS");
            if(wsUsr) {
                let uData = wsUsr.getDataRange().getValues();
                for(let i=1; i<uData.length; i++) { 
                    let nomeOriginal = String(uData[i][0]||"").trim(); // Coluna A (NOME)
                    if(!nomeOriginal) continue;
                    
                    let loginOriginal = String(uData[i][1]||"").trim(); // Coluna B (USUARIO)
                    let uSetor = String(uData[i][3]).trim().toUpperCase(); 
                    
                    let nomeLimpo = normalizarTexto(nomeOriginal); 
                    let loginLimpo = normalizarTexto(loginOriginal); 
                    
                    // Associa o Nome normalizado e o Login ao Nome Oficial de exibição
                    if(nomeLimpo) dictIdentidade[nomeLimpo] = nomeOriginal;
                    if(loginLimpo) dictIdentidade[loginLimpo] = nomeOriginal;
                    
                    // Tratamento extra para logins com pontos (ex: Gabriel.Santos vira GABRIEL SANTOS)
                    if(loginLimpo.includes(".")) dictIdentidade[loginLimpo.replace(/\./g, " ")] = nomeOriginal;
                    if(nomeLimpo.includes(".")) dictIdentidade[nomeLimpo.replace(/\./g, " ")] = nomeOriginal;
                    
                    let rawAcesso = uData[i][5];
                    let strAcesso = "Nunca";
                    if (rawAcesso instanceof Date) { 
                        strAcesso = Utilities.formatDate(rawAcesso, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"); 
                    } else if (rawAcesso) { 
                        strAcesso = String(rawAcesso); 
                    }
                    
                    if(!usersMap[uSetor]) usersMap[uSetor] = [];
                    usersMap[uSetor].push({ nomeDisplay: nomeOriginal, acesso: strAcesso });
                }
            }
        } catch(e) {}

        // 2. LÊ BACKUP
        if(wsBackup && wsBackup.getLastRow() > 1) {
            let bkpData = wsBackup.getDataRange().getValues();
            for(let i=1; i<bkpData.length; i++) { 
                let s = String(bkpData[i][COL.AREA]||"").trim().toUpperCase(); 
                if(s && s !== "ADMINISTRADOR") expectedSectors.add(s); 
            }
        }
        
        // 3. LÊ A BASE E CONTA OS LANÇAMENTOS USANDO A TRADUÇÃO DO DICIONÁRIO
        for(let i=1; i<data.length; i++) {
            let s = String(data[i][COL.AREA]||"").trim().toUpperCase();
            
            if(s && s !== "ADMINISTRADOR") { 
                if(!currentSectors[s]) currentSectors[s] = 0; 
                currentSectors[s]++; 
                
                let focalRaw = normalizarTexto(data[i][COL.FOCAL]);
                if(focalRaw) {
                    let focalAlt = focalRaw.replace(/\./g, " "); // Tira os pontos caso existam
                    
                    // Procura o nome oficial no dicionário usando a matrícula ou nome limpo
                    // Se não achar nada de correspondência, usa o texto original como fallback
                    let nomeOficial = dictIdentidade[focalRaw] || dictIdentidade[focalAlt] || focalRaw;
                    
                    if(!sectorFocalCount[s]) sectorFocalCount[s] = {};
                    if(!sectorFocalCount[s][nomeOficial]) sectorFocalCount[s][nomeOficial] = 0;
                    sectorFocalCount[s][nomeOficial]++;
                }
            }
        }

        // 4. GERA OS CRACHÁS VISUAIS DO ADMIN
        let todosSetores = new Set([...Array.from(expectedSectors), ...Object.keys(currentSectors)]);
        todosSetores.forEach(s => {
            let status = currentSectors[s] > 0 ? "LANÇADO" : "PENDENTE";
            
            let gestoresStr = "Sem Cadastro no Sistema<br><span style='color:var(--text-muted); font-size:10px;'>Último: -</span>";
            if (usersMap[s] && usersMap[s].length > 0) {
                gestoresStr = usersMap[s].map(u => {
                    // A busca agora é 100% precisa pois ambos (painel e contagem) usam o nomeDisplay
                    let count = (sectorFocalCount[s] && sectorFocalCount[s][u.nomeDisplay]) ? sectorFocalCount[s][u.nomeDisplay] : 0;
                    
                    // Fallback de contingência: se ainda estiver zerado, tenta cruzar pedaços do nome
                    if (count === 0 && sectorFocalCount[s]) {
                        let nomeLimpoDisplay = normalizarTexto(u.nomeDisplay);
                        Object.keys(sectorFocalCount[s]).forEach(k => {
                            if (nomeLimpoDisplay.includes(normalizarTexto(k))) {
                                count += sectorFocalCount[s][k];
                                sectorFocalCount[s][k] = 0; // Transfere a pontuação e zera
                            }
                        });
                    }
                    
                    let badge = count > 0 
                        ? `<span style="background:#d1fae5; color:#059669; padding:2px 6px; border-radius:4px; font-size:10px; font-weight:800; margin-left:6px; vertical-align:middle;">${count} lançamentos</span>` 
                        : `<span style="background:#fee2e2; color:#be123c; padding:2px 6px; border-radius:4px; font-size:10px; font-weight:800; margin-left:6px; vertical-align:middle;">0 lançamentos</span>`;
                        
                    return `<span style="color:var(--text-strong); font-weight:700;">${u.nomeDisplay}</span> ${badge}<br><span style="color:var(--text-muted); font-size:10px;">Último Acesso: ${u.acesso}</span>`;
                }).join('<div style="height:12px;"></div>');
            }

            statusEnvioSetores.push({ setor: s, total: currentSectors[s] || 0, status: status, gestoresStr: gestoresStr });
        });
        statusEnvioSetores.sort((a, b) => a.status.localeCompare(b.status)); 
    }
    
    return { 
        lista: objBase.lista, 
        stats: objBase.stats, 
        contratos: Array.from(objBase.contratos).sort(), 
        listaBackup: objBackup.lista, 
        statsBackup: objBackup.stats, 
        system: systemInfo, 
        lastUpdate: lastUpdate, 
        monitorEnvio: statusEnvioSetores 
    };
  } catch (e) { 
      return { erro: e.message || String(e) }; 
  }
}

function salvarColagemInteligente(dadosRecebidos) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000); // 🚦 Tolera até 25 segundos se o lote for pesado
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const wsBase = ss.getSheetByName(ABA_BASE); const wsBackup = ss.getSheetByName("BACKUP");
    const dadosDb = wsBase.getDataRange().getValues(); let dadosBkp = wsBackup ? wsBackup.getDataRange().getValues() : [];
    let alterados = 0;
    
    let mapaLinhasBase = {}; for(let i=1; i<dadosDb.length; i++) { let mat = String(dadosDb[i][COL.MATRICULA]).trim(); if(mat) mapaLinhasBase[mat] = i + 1; }
    let mapaLinhasBkp = {}; if(dadosBkp.length > 0) { for(let i=1; i<dadosBkp.length; i++) { let mat = String(dadosBkp[i][COL.MATRICULA]).trim(); if(mat) mapaLinhasBkp[mat] = i; } }

    dadosRecebidos.forEach(item => {
        let matStr = String(item.matricula).trim(); let linha = mapaLinhasBase[matStr];
        
        // Lógica original de puxar do backup se não existir
        if (!linha && mapaLinhasBkp[matStr]) {
            let linhaBkpIdx = mapaLinhasBkp[matStr]; let novaLinhaArray = dadosBkp[linhaBkpIdx].slice(); 
            for(let f=12; f<=42; f++) novaLinhaArray[f] = "";
            novaLinhaArray[COL.CARRO] = "A DEFINIR"; novaLinhaArray[COL.VAN] = "-"; novaLinhaArray[COL.MOT] = "-";
            wsBase.appendRow(novaLinhaArray); linha = wsBase.getLastRow(); mapaLinhasBase[matStr] = linha; 
        }
        if (linha) {
            let diasParaSalvar = item.dias; while (diasParaSalvar.length < 31) diasParaSalvar.push(""); diasParaSalvar = diasParaSalvar.slice(0, 31);
            wsBase.getRange(linha, 13, 1, diasParaSalvar.length).setValues([diasParaSalvar]); alterados++;
        }
    });
    SpreadsheetApp.flush();
    return { success: true, msg: `${alterados} escalas recuperadas/atualizadas com sucesso!` };
  } catch (e) { 
    return { success: false, msg: "Erro no servidor: " + e.message }; 
  } finally {
    lock.releaseLock();
  }
}

function salvarDados(d, nomeFocal, setorFocal, perfilUser) { 
  var lock = LockService.getScriptLock();
  try { 
    lock.waitLock(15000); // 🚦 Fila de espera
    const ss = SpreadsheetApp.getActive(); const prazos = getPrazoSistema(); 
    if (prazos.isLocked && perfilUser !== 'ADMIN') return "Sistema Fechado! Não é possível alterar a base após o prazo."; 
    
    let ws = ss.getSheetByName(ABA_BASE); 
    
    // VERIFICAÇÃO ABSOLUTA DE MATRÍCULA
    let dataMat = ws.getRange(1, COL.MATRICULA + 1, ws.getLastRow(), 1).getValues();
    let linhaEncontrada = -1;
    for (let i = 1; i < dataMat.length; i++) {
        if (String(dataMat[i][0]).trim() === String(d.matricula).trim()) {
            linhaEncontrada = i + 1;
            break;
        }
    }
    
    // Trava de segurança no servidor para cadastros duplos
    if (!d.isEdit && linhaEncontrada !== -1) return "Duplicado!";
    
    // Se for edição, confia na linha encontrada agora, e não na linha velha do Frontend
    let linha = d.isEdit ? (linhaEncontrada !== -1 ? linhaEncontrada : Number(d.linha)) : ws.getLastRow() + 1; 
    let dados = new Array(50).fill(""); 
    
    if(d.isEdit && linha <= ws.getLastRow()) { 
        let v = ws.getRange(linha, 1, 1, 48).getValues()[0]; 
        for(let k=0; k<v.length; k++) dados[k] = v[k]; 
    } 
    
    dados[COL.MATRICULA] = d.matricula; dados[COL.NOME] = String(d.nome).toUpperCase(); dados[COL.FONE] = d.telefone; dados[COL.CEP] = d.cep; dados[COL.RUA] = String(d.endereco).toUpperCase(); dados[COL.NUM] = d.numero; dados[COL.BAIRRO] = String(d.bairro).toUpperCase(); dados[COL.TIPO] = d.tipo; dados[COL.HORARIO] = d.horario; dados[COL.CONTRATO] = String(d.contrato).toUpperCase(); 
    
    if (!d.isEdit) { dados[COL.CARRO] = "A DEFINIR"; dados[COL.VAN] = "-"; dados[COL.MOT] = "-"; } 
    for(let f=12; f<=42; f++) dados[f] = ""; 
    if(d.folgas) d.folgas.forEach(dia => { let idx = 11 + parseInt(dia); if(idx <= 42) dados[idx] = "F"; }); 
    
    dados[COL.AREA] = dados[COL.AREA] || (perfilUser === 'ADMIN' && d.setorManual ? d.setorManual : setorFocal); 
    dados[COL.FOCAL] = dados[COL.FOCAL] || nomeFocal; 
    if (MAPA_GERENTES[dados[COL.AREA]]) dados[COL.GERENTE] = MAPA_GERENTES[dados[COL.AREA]]; 
    if (!dados[COL.DATA]) dados[COL.DATA] = new Date(); 
    
    let colDiff = dados.length - ws.getMaxColumns();
    if(colDiff > 0) ws.insertColumnsAfter(ws.getMaxColumns(), colDiff); 
    
    ws.getRange(linha, 1, 1, dados.length).setValues([dados]); 
    SpreadsheetApp.flush();
    return "Salvo com Sucesso!"; 
  } catch(e) { 
    return "ERRO: " + e.toString(); 
  } finally {
    lock.releaseLock(); // 🚦 Libera
  }
}

function executarViradaDeCiclo() {
  try {
    const ss = SpreadsheetApp.getActive(); const wsBase = ss.getSheetByName(ABA_BASE); let wsBackup = ss.getSheetByName("BACKUP");
    if (!wsBackup) { wsBackup = ss.insertSheet("BACKUP"); } else { wsBackup.clear(); }
    const data = wsBase.getDataRange().getValues();
    if (data.length > 0) { wsBackup.getRange(1, 1, data.length, data[0].length).setValues(data); }
    if (wsBase.getLastRow() > 1) { wsBase.getRange(2, 1, wsBase.getLastRow() - 1, wsBase.getMaxColumns()).clearContent(); }
    SpreadsheetApp.flush(); return "Virada de ciclo executada! Escala salva no BACKUP e Base zerada para o novo mês.";
  } catch(e) { return "Erro no servidor: " + e.message; }
}

function gerarExcelFormatado() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const originalSheet = ss.getSheetByName(ABA_BASE); const tempSheet = originalSheet.copyTo(ss); tempSheet.setName("TEMP_" + new Date().getTime()); let lastRow = tempSheet.getLastRow(); if (lastRow >= 2) { tempSheet.getRange(2, 1, lastRow - 1, 3).clearContent(); SpreadsheetApp.flush(); } const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx&gid=" + tempSheet.getSheetId(); const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } }; const blob = UrlFetchApp.fetch(url, params).getBlob().setName("Escala_WFS.xlsx"); ss.deleteSheet(tempSheet); return { base64: Utilities.base64Encode(blob.getBytes()), filename: "Escala_WFS.xlsx" }; 
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function buscarCEP(cep) { if(!cep) return null; cep = cep.replace(/\D/g, ''); try { let r = UrlFetchApp.fetch(`https://viacep.com.br/ws/${cep}/json/`); let j = JSON.parse(r.getContentText()); if(!j.erro) return { rua: j.logradouro, bairro: j.bairro }; } catch(e) {} try { let r2 = UrlFetchApp.fetch(`https://brasilapi.com.br/api/cep/v1/${cep}`); let j2 = JSON.parse(r2.getContentText()); if(j2.street) return { rua: j2.street, bairro: j2.neighborhood }; } catch(e){} return null; }
function remover(ln, mat) { 
  var lock = LockService.getScriptLock();
  try { 
    lock.waitLock(15000); // 🚦 Fila de espera
    const ss = SpreadsheetApp.getActive(); 
    let wsBase = ss.getSheetByName(ABA_BASE); 
    let wsHist = ss.getSheetByName(ABA_HIST); 
    
    let targetRow = Number(ln); 
    let matProcurada = (mat !== undefined && mat !== null) ? String(mat).trim() : "";
    
    // SISTEMA HÍBRIDO: Confere se a pessoa da linha é a pessoa certa. Se não for, caça pela matrícula!
    if (matProcurada !== "") {
        let currentMat = String(wsBase.getRange(targetRow, COL.MATRICULA + 1).getValue()).trim();
        if (currentMat !== matProcurada) {
            let dataMat = wsBase.getRange(1, COL.MATRICULA + 1, wsBase.getLastRow(), 1).getValues(); 
            let found = false;
            for (let i = 1; i < dataMat.length; i++) { 
                if (String(dataMat[i][0]).trim() === matProcurada) { 
                    targetRow = i + 1; 
                    found = true;
                    break; 
                } 
            }
            if (!found) return "Erro: Colaborador não encontrado na base para exclusão.";
        }
    }

    if (isNaN(targetRow) || targetRow < 2) return "Erro: Linha inválida para exclusão.";
    
    let lastCol = wsBase.getLastColumn(); 
    let dadosOriginais = wsBase.getRange(targetRow, 1, 1, lastCol).getValues()[0]; 
    
    // Arquivamento histórico seguro
    if(wsHist) { 
        let novoDado = []; 
        for(let i = 0; i < dadosOriginais.length; i++) novoDado.push(dadosOriginais[i]); 
        while(novoDado.length < 50) novoDado.push(""); 
        novoDado[50] = "EXCLUÍDO"; 
        
        let colDiff = 51 - wsHist.getMaxColumns();
        if (colDiff > 0) wsHist.insertColumnsAfter(wsHist.getMaxColumns(), colDiff); 
        
        wsHist.appendRow(novoDado); 
        let lr = wsHist.getLastRow(); 
        wsHist.getRange(lr, 1, 1, wsHist.getMaxColumns()).setBorder(true,true,true,true,true,true).setVerticalAlignment("middle"); 
        wsHist.getRange(lr, COL.DATA + 1).setNumberFormat("dd/MM/yyyy"); 
    } 
    
    wsBase.deleteRow(targetRow); 
    SpreadsheetApp.flush(); // Força a gravação síncrona
    return "Colaborador arquivado como EXCLUÍDO!"; 
  } catch(e) { 
    return "Erro no servidor: " + e.message; 
  } finally {
    lock.releaseLock(); 
  }
}
function removerLote(lista) { try { let sucessos = 0; for (let i = 0; i < lista.length; i++) { remover(lista[i].linha, lista[i].mat); sucessos++; } return `${sucessos} colaboradores foram arquivados!`; } catch(e) { return "Erro ao excluir: " + e.message; } }
function limparGradeFolgas() { try { const ss = SpreadsheetApp.getActive(); const ws = ss.getSheetByName(ABA_BASE); let data = ws.getDataRange().getValues(); let alterados = false; for(let i=1; i<data.length; i++) { for(let c=12; c<=42; c++) { if(data[i][c] !== "") { data[i][c] = ""; alterados = true; } } } if(alterados) { ws.getRange(1, 1, data.length, data[0].length).setValues(data); SpreadsheetApp.flush(); } return "Grade limpa!"; } catch(e) { return "Erro no servidor: " + e.message; } }
function buscarNomeAtivo(mat) { const strMat = String(mat).trim().toUpperCase(); const ss = SpreadsheetApp.getActive(); let wsBase = ss.getSheetByName(ABA_BASE); if(wsBase.getLastRow() > 1) { let f = wsBase.getRange(2, COL.MATRICULA+1, wsBase.getLastRow()-1, 2).getValues().find(r => String(r[0]).trim().toUpperCase() === strMat); if(f) return { status: 'EXISTE_BASE', nome: f[1] }; } try { const shAtivos = SpreadsheetApp.openById(ID_BASE_EFETIVOS).getSheetByName("ATIVOS"); if(shAtivos && shAtivos.getLastRow() > 1) { let f = shAtivos.getRange(2, 1, shAtivos.getLastRow()-1, 2).getValues().find(r => String(r[0]).trim().toUpperCase() === strMat); if(f) return { status: 'ENCONTRADO', nome: f[1] }; } } catch(e) {} return null; }
function obterUrlApp() { return ScriptApp.getService().getUrl(); }
