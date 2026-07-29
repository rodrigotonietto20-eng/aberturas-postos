/**
 * Backend do app "Abertura de Postos — Rede Sander".
 *
 * COMO USAR:
 * 1. Crie uma nova planilha Google Sheets (separada das planilhas do CRM e do Horas Extras).
 * 2. Crie duas abas com esses nomes exatos:
 *    - "Postos"  -> cabeçalho linha 1: ID | Nome | Bandeira | DataAbertura | CriadoEm
 *    - "Itens"   -> cabeçalho linha 1: ID | PostoID | Item | Fase | Setor | Responsavel | Status | Observacao | Historico | Anexos | Demandas
 *    (se a aba "Itens" já existir sem alguma coluna nova, adicione manualmente antes de
 *    reimplantar — sem isso o app ainda funciona (grava na posição certa mesmo sem cabeçalho
 *    nomeado), só fica sem nome visível pra quem olha a planilha direto. "Anexos" guarda
 *    [{nome,url,tipo,quando}], "Demandas" guarda [{id,texto,feito}] por item)
 *    - "Setores" -> NÃO precisa criar manualmente, o script cria sozinho (Key | Icone |
 *      ResponsavelPadrao) na primeira chamada depois que você colar este código novo.
 * 3. Menu Extensões -> Apps Script. Apague o conteúdo padrão e cole este arquivo inteiro.
 * 4. Menu Implantar -> Nova implantação -> tipo "App da Web".
 *    - Executar como: Eu (sua conta)
 *    - Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL gerada (termina em /exec) e cole no app: botão "configurar" no rodapé da
 *    lista de postos, ou direto na constante SCRIPT_URL_DEFAULT no topo do index.html.
 * 6. Sempre que reimplantar como NOVA implantação, a URL muda — atualize de novo.
 *    Reimplantar como "nova versão" na MESMA implantação não muda a URL.
 *
 * Mesmo padrão já validado no CRM de Visitas e no Horas Extras: mapa explícito
 * cabeçalho<->campo (CAMPOS_POSTO / CAMPOS_ITEM) em vez de comparar texto do cabeçalho
 * direto com o nome da propriedade — resistente a reformatação manual da planilha.
 */

var CAMPOS_POSTO = [
  {header:'ID',           key:'id'},
  {header:'Nome',         key:'nome'},
  {header:'Bandeira',     key:'bandeira'},
  {header:'DataAbertura', key:'dataAbertura', date:true},
  {header:'CriadoEm',     key:'criadoEm', num:true}
];

var CAMPOS_ITEM = [
  {header:'ID',           key:'id'},
  {header:'PostoID',      key:'postoId'},
  {header:'Item',         key:'item'},
  {header:'Fase',         key:'fase'},
  {header:'Setor',        key:'setor'},
  {header:'Responsavel',  key:'responsavel'},
  {header:'Status',       key:'status'},
  {header:'Observacao',   key:'observacao'},
  {header:'Historico',    key:'historico', json:true},
  {header:'Anexos',       key:'anexos', json:true},
  {header:'Demandas',     key:'demandas', json:true}
];

var CAMPOS_SETOR = [
  {header:'Key',              key:'key'},
  {header:'Icone',            key:'icone'},
  {header:'ResponsavelPadrao', key:'responsavel'}
];

var CAMPOS_DEMANDA_PADRAO = [
  {header:'ItemNome', key:'itemNome'},
  {header:'Textos',   key:'textos', json:true}
];

function getSheetPostos(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Postos'); }
function getSheetItens(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Itens'); }

// Aba nova (2026-07-27) — auto-criada na primeira chamada, não precisa passo manual.
function getSheetSetores(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Setores');
  if(!sh){
    sh = ss.insertSheet('Setores');
    sh.getRange(1,1,1,CAMPOS_SETOR.length).setValues([CAMPOS_SETOR.map(function(c){ return c.header; })]);
  }
  return sh;
}

// Aba nova (2026-07-27) — auto-criada, guarda as demandas (checkbox) que os usuários foram
// digitando em cada pergunta, pra virarem padrão nos próximos postos com a mesma pergunta.
function getSheetDemandasPadrao(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DemandasPadrao');
  if(!sh){
    sh = ss.insertSheet('DemandasPadrao');
    sh.getRange(1,1,1,CAMPOS_DEMANDA_PADRAO.length).setValues([CAMPOS_DEMANDA_PADRAO.map(function(c){ return c.header; })]);
  }
  return sh;
}

function lerLinhas(sheet, campos){
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var out = [];
  for(var i=1;i<rows.length;i++){
    var row = rows[i];
    if(!row[0]) continue; // sem ID, linha vazia
    var obj = {};
    campos.forEach(function(c, idx){
      var col = headers.indexOf(c.header);
      if(col<0) col = idx; // fallback posicional se o cabeçalho estiver diferente/em branco
      var val = row[col];
      if(c.json){ try{ val = val ? JSON.parse(val) : []; }catch(err){ val = []; } }
      else if(c.num){ val = val===''||val==null ? 0 : Number(val); }
      else if(c.date && Object.prototype.toString.call(val)==='[object Date]'){
        // O Sheets autoconverte a célula "DataAbertura" pra um valor de Data de verdade
        // quando recebe uma string tipo "2026-07-29" — sem isso, getValues() devolve um
        // objeto Date que serializa como "2026-07-29T03:00:00.000Z" na tela do app.
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[c.key] = val;
    });
    out.push(obj);
  }
  return out;
}

function doGet(e){
  var postos = lerLinhas(getSheetPostos(), CAMPOS_POSTO);
  var itens = lerLinhas(getSheetItens(), CAMPOS_ITEM);
  var setores = lerLinhas(getSheetSetores(), CAMPOS_SETOR);
  var demandasPadrao = lerLinhas(getSheetDemandasPadrao(), CAMPOS_DEMANDA_PADRAO);

  var itensPorPosto = {};
  itens.forEach(function(it){
    var pid = it.postoId;
    if(!itensPorPosto[pid]) itensPorPosto[pid] = [];
    delete it.postoId;
    itensPorPosto[pid].push(it);
  });
  postos.forEach(function(p){ p.itens = itensPorPosto[p.id] || []; });

  return ContentService.createTextOutput(JSON.stringify({ok:true, postos:postos, setores:setores, demandasPadrao:demandasPadrao}))
    .setMimeType(ContentService.MimeType.JSON);
}

// Atualiza em cima da linha existente (getRange().setValues(), rápido) em vez do padrão
// antigo de apagar todas as linhas com ID conhecido e reinserir tudo no fim (deleteRow() em
// loop, que fica muito lento com centenas de linhas — cada edição de um posto de ~220 itens
// reescrevia a aba inteira). Só faz append pros IDs que realmente ainda não existem na
// planilha. Efeito colateral bom: a ordem das linhas já existentes deixa de embaralhar a
// cada save, então dá pra editar a planilha manualmente sem os itens pularem de lugar.
function upsert(sheet, campos, registros, idKey){
  if(!registros.length) return;
  var headers = sheet.getDataRange().getValues()[0];
  var idCol = headers.indexOf(campos[0].header);
  if(idCol<0) idCol = 0;

  var data = sheet.getDataRange().getValues();
  var linhaPorId = {};
  for(var i=1;i<data.length;i++){
    if(data[i][idCol]) linhaPorId[data[i][idCol]] = i+1; // linha real na planilha (1-based)
  }

  var novos = [];
  registros.forEach(function(r){
    var linha = campos.map(function(c){
      var val = r[c.key];
      if(c.json) return JSON.stringify(val||[]);
      return val==null ? '' : val;
    });
    var linhaExistente = linhaPorId[r[idKey]];
    if(linhaExistente){
      sheet.getRange(linhaExistente, 1, 1, campos.length).setValues([linha]);
    } else {
      novos.push(linha);
    }
  });
  if(novos.length){
    sheet.getRange(sheet.getLastRow()+1, 1, novos.length, campos.length).setValues(novos);
  }
}

function excluirPosto(postoId){
  var sheetPostos = getSheetPostos();
  var dataP = sheetPostos.getDataRange().getValues();
  for(var i=dataP.length-1;i>=1;i--){ if(dataP[i][0]===postoId) sheetPostos.deleteRow(i+1); }

  var sheetItens = getSheetItens();
  var dataI = sheetItens.getDataRange().getValues();
  for(var i=dataI.length-1;i>=1;i--){ if(dataI[i][1]===postoId) sheetItens.deleteRow(i+1); }
}

// Item excluído pelo app (botão 🗑️ Excluir dentro do posto). Precisa de endpoint próprio
// porque o upsert() normal só insere/atualiza os IDs presentes no payload — nunca apaga —
// então sem isso o item voltaria assim que outro aparelho sincronizasse.
function excluirItemPorId(itemId){
  var sheet = getSheetItens();
  var data = sheet.getDataRange().getValues();
  for(var i=data.length-1;i>=1;i--){ if(data[i][0]===itemId) sheet.deleteRow(i+1); }
}

function getOrCreatePastaRaizAnexos(){
  var nome = 'Abertura de Postos - Anexos';
  var it = DriveApp.getFoldersByName(nome);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(nome);
}

function getOrCreatePastaPosto(postoId, postoNome){
  var raiz = getOrCreatePastaRaizAnexos();
  var nomePasta = (postoNome||'Posto') + ' - ' + postoId;
  var it = raiz.getFoldersByName(nomePasta);
  if(it.hasNext()) return it.next();
  return raiz.createFolder(nomePasta);
}

function uploadAnexo(payload){
  var pasta = getOrCreatePastaPosto(payload.postoId, payload.postoNome);
  var base64 = String(payload.base64||'');
  var virgula = base64.indexOf(',');
  if(virgula>=0) base64 = base64.substring(virgula+1);
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, payload.tipo || 'application/octet-stream', payload.nome || 'anexo');
  var file = pasta.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function doPost(e){
  var body = JSON.parse(e.postData.contents);

  if(body.uploadAnexo){
    var url = uploadAnexo(body.uploadAnexo);
    return ContentService.createTextOutput(JSON.stringify({ok:true, url:url})).setMimeType(ContentService.MimeType.JSON);
  }

  if(body.excluirPostoId){
    excluirPosto(body.excluirPostoId);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }

  if(body.excluirItemId){
    excluirItemPorId(body.excluirItemId);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }

  if(body.setores){
    upsert(getSheetSetores(), CAMPOS_SETOR, body.setores, 'key');
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }

  if(body.demandasPadrao){
    upsert(getSheetDemandasPadrao(), CAMPOS_DEMANDA_PADRAO, body.demandasPadrao, 'itemNome');
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }

  var postos = body.postos || [];
  if(!postos.length) return ContentService.createTextOutput(JSON.stringify({ok:true}));

  var postosSemItens = postos.map(function(p){
    return {id:p.id, nome:p.nome, bandeira:p.bandeira, dataAbertura:p.dataAbertura, criadoEm:p.criadoEm};
  });
  upsert(getSheetPostos(), CAMPOS_POSTO, postosSemItens, 'id');

  var todosItens = [];
  postos.forEach(function(p){
    (p.itens||[]).forEach(function(it){
      todosItens.push({
        id:it.id, postoId:p.id, item:it.item, fase:it.fase, setor:it.setor,
        responsavel:it.responsavel, status:it.status, observacao:it.observacao, historico:it.historico,
        anexos:it.anexos, demandas:it.demandas
      });
    });
  });
  upsert(getSheetItens(), CAMPOS_ITEM, todosItens, 'id');

  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}
