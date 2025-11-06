/** =========================
 * CONFIG
 * ========================= */
const SPREADSHEET_ID = '14TmgtzVvfYTTjf4oXklo74sqFMKFDRY1DZUL25gjOb0';

function getSS_() {
  try {
    if (SPREADSHEET_ID && !/^COLE_AQUI/.test(SPREADSHEET_ID)) {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  } catch (e) {
    Logger.log('Erro ao abrir planilha: ' + e.message);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** =========================
 * Helpers
 * ========================= */
function _normHeader_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim().toLowerCase().replace(/\s+/g, '_');
}

function _headerIndexMap_(ws) {
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0] || [];
  const map = {};
  headers.forEach((h, i) => map[_normHeader_(h)] = i + 1);
  return map;
}

function _findColByNames_(ws, candidates) {
  const map = _headerIndexMap_(ws);
  for (const name of candidates) {
    const idx = map[_normHeader_(name)];
    if (idx) return idx;
  }
  return null;
}

function _normName_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim().toLowerCase();
}

function _getSheetByNames_(ss, candidates) {
  const wanted = new Set(candidates.map(_normName_));
  for (const sh of ss.getSheets()) {
    if (wanted.has(_normName_(sh.getName()))) return sh;
  }
  return null;
}

function _normalizeId_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') {
    return String(value).replace(/\.0+$/, '');
  }
  return String(value).trim().replace(/\.0+$/, '');
}

function _isValidItemId_(value) {
  if (!value) return false;
  
  const str = String(value).trim();
  if (!str) return false;
  if (str.startsWith('=')) return false;
  
  // 🔧 ACEITA IDs com 1 ou mais caracteres
  if (str.length < 1) return false;
  
  const headersComuns = [
    'id_item', 'id', 'item', 'codigo', 'código', 'cor',
    'quantidade', 'qtd', 'qtde', 'estoque',
    'setor', 'setor_entrega', 'departamento', 'area', 'área',
    'nome', 'descricao', 'descrição', 'tipo'
  ];
  
  const strLower = str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (headersComuns.includes(strLower)) return false;
  if (['n/a', 'na', 'null', 'undefined', '-'].includes(strLower)) return false;
  
  return true;
}

function _serialize_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === 'number') {
    return value;
  }
  return String(value);
}

/** =========================
 * Roteamento
 * ========================= */
function doGet(e) {
  try {
    Logger.log('=== doGet chamado ===');
    Logger.log('Parâmetros: ' + JSON.stringify(e ? e.parameter : {}));
    
    const page = (e && e.parameter && e.parameter.page) || "";
    Logger.log('Página solicitada: "' + page + '"');
    
    let output;
    
    if (page === "app") {
      Logger.log('Carregando Index.html (aplicação)');
      output = HtmlService.createHtmlOutputFromFile("Index");
    } else {
      Logger.log('Carregando Login.html (tela de login)');
      output = HtmlService.createHtmlOutputFromFile("Login");
    }
    
    output.setTitle("Controle de Materiais");
    output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
    Logger.log('Página carregada com sucesso');
    return output;
    
  } catch (erro) {
    Logger.log('ERRO em doGet: ' + erro.message);
    Logger.log('Stack: ' + erro.stack);
    
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:Arial;padding:20px;">' +
      '<h1>Erro ao Carregar</h1>' +
      '<p>Erro: ' + erro.message + '</p>' +
      '<p>Verifique se os arquivos Login.html e Index.html existem no projeto.</p>' +
      '</body></html>'
    ).setTitle('Erro');
  }
}

function getWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    Logger.log('URL do Web App: ' + url);
    return url;
  } catch (e) {
    Logger.log('Erro ao obter URL: ' + e.message);
    throw new Error('Não foi possível obter a URL do Web App');
  }
}

/** =========================
 * Autenticação
 * ========================= */
function verificarLogin(usuario, senha) {
  try {
    Logger.log('=== verificarLogin chamado ===');
    Logger.log('Usuário: ' + usuario);
    
    const ss = getSS_();
    const ws = _getSheetByNames_(ss, ["Credenciais"]);
    
    if (!ws) {
      Logger.log('ERRO: Aba Credenciais não encontrada');
      throw new Error('Aba "Credenciais" não encontrada');
    }

    const last = ws.getLastRow();
    Logger.log('Linhas na aba Credenciais: ' + last);
    
    if (last < 2) {
      Logger.log('Aba Credenciais vazia');
      return null;
    }

    const data = ws.getRange(2, 1, last - 1, 4).getValues();
    const u = String(usuario || '').trim().toLowerCase();
    const s = String(senha || '').trim();

    for (let i = 0; i < data.length; i++) {
      const usuarioPlanilha = String(data[i][0]).trim().toLowerCase();
      const senhaPlanilha = String(data[i][1] || '').trim().replace(/\.0+$/, '');
      
      if (usuarioPlanilha === u && senhaPlanilha === s) {
        const nome = String(data[i][2] || '').trim();
        const funcao = String(data[i][3] || '').trim();
        
        const resultado = {
          usuario: data[i][0],
          nomeCompleto: nome || String(data[i][0]),
          funcao: funcao || ''
        };
        
        Logger.log('Login bem-sucedido: ' + JSON.stringify(resultado));
        return resultado;
      }
    }
    
    Logger.log('Login falhou: credenciais inválidas');
    return null;
    
  } catch (e) {
    Logger.log('ERRO em verificarLogin: ' + e.message);
    throw e;
  }
}

/** =========================
 * Dados auxiliares
 * ========================= */
function getSetoresCadastro() {
  try {
    const ss = getSS_();
    const ws = _getSheetByNames_(ss, ["Estoque"]);
    if (!ws || ws.getLastRow() < 2) return [];

    const colSetor = _findColByNames_(ws, ["Setor", "Setor_Entrega", "Departamento", "Area", "Área"]);
    if (!colSetor) return [];

    const maxRows = Math.min(ws.getLastRow() - 1, 1000);
    const colVals = ws.getRange(2, colSetor, maxRows, 1).getValues();

    const uniq = new Set();
    colVals.forEach(r => {
      const v = String(r[0] || '').trim();
      if (v && !v.startsWith('=')) uniq.add(v);
    });

    return Array.from(uniq).sort((a, b) => a.localeCompare(b));
  } catch (e) {
    Logger.log('Erro em getSetoresCadastro: ' + e.message);
    return [];
  }
}

/** =========================
 * 🔧 FUNÇÃO CORRIGIDA: Busca Dinâmica de Itens
 * Busca itens com priorização: começam com o termo primeiro
 * ========================= */
function buscarItens(termo, limite) {
  try {
    Logger.log('=== buscarItens: "' + termo + '" ===');
    
    // Validar parâmetros
    if (!termo || termo.length < 1) {
      Logger.log('Termo vazio, retornando vazio');
      return [];
    }
    
    const maxResultados = limite || 30;
    const termoNormalizado = String(termo)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim()
      .toLowerCase();
    
    const ss = getSS_();
    const ws = _getSheetByNames_(ss, ["Estoque"]);
    
    if (!ws || ws.getLastRow() < 2) {
      Logger.log('Aba Estoque vazia');
      return [];
    }

    const colId = _findColByNames_(ws, ["ID_Item", "ID", "Item", "Codigo", "Código", "Cor"]);
    if (!colId) {
      Logger.log('Coluna ID não encontrada');
      return [];
    }

    // 🔧 LER TODAS AS LINHAS (sem limite de 500)
    const totalRows = ws.getLastRow() - 1;
    Logger.log('Total de linhas no Estoque: ' + totalRows);
    
    const idsVals = ws.getRange(2, colId, totalRows, 1).getValues();
    
    const colSetor = _findColByNames_(ws, ["Setor", "Setor_Entrega", "Departamento", "Area", "Área"]);
    const setorVals = colSetor ? ws.getRange(2, colSetor, totalRows, 1).getValues() : null;

    // Carregar nomes da aba PRODUTOS
    const wsProdutos = _getSheetByNames_(ss, ["PRODUTOS", "Produtos", "produtos"]);
    const mapNome = new Map();
    
    if (wsProdutos && wsProdutos.getLastRow() >= 2) {
      const prodRows = wsProdutos.getRange(2, 1, wsProdutos.getLastRow() - 1, 2).getValues();
      prodRows.forEach(r => {
        const id = _normalizeId_(r[0]);
        if (id && _isValidItemId_(id)) {
          mapNome.set(id, String(r[1] || id));
        }
      });
    }

    // 🔧 BUSCA PRIORIZADA: itens que começam com o termo aparecem primeiro
    const resultadosExatos = [];  // Começam com o termo (startsWith)
    const resultadosOutros = [];  // Contêm o termo (includes)
    const vistos = new Set();
    
    for (let i = 0; i < totalRows; i++) {
      const rawValue = idsVals[i][0];
      const id = _normalizeId_(rawValue);
      
      if (!_isValidItemId_(id) || vistos.has(id)) {
        continue;
      }
      
      const nome = mapNome.get(id) || id;
      const setor = setorVals ? String(setorVals[i][0] || '').trim() : '';
      
      // Normalizar para busca
      const idNorm = id.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
      const nomeNorm = nome.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
      
      const item = {
        id: id,
        nome: nome,
        setor: setor,
        label: nome && nome !== id ? `${id} - ${nome}` : id
      };
      
      // Prioridade 1: ID ou nome começam com o termo
      if (idNorm.startsWith(termoNormalizado) || nomeNorm.startsWith(termoNormalizado)) {
        resultadosExatos.push(item);
        vistos.add(id);
      }
      // Prioridade 2: ID ou nome contêm o termo em qualquer posição
      else if (idNorm.includes(termoNormalizado) || nomeNorm.includes(termoNormalizado)) {
        resultadosOutros.push(item);
        vistos.add(id);
      }
    }
    
    // Ordenar cada categoria alfabeticamente por ID
    resultadosExatos.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true, sensitivity: 'base' }));
    resultadosOutros.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true, sensitivity: 'base' }));
    
    // Combinar: exatos primeiro, depois outros, respeitando o limite
    const resultados = [...resultadosExatos, ...resultadosOutros].slice(0, maxResultados);
    
    Logger.log('Encontrados ' + resultados.length + ' resultados para "' + termo + '"');
    Logger.log('  - Exatos (começam com termo): ' + resultadosExatos.length);
    Logger.log('  - Outros (contêm termo): ' + resultadosOutros.length);
    
    return resultados;
    
  } catch (e) {
    Logger.log('ERRO em buscarItens: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return [];
  }
}
/** =========================
 * Pedidos
 * ========================= */
function getPedidosComEstoque() {
  try {
    Logger.log('=== INÍCIO getPedidosComEstoque ===');
    const ss = getSS_();

    const wsPedidos = _getSheetByNames_(ss, ["MOVIMENTACOES", "Movimentacoes", "Movimentações"]);
    if (!wsPedidos) {
      Logger.log('ERRO: Aba MOVIMENTACOES não encontrada');
      return [];
    }

    if (wsPedidos.getLastRow() < 2) {
      Logger.log('Aba MOVIMENTACOES vazia');
      return [];
    }

    const rowsCount = wsPedidos.getLastRow() - 1;
    const numCols = Math.min(wsPedidos.getLastColumn(), 22);
    
    Logger.log('Lendo ' + rowsCount + ' pedidos com ' + numCols + ' colunas');
    
    let pedidosData = wsPedidos.getRange(2, 1, rowsCount, numCols).getValues();
    
    Logger.log('Pedidos lidos: ' + pedidosData.length);

    const idsUnicos = new Set();
    pedidosData.forEach(row => {
      const id = _normalizeId_(row[2]);
      if (id) idsUnicos.add(id);
    });

    Logger.log('IDs únicos: ' + idsUnicos.size);

    const wsProdutos = _getSheetByNames_(ss, ["PRODUTOS", "Produtos", "produtos"]);
    const produtosMap = new Map();
    if (wsProdutos && wsProdutos.getLastRow() >= 2) {
      const prodRows = wsProdutos.getRange(2, 1, wsProdutos.getLastRow() - 1, 2).getValues();
      prodRows.forEach(row => {
        const id = _normalizeId_(row[0]);
        if (id && idsUnicos.has(id)) {
          produtosMap.set(id, String(row[1] || id));
        }
      });
    }

    Logger.log('Produtos carregados: ' + produtosMap.size);

    const wsEstoque = _getSheetByNames_(ss, ["Estoque"]);
    const estoqueMap = new Map();
    
    if (wsEstoque && wsEstoque.getLastRow() >= 2) {
      const colId = _findColByNames_(wsEstoque, ["ID_Item", "ID", "Item", "Codigo", "Código"]);
      const colQtd = _findColByNames_(wsEstoque, ["Qtd", "Quantidade", "Estoque", "Qtd_Atual", "Qtde"]);
      
      if (colId && colQtd) {
        const maxRows = Math.min(wsEstoque.getLastRow() - 1, 1000);
        const ids = wsEstoque.getRange(2, colId, maxRows, 1).getValues();
        const qts = wsEstoque.getRange(2, colQtd, maxRows, 1).getValues();
        
        Logger.log('Carregando estoque: ' + maxRows + ' linhas');
        
        for (let i = 0; i < ids.length; i++) {
          const id = _normalizeId_(ids[i][0]);
          const qtdRaw = qts[i][0];
          
          if (id && idsUnicos.has(id)) {
            let qtdFinal = 0;
            
            if (typeof qtdRaw === 'number' && !isNaN(qtdRaw)) {
              qtdFinal = qtdRaw;
            } else if (typeof qtdRaw === 'string') {
              const parsed = parseFloat(qtdRaw);
              if (!isNaN(parsed)) {
                qtdFinal = parsed;
              }
            }
            
            estoqueMap.set(id, qtdFinal);
          }
        }
      }
    }

    Logger.log('Estoque carregado: ' + estoqueMap.size + ' itens');

    const resultado = [];
    
    for (let i = 0; i < pedidosData.length; i++) {
      const pedido = pedidosData[i];
      const arr = [];
      
      for (let j = 0; j < 22; j++) {
        if (j < pedido.length) {
          arr[j] = _serialize_(pedido[j]);
        } else {
          arr[j] = '';
        }
      }
      
      const idItem = _normalizeId_(pedido[2]);
      arr[2] = produtosMap.get(idItem) || idItem || 'Item Desconhecido';
      arr[19] = estoqueMap.get(idItem) || 0;
      arr[20] = idItem;
      arr[21] = pedido[21] || 0;
      
      resultado.push(arr);
    }

    resultado.sort((a, b) => {
      try {
        const dateA = a[5] ? new Date(a[5]).getTime() : 0;
        const dateB = b[5] ? new Date(b[5]).getTime() : 0;
        return dateB - dateA;
      } catch (e) {
        return 0;
      }
    });

    Logger.log('=== FIM - Retornando ' + resultado.length + ' pedidos ===');
    
    return resultado;
    
  } catch (e) {
    Logger.log('ERRO em getPedidosComEstoque: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return [];
  }
}

function criarNovoPedido(dados) {
  try {
    const ss = getSS_();
    const ws = _getSheetByNames_(ss, ["MOVIMENTACOES", "Movimentacoes", "Movimentações"]);
    if (!ws) throw new Error('Aba "MOVIMENTACOES" não encontrada');

    const itemId = _normalizeId_(dados.item);
    const kg = dados.kg || 0;
    
    Logger.log('Criando pedido para item: ' + itemId + ' | KG: ' + kg);

    const novaLinha = [
      'REQ-' + new Date().getTime(),
      'Novo',
      itemId,
      dados.quantidade,
      dados.setor,
      new Date(),
      dados.solicitante,
      '', '', '', '', '', '', '', '', '', '', '', '', '', '', kg
    ];
    
    ws.appendRow(novaLinha);
    
    Logger.log('Pedido criado: ' + novaLinha[0]);
    
    return "Novo pedido criado com sucesso!";
  } catch (e) {
    Logger.log('ERRO em criarNovoPedido: ' + e.message);
    throw e;
  }
}

function atualizarPedido(id, acao, valores) {
  try {
    const ss = getSS_();
    const ws = _getSheetByNames_(ss, ["MOVIMENTACOES", "Movimentacoes", "Movimentações"]);
    if (!ws) throw new Error('Aba "MOVIMENTACOES" não encontrada');

    const last = ws.getLastRow();
    if (last < 2) return "Erro: sem pedidos";

    const ids = ws.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] == id) {
        const L = i + 2;
        
        Logger.log('Atualizando pedido ' + id + ' - ação: ' + acao);
        
        switch (acao) {
          case 'enderecar':
            ws.getRange(L, 2).setValue('Aguardando Coleta');
            ws.getRange(L, 8).setValue(valores.usuario);
            ws.getRange(L, 9).setValue(valores.local);
            ws.getRange(L, 10).setValue(new Date());
            return "Pedido endereçado com sucesso.";
          case 'coletar':
            ws.getRange(L, 2).setValue('Em Trânsito');
            ws.getRange(L, 11).setValue(valores.usuario);
            ws.getRange(L, 12).setValue(valores.local);
            ws.getRange(L, 13).setValue(new Date());
            return "Coleta confirmada com sucesso.";
          case 'receber':
            ws.getRange(L, 2).setValue('Finalizado');
            ws.getRange(L, 14).setValue(new Date());
            ws.getRange(L, 19).setValue(valores.usuario);
            return "Recebimento confirmado.";
          case 'iniciarDevolucao':
            ws.getRange(L, 2).setValue('Aguardando Devolução');
            ws.getRange(L, 15).setValue(valores.usuario);
            return "Processo de devolução iniciado.";
          case 'coletarDevolucao':
            ws.getRange(L, 2).setValue('Devolução Finalizada');
            ws.getRange(L, 16).setValue(valores.usuario);
            ws.getRange(L, 17).setValue(valores.local);
            ws.getRange(L, 18).setValue(new Date());
            return "Devolução finalizada com sucesso.";
          default:
            return "Erro: Ação desconhecida.";
        }
      }
    }
    return "Erro: Pedido não encontrado.";
  } catch (e) {
    Logger.log('ERRO em atualizarPedido: ' + e.message);
    throw e;
  }
}

/** =========================
 * Debug
 * ========================= */
function debugResumo() {
  const ss = getSS_();
  const ws = _getSheetByNames_(ss, ["MOVIMENTACOES", "Movimentacoes", "Movimentações"]);
  const info = {
    ssId: ss.getId(),
    ssName: ss.getName(),
    sheets: ss.getSheets().map(s => s.getName()),
    mov: {
      exists: !!ws,
      name: ws ? ws.getName() : null,
      lastRow: ws ? ws.getLastRow() : 0,
      lastCol: ws ? ws.getLastColumn() : 0,
      headers: ws ? ws.getRange(1, 1, 1, Math.min(ws.getLastColumn(), 18)).getValues()[0] : []
    },
    sample: ws && ws.getLastRow() > 1
      ? ws.getRange(2, 1, Math.min(3, ws.getLastRow() - 1), Math.min(ws.getLastColumn(), 18)).getValues()
      : []
  };
  Logger.log('Debug: ' + JSON.stringify(info));
  return info;
}

function debugItensEstoque() {
  try {
    const ss = getSS_();
    const ws = _getSheetByNames_(ss, ["Estoque"]);
    
    if (!ws) {
      Logger.log('Aba Estoque não encontrada');
      return { erro: 'Aba Estoque não encontrada' };
    }
    
    const colId = _findColByNames_(ws, ["ID_Item", "ID", "Item", "Codigo", "Código", "Cor"]);
    
    if (!colId) {
      Logger.log('Coluna ID não encontrada');
      return { erro: 'Coluna ID não encontrada' };
    }
    
    const maxRows = Math.min(ws.getLastRow() - 1, 20);
    const dados = ws.getRange(1, 1, maxRows + 1, ws.getLastColumn()).getValues();
    
    const resultado = {
      nomeAba: ws.getName(),
      totalLinhas: ws.getLastRow(),
      totalColunas: ws.getLastColumn(),
      colunaID: colId,
      cabecalhos: dados[0],
      primeiras10Linhas: []
    };
    
    for (let i = 1; i <= Math.min(10, maxRows); i++) {
      const linha = dados[i];
      const idRaw = linha[colId - 1];
      const idNormalizado = _normalizeId_(idRaw);
      const valido = _isValidItemId_(idRaw);
      
      resultado.primeiras10Linhas.push({
        linha: i + 1,
        idRaw: idRaw,
        idNormalizado: idNormalizado,
        valido: valido,
        linhaCompleta: linha
      });
    }
    
    Logger.log('Debug Itens Estoque: ' + JSON.stringify(resultado, null, 2));
    return resultado;
    
  } catch (e) {
    Logger.log('ERRO em debugItensEstoque: ' + e.message);
    return { erro: e.message, stack: e.stack };
  }
}
