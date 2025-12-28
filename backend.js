/**
 * ============================================================================
 * DROPDOWN COLOR UI – BACKEND (CONTROLADOR) - VERSÃO CORRIGIDA
 * Garante: 1. Salvamento em DocumentProperties 2. Limpeza de Cache no Save
 * 3. Sanitização correta (não deleta se só subcategorias tiverem cor)
 * ============================================================================
 */

const PROP_KEY_CORES = 'DROPDOWN_CORES_OVERRIDE';

// --- HELPERS ---

// ✅ CORREÇÃO #6: Função idêntica ao DropdownManager (arrow function)
const _gerarIdLista = (nomeAba, nomeColuna) => {
  const norm = (str) => (str || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
  return `${norm(nomeAba)}_${norm(nomeColuna)}`;
};

function _lerOverrideCores_(propKey) {
  const doc = PropertiesService.getDocumentProperties();
  const script = PropertiesService.getScriptProperties();
  
  const rawDoc = doc.getProperty(propKey);
  if (rawDoc) { try { return { ok: true, data: JSON.parse(rawDoc) }; } catch (e) {} }

  const rawScript = script.getProperty(propKey);
  if (rawScript) { try { return { ok: true, data: JSON.parse(rawScript) }; } catch (e) {} }

  return { ok: true, data: {} };
}

// --- FUNÇÕES PÚBLICAS (RPC) ---

function getDadosConfiguracao() {
  try {
    if (typeof DropdownManager !== 'undefined') DropdownManager._limparCache();
    if (typeof Utils !== 'undefined') Utils.invalidateCache('all');

    const leitura = _lerOverrideCores_(PROP_KEY_CORES);
    const savedConfig = leitura.data || {};
    
    let todasListas = [];
    if (typeof DropdownManager !== 'undefined' && DropdownManager._obterListasConfig) {
      todasListas = DropdownManager._obterListasConfig();
    } else if (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) {
      Object.keys(CONFIG.SHEETS).forEach(aba => {
        (CONFIG.SHEETS[aba].LISTAS_SUSPENSAS || []).forEach(cfg => todasListas.push({ nomeAba: aba, ...cfg }));
      });
    }

    const resultado = {};

    todasListas.forEach(lista => {
      const listaId = _gerarIdLista(lista.nomeAba, lista.colunaCategoria);
      let dadosHierarquicos = {};
      
      if (typeof DropdownManager !== 'undefined' && DropdownManager.obterHierarquia) {
         dadosHierarquicos = DropdownManager.obterHierarquia(lista);
      } else {
         dadosHierarquicos = _lerHierarquiaLocal(lista);
      }

      let configSalva = savedConfig[listaId];
      if (!configSalva) {
         const k = Object.keys(savedConfig).find(key => key.toUpperCase() === listaId);
         if (k) configSalva = savedConfig[k];
      }
      configSalva = configSalva || { cores: {}, coresSubcategoria: {} };

      resultado[listaId] = {
        nomeExibicao: `${lista.nomeAba} (${lista.colunaCategoria})`,
        hierarquia: dadosHierarquicos,
        configuracaoAtual: {
           cores: { ...(lista.cores || {}), ...(configSalva.cores || {}) },
           coresSubcategoria: { ...(lista.coresSubcategoria || {}), ...(configSalva.coresSubcategoria || {}) },
           herdarCorSubcategoria: lista.herdarCorSubcategoria
        },
        configuracaoSalva: configSalva
      };
    });

    return { sucesso: true, dados: resultado };
  } catch (e) {
    console.error("Erro getDados:", e);
    return { sucesso: false, erro: e.message };
  }
}

function salvarConfiguracaoCores(novaConfigMapa) {
  try {
    const leitura = _lerOverrideCores_(PROP_KEY_CORES);
    let savedConfig = leitura.data || {};

    // ✅ CORREÇÃO: Helper para normalizar chaves do objeto
    const normalizarChaves = (obj) => {
      if (!obj || typeof obj !== 'object') return {};
      const normalizado = {};
      Object.keys(obj).forEach(key => {
        const keyNorm = _gerarIdLista('', key).split('_')[1]; // Reutiliza normalização
        normalizado[keyNorm] = obj[key];
      });
      return normalizado;
    };

    Object.entries(novaConfigMapa || {}).forEach(([listaId, dados]) => {
      const limpo = {
        cores: normalizarChaves(dados.cores || {}),
        coresSubcategoria: normalizarChaves(dados.coresSubcategoria || {})
      };
      
      const totalmenteVazio = (
        Object.keys(limpo.cores).length === 0 && 
        Object.keys(limpo.coresSubcategoria).length === 0
      );
      
      if (totalmenteVazio) {
        delete savedConfig[listaId];
      } else {
        savedConfig[listaId] = limpo;
      }
    });

    PropertiesService.getDocumentProperties().setProperty(PROP_KEY_CORES, JSON.stringify(savedConfig));
    
    if (typeof DropdownManager !== 'undefined') DropdownManager._limparCache();
    if (typeof Utils !== 'undefined') Utils.invalidateCache('all');

    return { sucesso: true };
  } catch (e) {
    console.error("Erro salvar:", e);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * EXPORTA todas as configs atuais (DocumentProperties) como JSON (texto)
 * Patch mínimo: não muda o modelo, só expõe leitura para UI.
 */
function exportarConfiguracaoCores() {
  try {
    const leitura = _lerOverrideCores_(PROP_KEY_CORES);
    const obj = leitura.data || {};
    return { sucesso: true, json: JSON.stringify(obj, null, 2) };
  } catch (e) {
    return { sucesso: false, erro: String(e && e.message ? e.message : e) };
  }
}

/**
 * IMPORTA JSON (texto) e grava tudo no DocumentProperties (overwrite total)
 * Patch mínimo: valida o básico e limpa cache após salvar.
 */
function importarConfiguracaoCores(jsonText) {
  try {
    if (!jsonText || typeof jsonText !== 'string') {
      return { sucesso: false, erro: 'JSON vazio ou inválido.' };
    }

    let obj;
    try {
      obj = JSON.parse(jsonText);
    } catch (e) {
      return { sucesso: false, erro: 'JSON malformado: ' + (e.message || e) };
    }

    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
      return { sucesso: false, erro: 'JSON deve ser um objeto no formato { LISTA_ID: { cores, coresSubcategoria } }.' };
    }

    // Validação mínima de estrutura (sem over-engineering)
    Object.keys(obj).forEach((listaId) => {
      const bloco = obj[listaId];
      if (!bloco || typeof bloco !== 'object') obj[listaId] = { cores: {}, coresSubcategoria: {} };
      if (!obj[listaId].cores || typeof obj[listaId].cores !== 'object') obj[listaId].cores = {};
      if (!obj[listaId].coresSubcategoria || typeof obj[listaId].coresSubcategoria !== 'object') obj[listaId].coresSubcategoria = {};
    });

    PropertiesService.getDocumentProperties().setProperty(PROP_KEY_CORES, JSON.stringify(obj));

    if (typeof DropdownManager !== 'undefined') DropdownManager._limparCache();
    if (typeof Utils !== 'undefined') Utils.invalidateCache('all');

    return { sucesso: true, totalListas: Object.keys(obj).length };
  } catch (e) {
    return { sucesso: false, erro: String(e && e.message ? e.message : e) };
  }
}


function _lerHierarquiaLocal(lista) {
  try {
    const fonte = lista.fonte || lista;
    const sheet = SpreadsheetApp.getActive().getSheetByName(fonte.nomeAba);
    if (!sheet) return {};
    
    const getIdx = (c) => {
        if(typeof c === 'number') return c;
        if(typeof Utils!=='undefined') return Utils.colunaParaIndice(c);
        let col=0,s=String(c).toUpperCase(); for(let i=0;i<s.length;i++) col=col*26+(s.charCodeAt(i)-64); return col;
    };

    const idxCat = getIdx(fonte.colunaCategoria);
    const idxSub = fonte.colunaSubcategoria ? getIdx(fonte.colunaSubcategoria) : null;
    const rowStart = fonte.linhaInicial || 4;
    const lastRow = Math.min(sheet.getLastRow(), rowStart + 3000);
    
    if (lastRow < rowStart) return {};

    const vals = sheet.getRange(rowStart, 1, lastRow - rowStart + 1, Math.max(idxCat, idxSub || 0)).getValues();
    const mapa = {}; 

    vals.forEach(row => {
      const cat = String(row[idxCat-1]||'').trim();
      if(!cat) return;
      if(!mapa[cat]) mapa[cat] = new Set();
      if(idxSub) {
        const sub = String(row[idxSub-1]||'').trim();
        if(sub) mapa[cat].add(sub);
      }
    });

    const res = {};
    Object.keys(mapa).sort().forEach(k => res[k] = Array.from(mapa[k]).sort());
    return res;
  } catch (e) { return {}; }
}