/**
 * regrasConfig.js — Fonte oficial (DSL) das Regras de Negócio
 * Consumo: RegrasNegocioV2 v1.1 (engine sem eval; condições simples por evento)
 */

var ACOES_SUPORTADAS = [
  "setar_data_se_vazio",
  "copiar_valor",
  "limpar_celula",
  "aplicar_formato",
  "mostrar_mensagem",
  "sincronizar_data_por_situacao",
  // NOVA (usada nas regras abaixo):
  "carimbar_data_por_preenchimento",
  // NOVAS ações usadas nas regras SEGUROS
  "restaurar_produto_seguros",
  "proteger_edicao_sensivel"
  ];

/**
 * REGRAS_DE_NEGOCIO — lista oficial de regras
 * Observação:
 *  - Normalização de comparações (case/acentos) é feita no engine.
 *  - Datas gravadas são objetos Date reais (00:00 local) + formato aplicado.
 */
var REGRAS_DE_NEGOCIO = [
  // SEGUROS — regra única I⟷M (duas faces da mesma moeda)
  //   I = "PAGO/CONTRATADO"  → M = hoje (se vazia)   [face ativada]
  //   I ≠ "PAGO/CONTRATADO" → M = limpar            [face desativada]
  {
    id: "SEG_I2M_SYNC_001",
    status: "ATIVO",
    abaAlvo: "SEGUROS🛡️",
    colunasAlvo: ["I"],               // gatilho: edição na coluna I
    prioridade: 1,                    // menor = executa primeiro
    acao: "sincronizar_data_por_situacao",
    parametros: {
      coluna_destino: "K",
      formato: "dd/MM/yyyy",
      valor_ativador: "PAGO/CONTRATADO",
      politica_quando_ativado: "SETAR_HOJE_SE_VAZIO", // SETAR_HOJE | SETAR_HOJE_SE_VAZIO | PRESERVAR
      politica_outros: "LIMPAR"                        // LIMPAR | PRESERVAR
    },
    descricao: "Sincroniza M com a situação de I: preenche hoje se I=PAGO/CONTRATADO (se vazia) e limpa M caso contrário."
  },

  // CADASTROS — Sincroniza I (Data) com F (Status)
  // F = "ATUALIZADO/CONCLUIDO" → I = hoje (se vazia)
  // F ≠ "ATUALIZADO/CONCLUIDO" → I = limpar
  {
    id: "CAD_F2I_SYNC_001",
    status: "ATIVO",
    abaAlvo: "CADASTROS🧑‍💻",
    colunasAlvo: ["F"],
    prioridade: 1,
    acao: "sincronizar_data_por_situacao",
    parametros: {
      coluna_destino: "I",
      formato: "dd/MM/yyyy",
      valor_ativador: "ATUALIZADO/CONCLUIDO",
      politica_quando_ativado: "SETAR_HOJE_SE_VAZIO",
      politica_outros: "LIMPAR"
    },
    descricao: "Sincroniza I com a situação de F: preenche data de hoje se F=ATUALIZADO/CONCLUIDO (se vazia) e limpa I caso contrário."
  },

  {
    id: "DEMANDAS_F2I_SYNC_001",
    status: "ATIVO",
    abaAlvo: "DEMANDAS DIVERSAS🔧",
    colunasAlvo: ["F"],
    prioridade: 1,
    acao: "sincronizar_data_por_situacao",
    parametros: {
      coluna_destino: "I",
      formato: "dd/MM/yyyy",
      valor_ativador: "CONCLUIDO",
      politica_quando_ativado: "SETAR_HOJE_SE_VAZIO",
      politica_outros: "LIMPAR"
    },
    descricao: "Sincroniza I com a situação de F: preenche data de hoje se F=ATUALIZADO/CONCLUIDO (se vazia) e limpa I caso contrário."
  },

  // INTERNALIZADO — carimba data em I quando QUALQUER entre A,B,D,E for preenchido; limpa se todos ficarem vazios
  {
    id: "INTL_PREENCHIMENTO_DATA_I_001",
    status: "ATIVO",
    abaAlvo: "INTERNALIZADO🎯",
    colunasAlvo: ["A","B","D","E"],   // gatilhos de edição
    prioridade: 1,
    acao: "carimbar_data_por_preenchimento",
    parametros: {
      coluna_destino: "I",
      formato: "dd/MM/yyyy",
      colunas_monitoradas: ["A","B","D","E"],
      criterio: "QUALQUER",                       // QUALQUER | TODAS
      politica_quando_preenchido: "SETAR_HOJE_SE_VAZIO",
      politica_quando_vazio: "LIMPAR"
    },
    descricao: "Carimba a data em I quando algum campo-chave (A,B,D,E) da linha é preenchido; limpa I se todos ficarem vazios."
  },

  {
    id: "DEMANDASL_PREENCHIMENTO_DATA_I_001",
    status: "ATIVO",
    abaAlvo: "DEMANDAS DIVERSAS🔧",
    colunasAlvo: ["A","B","D","E"],   // gatilhos de edição
    prioridade: 1,
    acao: "carimbar_data_por_preenchimento",
    parametros: {
      coluna_destino: "D",
      formato: "dd/MM/yyyy",
      colunas_monitoradas: ["A","B","D","E"],
      criterio: "QUALQUER",                       // QUALQUER | TODAS
      politica_quando_preenchido: "SETAR_HOJE_SE_VAZIO",
      politica_quando_vazio: "LIMPAR"
    },
    descricao: "Carimba a data em D quando algum campo-chave (A,B,D,E) da linha é preenchido; limpa I se todos ficarem vazios."
  },


  // EM ANALISE — carimba data em I quando QUALQUER entre A,B,D,E,F for preenchido; limpa se todos ficarem vazios
    {
    id: "ANALIS_PREENCHIMENTO_DATA_I_001",
    status: "ATIVO",
    abaAlvo: "EM ANALISE📊",
    colunasAlvo: ["A","B","D","E","F"],   // gatilhos de edição
    prioridade: 1,
    acao: "carimbar_data_por_preenchimento",
    parametros: {
      coluna_destino: "I",
      formato: "dd/MM/yyyy",
      colunas_monitoradas: ["A","B","D","E","F"],
      criterio: "QUALQUER",                       // QUALQUER | TODAS
      politica_quando_preenchido: "SETAR_HOJE_SE_VAZIO",
      politica_quando_vazio: "LIMPAR"
    },
    descricao: "Carimba a data em I quando algum campo-chave (A,B,D,E,F) da linha é preenchido; limpa I se todos ficarem vazios."
  },

    {
    id: "EMAN_PREENCHIMENTO_DATA_I_002",
    status: "ATIVO",
    abaAlvo: "SEGUROS🛡️",
    colunasAlvo: ["A","B","C", "D","E","F"], // gatilhos de edição
    prioridade: 1,
    acao: "carimbar_data_por_preenchimento",
    parametros: {
      coluna_destino: "H",
      formato: "dd/MM/yyyy",
      colunas_monitoradas: ["A","B", "C", "D","E","F"],
      criterio: "QUALQUER",
      politica_quando_preenchido: "SETAR_HOJE_SE_VAZIO",
      politica_quando_vazio: "LIMPAR"
    },
    descricao: "Carimba a data em H quando algum campo-chave (A,B,D,E,F) é preenchido; limpa H se todos ficarem vazios."
    },
  // CADASTROS — carimba data em D quando QUALQUER entre A,B,C for preenchido; limpa se todos ficarem vazios
    {
      id: "CAD_PREENCHIMENTO_DATA_D_001",
      status: "ATIVO",
      abaAlvo: "CADASTROS🧑‍💻",             // Aba alvo
      colunasAlvo: ["A","B","C"],       // Gatilhos de edição
      prioridade: 1,
      acao: "carimbar_data_por_preenchimento", // Ação suportada
      parametros: {
        coluna_destino: "D",
        formato: "dd/MM/yyyy",
        colunas_monitoradas: ["A","B","C"], // Colunas monitoradas
        criterio: "QUALQUER",               // QUALQUER | TODAS
        politica_quando_preenchido: "SETAR_HOJE_SE_VAZIO",
        politica_quando_vazio: "LIMPAR"     // Limpa D se A,B,C estiverem vazios
      },
      descricao: "Carimba a data em D quando algum campo-chave (A,B,C) da linha é preenchido; limpa D se todos ficarem vazios."
    },

   // ===== SEGUROS — R2: restauração se C foi apagado mas a origem ainda contém o produto
  {
    id: "SEGUROS_RESTAURAR_C_001",
    status: "INATIVO",
    abaAlvo: "SEGUROS🛡️",
    colunasAlvo: ["C"],
    prioridade: 1,
    acao: "restaurar_produto_seguros",
    parametros: {
      startRow: 4,
      coluna_produto: "C",
      origem_coluna_produtos: "J",             // EM ANALISE!J
      nota_prefixo: "origem=",                 // padrão do módulo PSS
      cor_pendencia: "#FFF2CC",
      nota_restauracao: "Registro restaurado automaticamente por divergência."
    },
    descricao: "Restaura C se usuário limpar e a origem ainda listar o produto."
  },

  // ===== SEGUROS — R3: edição sensível em E
{
  id: "SEGUROS_EDIT_SENSIVEL_E_001",
  status: "INATIVO",
  abaAlvo: "SEGUROS🛡️",
  colunasAlvo: ["E"],
  prioridade: 1,
  acao: "proteger_edicao_sensivel",
  parametros: {
    startRow: 4,
    coluna_sensivel: "E",
    cols_linha_preenchida: ["A","B","C","D","E","F"],
    confirmacao_2_toques: { ttlSegundos: 30 },
    cor_pendencia: "#FFF2CC",
    msg_aviso: "Linha sensível. Edite novamente em até 30s para confirmar."
  },
  descricao: "Protege E com 2 toques; destaca a linha; restaura visuais/formatos no 2º toque."
},

// SEGUROS — Validação Complexa: Operação Vinculada (D) e Vigência da Apólice (L vs K)
  // GATILHOS: Edição em C, D, I, K ou L
  // SEGUROS — Validação Complexa: Operação Vinculada (D) e Vigência da Apólice (L vs K)
  // GATILHOS: Edição em C, D, I, K ou L. Regra única para consolidar checagens de integridade.
    {
    id: "SEG_VALIDACAO_COMPLEXA_001",
    status: "ATIVO",
    abaAlvo: "SEGUROS🛡️",
    // ATUALIZADO: Gatilhos expandidos para incluir F e G
    colunasAlvo: ["C", "D", "F", "G", "I", "K", "L"], 
    prioridade: 1,
    acao: "validar_seguro_e_vigencia", 
    parametros: {
      // 1. Validação de D (Operação Vinculada)
      tipos_exigem_D: ["SEG. RD (EQUIPAMENTOS/VEICULO)", "SEG. VEICULO", "SEG. EMP. (PRÉDIO/BENS)"],
      coluna_produto: "C",
      coluna_operacao: "D",
      msg_erro_D: "Inserir número da operação vinculada, se houver.",

      // 2. NOVAS VALIDAÇÕES DE OBRIGATORIEDADE (F e G)
      coluna_valor_seguro: "F",
      msg_erro_F: "Selecionar modo de pagamento",
      coluna_numero_proposta: "G",
      msg_erro_G: "Inserir número do docsflow da op. vinculada, se houver.",
      
      // 3. Validação de L (Data Fim Vigência)
      coluna_status: "I",
      valor_ativador_L: "PAGO/CONTRATADO",
      coluna_data_inicio: "K", 
      coluna_data_fim: "L",
      msg_erro_L: "Inserir data fim de vigência da apólice",
      
      // Configuração de Tolerância (Múltiplos de Ano)
      tolerancia_dias: 30, 
      msg_erro_L_vigencia: "ERRO: Vigência Inválida. A apólice deve ser de 1 ano ou múltiplos (ex: 2, 3 anos), com tolerância de ±30 dias."
    },
    descricao: "Realiza validações complexas em C, D, F, G, I, K e L (Campos Obrigatórios e Vigência Anual/Plurianual)."
  },

];

/**
 * Diagnóstico mínimo (direto e estruturado)
 * - Verifica: ids únicos, status válido, ação suportada, colunasAlvo válidas e abaAlvo preenchida.
 * - Retorna { ok, erros, resumo, total } e loga um resumo se Logger existir.
 */
function RegrasConfig_diagnosticoMin() {
  function norm(s){ return (s||"").toString().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase(); }
  function colToIdx(c){
    if (typeof c === "number") return c;
    var s = (c||"").toString().trim().toUpperCase();
    var n = 0; for (var i=0;i<s.length;i++) n = n*26 + (s.charCodeAt(i)-64);
    return n;
  }

  var erros = [];
  var ids = Object.create(null);
  var resumo = Object.create(null);

  for (var i=0; i<REGRAS_DE_NEGOCIO.length; i++){
    var r = REGRAS_DE_NEGOCIO[i] || {};
    var id = r.id || ("<sem_id_"+i+">");
    var st = norm(r.status);
    var acao = r.acao || "";
    var aba = r.abaAlvo || "";
    var cols = Array.isArray(r.colunasAlvo) ? r.colunasAlvo : [];

    // id único
    if (ids[id]) erros.push("ID duplicado: "+id); else ids[id]=true;

    // status
    if (st !== "ATIVO" && st !== "INATIVO") erros.push(id+": status inválido ("+r.status+")");

    // ação suportada
    if (ACOES_SUPORTADAS.indexOf(acao) === -1) erros.push(id+": ação não suportada ("+acao+")");

    // abaAlvo
    if (!aba) erros.push(id+": abaAlvo vazia");

    // colunasAlvo
    if (cols.length === 0) {
      erros.push(id+": colunasAlvo vazia");
    } else {
      for (var k=0;k<cols.length;k++){
        var idx = colToIdx(cols[k]);
        if (!idx || idx < 1) erros.push(id+": coluna inválida ("+cols[k]+")");
      }
    }

    // resumo simples por aba
    var key = aba || "<sem_aba>";
    resumo[key] = (resumo[key] || 0) + 1;
  }

  var ok = (erros.length === 0);
  try {
    Logger.log("[RegrasConfig] OK? "+ok+" | Regras: "+REGRAS_DE_NEGOCIO.length+" | Por aba: "+JSON.stringify(resumo));
    if (!ok) Logger.log("[RegrasConfig] Erros: "+JSON.stringify(erros));
  } catch (_){}

  return { ok: ok, erros: erros, resumo: resumo, total: REGRAS_DE_NEGOCIO.length };
};
