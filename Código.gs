/***************************************************************
 * MOZAGRO — SISTEMA DE RASTREABILIDADE V5.1
 * Google Apps Script — Backend
 *
 * Compatível com:
 *   TelaInicial.html
 *   RASTREABILIDADE_MP A:AC
 *   ESTOQUE_LOTES A:K
 *
 * PRINCIPAIS REGRAS:
 *   Diferença = Pesado - Teórico
 *   % Diferença = Diferença / Teórico
 *   Tolerância = 1%
 *
 * RASTREABILIDADE_MP:
 * A  ID rastreabilidade
 * B  ID produção
 * C  ID batida
 * D  Data produção
 * E  Produto
 * F  Lote produto
 * G  Batida
 * H  Matéria-prima/Ingrediente
 * I  Fornecedor
 * J  Fabricante
 * K  Lote da MP
 * L  Validade da MP
 * M  Qtd. por batida (kg)
 * N  Nº de batidas
 * O  Qtd. calculada pela formulação (kg)
 * P  Qtd. realmente pesada (kg)
 * Q  Diferença (kg)
 * R  % diferença
 * S  Status
 * T  NF de origem
 * U  Data de recebimento da MP
 * V  Observações
 * W  Documento de origem
 * X  Controle do lote
 * Y  Recebimentos encontrados
 * Z  Saldo do lote após consumo
 * AA Chave MP + Lote
 * AB Chave MP + Lote + NF
 * AC Chave Produção + Batida + Lote
 ***************************************************************/


/* ============================================================
   CONFIGURAÇÃO
   ============================================================ */

const CONFIG = {
  SPREADSHEET_ID: '1wWqUvrG3TJ023BjYlHXlTrB-4VEsvUJh3yK-8zTxysw',

  ABAS: {
    RECEBIMENTOS: 'RECEBIMENTOS',
    PRODUCOES: 'PRODUÇÕES',
    FORMULACOES: 'FORMULAÇÕES',
    RASTREABILIDADE: 'RASTREABILIDADE_MP',
    ESTOQUE: 'ESTOQUE_LOTES',
    CONSULTA: 'CONSULTA_RASTREABILIDADE',
    FICHA: 'FICHA_RASTREABILIDADE',
    LISTAS: 'LISTAS'
  },

  LINHA_CABECALHO: 4,
  PRIMEIRA_LINHA_DADOS: 5,

  // 0.01 = 1%
  TOLERANCIA_PESAGEM: 0.01,

  GEMINI_MODEL: 'gemini-2.5-flash'
};


/* ============================================================
   ACESSO À PLANILHA
   ============================================================ */

function obterPlanilha() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}


function obterAba(nome) {
  const ss = obterPlanilha();
  const aba = ss.getSheetByName(nome);

  if (!aba) {
    throw new Error('Aba não encontrada: ' + nome);
  }

  return aba;
}


function doGet() {
  return HtmlService
    .createTemplateFromFile('TelaInicial')
    .evaluate()
    .setTitle('MOZAGRO — Rastreabilidade V5.1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function incluirArquivo(dados) {
  if (!dados || !dados.base64) {
    throw new Error('Arquivo não recebido.');
  }

  const bytes = Utilities.base64Decode(dados.base64);
  const blob = Utilities.newBlob(
    bytes,
    dados.mimeType || MimeType.BINARY,
    dados.nome || 'arquivo'
  );

  const arquivo = DriveApp.createFile(blob);

  return {
    sucesso: true,
    id: arquivo.getId(),
    nome: arquivo.getName(),
    url: arquivo.getUrl()
  };
}


/* ============================================================
   UTILITÁRIOS
   ============================================================ */

function normalizarTexto_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
}


function normalizarCabecalho_(valor) {
  return normalizarTexto_(valor)
    .replace(/[º°ª]/g, '')
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}


function texto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}


function numero_(valor) {
  if (valor === null || valor === undefined || valor === '') {
    return null;
  }

  if (typeof valor === 'number') {
    return isNaN(valor) ? null : valor;
  }

  let s = String(valor).trim();

  if (!s) return null;

  s = s.replace(/\s/g, '');

  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  }

  const n = Number(s);

  return isNaN(n) ? null : n;
}


function dataValida_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  if (!valor) return null;

  const s = String(valor).trim();

  // yyyy-mm-dd
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (m) {
    const d = new Date(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3])
    );

    return isNaN(d.getTime()) ? null : d;
  }

  // dd/mm/yyyy
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);

  if (m) {
    const d = new Date(
      Number(m[3]),
      Number(m[2]) - 1,
      Number(m[1])
    );

    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);

  return isNaN(d.getTime()) ? null : d;
}


function formatarData_(valor) {
  const d = dataValida_(valor);

  if (!d) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}


function formatarDataISO_(valor) {
  const d = dataValida_(valor);

  if (!d) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function formatarPercentual_(valor) {
  const n = numero_(valor);

  if (n === null) return '';

  return (n * 100).toFixed(2) + '%';
}


function gerarID(prefixo) {
  const agora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMddHHmmss'
  );

  const aleatorio = Math.floor(Math.random() * 9000) + 1000;

  return prefixo + '-' + agora + '-' + aleatorio;
}


function chaveNormalizada_(valor) {
  return normalizarTexto_(valor)
    .replace(/\s+/g, ' ');
}


function chaveMpLote_(mp, lote) {
  return chaveNormalizada_(mp) + '|' + chaveNormalizada_(lote);
}


function chaveMpLoteNf_(mp, lote, nf) {
  return chaveMpLote_(mp, lote) + '|' + chaveNormalizada_(nf);
}


function chaveProducao_(idProducao, batida, loteProduto) {
  return (
    chaveNormalizada_(idProducao) +
    '|' +
    chaveNormalizada_(batida) +
    '|' +
    chaveNormalizada_(loteProduto)
  );
}


/* ============================================================
   CABEÇALHOS / LEITURA GENÉRICA
   ============================================================ */

function lerRegistros_(nomeAba) {
  const sheet = obterAba(nomeAba);
  const lastCol = sheet.getLastColumn();

  if (lastCol < 1) {
    return {
      sheet: sheet,
      headers: [],
      map: {},
      registros: []
    };
  }

  const headers = sheet
    .getRange(CONFIG.LINHA_CABECALHO, 1, 1, lastCol)
    .getValues()[0];

  const map = {};

  headers.forEach(function(h, i) {
    const key = normalizarCabecalho_(h);

    if (key) {
      map[key] = i;
    }
  });

  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.PRIMEIRA_LINHA_DADOS) {
    return {
      sheet: sheet,
      headers: headers,
      map: map,
      registros: []
    };
  }

  const valores = sheet
    .getRange(
      CONFIG.PRIMEIRA_LINHA_DADOS,
      1,
      lastRow - CONFIG.PRIMEIRA_LINHA_DADOS + 1,
      lastCol
    )
    .getValues();

  const registros = valores.map(function(row, index) {
    return {
      linha: CONFIG.PRIMEIRA_LINHA_DADOS + index,
      valores: row,
      headers: headers,
      map: map
    };
  });

  return {
    sheet: sheet,
    headers: headers,
    map: map,
    registros: registros
  };
}


function indicePorAliases_(headers, aliases) {
  const normalizados = headers.map(normalizarCabecalho_);

  for (let i = 0; i < aliases.length; i++) {
    const alvo = normalizarCabecalho_(aliases[i]);

    const idx = normalizados.indexOf(alvo);

    if (idx >= 0) {
      return idx;
    }
  }

  return -1;
}


function valorCampo_(registro, aliases) {
  const idx = indicePorAliases_(registro.headers, aliases);

  if (idx < 0) return '';

  return registro.valores[idx];
}


function definirCampo_(valores, headers, aliases, valor) {
  const idx = indicePorAliases_(headers, aliases);

  if (idx >= 0) {
    valores[idx] = valor;
    return true;
  }

  return false;
}


function encontrarPrimeiraLinhaLivre(aba) {
  if (!aba) {
    throw new Error(
      'encontrarPrimeiraLinhaLivre exige uma aba como argumento.'
    );
  }

  const lastRow = aba.getLastRow();

  return Math.max(
    CONFIG.PRIMEIRA_LINHA_DADOS,
    lastRow + 1
  );
}


/* ============================================================
   RECEBIMENTOS
   ============================================================ */

function salvarRecebimento(dados) {
  if (!dados) {
    throw new Error('Dados do recebimento não informados.');
  }

  const materiaPrima = texto_(
    dados.materiaPrima ||
    dados.materiaPrimaRecebimento
  );

  const lote = texto_(
    dados.lote ||
    dados.loteRecebimento
  );

  const quantidade = numero_(
    dados.quantidade ||
    dados.quantidadeRecebimento
  );

  if (!materiaPrima) {
    throw new Error('Informe a matéria-prima.');
  }

  if (!lote) {
    throw new Error('Informe o lote da matéria-prima.');
  }

  if (quantidade === null || quantidade <= 0) {
    throw new Error('Informe uma quantidade recebida válida.');
  }

  const aba = obterAba(CONFIG.ABAS.RECEBIMENTOS);

  const lastCol = aba.getLastColumn();

  const headers = aba
    .getRange(CONFIG.LINHA_CABECALHO, 1, 1, lastCol)
    .getValues()[0];

  const linha = new Array(lastCol).fill('');

  definirCampo_(
    linha,
    headers,
    ['ID recebimento', 'ID', 'ID recebimento MP'],
    gerarID('REC')
  );

  definirCampo_(
    linha,
    headers,
    ['Data recebimento', 'Data de recebimento', 'Recebimento'],
    dataValida_(
      dados.dataRecebimento ||
      dados.data ||
      new Date()
    )
  );

  definirCampo_(
    linha,
    headers,
    [
      'Matéria-prima',
      'Materia prima',
      'Matéria prima',
      'MP',
      'Ingrediente'
    ],
    materiaPrima
  );

  definirCampo_(
    linha,
    headers,
    ['Lote', 'Lote da MP', 'Lote MP'],
    lote
  );

  definirCampo_(
    linha,
    headers,
    ['Fornecedor'],
    texto_(dados.fornecedor)
  );

  definirCampo_(
    linha,
    headers,
    ['Fabricante'],
    texto_(dados.fabricante)
  );

  definirCampo_(
    linha,
    headers,
    [
      'Quantidade recebida',
      'Qtd recebida',
      'Quantidade'
    ],
    quantidade
  );

  definirCampo_(
    linha,
    headers,
    ['Unidade', 'Un'],
    texto_(dados.unidade || 'KG')
  );

  definirCampo_(
    linha,
    headers,
    [
      'NF',
      'Nota fiscal',
      'Nº NF',
      'Numero NF',
      'NF recebimento'
    ],
    texto_(dados.nf)
  );

  definirCampo_(
    linha,
    headers,
    [
      'Validade',
      'Data de validade',
      'Validade da MP'
    ],
    dataValida_(dados.validade)
  );

  definirCampo_(
    linha,
    headers,
    [
      'Documento de origem',
      'Documento origem',
      'Documento'
    ],
    texto_(dados.documentoOrigem)
  );

  definirCampo_(
    linha,
    headers,
    ['Observações', 'Observacao', 'Obs', 'Obs.'],
    texto_(dados.observacoes)
  );

  // Validação mínima da estrutura
  if (
    indicePorAliases_(
      headers,
      ['Matéria-prima', 'Materia prima', 'MP']
    ) < 0
  ) {
    throw new Error(
      'Não encontrei a coluna de matéria-prima na aba RECEBIMENTOS.'
    );
  }

  if (
    indicePorAliases_(
      headers,
      ['Lote', 'Lote da MP', 'Lote MP']
    ) < 0
  ) {
    throw new Error(
      'Não encontrei a coluna de lote na aba RECEBIMENTOS.'
    );
  }

  aba
    .getRange(
      encontrarPrimeiraLinhaLivre(aba),
      1,
      1,
      lastCol
    )
    .setValues([linha]);

  try {
    atualizarEstoqueLotes();
  } catch (e) {
    console.log('Aviso ao atualizar estoque: ' + e.message);
  }

  return {
    sucesso: true,
    mensagem: 'Recebimento salvo com sucesso.',
    materiaPrima: materiaPrima,
    lote: lote,
    quantidade: quantidade
  };
}


/* ============================================================
   BUSCA DE RECEBIMENTOS
   ============================================================ */

function buscarRecebimentos_(materiaPrima, lote, nf) {
  const dados = lerRegistros_(CONFIG.ABAS.RECEBIMENTOS);

  const mpBusca = chaveNormalizada_(materiaPrima);
  const loteBusca = chaveNormalizada_(lote);
  const nfBusca = chaveNormalizada_(nf);

  return dados.registros.filter(function(r) {

    const mp = chaveNormalizada_(
      valorCampo_(r, [
        'Matéria-prima',
        'Materia prima',
        'MP',
        'Ingrediente'
      ])
    );

    const lot = chaveNormalizada_(
      valorCampo_(r, [
        'Lote',
        'Lote da MP',
        'Lote MP'
      ])
    );

    const nota = chaveNormalizada_(
      valorCampo_(r, [
        'NF',
        'Nota fiscal',
        'Nº NF',
        'Numero NF'
      ])
    );

    if (mp !== mpBusca) return false;

    if (loteBusca && lot !== loteBusca) {
      return false;
    }

    if (nfBusca && nota !== nfBusca) {
      return false;
    }

    return true;
  });
}


function buscarRecebimentoExato_(materiaPrima, lote, nf) {
  const encontrados = buscarRecebimentos_(
    materiaPrima,
    lote,
    nf
  );

  if (!encontrados.length) {
    return null;
  }

  // Com NF informado, deve existir correspondência exata.
  if (nf) {
    if (encontrados.length === 1) {
      return encontrados[0];
    }

    return null;
  }

  // Sem NF:
  // só aceita automaticamente se houver um único recebimento.
  if (encontrados.length === 1) {
    return encontrados[0];
  }

  return null;
}


function obterDadosRecebimento_(registro) {
  if (!registro) return null;

  return {
    linha: registro.linha,

    materiaPrima: texto_(
      valorCampo_(registro, [
        'Matéria-prima',
        'Materia prima',
        'MP',
        'Ingrediente'
      ])
    ),

    lote: texto_(
      valorCampo_(registro, [
        'Lote',
        'Lote da MP',
        'Lote MP'
      ])
    ),

    fornecedor: texto_(
      valorCampo_(registro, ['Fornecedor'])
    ),

    fabricante: texto_(
      valorCampo_(registro, ['Fabricante'])
    ),

    quantidade: numero_(
      valorCampo_(registro, [
        'Quantidade recebida',
        'Qtd recebida',
        'Quantidade'
      ])
    ) || 0,

    unidade: texto_(
      valorCampo_(registro, ['Unidade', 'Un'])
    ),

    nf: texto_(
      valorCampo_(registro, [
        'NF',
        'Nota fiscal',
        'Nº NF',
        'Numero NF'
      ])
    ),

    dataRecebimento:
      valorCampo_(registro, [
        'Data recebimento',
        'Data de recebimento',
        'Recebimento'
      ]),

    validade:
      valorCampo_(registro, [
        'Validade',
        'Data de validade',
        'Validade da MP'
      ]),

    documentoOrigem: texto_(
      valorCampo_(registro, [
        'Documento de origem',
        'Documento origem',
        'Documento'
      ])
    ),

    observacoes: texto_(
      valorCampo_(registro, [
        'Observações',
        'Observacao',
        'Obs',
        'Obs.'
      ])
    )
  };
}


/* ============================================================
   FORMULAÇÕES
   ============================================================ */

function obterQuantidadePorBatida_(registro) {
  return numero_(
    valorCampo_(registro, [
      'Qtd. por batida (kg)',
      'Qtd por batida (kg)',
      'Quantidade por batida (kg)',
      'Quantidade por batida',
      'Qtd por batida',
      'Quantidade',
      'Qtd'
    ])
  );
}


function localizarFormulacaoAtiva(produto, dataProducao) {
  const registros = lerRegistros_(
    CONFIG.ABAS.FORMULACOES
  );

  const produtoBusca = chaveNormalizada_(produto);

  const dataRef =
    dataValida_(dataProducao) ||
    new Date();

  const candidatos = [];

  registros.registros.forEach(function(r) {

    const prod = chaveNormalizada_(
      valorCampo_(r, [
        'Produto',
        'Produto acabado',
        'Produto final'
      ])
    );

    if (!prod || prod !== produtoBusca) {
      return;
    }

    const status = normalizarTexto_(
      valorCampo_(r, [
        'Status',
        'Situação',
        'Situacao'
      ])
    );

    if (
      status === 'INATIVA' ||
      status === 'INATIVO' ||
      status === 'CANCELADA' ||
      status === 'CANCELADO' ||
      status === 'NAO' ||
      status === 'NÃO'
    ) {
      return;
    }

    const inicio = dataValida_(
      valorCampo_(r, [
        'Data início',
        'Data inicio',
        'Início',
        'Inicio',
        'Vigência início',
        'Vigencia inicio'
      ])
    );

    const fim = dataValida_(
      valorCampo_(r, [
        'Data fim',
        'Fim',
        'Data final',
        'Vigência fim',
        'Vigencia fim'
      ])
    );

    if (inicio && dataRef < inicio) {
      return;
    }

    if (fim && dataRef > fim) {
      return;
    }

    const versao = texto_(
      valorCampo_(r, [
        'Versão',
        'Versao',
        'Versão da formulação',
        'Versao da formulacao'
      ])
    );

    candidatos.push({
      registro: r,
      versao: versao,
      inicio: inicio
    });
  });

  if (!candidatos.length) {
    return {
      encontrada: false,
      linhas: [],
      versao: '',
      mensagem:
        'Nenhuma formulação ativa encontrada para o produto na data informada.'
    };
  }

  // Agrupa por versão.
  const grupos = {};

  candidatos.forEach(function(c) {

    const chave =
      c.versao ||
      (
        c.inicio
          ? formatarDataISO_(c.inicio)
          : 'SEM_VERSAO'
      );

    if (!grupos[chave]) {
      grupos[chave] = [];
    }

    grupos[chave].push(c);
  });

  const gruposLista = Object.keys(grupos).map(function(chave) {
    return {
      chave: chave,
      itens: grupos[chave]
    };
  });

  gruposLista.sort(function(a, b) {

    const va = a.chave;
    const vb = b.chave;

    const na = Number(va);
    const nb = Number(vb);

    if (!isNaN(na) && !isNaN(nb)) {
      return nb - na;
    }

    const da =
      a.itens[0].inicio
        ? a.itens[0].inicio.getTime()
        : 0;

    const db =
      b.itens[0].inicio
        ? b.itens[0].inicio.getTime()
        : 0;

    if (db !== da) {
      return db - da;
    }

    return String(vb).localeCompare(String(va));
  });

  const escolhido = gruposLista[0];

  return {
    encontrada: true,
    versao: escolhido.itens[0].versao,
    chave: escolhido.chave,
    linhas: escolhido.itens.map(function(x) {
      return x.registro;
    }),
    mensagem: 'Formulação ativa encontrada.'
  };
}


function localizarLinhasFormulacao_(produto, dataProducao) {
  const resultado = localizarFormulacaoAtiva(
    produto,
    dataProducao
  );

  return resultado.linhas || [];
}


/* ============================================================
   ESTOQUE / CONSUMO
   ============================================================ */

function calcularEstoquePorMpLote_(materiaPrima, lote) {
  const recebimentos = buscarRecebimentos_(
    materiaPrima,
    lote,
    ''
  );

  let quantidadeRecebida = 0;
  let unidade = '';
  let primeiroRecebimento = null;

  const fornecedores = {};
  const nfs = {};

  recebimentos.forEach(function(r) {

    const d = obterDadosRecebimento_(r);

    quantidadeRecebida += d.quantidade || 0;

    if (!unidade && d.unidade) {
      unidade = d.unidade;
    }

    if (d.fornecedor) {
      fornecedores[d.fornecedor] = true;
    }

    if (d.nf) {
      nfs[d.nf] = true;
    }

    const data = dataValida_(d.dataRecebimento);

    if (
      data &&
      (
        !primeiroRecebimento ||
        data < primeiroRecebimento
      )
    ) {
      primeiroRecebimento = data;
    }
  });

  const rast = lerRegistros_(
    CONFIG.ABAS.RASTREABILIDADE
  );

  let quantidadeConsumida = 0;

  rast.registros.forEach(function(r) {

    const mp = chaveNormalizada_(
      r.valores[7]
    );

    const lot = chaveNormalizada_(
      r.valores[10]
    );

    if (
      mp === chaveNormalizada_(materiaPrima) &&
      lot === chaveNormalizada_(lote)
    ) {
      const pesado = numero_(r.valores[15]);

      if (pesado !== null) {
        quantidadeConsumida += pesado;
      }
    }
  });

  const saldo =
    quantidadeRecebida -
    quantidadeConsumida;

  let status = 'NÃO ENCONTRADO';

  if (recebimentos.length) {

    if (saldo < -0.000001) {
      status = 'SALDO NEGATIVO';
    } else if (Math.abs(saldo) < 0.000001) {
      status = 'ESGOTADO';
    } else {
      status = 'ATIVO';
    }
  }

  return {
    materiaPrima: materiaPrima,
    lote: lote,
    quantidadeRecebida: quantidadeRecebida,
    quantidadeConsumida: quantidadeConsumida,
    saldo: saldo,
    unidade: unidade || 'KG',
    primeiroRecebimento: primeiroRecebimento,
    fornecedor: Object.keys(fornecedores).join(', '),
    nf: Object.keys(nfs).join(', '),
    status: status,
    quantidadeRecebimentos: recebimentos.length
  };
}


function consultarSaldoLote(materiaPrima, lote) {
  if (!materiaPrima || !lote) {
    return {
      encontrado: false,
      mensagem: 'Informe matéria-prima e lote.'
    };
  }

  const resultado = calcularEstoquePorMpLote_(
    materiaPrima,
    lote
  );

  resultado.encontrado =
    resultado.quantidadeRecebimentos > 0;

  return resultado;
}


/* ============================================================
   DIFERENÇA DE PESAGEM
   ============================================================ */

function calcularDiferencaPesagem_(teorica, pesada) {
  const O = numero_(teorica);
  const P = numero_(pesada);

  if (O === null || P === null) {
    return {
      diferenca: null,
      percentual: null,
      status: ''
    };
  }

  const diferenca = P - O;

  if (O === 0) {
    return {
      diferenca: diferenca,
      percentual: null,
      status: 'VERIFICAR'
    };
  }

  const percentual = diferenca / O;

  const status =
    Math.abs(percentual) <=
    CONFIG.TOLERANCIA_PESAGEM
      ? 'OK'
      : 'VERIFICAR';

  return {
    diferenca: diferenca,
    percentual: percentual,
    status: status
  };
}


/* ============================================================
   RASTREABILIDADE — GERAÇÃO
   ============================================================ */

function gerarRastreabilidadeDaProducao_(
  dadosProducao,
  idProducao,
  idBatida
) {
  const linhasFormulacao =
    localizarLinhasFormulacao_(
      dadosProducao.produto,
      dadosProducao.dataProducao
    );

  if (!linhasFormulacao.length) {
    throw new Error(
      'Nenhuma linha de formulação ativa foi encontrada.'
    );
  }

  const aba = obterAba(
    CONFIG.ABAS.RASTREABILIDADE
  );

  const dataProducao =
    dataValida_(dadosProducao.dataProducao) ||
    new Date();

  const numeroBatidas =
    numero_(dadosProducao.numeroBatidas) || 1;

  const batida =
    texto_(dadosProducao.batida) ||
    'B' +
    String(numeroBatidas).padStart(2, '0');

  const linhas = [];

  linhasFormulacao.forEach(function(formula) {

    const materiaPrima = texto_(
      valorCampo_(formula, [
        'Matéria-prima/Ingrediente',
        'Matéria-prima',
        'Materia prima',
        'Ingrediente',
        'MP'
      ])
    );

    if (!materiaPrima) {
      return;
    }

    const qtdPorBatida =
      obterQuantidadePorBatida_(formula);

    if (
      qtdPorBatida === null ||
      qtdPorBatida === undefined
    ) {
      return;
    }

    const quantidadeTeorica =
      qtdPorBatida * numeroBatidas;

    /*
     * NÃO escolhemos silenciosamente o primeiro lote.
     *
     * Se existir somente um recebimento, podemos
     * relacioná-lo automaticamente.
     *
     * Se existirem vários, deixamos a origem
     * em aberto para evitar rastreabilidade falsa.
     */
    const recebimentos =
      buscarRecebimentos_(
        materiaPrima,
        '',
        ''
      );

    let recebimento = null;

    if (recebimentos.length === 1) {
      recebimento =
        obterDadosRecebimento_(
          recebimentos[0]
        );
    }

    let controleLote = '';

    if (recebimentos.length === 0) {
      controleLote = 'SEM RECEBIMENTO';
    } else if (recebimentos.length === 1) {
      controleLote = 'OK';
    } else {
      controleLote =
        'MÚLTIPLOS RECEBIMENTOS';
    }

    const loteMP =
      recebimento
        ? recebimento.lote
        : '';

    const nf =
      recebimento
        ? recebimento.nf
        : '';

    const validade =
      recebimento
        ? recebimento.validade
        : '';

    const fornecedor =
      recebimento
        ? recebimento.fornecedor
        : '';

    const fabricante =
      recebimento
        ? recebimento.fabricante
        : '';

    const dataRecebimento =
      recebimento
        ? recebimento.dataRecebimento
        : '';

    const documentoOrigem =
      texto_(dadosProducao.documentoOrigem);

    const observacoes =
      texto_(dadosProducao.observacoes);

    const linha = new Array(29).fill('');

    linha[0] = gerarID('RAST');
    linha[1] = idProducao;
    linha[2] = idBatida;
    linha[3] = dataProducao;
    linha[4] = texto_(dadosProducao.produto);
    linha[5] = texto_(dadosProducao.loteProduto);
    linha[6] = batida;
    linha[7] = materiaPrima;
    linha[8] = fornecedor;
    linha[9] = fabricante;
    linha[10] = loteMP;
    linha[11] = validade;
    linha[12] = qtdPorBatida;
    linha[13] = numeroBatidas;
    linha[14] = quantidadeTeorica;

    // P = quantidade realmente pesada.
    // Fica vazio até a pesagem real ser informada.
    linha[15] = '';

    // Q/R/S serão calculados pelas fórmulas.
    linha[16] = '';
    linha[17] = '';
    linha[18] = '';

    linha[19] = nf;
    linha[20] = dataRecebimento;
    linha[21] = observacoes;
    linha[22] = documentoOrigem;

    linha[23] = controleLote;
    linha[24] = recebimentos.length;
    linha[25] = '';

    linha[26] =
      chaveMpLote_(
        materiaPrima,
        loteMP
      );

    linha[27] =
      chaveMpLoteNf_(
        materiaPrima,
        loteMP,
        nf
      );

    linha[28] =
      chaveProducao_(
        idProducao,
        batida,
        dadosProducao.loteProduto
      );

    linhas.push(linha);
  });

  if (!linhas.length) {
    throw new Error(
      'A formulação não possui ingredientes válidos.'
    );
  }

  const primeiraLinha =
    encontrarPrimeiraLinhaLivre(aba);

  aba
    .getRange(
      primeiraLinha,
      1,
      linhas.length,
      29
    )
    .setValues(linhas);

  aplicarFormulasAutomaticas();

  return linhas.length;
}


/* ============================================================
   SALVAR PRODUÇÃO
   ============================================================ */

function salvarProducao(dados) {
  if (!dados) {
    throw new Error('Dados da produção não informados.');
  }

  const produto = texto_(
    dados.produto ||
    dados.produtoProducao
  );

  const loteProduto = texto_(
    dados.loteProduto ||
    dados.loteProdutoProducao
  );

  const dataProducao =
    dataValida_(
      dados.dataProducao ||
      dados.dataFabricacao ||
      dados.fabricacaoProducao
    ) || new Date();

  if (!produto) {
    throw new Error('Informe o produto.');
  }

  if (!loteProduto) {
    throw new Error('Informe o lote do produto.');
  }

  const formulacao =
    localizarFormulacaoAtiva(
      produto,
      dataProducao
    );

  if (!formulacao.encontrada) {
    return {
      sucesso: false,
      mensagem:
        'Não foi possível salvar a produção: ' +
        formulacao.mensagem
    };
  }

  const numeroBatidas =
    numero_(dados.numeroBatidas) || 1;

  const batida =
    texto_(dados.batida) ||
    'B' +
    String(numeroBatidas).padStart(2, '0');

  const idProducao =
    gerarID('PROD');

  const idBatida =
    gerarID('BAT');

  const aba =
    obterAba(CONFIG.ABAS.PRODUCOES);

  const lastCol =
    aba.getLastColumn();

  const headers =
    aba
      .getRange(
        CONFIG.LINHA_CABECALHO,
        1,
        1,
        lastCol
      )
      .getValues()[0];

  const linha =
    new Array(lastCol).fill('');

  definirCampo_(
    linha,
    headers,
    ['ID produção', 'ID producao', 'ID'],
    idProducao
  );

  definirCampo_(
    linha,
    headers,
    ['ID batida', 'ID lote batida'],
    idBatida
  );

  definirCampo_(
    linha,
    headers,
    [
      'Data produção',
      'Data producao',
      'Data fabricação',
      'Data fabricacao'
    ],
    dataProducao
  );

  definirCampo_(
    linha,
    headers,
    ['Produto', 'Produto acabado', 'Produto final'],
    produto
  );

  definirCampo_(
    linha,
    headers,
    ['Lote produto', 'Lote do produto', 'Lote'],
    loteProduto
  );

  definirCampo_(
    linha,
    headers,
    ['Espécie', 'Especie', 'Espécie destino', 'Destino'],
    texto_(dados.especie)
  );

  definirCampo_(
    linha,
    headers,
    [
      'Quantidade produzida',
      'Qtd produzida',
      'Quantidade'
    ],
    numero_(dados.quantidadeProduzida)
  );

  definirCampo_(
    linha,
    headers,
    [
      'Data validade',
      'Data de validade',
      'Validade'
    ],
    dataValida_(dados.dataValidade)
  );

  definirCampo_(
    linha,
    headers,
    ['Batida', 'Nº batida', 'Numero batida'],
    batida
  );

  definirCampo_(
    linha,
    headers,
    [
      'Nº batidas',
      'Numero batidas',
      'Número de batidas',
      'Qtd batidas'
    ],
    numeroBatidas
  );

  definirCampo_(
    linha,
    headers,
    ['Responsável', 'Responsavel'],
    texto_(dados.responsavel)
  );

  definirCampo_(
    linha,
    headers,
    [
      'Documento origem',
      'Documento de origem',
      'Documento'
    ],
    texto_(dados.documentoOrigem)
  );

  definirCampo_(
    linha,
    headers,
    ['Observações', 'Observacao', 'Obs', 'Obs.'],
    texto_(dados.observacoes)
  );

  const linhaLivre =
    encontrarPrimeiraLinhaLivre(aba);

  aba
    .getRange(
      linhaLivre,
      1,
      1,
      lastCol
    )
    .setValues([linha]);

  const qtdRastreabilidade =
    gerarRastreabilidadeDaProducao_(
      {
        produto: produto,
        loteProduto: loteProduto,
        dataProducao: dataProducao,
        numeroBatidas: numeroBatidas,
        batida: batida,
        documentoOrigem:
          dados.documentoOrigem,
        observacoes:
          dados.observacoes
      },
      idProducao,
      idBatida
    );

  try {
    atualizarEstoqueLotes();
  } catch (e) {
    console.log(
      'Aviso ao atualizar estoque: ' +
      e.message
    );
  }

  return {
    sucesso: true,
    mensagem:
      'Produção salva e rastreabilidade gerada.',
    idProducao: idProducao,
    idBatida: idBatida,
    linhasRastreabilidade:
      qtdRastreabilidade
  };
}


/* ============================================================
   CONSULTA DE PRODUÇÃO / RASTREABILIDADE
   ============================================================ */

function consultarProducao(produto, loteProduto) {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.RASTREABILIDADE
    );

  const produtoBusca =
    chaveNormalizada_(produto);

  const loteBusca =
    chaveNormalizada_(loteProduto);

  const resultado = [];

  dados.registros.forEach(function(r) {

    const row = r.valores;

    const produtoLinha =
      chaveNormalizada_(row[4]);

    const loteLinha =
      chaveNormalizada_(row[5]);

    if (
      produtoBusca &&
      produtoLinha !== produtoBusca
    ) {
      return;
    }

    if (
      loteBusca &&
      loteLinha !== loteBusca
    ) {
      return;
    }

    const teorica = numero_(row[14]);
    const pesada = numero_(row[15]);

    const calculo =
      calcularDiferencaPesagem_(
        teorica,
        pesada
      );

    const saldo =
      numero_(row[25]);

    resultado.push({
      linha: r.linha,

      idRastreabilidade:
        texto_(row[0]),

      idProducao:
        texto_(row[1]),

      idBatida:
        texto_(row[2]),

      dataProducao:
        formatarData_(row[3]),

      produto:
        texto_(row[4]),

      loteProduto:
        texto_(row[5]),

      batida:
        texto_(row[6]),

      materiaPrima:
        texto_(row[7]),

      fornecedor:
        texto_(row[8]),

      fabricante:
        texto_(row[9]),

      loteMP:
        texto_(row[10]),

      validadeMP:
        formatarData_(row[11]),

      quantidadePorBatida:
        numero_(row[12]),

      numeroBatidas:
        numero_(row[13]),

      quantidadeTeorica:
        teorica,

      quantidadePesada:
        pesada,

      diferenca:
        calculo.diferenca !== null
          ? calculo.diferenca
          : numero_(row[16]),

      percentual:
        calculo.percentual !== null
          ? formatarPercentual_(
              calculo.percentual
            )
          : formatarPercentual_(
              row[17]
            ),

      status:
        calculo.status ||
        texto_(row[18]),

      nfOrigem:
        texto_(row[19]),

      dataRecebimento:
        formatarData_(row[20]),

      observacoes:
        texto_(row[21]),

      documentoOrigem:
        texto_(row[22]),

      controleLote:
        texto_(row[23]),

      recebimentosEncontrados:
        numero_(row[24]) || 0,

      saldoLoteAposConsumo:
        saldo,

      chaveMpLote:
        texto_(row[26]),

      chaveMpLoteNf:
        texto_(row[27]),

      chaveProducao:
        texto_(row[28])
    });
  });

  return resultado;
}


/* ============================================================
   COMPLETAR ORIGEM DOS RECEBIMENTOS
   ============================================================ */

function completarDadosRecebimentoNaRastreabilidade_(
  linhaPlanilha,
  materiaPrima,
  lote,
  nf
) {
  const aba =
    obterAba(CONFIG.ABAS.RASTREABILIDADE);

  const recebimentos =
    buscarRecebimentos_(
      materiaPrima,
      lote || '',
      nf || ''
    );

  const range =
    aba.getRange(
      linhaPlanilha,
      1,
      1,
      29
    );

  const row =
    range.getValues()[0];

  row[24] =
    recebimentos.length;

  // Só preenche dados da origem se houver
  // uma correspondência não ambígua.
  let recebimento = null;

  if (nf && recebimentos.length === 1) {
    recebimento =
      obterDadosRecebimento_(
        recebimentos[0]
      );
  } else if (
    !nf &&
    lote &&
    recebimentos.length === 1
  ) {
    recebimento =
      obterDadosRecebimento_(
        recebimentos[0]
      );
  }

  if (recebimento) {

    row[8] =
      recebimento.fornecedor;

    row[9] =
      recebimento.fabricante;

    row[10] =
      recebimento.lote;

    row[11] =
      recebimento.validade;

    row[19] =
      recebimento.nf;

    row[20] =
      recebimento.dataRecebimento;

    row[22] =
      recebimento.documentoOrigem;

    row[23] =
      'OK';

  } else if (recebimentos.length > 1) {

    row[23] =
      'MÚLTIPLOS RECEBIMENTOS';

  } else {

    row[23] =
      'SEM RECEBIMENTO';
  }

  row[26] =
    chaveMpLote_(
      row[7],
      row[10]
    );

  row[27] =
    chaveMpLoteNf_(
      row[7],
      row[10],
      row[19]
    );

  row[28] =
    chaveProducao_(
      row[1],
      row[6],
      row[5]
    );

  range.setValues([row]);
}


function atualizarOrigemRastreabilidade() {
  const aba =
    obterAba(CONFIG.ABAS.RASTREABILIDADE);

  const lastRow =
    aba.getLastRow();

  if (
    lastRow <
    CONFIG.PRIMEIRA_LINHA_DADOS
  ) {
    return {
      sucesso: true,
      atualizadas: 0
    };
  }

  const valores =
    aba
      .getRange(
        CONFIG.PRIMEIRA_LINHA_DADOS,
        1,
        lastRow -
          CONFIG.PRIMEIRA_LINHA_DADOS +
          1,
        29
      )
      .getValues();

  let atualizadas = 0;

  valores.forEach(function(row, index) {

    const mp = texto_(row[7]);
    const lote = texto_(row[10]);
    const nf = texto_(row[19]);

    if (!mp) return;

    completarDadosRecebimentoNaRastreabilidade_(
      CONFIG.PRIMEIRA_LINHA_DADOS + index,
      mp,
      lote,
      nf
    );

    atualizadas++;
  });

  aplicarFormulasAutomaticas();

  return {
    sucesso: true,
    atualizadas: atualizadas
  };
}


/* ============================================================
   FÓRMULAS AUTOMÁTICAS
   ============================================================ */

function calcularSaldoAposLinha_(
  valores,
  indiceAtual,
  materiaPrima,
  lote
) {
  const mpBusca =
    chaveNormalizada_(materiaPrima);

  const loteBusca =
    chaveNormalizada_(lote);

  if (!mpBusca || !loteBusca) {
    return null;
  }

  let consumido = 0;

  for (
    let i = 0;
    i <= indiceAtual;
    i++
  ) {
    const row = valores[i];

    const mp =
      chaveNormalizada_(row[7]);

    const lot =
      chaveNormalizada_(row[10]);

    if (
      mp === mpBusca &&
      lot === loteBusca
    ) {
      const pesado =
        numero_(row[15]);

      if (pesado !== null) {
        consumido += pesado;
      }
    }
  }

  const estoque =
    calcularEstoquePorMpLote_(
      materiaPrima,
      lote
    );

  /*
   * O estoque retornado já desconta todo o consumo.
   * Para mostrar o saldo exatamente após a linha atual,
   * calculamos novamente usando o total recebido.
   */
  const saldo =
    estoque.quantidadeRecebida -
    consumido;

  return saldo;
}


function aplicarFormulasAutomaticas() {
  const aba =
    obterAba(CONFIG.ABAS.RASTREABILIDADE);

  const lastRow =
    aba.getLastRow();

  if (
    lastRow <
    CONFIG.PRIMEIRA_LINHA_DADOS
  ) {
    return {
      sucesso: true,
      linhas: 0
    };
  }

  const quantidade =
    lastRow -
    CONFIG.PRIMEIRA_LINHA_DADOS +
    1;

  const valores =
    aba
      .getRange(
        CONFIG.PRIMEIRA_LINHA_DADOS,
        1,
        quantidade,
        29
      )
      .getValues();

  const formulasQRS = [];
  const dadosXAC = [];

  valores.forEach(function(row, index) {

    const linha =
      CONFIG.PRIMEIRA_LINHA_DADOS +
      index;

    formulasQRS.push([
      '=IF(OR(O' + linha +
      '="",P' + linha +
      '=""),"",P' + linha +
      '-O' + linha + ')',

      '=IF(OR(O' + linha +
      '="",P' + linha +
      '="",O' + linha +
      '=0),"",Q' + linha +
      '/O' + linha + ')',

      '=IF(Q' + linha +
      '="","",IF(ABS(R' +
      linha +
      ')<=0.01,"OK","VERIFICAR"))'
    ]);

    const mp =
      texto_(row[7]);

    const lote =
      texto_(row[10]);

    const nf =
      texto_(row[19]);

    const recebimentos =
      mp
        ? buscarRecebimentos_(
            mp,
            lote,
            ''
          )
        : [];

    let controle =
      texto_(row[23]);

    if (mp) {
      if (recebimentos.length === 0) {
        controle = 'SEM RECEBIMENTO';
      } else if (
        recebimentos.length === 1
      ) {
        controle = 'OK';
      } else {
        controle =
          'MÚLTIPLOS RECEBIMENTOS';
      }
    }

    const saldo =
      calcularSaldoAposLinha_(
        valores,
        index,
        mp,
        lote
      );

    const chaveMP =
      chaveMpLote_(
        mp,
        lote
      );

    const chaveMPNF =
      chaveMpLoteNf_(
        mp,
        lote,
        nf
      );

    const chaveProd =
      chaveProducao_(
        row[1],
        row[6],
        row[5]
      );

    dadosXAC.push([
      controle,
      recebimentos.length,
      saldo === null ? '' : saldo,
      chaveMP,
      chaveMPNF,
      chaveProd
    ]);
  });

  // Q:R:S
  aba
    .getRange(
      CONFIG.PRIMEIRA_LINHA_DADOS,
      17,
      quantidade,
      3
    )
    .setFormulas(formulasQRS);

  // X:Y:Z:AA:AB:AC
  aba
    .getRange(
      CONFIG.PRIMEIRA_LINHA_DADOS,
      24,
      quantidade,
      6
    )
    .setValues(dadosXAC);

  // Formatação percentual
  aba
    .getRange(
      CONFIG.PRIMEIRA_LINHA_DADOS,
      18,
      quantidade,
      1
    )
    .setNumberFormat('0.00%');

  return {
    sucesso: true,
    linhas: quantidade
  };
}


/* ============================================================
   TESTE ESPECÍFICO DA DIFERENÇA DE PESAGEM
   ============================================================ */

function testeDiferencaPesagem() {
  const aba =
    obterAba(CONFIG.ABAS.RASTREABILIDADE);

  const linhaLivre =
    encontrarPrimeiraLinhaLivre(aba);

  const agora =
    new Date();

  const idProducao =
    gerarID('TEST-PROD');

  const idBatida =
    gerarID('TEST-BAT');

  const linha =
    new Array(29).fill('');

  linha[0] =
    gerarID('TEST-RAST');

  linha[1] =
    idProducao;

  linha[2] =
    idBatida;

  linha[3] =
    agora;

  linha[4] =
    'TESTE DIFERENCA';

  linha[5] =
    'LOTE-TESTE';

  linha[6] =
    'B01';

  linha[7] =
    'MP_TESTE';

  linha[8] =
    'FORNECEDOR TESTE';

  linha[9] =
    'FABRICANTE TESTE';

  linha[10] =
    'LOTE_TESTE';

  linha[11] =
    '';

  linha[12] =
    100;

  linha[13] =
    1;

  // Teórico
  linha[14] =
    100;

  // Realmente pesado
  linha[15] =
    95;

  // Q/R/S calculados
  linha[16] =
    '';

  linha[17] =
    '';

  linha[18] =
    '';

  linha[19] =
    'NF-TESTE';

  linha[20] =
    agora;

  linha[21] =
    'Teste automático da diferença de pesagem.';

  linha[22] =
    'TESTE';

  linha[23] =
    'TESTE';

  linha[24] =
    0;

  linha[25] =
    '';

  linha[26] =
    chaveMpLote_(
      'MP_TESTE',
      'LOTE_TESTE'
    );

  linha[27] =
    chaveMpLoteNf_(
      'MP_TESTE',
      'LOTE_TESTE',
      'NF-TESTE'
    );

  linha[28] =
    chaveProducao_(
      idProducao,
      'B01',
      'LOTE-TESTE'
    );

  aba
    .getRange(
      linhaLivre,
      1,
      1,
      29
    )
    .setValues([linha]);

  aplicarFormulasAutomaticas();

  const resultado =
    aba
      .getRange(
        linhaLivre,
        15,
        1,
        5
      )
      .getValues()[0];

  return {
    sucesso: true,
    linha: linhaLivre,
    teorico: resultado[0],
    pesado: resultado[1],
    diferenca: resultado[2],
    percentual:
      resultado[3],
    status:
      resultado[4],
    esperado: {
      diferenca: -5,
      percentual: -0.05,
      status: 'VERIFICAR'
    },
    mensagem:
      'Teste inserido na RASTREABILIDADE_MP. ' +
      'Teórico = 100 kg, pesado = 95 kg, ' +
      'diferença = -5 kg e status = VERIFICAR.'
  };
}


/* ============================================================
   RECONSTRUÇÃO DAS CHAVES
   ============================================================ */

function reconstruirChavesRastreabilidade() {
  const aba =
    obterAba(CONFIG.ABAS.RASTREABILIDADE);

  const lastRow =
    aba.getLastRow();

  if (
    lastRow <
    CONFIG.PRIMEIRA_LINHA_DADOS
  ) {
    return {
      sucesso: true,
      linhas: 0
    };
  }

  const quantidade =
    lastRow -
    CONFIG.PRIMEIRA_LINHA_DADOS +
    1;

  const valores =
    aba
      .getRange(
        CONFIG.PRIMEIRA_LINHA_DADOS,
        1,
        quantidade,
        29
      )
      .getValues();

  const saida = [];

  valores.forEach(function(row) {

    saida.push([
      chaveMpLote_(
        row[7],
        row[10]
      ),

      chaveMpLoteNf_(
        row[7],
        row[10],
        row[19]
      ),

      chaveProducao_(
        row[1],
        row[6],
        row[5]
      )
    ]);
  });

  // AA:AB:AC
  aba
    .getRange(
      CONFIG.PRIMEIRA_LINHA_DADOS,
      27,
      quantidade,
      3
    )
    .setValues(saida);

  return {
    sucesso: true,
    linhas: quantidade
  };
}


/* ============================================================
   ESTOQUE_LOTES
   ============================================================ */

function atualizarEstoqueLotes() {
  const abaEstoque =
    obterAba(CONFIG.ABAS.ESTOQUE);

  const recebimentos =
    lerRegistros_(
      CONFIG.ABAS.RECEBIMENTOS
    );

  const mapa = {};

  recebimentos.registros.forEach(function(r) {

    const d =
      obterDadosRecebimento_(r);

    if (!d.materiaPrima || !d.lote) {
      return;
    }

    const chave =
      chaveMpLote_(
        d.materiaPrima,
        d.lote
      );

    if (!mapa[chave]) {

      mapa[chave] = {
        materiaPrima:
          d.materiaPrima,

        lote:
          d.lote,

        quantidadeRecebida:
          0,

        unidade:
          d.unidade || 'KG',

        primeiroRecebimento:
          null,

        fornecedores: {},

        nfs: {}
      };
    }

    mapa[chave]
      .quantidadeRecebida +=
      d.quantidade || 0;

    if (d.fornecedor) {
      mapa[chave]
        .fornecedores[d.fornecedor] =
        true;
    }

    if (d.nf) {
      mapa[chave]
        .nfs[d.nf] =
        true;
    }

    const data =
      dataValida_(
        d.dataRecebimento
      );

    if (
      data &&
      (
        !mapa[chave]
          .primeiroRecebimento ||
        data <
        mapa[chave]
          .primeiroRecebimento
      )
    ) {
      mapa[chave]
        .primeiroRecebimento =
        data;
    }
  });

  const linhas =
    Object.keys(mapa).map(function(chave) {

      const item =
        mapa[chave];

      const estoque =
        calcularEstoquePorMpLote_(
          item.materiaPrima,
          item.lote
        );

      return [
        item.materiaPrima,
        item.lote,
        item.quantidadeRecebida,
        estoque.quantidadeConsumida,
        estoque.saldo,
        item.unidade,
        item.primeiroRecebimento || '',
        Object.keys(
          item.fornecedores
        ).join(', '),
        Object.keys(
          item.nfs
        ).join(', '),
        estoque.status,
        chave
      ];
    });

  const lastRow =
    abaEstoque.getLastRow();

  if (
    lastRow >=
    CONFIG.PRIMEIRA_LINHA_DADOS
  ) {
    abaEstoque
      .getRange(
        CONFIG.PRIMEIRA_LINHA_DADOS,
        1,
        lastRow -
          CONFIG.PRIMEIRA_LINHA_DADOS +
          1,
        11
      )
      .clearContent();
  }

  if (linhas.length) {

    abaEstoque
      .getRange(
        CONFIG.PRIMEIRA_LINHA_DADOS,
        1,
        linhas.length,
        11
      )
      .setValues(linhas);

    abaEstoque
      .getRange(
        CONFIG.PRIMEIRA_LINHA_DADOS,
        7,
        linhas.length,
        1
      )
      .setNumberFormat('dd/MM/yyyy');
  }

  return {
    sucesso: true,
    lotes: linhas.length
  };
}


/* ============================================================
   DIAGNÓSTICOS
   ============================================================ */

function testeAcessoPlanilha() {
  const ss =
    obterPlanilha();

  return {
    sucesso: true,
    nome:
      ss.getName(),
    id:
      ss.getId(),
    mensagem:
      'Acesso à planilha confirmado.'
  };
}


function listarAbas() {
  return obterPlanilha()
    .getSheets()
    .map(function(s) {
      return s.getName();
    });
}


function testeEstruturaCompleta() {
  const ss =
    obterPlanilha();

  const obrigatorias = [
    CONFIG.ABAS.RECEBIMENTOS,
    CONFIG.ABAS.PRODUCOES,
    CONFIG.ABAS.FORMULACOES,
    CONFIG.ABAS.RASTREABILIDADE,
    CONFIG.ABAS.ESTOQUE
  ];

  const existentes =
    ss
      .getSheets()
      .map(function(s) {
        return s.getName();
      });

  const faltantes =
    obrigatorias.filter(function(nome) {
      return existentes.indexOf(nome) < 0;
    });

  return {
    sucesso:
      faltantes.length === 0,

    abasEncontradas:
      existentes,

    abasObrigatorias:
      obrigatorias,

    faltantes:
      faltantes,

    mensagem:
      faltantes.length
        ? 'Existem abas obrigatórias faltando.'
        : 'Estrutura básica das abas OK.'
  };
}


function diagnosticoEstrutura() {
  return testeEstruturaCompleta();
}


function testeFormulasAutomaticas() {
  const resultado =
    aplicarFormulasAutomaticas();

  return {
    sucesso:
      resultado.sucesso,

    linhas:
      resultado.linhas,

    mensagem:
      'Cálculos Q/R/S e chaves Z/AA/AB/AC atualizados.'
  };
}


function testeGravacaoRecebimento() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.RECEBIMENTOS
    );

  const temMP =
    indicePorAliases_(
      dados.headers,
      [
        'Matéria-prima',
        'Materia prima',
        'MP'
      ]
    ) >= 0;

  const temLote =
    indicePorAliases_(
      dados.headers,
      [
        'Lote',
        'Lote da MP',
        'Lote MP'
      ]
    ) >= 0;

  return {
    sucesso:
      temMP && temLote,

    mensagem:
      temMP && temLote
        ? 'Estrutura para gravação de recebimentos OK.'
        : 'Verifique as colunas de matéria-prima e lote da aba RECEBIMENTOS.',

    registrosAtuais:
      dados.registros.length
  };
}


function testeMultiplosRecebimentos() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.RECEBIMENTOS
    );

  const mapa = {};

  dados.registros.forEach(function(r) {

    const d =
      obterDadosRecebimento_(r);

    if (!d.materiaPrima || !d.lote) {
      return;
    }

    const chave =
      chaveMpLote_(
        d.materiaPrima,
        d.lote
      );

    if (!mapa[chave]) {
      mapa[chave] = 0;
    }

    mapa[chave]++;
  });

  const multiplos =
    Object.keys(mapa)
      .filter(function(chave) {
        return mapa[chave] > 1;
      });

  return {
    sucesso: true,

    encontrados:
      multiplos.length,

    chaves:
      multiplos,

    mensagem:
      multiplos.length
        ? 'Foram encontrados lotes com múltiplos recebimentos. O sistema não escolhe o primeiro automaticamente.'
        : 'Nenhum lote com múltiplos recebimentos foi encontrado atualmente.'
  };
}


function testeNFRecebimento() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.RECEBIMENTOS
    );

  let exemplo = null;

  for (
    let i = 0;
    i < dados.registros.length;
    i++
  ) {
    const d =
      obterDadosRecebimento_(
        dados.registros[i]
      );

    if (
      d.materiaPrima &&
      d.lote &&
      d.nf
    ) {
      exemplo = d;
      break;
    }
  }

  if (!exemplo) {
    return {
      sucesso: true,
      encontrouExemplo: false,
      mensagem:
        'Não há recebimento com NF disponível para testar agora.'
    };
  }

  const encontrado =
    buscarRecebimentos_(
      exemplo.materiaPrima,
      exemplo.lote,
      exemplo.nf
    );

  return {
    sucesso:
      encontrado.length === 1,

    encontrouExemplo: true,

    materiaPrima:
      exemplo.materiaPrima,

    lote:
      exemplo.lote,

    nf:
      exemplo.nf,

    quantidadeEncontrada:
      encontrado.length,

    mensagem:
      encontrado.length === 1
        ? 'Busca exata por MP + lote + NF funcionando.'
        : 'A busca exata por NF retornou quantidade diferente de 1.'
  };
}


function testeEstoque() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.RECEBIMENTOS
    );

  for (
    let i = 0;
    i < dados.registros.length;
    i++
  ) {

    const d =
      obterDadosRecebimento_(
        dados.registros[i]
      );

    if (
      d.materiaPrima &&
      d.lote
    ) {

      const estoque =
        consultarSaldoLote(
          d.materiaPrima,
          d.lote
        );

      return {
        sucesso: true,
        exemplo: estoque,
        mensagem:
          'Consulta de estoque funcionando.'
      };
    }
  }

  return {
    sucesso: true,
    encontrouExemplo: false,
    mensagem:
      'Não há recebimentos cadastrados para testar estoque.'
  };
}


function testeFormulacaoAtiva() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.FORMULACOES
    );

  if (!dados.registros.length) {
    return {
      sucesso: false,
      encontrou: false,
      mensagem:
        'A aba FORMULAÇÕES não possui dados.'
    };
  }

  let produto = '';

  for (
    let i = 0;
    i < dados.registros.length;
    i++
  ) {
    produto =
      texto_(
        valorCampo_(
          dados.registros[i],
          [
            'Produto',
            'Produto acabado',
            'Produto final'
          ]
        )
      );

    if (produto) break;
  }

  if (!produto) {
    return {
      sucesso: false,
      encontrou: false,
      mensagem:
        'Não encontrei produto na aba FORMULAÇÕES.'
    };
  }

  const resultado =
    localizarFormulacaoAtiva(
      produto,
      new Date()
    );

  return {
    sucesso:
      resultado.encontrada,

    encontrou:
      resultado.encontrada,

    produto:
      produto,

    versao:
      resultado.versao,

    linhas:
      resultado.linhas
        ? resultado.linhas.length
        : 0,

    mensagem:
      resultado.mensagem
  };
}


function testeCompleto() {
  return {
    estrutura:
      testeEstruturaCompleta(),

    formulas:
      testeFormulasAutomaticas(),

    recebimento:
      testeGravacaoRecebimento(),

    multiplos:
      testeMultiplosRecebimentos(),

    diferenca:
      testeDiferencaPesagem(),

    nf:
      testeNFRecebimento(),

    estoque:
      testeEstoque(),

    formulacao:
      testeFormulacaoAtiva()
  };
}


/* ============================================================
   GEMINI
   ============================================================ */

function obterChaveGemini_() {
  const chave =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GEMINI_API_KEY'
      );

  if (!chave) {
    throw new Error(
      'GEMINI_API_KEY não configurada nas propriedades do projeto.'
    );
  }

  return chave;
}


function testarChaveGemini() {
  const chave =
    obterChaveGemini_();

  return {
    sucesso:
      !!chave,
    mensagem:
      'GEMINI_API_KEY encontrada nas propriedades do projeto.'
  };
}


function chamarGemini_(payload) {
  const chave =
    obterChaveGemini_();

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.GEMINI_MODEL +
    ':generateContent?key=' +
    encodeURIComponent(chave);

  const resposta =
    UrlFetchApp.fetch(
      url,
      {
        method: 'post',
        contentType: 'application/json',
        payload:
          JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

  const status =
    resposta.getResponseCode();

  const textoResposta =
    resposta.getContentText();

  if (
    status < 200 ||
    status >= 300
  ) {
    throw new Error(
      'Erro Gemini HTTP ' +
      status +
      ': ' +
      textoResposta.substring(0, 2000)
    );
  }

  return JSON.parse(
    textoResposta
  );
}


function extrairTextoGemini_(resposta) {
  try {
    const candidatos =
      resposta.candidates || [];

    if (!candidatos.length) {
      return '';
    }

    const parts =
      candidatos[0]
        .content
        .parts || [];

    return parts
      .map(function(p) {
        return p.text || '';
      })
      .join('');
  } catch (e) {
    return '';
  }
}


function limparJsonGemini_(texto) {
  let s =
    String(texto || '')
      .trim();

  s =
    s.replace(
      /^```json\s*/i,
      ''
    );

  s =
    s.replace(
      /^```\s*/i,
      ''
    );

  s =
    s.replace(
      /\s*```$/i,
      ''
    );

  const inicio =
    s.indexOf('{');

  const fim =
    s.lastIndexOf('}');

  if (
    inicio >= 0 &&
    fim > inicio
  ) {
    s =
      s.substring(
        inicio,
        fim + 1
      );
  }

  return s;
}


function testarGeminiAPI() {
  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          {
            text:
              'Responda somente com JSON: {"ok":true}'
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0,
      responseMimeType:
        'application/json'
    }
  };

  const resposta =
    chamarGemini_(payload);

  const texto =
    extrairTextoGemini_(resposta);

  return {
    sucesso: true,
    resposta: texto
  };
}


/* ============================================================
   OCR / ANÁLISE DE DOCUMENTOS COM GEMINI
   ============================================================ */

function analisarDocumentoGemini(arquivo) {
  if (!arquivo) {
    throw new Error(
      'Arquivo não recebido.'
    );
  }

  if (!arquivo.base64) {
    throw new Error(
      'O arquivo não contém base64.'
    );
  }

  const mimeType =
    arquivo.mimeType ||
    'application/octet-stream';

  const permitidos = [
    'application/pdf',
    'image/jpeg',
    'image/png',
    'image/webp'
  ];

  if (
    permitidos.indexOf(mimeType) < 0
  ) {
    throw new Error(
      'Tipo de arquivo não suportado: ' +
      mimeType
    );
  }

  /*
   * Evita enviar arquivos exageradamente grandes
   * para o endpoint.
   */
  if (
    arquivo.base64.length >
    15000000
  ) {
    throw new Error(
      'Arquivo muito grande para análise automática.'
    );
  }

  const prompt =
    [
      'Você é um assistente de conferência documental da MOZAGRO Nutrição Animal Ltda.',
      '',
      'Analise o documento anexado.',
      '',
      'Extraia informações de produção quando existirem.',
      '',
      'Retorne SOMENTE JSON válido, sem markdown.',
      '',
      'Formato obrigatório:',
      '{',
      '  "tipoDocumento": "",',
      '  "mes": "",',
      '  "responsavel": "",',
      '  "observacoes": "",',
      '  "documentoOrigem": "",',
      '  "linhas": [',
      '    {',
      '      "dataProducao": "",',
      '      "quantidadeProduzida": 0,',
      '      "produto": "",',
      '      "loteProduto": "",',
      '      "confianca": 0,',
      '      "revisaoNecessaria": false',
      '    }',
      '  ]',
      '}',
      '',
      'Use datas no formato dd/MM/yyyy quando possível.',
      'Não invente dados.',
      'Se uma informação não puder ser identificada, deixe vazia.',
      'Quantidades devem ser numéricas.',
      'Confianca deve ser de 0 a 1.'
    ].join('\n');

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          {
            text: prompt
          },
          {
            inline_data: {
              mime_type: mimeType,
              data: arquivo.base64
            }
          }
        ]
      }
    ],

    generationConfig: {
      temperature: 0.1,
      responseMimeType:
        'application/json'
    }
  };

  const resposta =
    chamarGemini_(payload);

  const texto =
    extrairTextoGemini_(resposta);

  if (!texto) {
    throw new Error(
      'O Gemini não retornou conteúdo.'
    );
  }

  const jsonTexto =
    limparJsonGemini_(texto);

  let resultado;

  try {
    resultado =
      JSON.parse(jsonTexto);
  } catch (e) {
    throw new Error(
      'Não foi possível interpretar o JSON retornado pelo Gemini: ' +
      jsonTexto.substring(0, 3000)
    );
  }

  if (!resultado.linhas) {
    resultado.linhas = [];
  }

  return resultado;
}


/* ============================================================
   SALVAR PRODUÇÃO CONFIRMADA PELA IA
   ============================================================ */

function salvarProducaoConfirmadaIA(resultado) {
  if (!resultado) {
    throw new Error(
      'Resultado da IA não recebido.'
    );
  }

  const linhas =
    resultado.linhas || [];

  if (!linhas.length) {
    return {
      sucesso: false,
      mensagem:
        'Nenhuma linha de produção encontrada.'
    };
  }

  const resultados = [];

  linhas.forEach(function(item, index) {

    const dados = {
      produto:
        texto_(item.produto),

      loteProduto:
        texto_(item.loteProduto),

      dataProducao:
        dataValida_(
          item.dataProducao
        ),

      quantidadeProduzida:
        numero_(
          item.quantidadeProduzida
        ),

      numeroBatidas:
        1,

      batida:
        'IA-' +
        String(index + 1)
          .padStart(2, '0'),

      responsavel:
        texto_(
          resultado.responsavel
        ),

      documentoOrigem:
        texto_(
          resultado.documentoOrigem
        ),

      observacoes:
        texto_(
          resultado.observacoes
        )
    };

    if (
      !dados.produto ||
      !dados.loteProduto
    ) {
      resultados.push({
        sucesso: false,
        mensagem:
          'Linha ' +
          (index + 1) +
          ' sem produto ou lote.'
      });

      return;
    }

    try {

      const salvo =
        salvarProducao(
          dados
        );

      resultados.push(
        salvo
      );

    } catch (e) {

      resultados.push({
        sucesso: false,
        mensagem:
          e.message
      });
    }
  });

  return {
    sucesso:
      resultados.some(function(r) {
        return r.sucesso;
      }),

    resultados:
      resultados,

    quantidade:
      resultados.length
  };
}


/* ============================================================
   MANUTENÇÃO
   ============================================================ */

function atualizarRastreabilidade() {
  const origem =
    atualizarOrigemRastreabilidade();

  const formulas =
    aplicarFormulasAutomaticas();

  return {
    sucesso: true,
    origem: origem,
    formulas: formulas
  };
}


function manutencaoGeral() {
  const origem =
    atualizarOrigemRastreabilidade();

  const chaves =
    reconstruirChavesRastreabilidade();

  const formulas =
    aplicarFormulasAutomaticas();

  const estoque =
    atualizarEstoqueLotes();

  return {
    sucesso: true,
    origem: origem,
    chaves: chaves,
    formulas: formulas,
    estoque: estoque
  };
}


/* ============================================================
   LISTAS PARA A INTERFACE
   ============================================================ */

function obterListaMateriasPrimas() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.RECEBIMENTOS
    );

  const mapa = {};

  dados.registros.forEach(function(r) {

    const mp =
      texto_(
        valorCampo_(
          r,
          [
            'Matéria-prima',
            'Materia prima',
            'MP',
            'Ingrediente'
          ]
        )
      );

    if (mp) {
      mapa[
        normalizarTexto_(mp)
      ] = mp;
    }
  });

  return Object.keys(mapa)
    .sort()
    .map(function(chave) {
      return mapa[chave];
    });
}


function obterListaProdutos() {
  const dados =
    lerRegistros_(
      CONFIG.ABAS.PRODUCOES
    );

  const mapa = {};

  dados.registros.forEach(function(r) {

    const produto =
      texto_(
        valorCampo_(
          r,
          [
            'Produto',
            'Produto acabado',
            'Produto final'
          ]
        )
      );

    if (produto) {
      mapa[
        normalizarTexto_(produto)
      ] = produto;
    }
  });

  // Também pega produtos da formulação.
  const formulas =
    lerRegistros_(
      CONFIG.ABAS.FORMULACOES
    );

  formulas.registros.forEach(function(r) {

    const produto =
      texto_(
        valorCampo_(
          r,
          [
            'Produto',
            'Produto acabado',
            'Produto final'
          ]
        )
      );

    if (produto) {
      mapa[
        normalizarTexto_(produto)
      ] = produto;
    }
  });

  return Object.keys(mapa)
    .sort()
    .map(function(chave) {
      return mapa[chave];
    });
}


/* ============================================================
   RESUMO DO SISTEMA
   ============================================================ */

function obterResumoSistema() {
  const resumo = {
    planilha:
      obterPlanilha().getName(),

    abas: {},

    recebimentos: 0,

    producoes: 0,

    rastreabilidades: 0,

    lotesEstoque: 0
  };

  [
    'RECEBIMENTOS',
    'PRODUCOES',
    'FORMULACOES',
    'RASTREABILIDADE',
    'ESTOQUE'
  ].forEach(function(chave) {

    const nome =
      CONFIG.ABAS[chave];

    try {

      const dados =
        lerRegistros_(nome);

      resumo.abas[nome] =
        dados.registros.length;

    } catch (e) {

      resumo.abas[nome] =
        'ERRO';
    }
  });

  resumo.recebimentos =
    resumo.abas[
      CONFIG.ABAS.RECEBIMENTOS
    ] || 0;

  resumo.producoes =
    resumo.abas[
      CONFIG.ABAS.PRODUCOES
    ] || 0;

  resumo.rastreabilidades =
    resumo.abas[
      CONFIG.ABAS.RASTREABILIDADE
    ] || 0;

  resumo.lotesEstoque =
    resumo.abas[
      CONFIG.ABAS.ESTOQUE
    ] || 0;

  return resumo;
}


function verificarSistema() {
  const estrutura =
    testeEstruturaCompleta();

  let gemini;

  try {
    gemini =
      testarChaveGemini();
  } catch (e) {
    gemini = {
      sucesso: false,
      mensagem: e.message
    };
  }

  return {
    sucesso:
      estrutura.sucesso &&
      gemini.sucesso,

    estrutura:
      estrutura,

    gemini:
      gemini
  };
}


function testeRapido() {
  return {
    acesso:
      testeAcessoPlanilha(),

    estrutura:
      testeEstruturaCompleta(),

    resumo:
      obterResumoSistema()
  };
}


/* ============================================================
   FUNÇÕES DE COMPATIBILIDADE
   ============================================================ */

function obterAbaPorNome(nome) {
  return obterAba(nome);
}


function calcularEstoque(materiaPrima, lote) {
  return calcularEstoquePorMpLote_(
    materiaPrima,
    lote
  );
}


/* ============================================================
   FIM DO CODIGO.GS
   ============================================================ */