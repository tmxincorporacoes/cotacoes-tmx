/**
 * Casamento de itens da cotação com os itens da OC (seção 5 do guia).
 *
 * Três camadas, nesta ordem:
 *   1. Dicionário exato (fornecedor + código, ou fornecedor + descrição normalizada) → 'confere'.
 *   2. Similaridade textual → 'confere' só se a cotação cobre todas as palavras da OC e não há
 *      conflito de marca, material ou bitola; caso contrário 'confirmar' (com notas do motivo).
 *   3. IA (claude.ai, passo manual na fase gratuita): montarPromptIA() / aplicarRespostaIA().
 *      Tudo que vem da IA entra como 'confirmar'.
 *
 * Regra principal: nunca 'confere' quando marca, material ou bitola divergem da OC.
 * Funções puras, sem acesso a navegador ou planilha.
 */
(function (raiz) {
  'use strict';

  // ------------------------------------------------------------ vocabulário

  const ABREV = {
    valv: 'valvula', sanf: 'sanfonado', bco: 'branco', lav: 'lavatorio', amer: 'americana', americano: 'americana',
    adap: 'adaptador', cx: 'caixa', sanit: 'sanitario', univ: 'universal', crom: 'cromado', cr: 'cromado',
    met: 'metal', lad: 'ladrao', s: 'sem', acab: 'acabamento', reg: 'registro', ass: 'assento', tee: 'te', borrac: 'borracha', esg: 'esgoto'
  };
  const FRASES = [[/pia de banheiro/g, 'lavatorio']];
  // Palavra da OC que o fornecedor não escreve, mas que a cotação já diz por outro meio
  // (a JK chama de "ANEL BORRACHA" o que a OC chama de "anel de vedação").
  const IMPLICITOS = { vedacao: ['borracha'] };
  const PARADAS = new Set(['de', 'da', 'do', 'das', 'dos', 'para', 'a', 'o', 'e', 'em', 'com', 'c', 'un', 'und', 'unid', 'unidade', 'unidades', 'pc', 'pcs', 'peca', 'pecas']);

  // marca → forma canônica; inclui grafias truncadas que aparecem nos PDFs ("GHELPLU.", "BLUKT")
  const MARCAS = {
    furkin: 'furkin', amanco: 'amanco', ghelplus: 'ghelplus', ghelplu: 'ghelplus', blukit: 'blukit', blukt: 'blukit',
    krona: 'krona', hervy: 'hervy', tigre: 'tigre', deca: 'deca', docol: 'docol', astra: 'astra', fortlev: 'fortlev', tramontina: 'tramontina'
  };
  // material → grupo (cromado e metal são o mesmo grupo)
  const MATERIAIS = {
    pvc: 'pvc', plastico: 'pvc', abs: 'pvc', metal: 'metal', cromado: 'metal', latao: 'metal', aco: 'metal',
    inox: 'inox', cera: 'cera', massa: 'massa', borracha: 'borracha', ceramica: 'ceramica', porcelana: 'ceramica'
  };
  // equivalência comercial PVC soldável: mm ↔ polegada
  const MM_POL = { 20: '1/2', 25: '3/4', 32: '1', 40: '1 1/4', 50: '1 1/2', 60: '2', 75: '2 1/2', 85: '3', 110: '4' };
  const POL_MM = Object.keys(MM_POL).reduce((o, k) => { o[MM_POL[k]] = k + 'mm'; return o; }, {});

  // ------------------------------------------------------------ análise de descrição

  const semAcento = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '');

  /** Extrai medidas (bitolas) e as remove do texto. Devolve { texto, canon:Set, exibir:[] }. */
  function extrairMedidas(s) {
    const canon = new Set(), exibir = [];
    const add = (canonico, mostrar) => {
      canon.add(canonico);
      if (MM_POL[parseInt(canonico, 10)] && /mm$/.test(canonico)) canon.add(MM_POL[parseInt(canonico, 10)]);
      if (POL_MM[canonico]) canon.add(POL_MM[canonico]);
      exibir.push(mostrar);
    };
    s = s.replace(/(\d+)\s*mm\b/g, (m, n) => { add(n + 'mm', n + ' mm'); return ' '; });
    s = s.replace(/(\d+(?:\s+\d\/\d)?|\d\/\d)\s*"/g, (m, n) => { add(n.replace(/\s+/g, ' '), n + '"'); return ' '; });
    s = s.replace(/\b(\d+)\s+(\d)\/(\d)\b/g, (m, a, b, c) => { add(a + ' ' + b + '/' + c, a + ' ' + b + '/' + c + '"'); return ' '; });
    s = s.replace(/(^|[^\d])(\d)1\/([24])\b/g, (m, pre, a, c) => { add(a + ' 1/' + c, a + ' 1/' + c + '"'); return pre + ' '; });
    s = s.replace(/(^|[^\d\/])(\d)\/(\d)\b/g, (m, pre, b, c) => { add(b + '/' + c, b + '/' + c + '"'); return pre + ' '; });
    return { texto: s, canon, exibir };
  }

  /** Decompõe a descrição em conteúdo (palavras), marcas, materiais e medidas. */
  function analisar(desc) {
    let s = semAcento(String(desc || '').toLowerCase());
    s = s.replace(/cod\.?\s*ncm:?\s*\d+/g, ' ').replace(/cod\.?\s*barras:?[\d\s]*/g, ' ').replace(/marca:?/g, ' ');
    s = s.replace(/\bc\/\s*0*\d+\s*un\b/g, ' ').replace(/\b\d{7,}\b/g, ' ');
    // medida composta "100x100mm" / "45x100mm": separa os números (a unidade fica no último), igual em qualquer escrita;
    // se um dos lados tem 1 dígito (placa 4x2, disjuntor 1x16A) mantém junto, para não perder a diferença
    s = s.replace(/([a-z])(?=\d+(?:[.,]\d+)?\s*(?:x\s*\d+(?:[.,]\d+)?\s*)?mm\b)/g, '$1 ').replace(/([a-z])(?=\d{2,}\s*x\s*\d{2,})/g, '$1 ');
    s = s.replace(/×/g, 'x').replace(/(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)/g, (m, a, b) => (a.replace(/\D/g, '').length >= 2 && b.replace(/\D/g, '').length >= 2) ? (a === b ? a : a + ' ' + b) : a + 'x' + b);
    // peso e volume por embalagem (pasta 160 g × 300 g)
    const pesos = new Set(), pesosTxt = [];
    s = s.replace(/(\d+(?:[.,]\d+)?)\s*(kg|g|ml|l)\b/g, (m, n, u) => {
      const v = Number(n.replace(',', '.')) * (u === 'kg' || u === 'l' ? 1000 : 1);
      pesos.add(String(Math.round(v)) + (u === 'ml' || u === 'l' ? 'ml' : 'g')); pesosTxt.push(n + ' ' + u); return ' ';
    });
    const med = extrairMedidas(s);
    s = med.texto;
    FRASES.forEach(([re, por]) => { s = s.replace(re, ' ' + por + ' '); });
    const conteudo = new Set(), marcas = new Set(), materiais = new Set();
    s.split(/[^a-z0-9]+/).filter(Boolean).forEach(t => {
      t = ABREV[t] || t;
      if (PARADAS.has(t) || t.length < 2) return;
      if (MARCAS[t]) marcas.add(MARCAS[t]);
      else if (MATERIAIS[t]) materiais.add(MATERIAIS[t]);
      else conteudo.add(t);
    });
    return { conteudo, marcas, materiais, medidas: med.canon, medidasTxt: med.exibir, pesos, pesosTxt };
  }

  /** Descrição normalizada, usada como chave do dicionário. */
  function normalizar(desc) {
    const a = analisar(desc);
    return [].concat([...a.conteudo], [...a.marcas], [...a.materiais], [...a.medidas], [...a.pesos]).sort().join(' ');
  }

  const inter = (a, b) => [...a].filter(x => b.has(x));
  const disjuntos = (a, b) => a.size > 0 && b.size > 0 && inter(a, b).length === 0;

  /** Compara a descrição de um item da OC com a de uma linha cotada. */
  function comparar(descOC, descCot) {
    const o = analisar(descOC), c = analisar(descCot);
    const implicito = t => !c.conteudo.has(t) && (IMPLICITOS[t] || []).some(m => c.materiais.has(MATERIAIS[m] || m));
    const cobertos = [...o.conteudo].filter(t => c.conteudo.has(t) || implicito(t));
    const comuns = cobertos, nImplicitos = cobertos.filter(implicito).length;
    const faltando = [...o.conteudo].filter(t => !c.conteudo.has(t) && !implicito(t));
    const total = o.conteudo.size + c.conteudo.size + nImplicitos;
    const conflitos = [];
    if (disjuntos(o.medidas, c.medidas)) conflitos.push('Bitola: OC ' + o.medidasTxt.join(' / ') + ' × cotação ' + c.medidasTxt.join(' / '));
    if (disjuntos(o.pesos, c.pesos)) conflitos.push('Quantidade na embalagem: OC ' + o.pesosTxt.join(' / ') + ' × cotação ' + c.pesosTxt.join(' / '));
    if (disjuntos(o.materiais, c.materiais)) conflitos.push('Material: OC ' + [...o.materiais].join('/') + ' × cotação ' + [...c.materiais].join('/'));
    if (disjuntos(o.marcas, c.marcas)) conflitos.push('Marca: OC ' + [...o.marcas].join('/') + ' × cotação ' + [...c.marcas].join('/'));
    const marcaFalta = [...o.marcas].filter(m => !c.marcas.has(m));
    return {
      compartilhados: comuns.length,
      cobertura: o.conteudo.size ? comuns.length / o.conteudo.size : 0,
      score: total ? (2 * comuns.length) / total : 0,
      faltando, conflitos,
      marcaFalta: marcaFalta.length && !conflitos.some(x => /^Marca/.test(x)) ? marcaFalta : []
    };
  }

  function notasDe(r) {
    const n = r.conflitos.slice();
    if (r.marcaFalta.length) n.push('Marca pedida na OC (' + r.marcaFalta.join('/') + ') não aparece na cotação.');
    if (r.faltando.length) n.push('Palavras da OC que a cotação não traz: ' + r.faltando.join(', ') + '.');
    return n;
  }

  // ------------------------------------------------------------ fator de unidade fixo por produto

  // Produtos que sempre são vendidos em embalagem com mais de uma peça (decisão do usuário, 21/10/2026).
  // Só vale quando o fornecedor não informou o fator; o usuário ainda pode ajustar na célula do mapa.
  const FATORES_FIXOS = [{ palavras: ['parafuso', 'castelo'], fator: 2, motivo: 'parafuso castelo é sempre kit com 2' }];

  /** Fator de unidade fixo para a descrição, ou null se nenhuma regra se aplica. */
  function fatorFixo(desc) {
    const t = analisar(desc).conteudo;
    const r = FATORES_FIXOS.find(x => x.palavras.every(w => t.has(w)));
    return r ? { fator: r.fator, motivo: r.motivo } : null;
  }

  // ------------------------------------------------------------ casamento

  /**
   * @param {object} p
   *   p.ocItens     [{n, descricao, material_id?}]
   *   p.itens       itens extraídos da cotação [{codigo, descricao, descricao_curta?}]
   *   p.fornecedorId
   *   p.dicionario  linhas da aba Dicionario (opcional)
   *   p.forcarConfirmar  true quando a cotação foi lida por IA: só o dicionário pode dar 'confere'
   * @returns {{vinculos, sem_cotacao, fora_da_oc}}
   *   vinculos: [{oc_n, idx, status:'confere'|'confirmar', origem:'dicionario'|'similaridade', score, notas[]}]
   */
  function casar(p) {
    const dic = (p.dicionario || []).filter(d => String(d.fornecedor_id) === String(p.fornecedorId));
    const vinculos = [], ocLivres = new Set(p.ocItens.map(i => i.n)), linhasLivres = new Set(p.itens.map((_, k) => k));
    const textoLinha = it => it.descricao_curta || it.descricao;

    // camada 0: planilha padrão da TMX. O código do Obra Prima (ou o nº do item na OC) liga a linha ao item
    // com certeza; se o fornecedor trocou o produto (marca, material, bitola), pede confirmação.
    p.itens.forEach((it, k) => {
      if (!linhasLivres.has(k)) return;
      let oc = it.codigo_tmx ? p.ocItens.find(i => i.material_id && String(i.material_id) === String(it.codigo_tmx) && ocLivres.has(i.n)) : null;
      if (!oc && it.n_oc !== undefined && it.n_oc !== null && it.n_oc !== '') oc = p.ocItens.find(i => i.n === Number(it.n_oc) && ocLivres.has(i.n));
      if (!oc) return;
      const informado = it.descricao_pedida && it.descricao !== it.descricao_pedida;
      let notas = [];
      if (informado) {
        const r = comparar(oc.descricao, it.descricao);
        notas = r.conflitos.slice();
        if (r.marcaFalta.length) notas.push('Marca pedida na OC (' + r.marcaFalta.join('/') + ') não aparece no produto cotado.');
      }
      vinculos.push({ oc_n: oc.n, idx: k, status: notas.length ? 'confirmar' : 'confere', origem: 'oc_padrao', score: 1, notas });
      ocLivres.delete(oc.n); linhasLivres.delete(k);
    });

    // camada 1: dicionário exato
    p.itens.forEach((it, k) => {
      const chave = normalizar(textoLinha(it));
      const d = dic.find(x => x.codigo_fornecedor && String(x.codigo_fornecedor) === String(it.codigo)) ||
                dic.find(x => !x.codigo_fornecedor && x.descricao_normalizada && String(x.descricao_normalizada) === chave);
      if (!d) return;
      const oc = p.ocItens.find(i => i.material_id && String(i.material_id) === String(d.material_id) && ocLivres.has(i.n));
      if (!oc) return;
      vinculos.push({ oc_n: oc.n, idx: k, status: 'confere', origem: 'dicionario', score: 1, notas: [] });
      ocLivres.delete(oc.n); linhasLivres.delete(k);
    });

    // camada 2: similaridade textual, melhores pares primeiro, um para um
    const pares = [];
    p.ocItens.forEach(oc => {
      if (!ocLivres.has(oc.n)) return;
      p.itens.forEach((it, k) => {
        if (!linhasLivres.has(k)) return;
        const r = comparar(oc.descricao, textoLinha(it));
        const aceita = (r.compartilhados >= 2 && r.score >= 0.3) || (r.cobertura === 1 && r.compartilhados >= 1);
        if (aceita) pares.push({ oc, k, r });
      });
    });
    pares.sort((a, b) => b.r.score - a.r.score || b.r.compartilhados - a.r.compartilhados || a.oc.n - b.oc.n);
    pares.forEach(({ oc, k, r }) => {
      if (!ocLivres.has(oc.n) || !linhasLivres.has(k)) return;
      const confere = !p.forcarConfirmar && r.cobertura === 1 && !r.conflitos.length && !r.marcaFalta.length;
      vinculos.push({ oc_n: oc.n, idx: k, status: confere ? 'confere' : 'confirmar', origem: 'similaridade', score: Math.round(r.score * 100) / 100, notas: notasDe(r) });
      ocLivres.delete(oc.n); linhasLivres.delete(k);
    });

    vinculos.sort((a, b) => a.oc_n - b.oc_n);
    return { vinculos, sem_cotacao: [...ocLivres].sort((a, b) => a - b), fora_da_oc: [...linhasLivres].sort((a, b) => a - b) };
  }

  // ------------------------------------------------------------ camada 3 (IA, manual)

  /** Texto para colar no claude.ai com o que sobrou. A resposta volta como JSON. */
  function montarPromptIA(ocItens, itens, restoOC, restoLinhas) {
    const oc = restoOC.map(n => { const i = ocItens.find(x => x.n === n); return { oc_n: n, descricao: i.descricao }; });
    const linhas = restoLinhas.map(k => ({ linha: k, codigo: itens[k].codigo, descricao: itens[k].descricao_curta || itens[k].descricao }));
    return [
      'Você ajuda a casar itens de uma cotação de fornecedor com os itens de uma ordem de compra (OC) de materiais de construção.',
      'Para cada linha da cotação, diga qual item da OC ela corresponde, ou null se nenhum corresponde.',
      'Regras: não force o casamento; se marca, material ou bitola divergem da OC, ainda assim indique o item mais próximo mas com confianca "baixa" e explique a divergência.',
      'Responda SOMENTE com um JSON, sem texto fora dele, no formato:',
      '[{"linha": 0, "oc_n": 3, "confianca": "alta|media|baixa", "justificativa": "frase curta"}]',
      '',
      'ITENS DA OC:', JSON.stringify(oc, null, 1), '',
      'LINHAS DA COTAÇÃO:', JSON.stringify(linhas, null, 1)
    ].join('\n');
  }

  /** Lê a resposta colada. Erro claro se inválida; não grava nada. Tudo entra como 'confirmar'. */
  function aplicarRespostaIA(resposta, ocLivres, linhasLivres) {
    let t = String(resposta || '').trim();
    const cerca = /```(?:json)?\s*([\s\S]*?)```/.exec(t);
    if (cerca) t = cerca[1].trim();
    let dados;
    try { dados = JSON.parse(t); } catch (e) { throw new Error('A resposta colada não é um JSON válido. Copie só o JSON que o claude.ai devolveu.'); }
    if (!Array.isArray(dados)) throw new Error('O JSON deve ser uma lista de vínculos.');
    const usadosOC = new Set(), usadasLinha = new Set(), vinculos = [];
    dados.forEach((d, i) => {
      const r = 'Vínculo ' + (i + 1) + ': ';
      if (!Number.isInteger(d.linha) || linhasLivres.indexOf(d.linha) < 0) throw new Error(r + 'linha inexistente ou já vinculada.');
      if (usadasLinha.has(d.linha)) throw new Error(r + 'linha repetida.');
      usadasLinha.add(d.linha);
      if (d.oc_n === null || d.oc_n === undefined) return;
      if (!Number.isInteger(d.oc_n) || ocLivres.indexOf(d.oc_n) < 0) throw new Error(r + 'item da OC inexistente ou já vinculado.');
      if (usadosOC.has(d.oc_n)) throw new Error(r + 'o item ' + d.oc_n + ' da OC foi usado duas vezes.');
      usadosOC.add(d.oc_n);
      vinculos.push({ oc_n: d.oc_n, idx: d.linha, status: 'confirmar', origem: 'ia', score: null, notas: [String(d.justificativa || '').trim(), d.confianca ? 'Confiança da IA: ' + d.confianca + '.' : ''].filter(Boolean) });
    });
    return vinculos;
  }

  // ------------------------------------------------------------ dicionário

  /** Registro para gravar na aba Dicionario quando o usuário confirma ou corrige um vínculo. */
  function registroDicionario(fornecedorId, linha, ocItem, criadoPor) {
    if (!ocItem.material_id) throw new Error('O item ' + ocItem.n + ' da OC não tem código de material; sem ele não dá para gravar no dicionário.');
    return {
      fornecedor_id: fornecedorId, codigo_fornecedor: linha.codigo || '',
      descricao_normalizada: normalizar(linha.descricao_curta || linha.descricao),
      material_id: ocItem.material_id, fator_unidade: linha.fator_unidade || 1, criado_por: criadoPor || ''
    };
  }

  const api = { analisar, normalizar, comparar, fatorFixo, casar, montarPromptIA, aplicarRespostaIA, registroDicionario };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.Casamento = api;
})(typeof window !== 'undefined' ? window : this);
