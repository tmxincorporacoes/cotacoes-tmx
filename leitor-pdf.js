/**
 * Leitor de PDF com texto, sem IA e sem custo (roda no navegador, usa pdf.js).
 *
 * Fluxo: lerArquivo() extrai os textos com posição → agruparLinhas() monta as
 * linhas da página → detectar() reconhece o formato → o leitor do formato devolve
 * o JSON padrão do guia (seção 4), com campos extras: fator_unidade,
 * descricao_curta, total_documento, validade_ate. validar() confere as contas.
 *
 * Formatos conhecidos: orçamento da J K Materiais ('jk'), orçamento da Tambasa
 * ('tambasa') e a ordem de compra do Obra Prima ('oc'). Qualquer outro PDF
 * devolve null em detectar(): a tela deve oferecer o caminho do claude.ai.
 *
 * As funções de leitura recebem só as linhas, então não dependem do navegador.
 */
(function (raiz) {
  'use strict';

  // ------------------------------------------------------------ utilitários

  function numBR(s) {
    if (s === null || s === undefined) return null;
    const t = String(s).replace('%', '').trim();
    if (!/^-?\d{1,3}(\.\d{3})*(,\d+)?$|^-?\d+(,\d+)?$/.test(t)) return null;
    return Number(t.replace(/\./g, '').replace(',', '.'));
  }
  const ehNum = s => numBR(s) !== null;
  const r2 = v => Math.round((v + Number.EPSILON) * 100) / 100;

  function dataISO(s) {
    const m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(s || '');
    return m ? m[3] + '-' + m[2] + '-' + m[1] : null;
  }
  function somaDias(iso, n) {
    const d = new Date(iso + 'T12:00:00Z');
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
  }

  const texto = l => l.itens.map(i => i.s).join(' ');
  const achar = (linhas, re) => linhas.findIndex(l => re.test(texto(l)));
  const doTipo = (linhas, re) => linhas.filter(l => re.test(texto(l)));
  const ultimoNumero = l => { const n = l.itens.map(i => numBR(i.s)).filter(v => v !== null); return n.length ? n[n.length - 1] : null; };
  const itemDepois = (l, rotulo) => {
    const k = l.itens.findIndex(i => rotulo.test(i.s));
    return k >= 0 && l.itens[k + 1] ? l.itens[k + 1].s : null;
  };

  // ------------------------------------------------------------ extração

  /** Agrupa os itens de texto de uma página em linhas (de cima para baixo). */
  function agruparLinhas(itens, pag) {
    const ord = itens.filter(i => i.s.trim()).sort((a, b) => b.y - a.y || a.x - b.x);
    const linhas = [];
    ord.forEach(i => {
      const ult = linhas[linhas.length - 1];
      if (ult && Math.abs(ult.y - i.y) <= 2.5) ult.itens.push(i);
      else linhas.push({ y: i.y, pag, itens: [i] });
    });
    linhas.forEach(l => l.itens.sort((a, b) => a.x - b.x));
    return linhas;
  }

  /** Lê um PDF (ArrayBuffer ou URL) com pdf.js e devolve as linhas de todas as páginas. */
  async function lerArquivo(fonte) {
    const pdfjs = raiz.pdfjsLib;
    if (!pdfjs) throw new Error('pdf.js não carregado.');
    const pdf = await pdfjs.getDocument(fonte instanceof ArrayBuffer ? { data: fonte } : fonte).promise;
    let linhas = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const tc = await (await pdf.getPage(p)).getTextContent();
      const itens = tc.items.map(i => ({ s: i.str, x: i.transform[4], y: i.transform[5] }));
      linhas = linhas.concat(agruparLinhas(itens, p));
    }
    return linhas;
  }

  // ------------------------------------------------------------ detecção

  function detectar(linhas) {
    const t = linhas.map(texto).join('\n');
    if (/Documento Auxiliar de Venda/i.test(t) && /Vl\. Total/.test(t)) return 'jk';
    if (/ORÇAMENTO N°/.test(t) && /LIQUIDO/.test(t) && /Raz[aã]o Empresa/i.test(t)) return 'tambasa';
    if (/Descrição do Produto:/.test(t) && /Vlr\. Unit\./.test(t) && /Vlr\. Total/.test(t) && /Qntd\. itens:/.test(t)) return 'felt';
    if (/ORDEM DE COMPRA\s+\d+/.test(t) && /Subtotal \(R\$\)/.test(t)) return 'oc';
    return null;
  }

  function ler(linhas) {
    const tipo = detectar(linhas);
    if (tipo === 'jk') return lerJK(linhas);
    if (tipo === 'tambasa') return lerTambasa(linhas);
    if (tipo === 'felt') return lerFelt(linhas);
    if (tipo === 'oc') return lerOC(linhas);
    return null;
  }

  // ------------------------------------------------------------ J K Materiais

  function lerJK(linhas) {
    const cab = linhas.find(l => /^Item$/.test(l.itens[0].s) && l.itens.some(i => /^Descrição$/.test(i.s)));
    if (!cab) throw new Error('Cabeçalho da tabela não encontrado (J K).');
    const x = nome => cab.itens.find(i => i.s === nome).x;
    const xDesc = x('Descrição') - 3, xCod = x('Código') - 6, xQt = x('Qt.') - 12;
    const iCab = linhas.indexOf(cab);

    const itens = [];
    for (const l of linhas.slice(iCab + 1)) {
      const no = l.itens[0];
      if (!no || !/^\d{3}$/.test(no.s) || no.x > 40) continue;
      const cod = l.itens.filter(i => i.x >= xCod && i.x < xDesc).map(i => i.s).join(' ');
      const desc = l.itens.filter(i => i.x >= xDesc && i.x < x('Emb.') - 3).map(i => i.s).join(' ');
      const un = (l.itens.find(i => i.x >= x('Emb.') - 3 && i.x < x('Emb.') + 20) || {}).s || null;
      const nums = l.itens.filter(i => i.x >= xQt).map(i => numBR(i.s));
      if (nums.length !== 3 || nums.some(v => v === null)) throw new Error('Linha ' + no.s + ' (J K): colunas numéricas não reconhecidas.');
      const kit = /C\/0*(\d+)\s*UN/i.exec(desc);
      itens.push({
        codigo: cod, descricao: desc, descricao_curta: desc, un,
        qtd: nums[0], preco_unit: nums[1], desconto_pct: null, total: nums[2],
        fator_unidade: kit ? Number(kit[1]) : 1,
        embalagem: kit ? 'kit com ' + Number(kit[1]) + ' (' + kit[0] + ')' : null
      });
    }

    const linhaDe = re => linhas.find(l => re.test(texto(l)));
    const emissaoL = linhaDe(/Data Emiss/);
    const emissao = emissaoL ? dataISO(texto(emissaoL)) : null;
    const valL = linhaDe(/^Validade/);
    const dias = valL ? /(\d+)\s*dias?/i.exec(texto(valL)) : null;
    const iNC = achar(linhas, /Não comprova pagamento/);
    const planoL = linhaDe(/Plano de Pagto/);
    const freteL = linhaDe(/Vl\.Frete/);
    const totalL = linhaDe(/Total Geral/);
    const obs = linhaDe(/^Observações/);
    const obsTxt = obs ? obs.itens.slice(1).map(i => i.s).join(' ').trim() : '';

    return {
      tipo: 'jk',
      fornecedor_curto: 'J K Materiais',
      fornecedor: iNC >= 0 ? texto(linhas[iNC + 1]) : null,
      data_emissao: emissao,
      validade: valL ? texto(valL).replace(/\s+/g, ' ').replace(/\.$/, '') : null,
      validade_ate: emissao && dias ? somaDias(emissao, Number(dias[1])) : null,
      pagamento: planoL ? planoL.itens.slice(1).map(i => i.s).join(' ').replace(/^Plano de Pagto:\s*/, '') : null,
      frete: freteL ? ultimoNumero(freteL) : null,
      despesas_acessorias: null,
      observacoes: obsTxt ? [obsTxt] : [],
      total_documento: totalL ? ultimoNumero(totalL) : null,
      itens
    };
  }

  // ------------------------------------------------------------ Tambasa

  function lerTambasa(linhas) {
    const cab = linhas.find(l => l.itens.some(i => i.s === 'LIQUIDO') && l.itens.some(i => i.s === 'PRODUTO'));
    if (!cab) throw new Error('Cabeçalho da tabela não encontrado (Tambasa).');
    const x = nome => cab.itens.find(i => i.s === nome).x;
    const xUN = x('UN') - 10, xDesc = x('PRODUTO') - 5;
    const iCab = linhas.indexOf(cab);
    const iFim = linhas.findIndex((l, k) => k > iCab && /^Qtd\. Total:/.test(texto(l)));
    const corpo = linhas.slice(iCab + 1, iFim < 0 ? undefined : iFim);

    const itens = [];
    let atual = null;
    for (const l of corpo) {
      const cod = l.itens[0];
      if (/^\d{6}-\d$/.test(cod.s) && cod.x < xDesc) {
        const nums = l.itens.filter(i => i.x >= xUN).map(i => i.s);
        // un, qtd, desc %, líquido, unit, total
        if (nums.length !== 6 || ![1, 2, 3, 4, 5].every(k => ehNum(nums[k]))) throw new Error('Item ' + cod.s + ' (Tambasa): colunas não reconhecidas.');
        atual = { codigo: cod.s, linhas: [l.itens.filter(i => i.x >= xDesc && i.x < xUN).map(i => i.s).join(' ')], nums };
        itens.push(atual);
      } else if (atual) {
        const t = l.itens.filter(i => i.x >= xDesc && i.x < xUN).map(i => i.s).join(' ');
        if (t) atual.linhas.push(t);
      }
    }

    const saida = itens.map(it => {
      const [un, qtd, desc, liq, unit, total] = [it.nums[0], numBR(it.nums[1]), numBR(it.nums[2]), numBR(it.nums[3]), numBR(it.nums[4]), numBR(it.nums[5])];
      const iMarca = it.linhas.findIndex(s => /^MARCA:/.test(s));
      const curta = (iMarca < 0 ? it.linhas : it.linhas.slice(0, iMarca)).join(' ');
      // Embalagem fechada: o líquido é o preço da embalagem e o unit. o preço da peça.
      let fator = 1, emb = null;
      if (Math.abs(liq - unit) > 0.005 && unit > 0) {
        const razao = liq / unit, dica = /\bC(\d{1,3})\b/.exec(curta), cand = dica ? Number(dica[1]) : Math.round(razao);
        // o unit. impresso é arredondado em centavos (0,77 × 50 ≠ 38,29), então a razão sai imprecisa (49,73); a caixa "C50" tem 50
        fator = (cand >= 2 && Math.abs(razao - cand) / cand <= 0.02) ? cand : r2(razao);
        emb = 'embalagem com ' + fator + ' (líquido ' + liq.toFixed(2).replace('.', ',') + ' / unit. ' + unit.toFixed(2).replace('.', ',') + ')';
      } else if (un === 'PR') { fator = 2; emb = 'par (2 peças)'; }
      return {
        codigo: it.codigo, descricao: it.linhas.join(' '), descricao_curta: curta, un,
        qtd, preco_unit: liq, desconto_pct: desc, total,
        fator_unidade: fator, embalagem: emb
      };
    });

    const linhaDe = re => linhas.find(l => re.test(texto(l)));
    const razao = linhaDe(/^Raz[aã]o Empresa:/);
    const pedido = linhaDe(/Data Pedido:/);
    const linhaCond = linhas.find(l => /Pz:/.test(texto(l)));
    const totalL = linhaDe(/Valor Total:/);
    const rod = linhaDe(/DESPESAS ACESS/i);
    const desp = rod ? /DESPESAS ACESS\S*\s*R\$\s*([\d.,]+)/i.exec(texto(rod)) : null;
    const obs = [];
    const iRod = linhas.findIndex(l => /^ORCAMENTO SUJEITO/.test(texto(l)));
    if (iRod >= 0) {
      let buf = '';
      linhas.slice(iRod).forEach(l => {
        const t = texto(l);
        buf += (buf ? ' ' : '') + t;
        if (/[.)]$/.test(t) && !/^FILIAL$|^BA \(/.test(t)) { obs.push(buf); buf = ''; }
      });
      if (buf) obs.push(buf);
    }

    return {
      tipo: 'tambasa',
      fornecedor_curto: 'Tambasa',
      fornecedor: razao ? texto(razao).replace(/^Raz[aã]o Empresa:\s*/, '') : null,
      data_emissao: pedido ? dataISO(texto(pedido)) : null,
      validade: null,
      validade_ate: null,
      pagamento: linhaCond ? linhaCond.itens.filter(i => i.x > 100 && i.x < 320).map(i => i.s).join(' ') : null,
      frete: null,
      despesas_acessorias: desp ? numBR(desp[1]) : null,
      observacoes: obs,
      total_documento: totalL ? ultimoNumero(totalL) : null,
      itens: saida
    };
  }


  // ------------------------------------------------------------ Felt Elétrica (e outras lojas do mesmo sistema)

  function lerFelt(linhas) {
    const cab = linhas.find(l => l.itens.some(i => i.s === 'Vlr. Total') && l.itens.some(i => i.s === 'Vlr. Unit.') && l.itens.some(i => i.s === 'Qtd.'));
    if (!cab) throw new Error('Cabeçalho da tabela não encontrado (Felt).');
    const x = nome => cab.itens.find(i => i.s === nome).x;
    const xNcm = x('NCM') - 8, xMarca = x('Marca') - 5, xQtd = x('Qtd.') - 15, xUnit = x('Vlr. Unit.') - 8, xTot = x('Vlr. Total') - 17;
    const dinheiro = v => numBR(String(v).replace(/^R\$\s*/, ''));
    const entre = (l, a, b) => l.itens.filter(i => i.x >= a && i.x < b).map(i => i.s).join(' ').trim();
    const iCab = linhas.indexOf(cab);

    const itens = [];
    let fimTabela = false, atual = null;
    for (const l of linhas.slice(iCab + 1)) {
      if (fimTabela) break;
      const t = texto(l);
      if (/^Qntd\. itens:/.test(t)) { fimTabela = true; continue; }
      const cod = l.itens[0];
      if (/^\d{1,9}$/.test(cod.s) && cod.x < 62) {
        const desc = entre(l, 62, xNcm), marca = entre(l, xMarca, xQtd);
        const qtd = numBR(entre(l, xQtd, xUnit)), unit = dinheiro(entre(l, xUnit, xTot)), total = dinheiro(entre(l, xTot, 9999));
        if (qtd === null || unit === null || total === null) throw new Error('Linha do código ' + cod.s + ' (Felt): colunas numéricas não reconhecidas.');
        // o preço unitário impresso é arredondado em centavos; o real sai do total
        const real = total / qtd;
        atual = { codigo: cod.s, desc, marca, ncm: entre(l, xNcm, xMarca), qtd, unit, total, real, consistente: Math.abs(real - unit) <= 0.0051 };
        itens.push(atual);
      } else if (atual && l.itens.every(i => i.x >= 62 && i.x < xNcm)) {
        atual.desc += ' ' + t;          // linha extra da descrição
      }
    }

    const saida = itens.map(it => {
      const nome = (it.desc + (it.marca ? ' · ' + it.marca : '')).trim();
      return {
        codigo: it.codigo, descricao: nome, descricao_curta: nome, un: null, qtd: it.qtd,
        // se o preço impresso não é coerente com o total, mantém o impresso para a validação acusar
        preco_unit: it.consistente ? Math.round(it.real * 1e6) / 1e6 : it.unit,
        desconto_pct: null, total: it.total, fator_unidade: 1, embalagem: null
      };
    });

    // dados do documento, por posição (cada rótulo tem o valor logo abaixo)
    const todos = [];
    linhas.forEach(l => l.itens.forEach(i => todos.push({ s: i.s, x: i.x, y: i.y, pag: l.pag })));
    const rotulo = nome => todos.find(i => i.s === nome);
    const datas = todos.filter(i => /^\d{2}\/\d{2}\/\d{4}( \d{2}:\d{2}:\d{2})?$/.test(i.s));
    const valorAbaixo = nome => {
      const r = rotulo(nome); if (!r) return null;
      const c = datas.filter(d => d.pag === r.pag && d.y < r.y && r.y - d.y < 20).sort((a, b) => Math.abs(a.x - r.x) - Math.abs(b.x - r.x))[0];
      return c && Math.abs(c.x - r.x) < 60 ? c.s : null;
    };
    const emissao = valorAbaixo('Data emissão:'), validade = valorAbaixo('Data de Validade:');
    const topo = linhas.filter(l => l.pag === 1 && l.y > 815).map(l => l.itens.filter(i => i.x >= 0 && i.s.length >= 3).map(i => i.s).join(' ')).filter(Boolean)[0] || null;
    const linhaDe = re => linhas.find(l => re.test(texto(l)));
    const iPg = linhas.findIndex(l => /^Forma de Pagamento:/.test(texto(l)));
    const pgto = iPg >= 0 ? linhas.slice(iPg + 1, iPg + 3).map(l => l.itens.filter(i => i.x < 400).map(i => i.s).join(' ').trim()).filter(Boolean) : [];
    const totalL = linhas.find(l => l.itens.some(i => i.s === 'TOTAL:'));
    const declL = linhaDe(/^Qntd\. itens:/);
    const iObs = linhas.findIndex(l => l.itens.some(i => i.s === 'Observação:'));
    const orc = linhaDe(/^Orçamento:/);
    const obs = [];
    if (orc) obs.push(texto(orc).replace(/^Orçamento:\s*/, 'Orçamento nº '));
    if (iObs >= 0 && linhas[iObs + 1]) obs.push('Observação: ' + linhas[iObs + 1].itens.filter(i => i.x < 400).map(i => i.s).join(' '));
    linhas.forEach(l => { const t = texto(l); if (/^- /.test(t)) obs.push(t.replace(/^- /, '')); });

    return {
      tipo: 'felt', fornecedor_curto: topo, fornecedor: topo,
      data_emissao: emissao ? dataISO(emissao) : null,
      validade: validade ? 'Válido até ' + validade : null, validade_ate: validade ? dataISO(validade) : null,
      pagamento: pgto.length ? pgto.join(' · ') : null, frete: null, despesas_acessorias: null, observacoes: obs,
      total_documento: totalL ? dinheiro(totalL.itens[totalL.itens.length - 1].s) : null,
      itens_declarados: declL ? numBR(declL.itens[declL.itens.length - 1].s) : null,
      itens: saida
    };
  }

  // ------------------------------------------------------------ OC (Obra Prima)

  function lerOC(linhas) {
    const num = /ORDEM DE COMPRA\s+(\d+)/.exec(linhas.map(texto).join('\n'));
    const linhaDe = re => linhas.find(l => re.test(texto(l)));
    const obraL = linhaDe(/^OBRA\/CENTRO DE CUSTO:/);
    const dataL = linhaDe(/^Data:/);
    const iForn = achar(linhas, /^DADOS DO FORNECEDOR/);
    const fornL = iForn >= 0 ? linhas.slice(iForn).find(l => /^Nome:/.test(texto(l))) : null;
    const condL = linhaDe(/^Cond\. pgto\.:/);
    const prevL = linhaDe(/Previsão da entrega:/);

    const itens = [];
    let atual = null, emTabela = false, pag = 0, rodape = {};
    for (const l of linhas) {
      if (l.pag !== pag) { pag = l.pag; atual = null; emTabela = false; }
      if (l.itens.some(i => i.s === 'Item') && l.itens.some(i => i.s === 'Qtd.')) { emTabela = true; continue; }
      if (!emTabela) continue;
      const t0 = l.itens[0];
      if (/^(Subtotal|Frete|Total)$/.test(t0.s) && t0.x > 250) {
        rodape[t0.s] = l.itens.slice(1).map(i => numBR(i.s)).filter(v => v !== null);
        atual = null;
        continue;
      }
      if (/^\d{1,3}$/.test(t0.s) && t0.x < 35) {
        const nums = l.itens.filter(i => i.x >= 235).map(i => numBR(i.s));
        if (nums.length !== 5 || nums.some(v => v === null)) throw new Error('Item ' + t0.s + ' (OC): colunas numéricas não reconhecidas.');
        atual = { n: Number(t0.s), desc: [l.itens.filter(i => i.x >= 40 && i.x < 235).map(i => i.s).join(' ')], un: null, nums };
        itens.push(atual);
      } else if (atual) {
        const d = l.itens.filter(i => i.x >= 40 && i.x < 235).map(i => i.s).join(' ');
        if (d) atual.desc.push(d);
        const u = l.itens.find(i => i.x >= 235 && /^[A-Za-zçÇ]+$/.test(i.s));
        if (u) atual.un = /^unidades?$/i.test(u.s) ? 'UN' : u.s;
      }
    }

    const sub = rodape.Subtotal || [];
    return {
      tipo: 'oc',
      numero: num ? num[1] : null,
      obra: obraL ? obraL.itens.slice(1).map(i => i.s).filter(s => !/^CNO:?$/.test(s)).join(' ').trim() : null,
      data_emissao: dataL ? dataISO(texto(dataL)) : null,
      fornecedor: fornL ? fornL.itens.filter(i => i.x > 85 && i.x < 340).map(i => i.s).join(' ') : null,
      cond_pgto: condL ? condL.itens.filter(i => i.x > 90 && i.x < 340).map(i => i.s).join(' ').trim() || null : null,
      previsao_entrega: prevL ? prevL.itens.filter(i => i.x > 400).map(i => i.s).join(' ').trim() || null : null,
      subtotal: sub[0] === undefined ? null : sub[0],
      desconto_total: sub[1] === undefined ? null : sub[1],
      total_documento: sub[2] === undefined ? null : sub[2],
      frete: (rodape.Frete || [])[0] === undefined ? null : rodape.Frete[0],
      itens: itens.map(it => ({
        n: it.n, descricao: it.desc.join(' '), qtd: it.nums[0], un: it.un || 'UN', preco_unit: it.nums[1],
        subtotal: it.nums[2], desconto_valor: it.nums[3], total: it.nums[4]
      }))
    };
  }


  // ------------------------------------------------------------ PDF novo: claude.ai (manual)

  /** Texto para colar no claude.ai junto com o PDF. A resposta volta no JSON padrão do guia. */
  function montarPromptExtracao() {
    return [
      'Leia o documento anexo, que é uma cotação (orçamento) de fornecedor de materiais de construção, e devolva SOMENTE um JSON, sem texto fora dele, neste formato:',
      '{',
      '  "fornecedor": "texto como aparece no documento",',
      '  "data_emissao": "AAAA-MM-DD",',
      '  "validade": "texto da validade como aparece, ou null",',
      '  "pagamento": "condição de pagamento como aparece, ou null",',
      '  "frete": 0.0,',
      '  "despesas_acessorias": 0.0,',
      '  "observacoes": ["texto livre relevante do rodapé"],',
      '  "itens": [',
      '    { "codigo": "código do fornecedor", "descricao": "descrição como aparece", "un": "UN", "qtd": 0,',
      '      "preco_unit": 0.0, "desconto_pct": 0, "total": 0.0, "embalagem": null }',
      '  ]',
      '}',
      'Regras: não invente campo ausente, use null. Marca, NCM e código de barras que vêm dentro da descrição ficam em "descricao" sem alteração.',
      'preco_unit é o preço por unidade comercial vendida (se vende saco, caixa ou kit fechado, é o preço dessa embalagem). Se a coluna "líquido" difere da "unit.", registre a diferença em "embalagem".',
      'Números em formato JSON (ponto decimal, sem separador de milhar). Datas em AAAA-MM-DD.'
    ].join('\n');
  }

  /** Lê o JSON colado do claude.ai e devolve no mesmo formato dos leitores. Erro claro, não grava nada. */
  function lerJSONColado(entrada) {
    let t = String(entrada || '').trim();
    const cerca = /```(?:json)?\s*([\s\S]*?)```/.exec(t);
    if (cerca) t = cerca[1].trim();
    let d;
    try { d = JSON.parse(t); } catch (e) { throw new Error('O texto colado não é um JSON válido. Copie só o JSON que o claude.ai devolveu.'); }
    if (!d || typeof d !== 'object' || Array.isArray(d) || !Array.isArray(d.itens) || !d.itens.length) throw new Error('O JSON precisa ter a lista "itens" com pelo menos um item.');
    const numero = (v, campo, obrigatorio) => {
      if (v === null || v === undefined || v === '') { if (obrigatorio) throw new Error('Campo ausente: ' + campo + '.'); return null; }
      const n = typeof v === 'number' ? v : numBR(v);
      if (n === null || !isFinite(n)) throw new Error('Número inválido em ' + campo + ': ' + v + '.');
      return n;
    };
    const iso = v => (!v ? null : /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : dataISO(v));
    const itens = d.itens.map((i, k) => {
      const r = 'Item ' + (k + 1) + ': ';
      if (!i || !String(i.descricao || '').trim()) throw new Error(r + 'sem descrição.');
      const e = x => { try { return numero(x.v, x.c, x.o); } catch (er) { throw new Error(r + er.message); } };
      return {
        codigo: i.codigo === null || i.codigo === undefined ? '' : String(i.codigo),
        descricao: String(i.descricao).trim(), descricao_curta: String(i.descricao).trim(), un: i.un || null,
        qtd: e({ v: i.qtd, c: 'qtd', o: true }), preco_unit: e({ v: i.preco_unit, c: 'preco_unit', o: true }),
        desconto_pct: e({ v: i.desconto_pct, c: 'desconto_pct' }), total: e({ v: i.total, c: 'total', o: true }),
        fator_unidade: e({ v: i.fator_unidade, c: 'fator_unidade' }) || 1, embalagem: i.embalagem || null
      };
    });
    const emissao = iso(d.data_emissao);
    const dias = d.validade ? /(\d+)\s*dias?/i.exec(String(d.validade)) : null;
    const total = itens.reduce((s, i) => s + i.total, 0);
    return {
      tipo: 'ia', fornecedor_curto: null, fornecedor: d.fornecedor ? String(d.fornecedor) : null,
      data_emissao: emissao, validade: d.validade || null, validade_ate: emissao && dias ? somaDias(emissao, Number(dias[1])) : null,
      pagamento: d.pagamento || null, frete: numero(d.frete, 'frete'), despesas_acessorias: numero(d.despesas_acessorias, 'despesas_acessorias'),
      observacoes: Array.isArray(d.observacoes) ? d.observacoes.map(String) : [],
      total_documento: r2(total), itens
    };
  }

  // ------------------------------------------------------------ validação

  /**
   * Confere as contas. Retorna { ok, avisos:[{nivel,msg}] }. nivel 'erro' = conta
   * que não fecha; 'aviso' = diferença pequena ou ponto a observar.
   */
  function validar(d) {
    const avisos = [];
    const add = (nivel, msg) => avisos.push({ nivel, msg });
    if (!d.itens.length) add('erro', 'Nenhum item lido.');

    if (d.tipo === 'oc') {
      d.itens.forEach(i => {
        if (Math.abs(r2(i.qtd * i.preco_unit) - i.subtotal) > 0.011) add('erro', 'Item ' + i.n + ': qtd × unit. ≠ subtotal.');
        if (Math.abs(r2(i.subtotal - i.desconto_valor) - i.total) > 0.011) add('erro', 'Item ' + i.n + ': subtotal − desconto ≠ total.');
      });
      const somaT = r2(d.itens.reduce((s, i) => s + i.total, 0));
      const dif = r2(d.total_documento - somaT);
      if (Math.abs(dif) > 0.05) add('erro', 'Soma das linhas ' + somaT + ' difere do total ' + d.total_documento + '.');
      else if (dif !== 0) add('aviso', 'Soma das linhas ' + somaT.toFixed(2) + ' × total impresso ' + d.total_documento.toFixed(2) + ': diferença de arredondamento do PDF.');
      if (!d.cond_pgto) add('aviso', 'OC sem condição de pagamento.');
      if (!d.previsao_entrega) add('aviso', 'OC sem previsão de entrega.');
    } else {
      d.itens.forEach(i => {
        if (Math.abs(r2(i.qtd * i.preco_unit) - i.total) > 0.011) add('erro', 'Item ' + i.codigo + ': qtd × preço ≠ total.');
      });
      if (d.itens_declarados !== undefined && d.itens_declarados !== null && d.itens_declarados !== d.itens.length) add('erro', 'O documento declara ' + d.itens_declarados + ' itens, mas foram lidos ' + d.itens.length + '.');
      const soma = r2(d.itens.reduce((s, i) => s + i.total, 0));
      const alvo = d.tipo === 'jk' ? d.total_documento - (d.frete || 0) : d.total_documento;
      if (alvo !== null && Math.abs(soma - alvo) > 0.011) add('erro', 'Soma dos itens ' + soma + ' difere do total do documento ' + alvo + '.');
      if (!d.validade) add('aviso', 'Cotação sem validade.');
      if (d.despesas_acessorias) add('aviso', 'Despesas acessórias de R$ ' + d.despesas_acessorias.toFixed(2).replace('.', ',') + ' entram no total do fornecedor, não por item.');
      d.itens.forEach(i => { if (i.fator_unidade > 1) add('aviso', 'Item ' + i.codigo + ': vendido em embalagem com ' + i.fator_unidade + ' peças.'); });
    }
    return { ok: !avisos.some(a => a.nivel === 'erro'), avisos };
  }

  const api = { numBR, dataISO, agruparLinhas, lerArquivo, detectar, ler, lerJK, lerTambasa, lerFelt, lerOC, validar, montarPromptExtracao, lerJSONColado };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.LeitorPDF = api;
})(typeof window !== 'undefined' ? window : this);
