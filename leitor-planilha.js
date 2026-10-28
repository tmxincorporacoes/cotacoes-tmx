/**
 * Leitor de cotação em Excel (.xlsx, .xls) e CSV, sem IA e sem custo.
 *
 * Acha sozinho a linha de cabeçalho e as colunas pelo dicionário de nomes (PAPEIS),
 * lê o que estiver acima e abaixo da tabela (fornecedor, validade, pagamento, frete,
 * total) e devolve o mesmo JSON dos outros leitores (leitor-pdf.js), para passar
 * pela mesma validação e pelo mesmo casamento. Se não achar as colunas, falha com
 * err.codigo === 'sem_cabecalho' e a tela pede o mapeamento manual (extrair com opcoes).
 *
 * Funções puras, exceto lerArquivo (usa File, TextDecoder e o SheetJS, global XLSX).
 */
(function (raiz) {
  'use strict';

  const r2 = v => Math.round((v + Number.EPSILON) * 100) / 100;
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9%]+/g, ' ').trim();

  // Nomes de coluna que cada fornecedor usa para o mesmo campo (normalizados: sem acento, minúsculas).
  // "fraco" = nome ambíguo; só vale se não houver um nome forte para o mesmo papel.
  const PAPEIS = {
    codigo: { forte: ['codigo', 'cod', 'cod produto', 'codigo produto', 'cod prod', 'codigo do produto', 'ref', 'referencia', 'sku', 'cod fab', 'codigo fab', 'codigo do fornecedor', 'cod fornecedor', 'codigo fornecedor', 'ref fornecedor'], fraco: [] },
    // colunas da OC padrão da TMX (oc-padrao.js)
    codigo_tmx: { forte: ['codigo obra prima', 'cod obra prima', 'codigo tmx', 'cod tmx', 'codigo interno', 'cod interno', 'codigo insumo', 'cod insumo'], fraco: [] },
    n_oc: { forte: ['n do item na oc', 'n item oc', 'item oc', 'n do item', 'numero do item', 'n item'], fraco: [] },
    descricao_pedida: { forte: ['descricao pedida', 'descricao na oc', 'descricao oc', 'material pedido', 'item pedido'], fraco: [] },
    descricao_cotada: { forte: ['produto cotado', 'descricao cotada', 'produto ofertado', 'marca modelo'], fraco: [] },
    descricao: { forte: ['descricao', 'descricao produto', 'descricao do produto', 'descricao do item', 'produto', 'material', 'mercadoria', 'nome', 'nome produto'], fraco: ['item'] },
    un: { forte: ['un', 'und', 'unid', 'unidade', 'um', 'unid medida'], fraco: ['emb', 'embalagem'] },
    qtd: { forte: ['qtd', 'qtde', 'quant', 'quantidade', 'qt', 'qtd cotada', 'quantidade cotada'], fraco: [] },
    preco_unit: { forte: ['preco', 'preco unit', 'preco unitario', 'valor unit', 'valor unitario', 'vl unit', 'vlr unit', 'p unit', 'punit', 'preco un', 'custo unit', 'preco liquido'], fraco: ['unitario', 'unit', 'valor', 'liquido'] },
    desconto_pct: { forte: ['desc', 'desconto', 'desc %', '% desc', 'desconto %', 'desc pct', 'desconto pct'], fraco: [] },
    total: { forte: ['total', 'valor total', 'vl total', 'vlr total', 'total item', 'preco total', 'subtotal'], fraco: [] },
    ignorar: { forte: ['ean', 'cod barras', 'codigo de barras', 'ncm', 'marca', 'cod ean', 'gtin'], fraco: [] }
  };
  const NUMERICOS = ['qtd', 'preco_unit', 'desconto_pct', 'total'];
  const ROTULOS = { codigo: 'Código', codigo_tmx: 'Código Obra Prima', n_oc: 'Nº do item na OC', descricao_pedida: 'Descrição pedida', descricao_cotada: 'Produto cotado', descricao: 'Descrição', un: 'Unidade', qtd: 'Quantidade', preco_unit: 'Preço unitário', desconto_pct: 'Desconto %', total: 'Total da linha' };

  // ------------------------------------------------------------ números e datas

  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let t = String(v).replace(/R\$|%/g, '').replace(/\s/g, '');
    if (t === '' || !/^-?[\d.,]+$/.test(t)) return null;
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, '');
    const n = Number(t);
    return isFinite(n) ? n : null;
  }

  function dataISO(v) {
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    if (typeof v === 'number' && v > 30000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 86400000)); return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0'); }
    const s = String(v == null ? '' : v);
    let m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(s); if (m) return m[3] + '-' + m[2] + '-' + m[1];
    m = /(\d{4})-(\d{2})-(\d{2})/.exec(s); return m ? m[0] : null;
  }
  function somaDias(iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }

  // ------------------------------------------------------------ CSV

  /** Lê CSV com delimitador ; , ou tabulação (detectado), aspas e quebras de linha dentro de aspas. */
  function parseCSV(texto) {
    texto = String(texto).replace(/^﻿/, '');
    const amostra = texto.split(/\r?\n/).filter(l => l.trim()).slice(0, 8).join('\n').replace(/"[^"]*"/g, '');
    const cont = d => (amostra.match(new RegExp(d === '\t' ? '\t' : '\\' + d, 'g')) || []).length;
    const delim = [';', '\t', ','].map(d => [d, cont(d)]).sort((a, b) => b[1] - a[1])[0][0];
    const linhas = []; let campo = '', linha = [], aspas = false;
    for (let i = 0; i < texto.length; i++) {
      const c = texto[i];
      if (aspas) {
        if (c === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else aspas = false; } else campo += c;
      } else if (c === '"') aspas = true;
      else if (c === delim) { linha.push(campo); campo = ''; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && texto[i + 1] === '\n') i++; linha.push(campo); linhas.push(linha); linha = []; campo = ''; }
      else campo += c;
    }
    if (campo !== '' || linha.length) { linha.push(campo); linhas.push(linha); }
    return linhas;
  }

  // ------------------------------------------------------------ cabeçalho

  function papelDe(cel) {
    const h = norm(cel);
    if (!h) return null;
    let melhor = null;
    Object.keys(PAPEIS).forEach(papel => ['forte', 'fraco'].forEach(nivel => PAPEIS[papel][nivel].forEach(syn => {
      const ok = h === syn || (' ' + h + ' ').includes(' ' + syn + ' ');
      if (ok && (!melhor || syn.length > melhor.syn.length || (syn.length === melhor.syn.length && h === syn))) melhor = { papel, nivel, syn };
    })));
    return melhor;
  }

  /** Acha a linha de cabeçalho e as colunas. Devolve null se não achar descrição, quantidade e preço. */
  function detectarCabecalho(rows, opt) {
    const semPreco = !!(opt && opt.semPreco);
    let melhor = null;
    rows.slice(0, 60).forEach((linha, i) => {
      const achados = {};
      linha.forEach((cel, c) => {
        const p = papelDe(cel);
        if (!p || p.papel === 'ignorar') return;
        const atual = achados[p.papel];
        if (!atual || (atual.nivel === 'fraco' && p.nivel === 'forte')) achados[p.papel] = { col: c, nivel: p.nivel };
      });
      const tem = k => achados[k] !== undefined;
      if (!((tem('descricao') || tem('descricao_pedida') || tem('descricao_cotada')) && tem('qtd') && (semPreco || tem('preco_unit')))) return;
      const score = Object.keys(achados).length;
      if (!melhor || score > melhor.score) melhor = { score, linhaCab: i, colunas: Object.keys(achados).reduce((o, k) => { o[k] = achados[k].col; return o; }, {}) };
    });
    return melhor ? { linhaCab: melhor.linhaCab, colunas: melhor.colunas } : null;
  }

  // ------------------------------------------------------------ extração

  const ROT_META = [
    ['fornecedor', /^(fornecedor|raz[aã]o social|empresa|cota[cç][aã]o de|proponente)\b/i],
    ['validade', /^(validade|v[aá]lid[ao]|prazo de validade)\b/i],
    ['pagamento', /^(pagamento|cond[^\s]*\.? ?(de )?(pagamento|pgto)|forma de pagamento|prazo de pagamento)\b/i],
    ['frete', /^frete\b/i],
    ['despesas', /^(despesas|desp\.? acess)/i],
    ['emissao', /^(data( da cota[cç][aã]o| de emiss[aã]o)?|emiss[aã]o)\b/i],
    ['obs', /^(observa[cç][oõ]es?|obs\.?)\b/i],
    ['entrega', /^(prazo de entrega|previs[aã]o de entrega|entrega)\b/i],
    ['oc_numero', /^(ordem de compra|oc\b|n[º°o.]*\s*(da )?oc\b|numero da oc|pedido de compra)/i],
    ['obra', /^(obra|centro de custo|obra\/centro de custo)\b/i],
    ['total', /^(total geral|valor total( geral)?|total da cota[cç][aã]o)\b/i]
  ];

  function valorDoRotulo(linha, c) {
    const t = String(linha[c]).trim();
    const dp = t.indexOf(':');
    if (dp > 0 && t.slice(dp + 1).trim() !== '') return t.slice(dp + 1).trim();
    for (let k = c + 1; k < linha.length; k++) if (linha[k] !== '' && linha[k] != null) return linha[k];
    return null;
  }

  /**
   * @param rows   matriz de linhas (arrays de células)
   * @param opcoes { linhaCab (0-based), colunas:{papel: índice} } para mapeamento manual
   */
  function extrair(rows, opcoes) {
    rows = rows.map(l => l.map(c => (c === null || c === undefined ? '' : c)));
    const cab = opcoes && opcoes.colunas ? { linhaCab: opcoes.linhaCab, colunas: opcoes.colunas } : detectarCabecalho(rows);
    const temDesc = cab && (cab.colunas.descricao !== undefined || cab.colunas.descricao_pedida !== undefined || cab.colunas.descricao_cotada !== undefined);
    if (!cab || !temDesc || cab.colunas.qtd === undefined || cab.colunas.preco_unit === undefined) {
      const e = new Error('Não encontrei as colunas de descrição, quantidade e preço unitário nesta planilha.');
      e.codigo = 'sem_cabecalho'; e.rows = rows; throw e;
    }
    const col = cab.colunas, meta = {}, itens = [], usadas = new Set();
    let totalCalculado = false;

    rows.forEach((linha, i) => {
      if (i <= cab.linhaCab) {
        // acima do cabeçalho: dados do documento
        linha.forEach((cel, c) => { if (typeof cel === 'string') ROT_META.forEach(([k, re]) => { if (meta[k] === undefined && re.test(cel.trim())) meta[k] = valorDoRotulo(linha, c); }); });
        return;
      }
      const txt = c => (c === undefined ? '' : String(linha[c]).trim());
      const pedida = txt(col.descricao_pedida), cotada = txt(col.descricao_cotada);
      const desc = cotada || pedida || txt(col.descricao), q = num(linha[col.qtd]), p = num(linha[col.preco_unit]);
      if (desc && q !== null && q > 0 && p !== null && p >= 0 && !/^total\b/i.test(desc)) {
        const d = col.desconto_pct !== undefined ? num(linha[col.desconto_pct]) : null;
        let tot = col.total !== undefined ? num(linha[col.total]) : null;
        let preco = p;
        if (d && d > 0 && tot !== null && Math.abs(r2(q * p * (1 - d / 100)) - tot) <= 0.011) preco = Math.round(p * (1 - d / 100) * 10000) / 10000;
        if (tot === null) { tot = r2(q * preco); if (col.total === undefined) totalCalculado = true; }
        const un = col.un !== undefined ? String(linha[col.un]).trim() : '';
        const kit = /C\/\s*0*(\d+)\s*UN/i.exec(desc);
        itens.push({
          codigo: col.codigo !== undefined && linha[col.codigo] !== '' ? String(linha[col.codigo]).trim() : '', descricao: desc, descricao_curta: desc, un: un || null,
          codigo_tmx: col.codigo_tmx !== undefined && linha[col.codigo_tmx] !== '' ? String(linha[col.codigo_tmx]).trim() : '',
          n_oc: col.n_oc !== undefined && num(linha[col.n_oc]) !== null ? num(linha[col.n_oc]) : null,
          descricao_pedida: pedida || null,
          qtd: q, preco_unit: preco, desconto_pct: d, total: tot,
          fator_unidade: kit ? Number(kit[1]) : (/^PR$/i.test(un) ? 2 : 1),
          embalagem: kit ? 'kit com ' + kit[1] + ' (' + kit[0] + ')' : (/^PR$/i.test(un) ? 'par (2 peças)' : (d && preco !== p ? 'preço líquido, já com o desconto de ' + d + '%' : null))
        });
        usadas.add(i);
        return;
      }
      // abaixo da tabela ou linha de apoio: frete, total, validade…
      linha.forEach((cel, c) => { if (typeof cel === 'string') ROT_META.forEach(([k, re]) => { if (meta[k] === undefined && re.test(cel.trim())) meta[k] = valorDoRotulo(linha, c); }); });
    });

    if (!itens.length) { const e = new Error('A tabela foi encontrada, mas nenhuma linha tem descrição, quantidade e preço válidos.'); e.codigo = 'sem_itens'; throw e; }
    const emissao = meta.emissao != null ? dataISO(meta.emissao) : null;
    const dias = meta.validade != null ? /(\d+)\s*dias?/i.exec(String(meta.validade)) : null;
    const validadeISO = meta.validade != null ? dataISO(meta.validade) : null;
    const soma = r2(itens.reduce((s, i) => s + i.total, 0));
    const colunasLidas = Object.keys(col).filter(k => ROTULOS[k]).map(k => ROTULOS[k] + ' ← "' + String(rows[cab.linhaCab][col[k]]).trim() + '"');
    return {
      tipo: 'planilha', fornecedor_curto: null,
      fornecedor: meta.fornecedor != null ? String(meta.fornecedor).trim() : null,
      data_emissao: emissao, validade: meta.validade != null ? String(meta.validade instanceof Date ? dataISO(meta.validade) : meta.validade).trim() : null,
      validade_ate: validadeISO || (emissao && dias ? somaDias(emissao, Number(dias[1])) : null),
      pagamento: meta.pagamento != null ? String(meta.pagamento).trim() : null,
      frete: meta.frete != null ? num(meta.frete) : null, despesas_acessorias: meta.despesas != null ? num(meta.despesas) : null,
      oc_numero: meta.oc_numero != null ? String(meta.oc_numero).trim() : null,
      observacoes: [meta.entrega != null && String(meta.entrega).trim() ? 'Prazo de entrega: ' + String(meta.entrega).trim() : null, meta.obs != null ? String(meta.obs).trim() : null, totalCalculado ? 'A planilha não tem coluna de total; o total de cada linha foi calculado (qtd × preço).' : null].filter(Boolean),
      total_documento: soma, total_impresso: meta.total != null ? num(meta.total) : null,
      colunas_lidas: colunasLidas, itens
    };
  }


  // ------------------------------------------------------------ ordem de compra em Excel/CSV

  /**
   * Lê uma OC (itens a comprar) de uma planilha. Só descrição e quantidade são obrigatórias;
   * unidade, código do insumo e preço são lidos se existirem. Devolve o mesmo formato do leitor
   * de PDF da OC (tipo 'oc'), com material_id por item. Se não achar as colunas, falha com
   * err.codigo === 'sem_cabecalho' (a tela abre o mapeamento manual).
   */
  function extrairOC(rows, opcoes) {
    rows = rows.map(l => l.map(c => (c === null || c === undefined ? '' : c)));
    const cab = opcoes && opcoes.colunas ? { linhaCab: opcoes.linhaCab, colunas: opcoes.colunas } : detectarCabecalho(rows, { semPreco: true });
    const col = cab && cab.colunas;
    const temDesc = col && (col.descricao !== undefined || col.descricao_pedida !== undefined || col.descricao_cotada !== undefined);
    if (!cab || !temDesc || col.qtd === undefined) {
      const e = new Error('Não encontrei as colunas de descrição e quantidade nesta planilha.');
      e.codigo = 'sem_cabecalho'; e.rows = rows; throw e;
    }
    const meta = {};
    const lerMeta = linha => linha.forEach((cel, c) => { if (typeof cel === 'string') ROT_META.forEach(([k, re]) => { if (meta[k] === undefined && re.test(cel.trim())) meta[k] = valorDoRotulo(linha, c); }); });
    const txt = (linha, c) => (c === undefined ? '' : String(linha[c]).trim());
    const itens = [];
    rows.forEach((linha, i) => {
      if (i <= cab.linhaCab) { lerMeta(linha); return; }
      const desc = txt(linha, col.descricao_pedida !== undefined ? col.descricao_pedida : col.descricao) || txt(linha, col.descricao_cotada);
      const q = num(linha[col.qtd]);
      if (desc && q !== null && q > 0 && !/^total\b/i.test(desc)) {
        const p = col.preco_unit !== undefined ? num(linha[col.preco_unit]) : null;
        const cod = col.codigo_tmx !== undefined ? txt(linha, col.codigo_tmx) : (col.codigo !== undefined ? txt(linha, col.codigo) : '');
        itens.push({ n: itens.length + 1, material_id: cod, descricao: desc, qtd: q, un: (txt(linha, col.un) || 'UN'), preco_unit: p, subtotal: p !== null ? r2(q * p) : null, desconto_valor: null, total: p !== null ? r2(q * p) : null });
      } else lerMeta(linha);
    });
    if (!itens.length) { const e = new Error('A tabela foi encontrada, mas nenhuma linha tem descrição e quantidade válidas.'); e.codigo = 'sem_itens'; throw e; }
    const emissao = meta.emissao != null ? dataISO(meta.emissao) : null;
    const soma = r2(itens.reduce((s, i) => s + (i.total || 0), 0));
    return {
      tipo: 'oc', origem: 'planilha',
      numero: meta.oc_numero != null ? String(meta.oc_numero).trim() : '', obra: meta.obra != null ? String(meta.obra).trim() : '',
      data_emissao: emissao, fornecedor: meta.fornecedor != null ? String(meta.fornecedor).trim() : null,
      cond_pgto: meta.pagamento != null ? String(meta.pagamento).trim() : null, previsao_entrega: meta.entrega != null ? String(meta.entrega).trim() : null,
      subtotal: col.preco_unit !== undefined ? soma : null, desconto_total: null, total_documento: col.preco_unit !== undefined ? soma : null, frete: null,
      colunas_lidas: Object.keys(col).filter(k => ROTULOS[k]).map(k => ROTULOS[k] + ' ← "' + String(rows[cab.linhaCab][col[k]]).trim() + '"'),
      itens
    };
  }

  // ------------------------------------------------------------ arquivo

  function textoDe(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { return new TextDecoder('windows-1252').decode(buf); }
  }

  /** Lê o arquivo escolhido e devolve a matriz de linhas (a primeira aba que tiver a tabela, no Excel). */
  async function lerArquivo(file, opt) {
    const buf = await file.arrayBuffer();
    if (/\.csv$|\.txt$/i.test(file.name)) return parseCSV(textoDe(buf));
    if (!raiz.XLSX) throw new Error('A biblioteca de planilhas não carregou. Confira a conexão com a internet.');
    const wb = raiz.XLSX.read(buf, { type: 'array', cellDates: true });
    let primeira = null;
    for (const nome of wb.SheetNames) {
      const rows = raiz.XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, raw: true, defval: '' });
      if (!primeira && rows.length) primeira = rows;
      if (detectarCabecalho(rows, opt)) return rows;
    }
    return primeira || [];
  }


  /** Todas as abas do arquivo: [{ nome, rows }]. CSV tem uma só (o nome do arquivo). */
  async function lerAbas(file) {
    const buf = await file.arrayBuffer();
    if (/\.csv$|\.txt$/i.test(file.name)) return [{ nome: file.name.replace(/\.[^.]+$/, ''), rows: parseCSV(textoDe(buf)) }];
    if (!raiz.XLSX) throw new Error('A biblioteca de planilhas não carregou. Confira a conexão com a internet.');
    const wb = raiz.XLSX.read(buf, { type: 'array', cellDates: true });
    return wb.SheetNames.map(nome => ({ nome, rows: raiz.XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, raw: true, defval: '' }) }));
  }

  const api = { num, dataISO, parseCSV, detectarCabecalho, extrair, extrairOC, lerArquivo, lerAbas, ROTULOS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.LeitorPlanilha = api;
})(typeof window !== 'undefined' ? window : this);
