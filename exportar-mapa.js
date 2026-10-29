/**
 * Exportação do mapa de cotação em Excel (.xlsx), com quatro abas:
 *   Mapa           itens da OC lado a lado, com um bloco de colunas por fornecedor e totais
 *   Fornecedores   resumo de cada cotação (total, ajustes, validade, pagamento, entrega)
 *   Pontos         os pontos a confirmar antes de fechar
 *   Fora da OC     linhas que o fornecedor cotou e que não estão na OC
 * Os números saem como números (somam e formatam no Excel). montar() é pura e testável;
 * baixar() usa o SheetJS (XLSX) e o navegador. Depende de calculo.js.
 */
(function (raiz) {
  'use strict';

  const C = () => raiz.Calculo || require('./calculo.js');
  const ROTULO_STATUS = { confere: 'Confere', confirmar: 'Confirmar produto', manual: 'Vínculo manual', sem_cotacao: 'Fora da OC' };
  const dataBR = iso => (/^\d{4}-\d{2}-\d{2}/.test(iso || '') ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : (iso || ''));
  const numOuVazio = v => (v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v));
  const COLS_FORN = ['Código', 'Descrição cotada', 'Preço cotado (R$)', 'Fator', 'Total (R$)', 'Status', 'Observação'];

  /**
   * @param p { oc, cotacoes, nomeForn(id), alertas:[{nivel,titulo,texto}], hoje (ISO) }
   * @returns { abas:[{nome, aoa, merges, largs, dinheiro:[{c, de, ate}]}], nomeArquivo }
   */
  function montar(p) {
    const { oc, cotacoes: cots, nomeForn } = p, K = C();
    if (!cots.length) throw new Error('Não há cotações para exportar.');
    const nf = cots.map(c => nomeForn(c.fornecedor_id));
    const tot = cots.map(c => K.totalCot(c, oc));

    // ---- aba Mapa
    const NB = COLS_FORN.length, base = 4;
    const aoa = [
      ['Mapa de cotação · OC ' + oc.numero + (oc.obra ? ' · ' + oc.obra : '')],
      ['Gerado em ' + dataBR(p.hoje) + '. Os totais são calculados sobre as quantidades da OC, não sobre o que cada fornecedor cotou. Preço dividido pelo fator de unidade (kit, par, saco).'],
      []
    ];
    const grupo = ['', '', '', ''], cab = ['Item', 'Descrição da OC', 'Qtd. OC', 'Un.'];
    nf.forEach(n => { grupo.push(n); for (let i = 1; i < NB; i++) grupo.push(''); COLS_FORN.forEach(h => cab.push(h)); });
    grupo.push(''); cab.push('Menor preço');
    aoa.push(grupo, cab);
    const linhaCab = aoa.length;                       // 1-based da linha de cabeçalho = aoa.length
    oc.itens.forEach(it => {
      const l = [it.n, it.descricao, numOuVazio(it.qtd), it.unidade || ''];
      cots.forEach(c => {
        const lk = K.linhaDe(c, it.n);
        if (!lk) { l.push('', 'Não cotado', '', '', '', '', ''); return; }
        l.push(lk.codigo || '', lk.descricao_curta || lk.descricao_original || '', numOuVazio(lk.preco_unit), numOuVazio(lk.fator_unidade) || 1,
          K.custo(c, lk, it.qtd), ROTULO_STATUS[lk.status] || lk.status, [lk.nota, lk.embalagem ? 'Embalagem: ' + lk.embalagem : '', (() => { const cx = K.caixas(c, lk, it.qtd); return cx ? 'Compra de ' + cx.caixas + ' emb. de ' + cx.fator + ' = ' + cx.pecas + ' peças' + (cx.sobra > 0 ? ' (sobram ' + cx.sobra + ')' : '') : ''; })()].filter(Boolean).join(' | '));
      });
      const m = K.menorPreco(it, cots);
      l.push(m ? nf[m.indice] : '');
      aoa.push(l);
    });
    const colTotal = k => base + k * NB + 4;           // índice (0-based) da coluna "Total (R$)" do fornecedor k
    const linhaTotais = (rotulo, f, fmtPct) => {
      const l = new Array(base + cots.length * NB + 1).fill('');
      l[1] = rotulo;
      cots.forEach((c, k) => { l[colTotal(k)] = f(c, tot[k]); });
      aoa.push(l);
    };
    aoa.push([]);
    linhaTotais('Itens cotados (de ' + oc.itens.length + ')', (c, t) => t.n);
    linhaTotais('Desconto (%)', c => numOuVazio(c.desconto_pct));
    linhaTotais('Acréscimo / ICMS (%)', c => numOuVazio(c.acrescimo_pct));
    linhaTotais('Subtotal dos itens (R$)', (c, t) => t.itens);
    linhaTotais('Frete (R$)', (c, t) => t.frete);
    linhaTotais('Despesas acessórias (R$)', (c, t) => t.desp);
    linhaTotais('TOTAL DO FORNECEDOR (R$)', (c, t) => t.total);
    const ultima = aoa.length;

    const merges = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } }];
    nf.forEach((n, k) => merges.push({ s: { r: 3, c: base + k * NB }, e: { r: 3, c: base + k * NB + NB - 1 } }));
    const largs = [6, 44, 9, 6];
    nf.forEach(() => [14, 34, 14, 7, 14, 18, 40].forEach(w => largs.push(w)));
    largs.push(18);
    const dinheiro = [];
    cots.forEach((c, k) => { dinheiro.push({ c: base + k * NB + 2, de: linhaCab, ate: linhaCab + oc.itens.length }, { c: colTotal(k), de: linhaCab, ate: ultima }); });

    // ---- aba Fornecedores
    const hojeIso = p.hoje;
    const resumo = [['Fornecedor', 'Itens cotados', 'Subtotal dos itens (R$)', 'Desconto (%)', 'Acréscimo / ICMS (%)', 'Frete (R$)', 'Despesas (R$)', 'Total (R$)', 'Emissão', 'Validade', 'Situação da validade', 'Pagamento', 'Prazo de entrega', 'Arquivo', 'Embalagem fechada']];
    cots.forEach((c, k) => {
      const venc = c.validade_ate && c.validade_ate < hojeIso;
      resumo.push([nf[k], tot[k].n, tot[k].itens, numOuVazio(c.desconto_pct), numOuVazio(c.acrescimo_pct), tot[k].frete, tot[k].desp, tot[k].total, dataBR(c.data_emissao),
        c.validade || '', c.validade_ate ? (venc ? 'VENCIDA em ' + dataBR(c.validade_ate) : 'Válida até ' + dataBR(c.validade_ate)) : (c.validade ? 'Validade só em texto, sem data' : 'Sem validade'), c.pagamento || '',
        ((c.observacoes || []).find(o => /^Prazo de entrega:/.test(o)) || '').replace(/^Prazo de entrega:\s*/, ''), c.arquivo_origem || '', c.itens.some(i => i.item_oc_n !== '' && Number(i.fator_unidade) > 1) ? (c.embalagem_fechada === false ? 'Não (por peça)' : 'Sim (embalagens inteiras)') : '—']);
    });

    // ---- aba Pontos
    const pontos = [['Gravidade', 'Ponto', 'Detalhe']].concat((p.alertas || []).map(a => [a.nivel === 'erro' ? 'Erro' : 'Atenção', a.titulo, a.texto]));
    if (pontos.length === 1) pontos.push(['', 'Nenhum ponto pendente.', '']);

    // ---- aba Fora da OC
    const fora = [['Fornecedor', 'Código', 'Descrição cotada', 'Qtd. cotada', 'Un.', 'Preço (R$)', 'Fator', 'Embalagem']];
    cots.forEach((c, k) => c.itens.filter(i => i.item_oc_n === '' || i.item_oc_n === null || i.item_oc_n === undefined).forEach(i =>
      fora.push([nf[k], i.codigo || '', i.descricao_curta || i.descricao_original || '', numOuVazio(i.qtd_cotada), i.un || '', numOuVazio(i.preco_unit), numOuVazio(i.fator_unidade) || 1, i.embalagem || ''])));
    if (fora.length === 1) fora.push(['', '', 'Todas as linhas cotadas estão ligadas a um item da OC.', '', '', '', '', '']);

    return {
      nomeArquivo: 'Mapa-OC-' + String(oc.numero).replace(/[^\w-]+/g, '') + '.xlsx',
      abas: [
        { nome: 'Mapa', aoa, merges, largs, dinheiro },
        { nome: 'Fornecedores', aoa: resumo, largs: [26, 10, 18, 11, 16, 11, 13, 14, 12, 22, 24, 20, 20, 24, 24], dinheiro: [{ c: 2, de: 2, ate: resumo.length }, { c: 5, de: 2, ate: resumo.length }, { c: 6, de: 2, ate: resumo.length }, { c: 7, de: 2, ate: resumo.length }] },
        { nome: 'Pontos a confirmar', aoa: pontos, largs: [11, 44, 110] },
        { nome: 'Fora da OC', aoa: fora, largs: [24, 14, 44, 12, 6, 12, 7, 24], dinheiro: [{ c: 5, de: 2, ate: fora.length }] }
      ]
    };
  }

  /** Versão simples (sem formatação), usada só se o ExcelJS não carregar. Devolve o buffer do .xlsx. */
  function baixarPlano(p) {
    if (!raiz.XLSX) throw new Error('A biblioteca de planilhas não carregou. Confira a conexão com a internet.');
    const m = montar(p), X = raiz.XLSX, wb = X.utils.book_new();
    m.abas.forEach(a => {
      const ws = X.utils.aoa_to_sheet(a.aoa);
      ws['!cols'] = a.largs.map(w => ({ wch: w }));
      if (a.merges) ws['!merges'] = a.merges;
      (a.dinheiro || []).forEach(d => { for (let r = d.de; r <= d.ate; r++) { const cel = ws[X.utils.encode_cell({ r: r - 1, c: d.c })]; if (cel && cel.t === 'n') cel.z = '#,##0.00'; } });
      X.utils.book_append_sheet(wb, ws, a.nome);
    });
    return X.write(wb, { type: 'array', bookType: 'xlsx' });
  }

  // ------------------------------------------------------------ Excel formatado (ExcelJS)

  const EXCELJS_URL = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
  const EXCELJS_SRI = 'sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz';

  /** Carrega o ExcelJS (versão e integridade fixas) só na hora de exportar. */
  function carregarExcelJS() {
    if (raiz.ExcelJS) return Promise.resolve(raiz.ExcelJS);
    return new Promise((ok, erro) => {
      const el = raiz.document.createElement('script');
      el.src = EXCELJS_URL; el.integrity = EXCELJS_SRI; el.crossOrigin = 'anonymous';
      el.onload = () => (raiz.ExcelJS ? ok(raiz.ExcelJS) : erro(new Error('ExcelJS não carregou.')));
      el.onerror = () => erro(new Error('Não foi possível carregar o gerador de Excel formatado.'));
      raiz.document.head.appendChild(el);
    });
  }

  const COR = {
    azul: 'FF2D5F8B', azul2: 'FF3F7BA8', azulClaro: 'FFDCE8F2', azulMuito: 'FFEEF3F8', laranja: 'FFFFE9B8', cinzaLinha: 'FFD5DDDF', cinzaFundo: 'FFF4F6F8',
    verdeFundo: 'FFDFF1E6', verde: 'FF1B6F45', ambarFundo: 'FFFBECD0', ambar: 'FF8A5200', vermelhoFundo: 'FFF9DCDF', vermelho: 'FFA3262F', cinza: 'FF5D6E73', cinzaStatus: 'FFE6EBED', branco: 'FFFFFFFF', tinta: 'FF1F2D3A'
  };
  const FUNDO = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const BORDA = { top: { style: 'thin', color: { argb: COR.cinzaLinha } }, left: { style: 'thin', color: { argb: COR.cinzaLinha } }, bottom: { style: 'thin', color: { argb: COR.cinzaLinha } }, right: { style: 'thin', color: { argb: COR.cinzaLinha } } };
  const STATUS_COR = {
    'Confere': { f: COR.verdeFundo, t: COR.verde }, 'Confirmar produto': { f: COR.ambarFundo, t: COR.ambar },
    'Vínculo manual': { f: COR.azulClaro, t: COR.azul }, 'Fora da OC': { f: COR.cinzaStatus, t: COR.cinza }, 'Não cotado': { f: COR.cinzaStatus, t: COR.cinza }
  };
  const MOEDA = '#,##0.00';
  const PCT = 'General"%"';

  function cabecalho(ws, linha, de, ate, cor, corTexto) {
    for (let c = de; c <= ate; c++) {
      const cel = ws.getCell(linha, c);
      cel.fill = FUNDO(cor || COR.azul); cel.font = { bold: true, color: { argb: corTexto || COR.branco }, size: 10 };
      cel.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; cel.border = BORDA;
    }
  }

  /** Monta o workbook formatado (ExcelJS) a partir dos mesmos dados de montar(). */
  function criarWorkbook(ExcelJS, p) {
    const m = montar(p), K = C(), cots = p.cotacoes, oc = p.oc;
    const NB = COLS_FORN.length, base = 4, wb = new ExcelJS.Workbook();
    wb.creator = 'TMX · Mapa de Cotação'; wb.created = new Date();

    // ---- Mapa
    const a0 = m.abas[0], ws = wb.addWorksheet('Mapa', { views: [{ state: 'frozen', xSplit: 2, ySplit: 5, showGridLines: false }], pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    a0.aoa.forEach(l => ws.addRow(l));
    a0.largs.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    a0.merges.forEach(g => ws.mergeCells(g.s.r + 1, g.s.c + 1, g.e.r + 1, g.e.c + 1));
    const ultCol = a0.largs.length, nItens = oc.itens.length, linCab = 5, linIni = 6, linFim = linCab + nItens;
    // título e subtítulo
    for (let c = 1; c <= 13; c++) ws.getCell(1, c).fill = FUNDO(COR.azul);
    ws.getCell(1, 1).font = { bold: true, size: 15, color: { argb: COR.branco } }; ws.getCell(1, 1).alignment = { vertical: 'middle' }; ws.getRow(1).height = 28;
    ws.getCell(2, 1).font = { italic: true, size: 9, color: { argb: COR.cinza } }; ws.getCell(2, 1).alignment = { wrapText: true, vertical: 'top' }; ws.getRow(2).height = 28;
    // linha de fornecedores
    cots.forEach((c, k) => cabecalho(ws, 4, base + k * NB + 1, base + k * NB + NB, k % 2 ? COR.azul2 : COR.azul));
    ws.getRow(4).height = 22;
    // cabeçalho das colunas
    cabecalho(ws, linCab, 1, ultCol, COR.azulClaro, COR.tinta);
    ws.getRow(linCab).height = 30;
    for (let c = 1; c <= ultCol; c++) ws.getCell(linCab, c).border = Object.assign({}, BORDA, { bottom: { style: 'medium', color: { argb: COR.azul } } });
    // itens
    oc.itens.forEach((it, i) => {
      const r = linIni + i, melhor = K.menorPreco(it, cots);
      for (let c = 1; c <= ultCol; c++) { const cel = ws.getCell(r, c); cel.border = BORDA; cel.alignment = { vertical: 'top', wrapText: c === 2 }; cel.font = { size: 10, color: { argb: COR.tinta } }; }
      ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'top' }; ws.getCell(r, 2).font = { size: 10, bold: true, color: { argb: COR.tinta } };
      ws.getCell(r, 3).numFmt = 'General'; ws.getCell(r, 3).alignment = { horizontal: 'right', vertical: 'top' }; ws.getCell(r, 4).alignment = { horizontal: 'center', vertical: 'top' };
      cots.forEach((c, k) => {
        const b = base + k * NB + 1, lk = K.linhaDe(c, it.n);
        if (!lk) {
          for (let j = 0; j < NB; j++) { const cel = ws.getCell(r, b + j); cel.fill = FUNDO(COR.cinzaFundo); }
          ws.getCell(r, b + 1).font = { italic: true, size: 10, color: { argb: COR.cinza } };
          return;
        }
        ws.getCell(r, b + 1).alignment = { wrapText: true, vertical: 'top' };
        ws.getCell(r, b + 2).numFmt = MOEDA; ws.getCell(r, b + 3).numFmt = 'General'; ws.getCell(r, b + 3).alignment = { horizontal: 'center', vertical: 'top' };
        const tot = ws.getCell(r, b + 4); tot.numFmt = MOEDA; tot.font = { bold: true, size: 10, color: { argb: COR.tinta } };
        if (melhor && melhor.indice === k) { tot.fill = FUNDO(COR.laranja); tot.font = { bold: true, size: 11, color: { argb: COR.tinta } }; }
        const st = ws.getCell(r, b + 5), cor = STATUS_COR[st.value];
        if (cor) { st.fill = FUNDO(cor.f); st.font = { bold: true, size: 9, color: { argb: cor.t } }; st.alignment = { horizontal: 'center', vertical: 'top', wrapText: true }; }
        ws.getCell(r, b + 6).alignment = { wrapText: true, vertical: 'top' }; ws.getCell(r, b + 6).font = { size: 9, color: { argb: COR.cinza } };
      });
      const mp = ws.getCell(r, ultCol); mp.font = { bold: true, size: 10, color: { argb: COR.azul } }; mp.alignment = { vertical: 'top', wrapText: true };
    });
    // totais
    const fmtLinha = { 'Itens cotados': '0', 'Desconto': PCT, 'Acréscimo': PCT };
    for (let r = linFim + 2; r <= a0.aoa.length; r++) {
      const rot = String(ws.getCell(r, 2).value || ''), final = /^TOTAL/.test(rot);
      ws.getCell(r, 2).font = { bold: true, size: 10, color: { argb: final ? COR.branco : COR.tinta } };
      ws.getCell(r, 2).alignment = { horizontal: 'right' };
      for (let c = 1; c <= 4; c++) { ws.getCell(r, c).fill = FUNDO(final ? COR.azul : COR.cinzaFundo); }
      cots.forEach((c, k) => {
        const b = base + k * NB + 1;
        for (let j = 0; j < NB; j++) ws.getCell(r, b + j).fill = FUNDO(final ? COR.azul : COR.cinzaFundo);
        const cel = ws.getCell(r, b + 4), chave = Object.keys(fmtLinha).find(x => rot.indexOf(x) === 0);
        cel.numFmt = chave ? fmtLinha[chave] : MOEDA; cel.alignment = { horizontal: 'right' };
        cel.font = { bold: true, size: final ? 12 : 10, color: { argb: final ? COR.branco : COR.tinta } };
      });
      ws.getCell(r, ultCol).fill = FUNDO(final ? COR.azul : COR.cinzaFundo);
      if (final) ws.getRow(r).height = 24;
    }
    ws.autoFilter = { from: { row: linCab, column: 1 }, to: { row: linFim, column: ultCol } };

    // ---- demais abas
    const demais = (a, cabCor) => {
      const w = wb.addWorksheet(a.nome, { views: [{ state: 'frozen', ySplit: 1, showGridLines: false }] });
      a.aoa.forEach(l => w.addRow(l)); a.largs.forEach((x, i) => { w.getColumn(i + 1).width = x; });
      cabecalho(w, 1, 1, a.largs.length, cabCor || COR.azul); w.getRow(1).height = 30;
      for (let r = 2; r <= a.aoa.length; r++) for (let c = 1; c <= a.largs.length; c++) { const cel = w.getCell(r, c); cel.border = BORDA; cel.alignment = { vertical: 'top', wrapText: true }; cel.font = { size: 10, color: { argb: COR.tinta } }; }
      w.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(2, a.aoa.length), column: a.largs.length } };
      return w;
    };
    const wf = demais(m.abas[1]);
    for (let r = 2; r <= m.abas[1].aoa.length; r++) {
      [3, 6, 7, 8].forEach(c => { wf.getCell(r, c).numFmt = MOEDA; });
      [4, 5].forEach(c => { wf.getCell(r, c).numFmt = PCT; });
      wf.getCell(r, 1).font = { bold: true, size: 10, color: { argb: COR.tinta } }; wf.getCell(r, 8).font = { bold: true, size: 10, color: { argb: COR.tinta } };
      const sit = wf.getCell(r, 11), txt = String(sit.value || '');
      if (/^VENCIDA/.test(txt)) { sit.fill = FUNDO(COR.vermelhoFundo); sit.font = { bold: true, size: 10, color: { argb: COR.vermelho } }; }
      else if (/^Válida/.test(txt)) { sit.fill = FUNDO(COR.verdeFundo); sit.font = { bold: true, size: 10, color: { argb: COR.verde } }; }
      else { sit.fill = FUNDO(COR.ambarFundo); sit.font = { bold: true, size: 10, color: { argb: COR.ambar } }; }
    }
    const wp = demais(m.abas[2]);
    for (let r = 2; r <= m.abas[2].aoa.length; r++) {
      const g = wp.getCell(r, 1), erro = g.value === 'Erro';
      if (g.value) { g.fill = FUNDO(erro ? COR.vermelhoFundo : COR.ambarFundo); g.font = { bold: true, size: 10, color: { argb: erro ? COR.vermelho : COR.ambar } }; g.alignment = { horizontal: 'center', vertical: 'top' }; }
      wp.getCell(r, 2).font = { bold: true, size: 10, color: { argb: COR.tinta } };
    }
    const wo = demais(m.abas[3]);
    for (let r = 2; r <= m.abas[3].aoa.length; r++) wo.getCell(r, 6).numFmt = MOEDA;
    return wb;
  }

  /** Excel formatado; se o ExcelJS não carregar, cai no arquivo simples (sem cores). */
  async function baixar(p) {
    const m = montar(p);
    let buf, formatado = true;
    try { buf = await (criarWorkbook(await carregarExcelJS(), p)).xlsx.writeBuffer(); }
    catch (e) { formatado = false; buf = baixarPlano(p, true); }
    const url = raiz.URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = raiz.document.createElement('a');
    a.href = url; a.download = m.nomeArquivo; raiz.document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => raiz.URL.revokeObjectURL(url), 4000);
    return { nomeArquivo: m.nomeArquivo, formatado };
  }

  const api = { montar, baixar, baixarPlano, criarWorkbook, carregarExcelJS, ROTULO_STATUS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.ExportarMapa = api;
})(typeof window !== 'undefined' ? window : this);
