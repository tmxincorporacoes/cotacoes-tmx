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
          K.custo(c, lk, it.qtd), ROTULO_STATUS[lk.status] || lk.status, [lk.nota, lk.embalagem ? 'Embalagem: ' + lk.embalagem : ''].filter(Boolean).join(' | '));
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
    const resumo = [['Fornecedor', 'Itens cotados', 'Subtotal dos itens (R$)', 'Desconto (%)', 'Acréscimo / ICMS (%)', 'Frete (R$)', 'Despesas (R$)', 'Total (R$)', 'Emissão', 'Validade', 'Situação da validade', 'Pagamento', 'Prazo de entrega', 'Arquivo']];
    cots.forEach((c, k) => {
      const venc = c.validade_ate && c.validade_ate < hojeIso;
      resumo.push([nf[k], tot[k].n, tot[k].itens, numOuVazio(c.desconto_pct), numOuVazio(c.acrescimo_pct), tot[k].frete, tot[k].desp, tot[k].total, dataBR(c.data_emissao),
        c.validade || '', c.validade ? (venc ? 'VENCIDA em ' + dataBR(c.validade_ate) : 'Válida até ' + dataBR(c.validade_ate)) : 'Sem validade', c.pagamento || '',
        ((c.observacoes || []).find(o => /^Prazo de entrega:/.test(o)) || '').replace(/^Prazo de entrega:\s*/, ''), c.arquivo_origem || '']);
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
        { nome: 'Fornecedores', aoa: resumo, largs: [26, 10, 18, 11, 16, 11, 13, 14, 12, 22, 24, 20, 20, 24], dinheiro: [{ c: 2, de: 2, ate: resumo.length }, { c: 5, de: 2, ate: resumo.length }, { c: 6, de: 2, ate: resumo.length }, { c: 7, de: 2, ate: resumo.length }] },
        { nome: 'Pontos a confirmar', aoa: pontos, largs: [11, 44, 110] },
        { nome: 'Fora da OC', aoa: fora, largs: [24, 14, 44, 12, 6, 12, 7, 24], dinheiro: [{ c: 5, de: 2, ate: fora.length }] }
      ]
    };
  }

  /** Gera o .xlsx e dispara o download. */
  function baixar(p) {
    if (!raiz.XLSX) throw new Error('A biblioteca de planilhas não carregou. Confira a conexão com a internet.');
    const m = montar(p), X = raiz.XLSX, wb = X.utils.book_new();
    m.abas.forEach(a => {
      const ws = X.utils.aoa_to_sheet(a.aoa);
      ws['!cols'] = a.largs.map(w => ({ wch: w }));
      if (a.merges) ws['!merges'] = a.merges;
      (a.dinheiro || []).forEach(d => { for (let r = d.de; r <= d.ate; r++) { const cel = ws[X.utils.encode_cell({ r: r - 1, c: d.c })]; if (cel && cel.t === 'n') cel.z = '#,##0.00'; } });
      X.utils.book_append_sheet(wb, ws, a.nome);
    });
    const buf = X.write(wb, { type: 'array', bookType: 'xlsx' });
    const url = raiz.URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = raiz.document.createElement('a');
    a.href = url; a.download = m.nomeArquivo; raiz.document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => raiz.URL.revokeObjectURL(url), 4000);
    return { nomeArquivo: m.nomeArquivo };
  }

  const api = { montar, baixar, ROTULO_STATUS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.ExportarMapa = api;
})(typeof window !== 'undefined' ? window : this);
