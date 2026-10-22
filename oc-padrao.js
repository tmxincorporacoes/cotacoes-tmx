/**
 * OC padrão em Excel: a planilha que a TMX envia ao fornecedor para ele preencher e devolver.
 *
 * Já vem com os itens e as quantidades; o fornecedor preenche só as colunas marcadas com ✎
 * (preço, produto cotado, prazo, pagamento, frete…). Cada linha leva o código do Obra Prima e
 * o número do item na OC, então, ao voltar, o leitor (leitor-planilha.js) liga cada linha ao
 * item certo da OC sem comparar descrições. O formato é lido de volta pelo mesmo leitor.
 *
 * gerarMatriz() é pura (testável); baixar() usa o SheetJS (XLSX) e o navegador.
 */
(function (raiz) {
  'use strict';

  const dataBR = iso => (/^\d{4}-\d{2}-\d{2}/.test(iso || '') ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : (iso || ''));

  /** Colunas da tabela, na ordem. "preenche" = o fornecedor escreve aqui. */
  const COLUNAS = [
    { cab: 'Nº do item na OC', larg: 10 },
    { cab: 'Código Obra Prima', larg: 16 },
    { cab: 'Descrição pedida', larg: 46 },
    { cab: 'Quantidade', larg: 11 },
    { cab: 'Unidade', larg: 9 },
    { cab: 'Produto cotado (marca/modelo) ✎', larg: 34, preenche: true },
    { cab: 'Código do fornecedor ✎', larg: 16, preenche: true },
    { cab: 'Preço unitário (R$) ✎', larg: 16, preenche: true },
    { cab: 'Total (R$)', larg: 14 },
    { cab: 'Observação ✎', larg: 28, preenche: true }
  ];
  const COL_QTD = 'D', COL_PRECO = 'H';

  /** @param oc {numero, obra, data, itens:[{n, material_id, descricao, qtd, unidade}]} */
  function gerarMatriz(oc) {
    if (!oc || !oc.itens || !oc.itens.length) throw new Error('A OC não tem itens.');
    const topo = [
      ['Pedido de cotação'],
      ['Preencha somente as colunas marcadas com ✎ e devolva este mesmo arquivo, sem mudar a ordem nem o nome das colunas.'],
      ['Se não for cotar um item, deixe o preço em branco. Se oferecer outro produto, escreva-o em "Produto cotado".'],
      [],
      ['Ordem de compra', String(oc.numero)],
      ['Obra', oc.obra || ''],
      ['OC emitida em', dataBR(oc.data)],
      [],
      ['Fornecedor ✎', ''],
      ['Data da cotação ✎', ''],
      ['Validade da cotação ✎', ''],
      ['Condição de pagamento ✎', ''],
      ['Prazo de entrega ✎', ''],
      ['Frete (R$) ✎', ''],
      ['Despesas acessórias (R$) ✎', ''],
      []
    ];
    const linhaCab = topo.length;                       // 0-based
    const aoa = topo.concat([COLUNAS.map(c => c.cab)]);
    oc.itens.forEach((it, k) => {
      const r = linhaCab + 2 + k;                       // número da linha na planilha (1-based)
      aoa.push([it.n, it.material_id || '', it.descricao, it.qtd, it.unidade || 'UN', '', '', '',
        { t: 'n', f: 'IF(' + COL_PRECO + r + '="","",' + COL_QTD + r + '*' + COL_PRECO + r + ')' }, '']);
    });
    return { aoa, linhaCab, colunas: COLUNAS, nomeArquivo: 'OC-' + String(oc.numero).replace(/[^\w-]+/g, '') + '-pedido-de-cotacao.xlsx' };
  }

  /** Gera o .xlsx e dispara o download. Devolve { nomeArquivo, itensSemCodigo }. */
  function baixar(oc) {
    if (!raiz.XLSX) throw new Error('A biblioteca de planilhas não carregou. Confira a conexão com a internet.');
    const m = gerarMatriz(oc);
    const ws = raiz.XLSX.utils.aoa_to_sheet(m.aoa);
    ws['!cols'] = m.colunas.map(c => ({ wch: c.larg }));
    const wb = raiz.XLSX.utils.book_new();
    raiz.XLSX.utils.book_append_sheet(wb, ws, 'Cotação');
    const buf = raiz.XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const url = raiz.URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = raiz.document.createElement('a');
    a.href = url; a.download = m.nomeArquivo; raiz.document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => raiz.URL.revokeObjectURL(url), 4000);
    return { nomeArquivo: m.nomeArquivo, itensSemCodigo: oc.itens.filter(i => !String(i.material_id || '').trim()).length };
  }

  const api = { gerarMatriz, baixar, COLUNAS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.OCPadrao = api;
})(typeof window !== 'undefined' ? window : this);
