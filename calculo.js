/**
 * Cálculo do mapa de cotação (puro, sem navegador). Usado pela tela e pela exportação em Excel.
 *
 * O total de cada fornecedor é calculado sobre as quantidades da OC, não sobre o que ele cotou:
 *   custo da linha = qtd da OC × (preço cotado ÷ fator de unidade) × (1 − desconto) × (1 + acréscimo)
 * Frete e despesas acessórias entram uma vez no total do fornecedor, nunca por item.
 */
(function (raiz) {
  'use strict';

  const r2 = v => Math.round((v + Number.EPSILON) * 100) / 100;

  /** Linha da cotação ligada ao item n da OC (ou undefined). */
  const linhaDe = (cot, n) => cot.itens.find(i => Number(i.item_oc_n) === n);

  /**
   * Embalagem fechada: quando o item vem em caixa, saco ou kit (fator > 1) e a cotação não desligou a opção,
   * compra-se em embalagens inteiras: ceil(qtd ÷ fator) embalagens. Devolve null se não se aplica.
   * { caixas, fator, pecas (total comprado), sobra (peças além do pedido) }
   */
  function caixas(cot, lk, qtd) {
    const fator = Number(lk.fator_unidade) || 1;
    if (fator <= 1 || cot.embalagem_fechada === false) return null;
    const n = Math.ceil(qtd / fator - 1e-9);
    return { caixas: n, fator, pecas: r2(n * fator), sobra: r2(n * fator - qtd) };
  }

  function custo(cot, lk, qtd) {
    const fator = Number(lk.fator_unidade) || 1;
    const f = (1 - (Number(cot.desconto_pct) || 0) / 100) * (1 + (Number(cot.acrescimo_pct) || 0) / 100);
    const cx = caixas(cot, lk, qtd);
    if (cx) return r2(cx.caixas * Number(lk.preco_unit) * f);      // embalagens inteiras × preço da embalagem
    return r2(qtd * (Number(lk.preco_unit) / fator) * f);          // por peça
  }

  function totalCot(cot, oc) {
    let s = 0, n = 0;
    oc.itens.forEach(it => { const lk = linhaDe(cot, it.n); if (lk) { s += custo(cot, lk, it.qtd); n++; } });
    const desp = (cot.incluir_despesas !== false && n > 0) ? (Number(cot.despesas) || 0) : 0;
    const frete = n > 0 ? (Number(cot.frete) || 0) : 0;
    return { itens: r2(s), n, desp, frete, total: r2(s + desp + frete) };
  }

  /** Menor preço de um item entre as cotações que o cotaram: { indice, valor } ou null se empate ou menos de 2 cotações. */
  function menorPreco(item, cots) {
    const custos = cots.map(c => { const lk = linhaDe(c, item.n); return lk ? custo(c, lk, item.qtd) : null; });
    const validos = custos.filter(v => v !== null);
    if (validos.length < 2) return null;
    const min = Math.min(...validos);
    if (validos.filter(v => v === min).length > 1) return null;
    return { indice: custos.indexOf(min), valor: min };
  }

  const api = { r2, linhaDe, caixas, custo, totalCot, menorPreco };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.Calculo = api;
})(typeof window !== 'undefined' ? window : this);
