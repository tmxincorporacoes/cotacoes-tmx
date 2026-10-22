/**
 * Camada de dados da tela: duas implementações com o MESMO contrato.
 *   criarLocal(storage)   cópia local no navegador (localStorage), para testar sem planilha
 *   criarRemoto(opcoes)   planilha do Google, via o Web App do Apps Script (apps-script/Code.gs)
 * Os testes em teste/teste-banco.html rodam as mesmas verificações nas duas.
 * Este arquivo não contém dados nem chaves: URL e token vêm da janela de conexão.
 */
(function (raiz) {
  'use strict';

  const clonar = o => JSON.parse(JSON.stringify(o));
  const rid = p => p + '_' + Math.random().toString(36).slice(2, 12);
  const hoje = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  const vazio = () => ({ Fornecedores: [], Dicionario: [], OCs: [], ItensOC: [], Cotacoes: [], ItensCotacao: [] });

  // ------------------------------------------------------------ local

  function criarLocal(storage, chave) {
    chave = chave || 'tmx_cotacoes_local_v1';
    let db;
    try { const t = storage.getItem(chave); db = t ? Object.assign(vazio(), JSON.parse(t)) : vazio(); } catch (e) { db = vazio(); }
    const gravar = () => { try { storage.setItem(chave, JSON.stringify(db)); } catch (e) { throw new Error('Não foi possível gravar na cópia local do navegador.'); } };
    const achaDic = reg => db.Dicionario.findIndex(x => x.fornecedor_id === reg.fornecedor_id &&
      (reg.codigo_fornecedor ? String(x.codigo_fornecedor) === String(reg.codigo_fornecedor) : !x.codigo_fornecedor && x.descricao_normalizada === reg.descricao_normalizada));

    const B = {
      modo: 'local',
      async listarOCs() {
        return clonar(db.OCs).map(o => Object.assign(o, {
          itens: clonar(db.ItensOC.filter(i => i.oc_id === o.id)).sort((a, b) => a.n - b.n),
          n_cotacoes: db.Cotacoes.filter(c => c.oc_id === o.id).length
        }));
      },
      async salvarOC({ oc, itens }) {
        if (!oc || !String(oc.numero || '').trim()) throw new Error('OC sem número.');
        if (!itens || !itens.length) throw new Error('OC sem itens.');
        const vistos = new Set();
        itens.forEach(i => {
          if (!(Number(i.qtd) > 0)) throw new Error('Item ' + i.n + ': quantidade inválida.');
          if (!String(i.descricao || '').trim()) throw new Error('Item ' + i.n + ': sem descrição.');
          if (vistos.has(Number(i.n))) throw new Error('Item ' + i.n + ' repetido.');
          vistos.add(Number(i.n));
        });
        if (oc.id && !db.OCs.some(o => o.id === oc.id)) throw new Error('OC não encontrada.');
        if (!oc.id && db.OCs.some(o => String(o.numero).trim() === String(oc.numero).trim())) throw new Error('Já existe uma OC com o número ' + String(oc.numero).trim() + '.');
        const id = oc.id || rid('oc');
        db.OCs = db.OCs.filter(o => o.id !== id); db.ItensOC = db.ItensOC.filter(i => i.oc_id !== id);
        const o = clonar(oc); delete o.itens; delete o.n_cotacoes;
        db.OCs.push(Object.assign({ status: 'aberta', cond_pgto: '', previsao_entrega: '', fornecedor: '', subtotal: '', desconto_total: '', total_documento: '' }, o, { id }));
        itens.forEach(i => db.ItensOC.push({
          oc_id: id, n: Number(i.n), material_id: String(i.material_id || '').trim(), descricao: String(i.descricao).trim(), qtd: Number(i.qtd),
          unidade: i.unidade || 'UN', preco_unit: i.preco_unit == null ? '' : i.preco_unit, desconto_valor: i.desconto_valor == null ? '' : i.desconto_valor
        }));
        gravar(); return { id };
      },
      async excluirOC(id) {
        if (!db.OCs.some(o => o.id === id)) throw new Error('OC não encontrada.');
        db.Cotacoes.filter(c => c.oc_id === id).forEach(c => { db.ItensCotacao = db.ItensCotacao.filter(i => i.cotacao_id !== c.id); });
        db.Cotacoes = db.Cotacoes.filter(c => c.oc_id !== id); db.ItensOC = db.ItensOC.filter(i => i.oc_id !== id); db.OCs = db.OCs.filter(o => o.id !== id);
        gravar(); return { id };
      },
      async listarFornecedores() { return clonar(db.Fornecedores); },
      async salvarFornecedor(f) {
        if (!String(f.nome || '').trim()) throw new Error('Fornecedor sem nome.');
        if (f.id && !db.Fornecedores.some(x => x.id === f.id)) throw new Error('Fornecedor não encontrado.');
        const id = f.id || rid('f'); db.Fornecedores = db.Fornecedores.filter(x => x.id !== id);
        db.Fornecedores.push({ id, nome: f.nome.trim(), regra_icms: f.regra_icms || '', contato: f.contato || '' }); gravar(); return { id };
      },
      async listarCotacoes(ocId) {
        return clonar(db.Cotacoes.filter(c => c.oc_id === ocId)).map(c => Object.assign(c, { itens: clonar(db.ItensCotacao.filter(i => i.cotacao_id === c.id)).sort((a, b) => a.linha - b.linha) }));
      },
      async salvarCotacao({ cotacao, itens }) {
        if (!db.OCs.some(o => o.id === cotacao.oc_id)) throw new Error('OC não encontrada.');
        if (!db.Fornecedores.some(f => f.id === cotacao.fornecedor_id)) throw new Error('Fornecedor não encontrado.');
        const STATUS = ['confere', 'confirmar', 'manual', 'sem_cotacao'];
        itens.forEach((i, k) => { if (STATUS.indexOf(i.status) < 0) throw new Error('Linha ' + (k + 1) + ': status inválido.'); });
        const id = cotacao.id || rid('cot');
        db.Cotacoes = db.Cotacoes.filter(c => c.id !== id); db.ItensCotacao = db.ItensCotacao.filter(i => i.cotacao_id !== id);
        const c = clonar(cotacao); delete c.itens;
        db.Cotacoes.push(Object.assign(c, { id }));
        itens.forEach((i, k) => db.ItensCotacao.push(Object.assign(clonar(i), { cotacao_id: id, linha: i.linha == null ? k : i.linha })));
        gravar(); return { id, itens: itens.length };
      },
      async excluirCotacao(id) { db.ItensCotacao = db.ItensCotacao.filter(i => i.cotacao_id !== id); db.Cotacoes = db.Cotacoes.filter(c => c.id !== id); gravar(); return { id }; },
      async lerDicionario() { return clonar(db.Dicionario); },
      async gravarDicionario(vinculos) {
        if (!vinculos.length) throw new Error('Nenhum vínculo para gravar.');
        vinculos.forEach((v, i) => {
          const r = 'Vínculo ' + (i + 1) + ': ';
          if (!String(v.fornecedor_id || '').trim() || !String(v.material_id || '').trim()) throw new Error(r + 'fornecedor ou material ausente.');
          if (!String(v.codigo_fornecedor || '').trim() && !String(v.descricao_normalizada || '').trim()) throw new Error(r + 'informe código ou descrição.');
          if (v.fator_unidade != null && v.fator_unidade !== '' && !(Number(v.fator_unidade) > 0)) throw new Error(r + 'fator de unidade inválido.');
        });
        let novos = 0, atualizados = 0;
        vinculos.forEach(v => {
          const reg = Object.assign({}, v, { data: hoje(), fator_unidade: v.fator_unidade || 1 });
          const k = achaDic(reg);
          if (k >= 0) { db.Dicionario[k] = reg; atualizados++; } else { db.Dicionario.push(reg); novos++; }
        });
        gravar(); return { novos, atualizados };
      },
      async excluirDicionario(chaveReg) {
        const k = achaDic(chaveReg);
        if (k < 0) throw new Error('Vínculo não encontrado.');
        db.Dicionario.splice(k, 1); gravar(); return { removido: true };
      },
      /** Tudo que a tela precisa para abrir: uma chamada só. */
      async carregar(ocId) {
        return { ocs: await B.listarOCs(), fornecedores: await B.listarFornecedores(), dicionario: await B.lerDicionario(), cotacoes: ocId ? await B.listarCotacoes(ocId) : [] };
      },
      async limpar() { db = vazio(); gravar(); }
    };
    return B;
  }

  // ------------------------------------------------------------ remoto (Apps Script)

  function criarRemoto({ url, token, fetchImpl, aoOcupar, limiteMs }) {
    if (!url || !token) throw new Error('Informe a URL e o token da planilha.');
    const f = fetchImpl || ((...a) => raiz.fetch(...a));
    async function chamar(action, dados) {
      if (aoOcupar) aoOcupar(1);
      const ctl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const t = ctl ? setTimeout(() => ctl.abort(), limiteMs || 60000) : null;
      try {
        // text/plain evita a checagem prévia de CORS, que o Apps Script não responde
        const r = await f(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ token, action }, dados || {})), signal: ctl ? ctl.signal : undefined });
        if (!r.ok) throw new Error('Falha de rede (' + r.status + ').');
        let j;
        try { j = await r.json(); } catch (e) { throw new Error('Resposta inválida do servidor. Confira a URL e se a implantação está como "Qualquer pessoa".'); }
        if (!j.ok) throw new Error(j.erro || 'Erro no servidor.');
        return j.data;
      } catch (e) {
        if (e && e.name === 'AbortError') throw new Error('O servidor demorou demais para responder. Tente de novo.');
        if (e instanceof TypeError) throw new Error('Não foi possível falar com a planilha. Confira a internet e a URL.');
        throw e;
      } finally { if (t) clearTimeout(t); if (aoOcupar) aoOcupar(-1); }
    }
    return {
      modo: 'planilha',
      listarOCs: () => chamar('listarOCs'),
      salvarOC: ({ oc, itens }) => chamar('salvarOC', { oc, itens }),
      excluirOC: id => chamar('excluirOC', { id }),
      listarFornecedores: () => chamar('listarFornecedores'),
      salvarFornecedor: fo => chamar('salvarFornecedor', fo),
      listarCotacoes: ocId => chamar('listarCotacoes', { oc_id: ocId }),
      salvarCotacao: ({ cotacao, itens }) => chamar('salvarCotacao', { cotacao, itens }),
      excluirCotacao: id => chamar('excluirCotacao', { id }),
      lerDicionario: () => chamar('lerDicionario'),
      gravarDicionario: vinculos => chamar('gravarDicionario', { vinculos }),
      excluirDicionario: chave => chamar('excluirDicionario', chave),
      carregar: ocId => chamar('carregar', { oc_id: ocId || '' }),
      testar: () => chamar('listarFornecedores'),
      criarAbas: () => chamar('setup')
    };
  }

  const api = { criarLocal, criarRemoto };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.TMXBanco = api;
})(typeof window !== 'undefined' ? window : this);
