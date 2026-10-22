/**
 * Curiosa House | Briefing de entrada
 * Recebe o formulário, grava na planilha, gera o resumo, manda os e-mails e serve o painel da Jami.
 *
 * Onde colar: planilha "Briefings de entrada Curiosa" > Extensões > Apps Script.
 * Arquivos do projeto: Code.gs (este) e Painel.html (Arquivo > + > HTML, nome exato "Painel").
 * Implantar > Nova implantação > App da Web | Executar como: Eu | Quem pode acessar: Qualquer pessoa.
 */

/* ============ CONFIG ============ */
const NOTIFY_EMAIL = 'eujamileferraz@outlook.com';          // quem recebe cada briefing (pode pôr mais de um, separado por vírgula)
const REPLY_TO     = 'officecuriosa@gmail.com';             // para onde o cliente responde
const PAINEL_KEY   = 'TROQUE-ESTA-CHAVE';                   // senha do painel. Troque por algo só seu, sem espaços.
const SEND_CLIENT_COPY = true;                              // manda confirmação com as respostas para o cliente
const WHATS_TXT = '(19) 98146-0819';
const SHEET_NAME = 'Briefings';

const STATUS = ['Novo', 'Reunião marcada', 'Proposta enviada', 'Fechado', 'Perdido'];
const HEADERS = ['ID', 'Recebido em', 'Status', 'Temperatura', 'Interesse', 'Nome', 'Empresa', 'Cargo', 'E-mail', 'WhatsApp',
  'Instagram', 'Site', 'Origem', 'Prazo', 'Investimento', 'Decisor', 'Resumo', 'Pontos para a reunião', 'Respostas (JSON)',
  'Campos (JSON)', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'Página', 'Notas'];

/* ============ RECEBER ============ */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const raw = (e && e.parameter && e.parameter.payload) || (e && e.postData && e.postData.contents) || '{}';
    const data = JSON.parse(raw);
    const c = data.campos || {};
    if (!c.nome || !c.email) throw new Error('Payload sem nome ou e-mail');

    const sh = getSheet_();
    const id = 'BR-' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyMMdd-HHmmss');
    const r = montarResumo(c);
    const clean = v => String(v == null ? '' : v).slice(0, 45000).replace(/^[=+\-@]/, "'$&");
    const utm = data.utm || {};

    sh.appendRow([
      id, new Date(), 'Novo', r.temperatura, c.interesse_nome, clean(c.nome), clean(c.empresa), clean(c.cargo), clean(c.email),
      clean(c.whatsapp), clean(c.instagram), clean(c.site), clean(c.origem), clean(c.prazo), clean(c.investimento), clean(c.decisor),
      r.resumo, r.pontos.join('\n'), JSON.stringify(data.respostas || []), JSON.stringify(c),
      clean(utm.utm_source), clean(utm.utm_medium), clean(utm.utm_campaign), clean(utm.utm_content), clean(data.pagina), ''
    ]);

    const painelUrl = ScriptApp.getService().getUrl() + '?k=' + encodeURIComponent(PAINEL_KEY) + '#' + id;
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        replyTo: c.email,
        subject: `${r.emoji} Briefing ${r.temperatura.toLowerCase()} · ${c.empresa || c.nome} · ${c.interesse_nome}`,
        htmlBody: emailJami_(id, c, r, data.respostas || [], painelUrl),
        name: 'Briefing Curiosa'
      });
    }
    if (SEND_CLIENT_COPY && /@/.test(c.email)) {
      MailApp.sendEmail({
        to: c.email,
        replyTo: REPLY_TO,
        subject: 'Seu briefing chegou na mesa da Curiosa',
        htmlBody: emailCliente_(c, data.respostas || []),
        name: 'Jami | Curiosa House'
      });
    }
    return json_({ ok: true, id });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ============ PAINEL ============ */
function doGet(e) {
  const k = e && e.parameter && e.parameter.k;
  if (k !== PAINEL_KEY) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px">Página privada da Curiosa House.</p>')
      .setTitle('Curiosa House');
  }
  const t = HtmlService.createTemplateFromFile('Painel');
  const safe = o => JSON.stringify(o).replace(/</g, '\\u003c');
  t.dataJson = safe(listar_());
  t.keyJson = safe(PAINEL_KEY);
  t.statusJson = safe(STATUS);
  return t.evaluate().setTitle('Briefings Curiosa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// chamadas do painel (google.script.run)
function painelAtualizar(key, id, status, notas) {
  if (key !== PAINEL_KEY) throw new Error('Chave inválida');
  const sh = getSheet_();
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
  const i = ids.indexOf(id);
  if (i < 0) throw new Error('Briefing não encontrado');
  const row = i + 2;
  if (status != null && STATUS.indexOf(status) >= 0) sh.getRange(row, HEADERS.indexOf('Status') + 1).setValue(status);
  if (notas != null) sh.getRange(row, HEADERS.indexOf('Notas') + 1).setValue(String(notas).slice(0, 5000));
  return true;
}
function painelListar(key) {
  if (key !== PAINEL_KEY) throw new Error('Chave inválida');
  return listar_();
}

function listar_() {
  const sh = getSheet_();
  if (sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  const col = n => HEADERS.indexOf(n);
  return vals.filter(v => v[0]).map(v => ({
    id: v[col('ID')],
    recebido: v[col('Recebido em')] instanceof Date ? v[col('Recebido em')].toISOString() : String(v[col('Recebido em')]),
    status: v[col('Status')], temperatura: v[col('Temperatura')], interesse: v[col('Interesse')],
    nome: v[col('Nome')], empresa: v[col('Empresa')], cargo: v[col('Cargo')], email: v[col('E-mail')], whatsapp: v[col('WhatsApp')],
    instagram: v[col('Instagram')], site: v[col('Site')], origem: v[col('Origem')],
    prazo: v[col('Prazo')], investimento: v[col('Investimento')], decisor: v[col('Decisor')],
    resumo: v[col('Resumo')], pontos: String(v[col('Pontos para a reunião')] || '').split('\n').filter(Boolean),
    respostas: safeParse_(v[col('Respostas (JSON)')], []), notas: v[col('Notas')],
    utm: v[col('utm_source')] ? [v[col('utm_source')], v[col('utm_medium')], v[col('utm_campaign')]].filter(Boolean).join(' / ') : ''
  })).reverse();
}

/* ============ RESUMO AUTOMÁTICO ============ */
function montarResumo(c) {
  // temperatura: prazo + investimento + quem decide
  const pPrazo = { 'O quanto antes': 3, 'No próximo mês': 2, 'Em 2 a 3 meses': 1, 'Ainda sem data': 0 }[c.prazo] || 0;
  const pInv = { 'Acima de R$ 40 mil': 3, 'De R$ 15 mil a R$ 40 mil': 3, 'De R$ 5 mil a R$ 15 mil': 2, 'Até R$ 5 mil': 1, 'Ainda não definimos': 0 }[c.investimento] || 0;
  const pDec = { 'Eu': 2, 'Eu e mais alguém (sócio ou diretoria)': 1, 'Um comitê ou outra área': 0 }[c.decisor] || 0;
  const score = pPrazo + pInv + pDec;
  const temperatura = score >= 6 ? 'Quente' : score >= 3 ? 'Morno' : 'Frio';
  const emoji = { Quente: '🔥', Morno: '🌤', Frio: '❄️' }[temperatura];

  const quem = [c.nome, [c.cargo, c.empresa].filter(Boolean).join(' na ')].filter(Boolean).join(', ');
  let frase = `${quem} procura ${lc_(c.interesse_nome)}.`;
  const pontos = [];

  if (c.interesse === 'evento') {
    const dias = diasAte_(c.ev_data);
    frase += ` Evento "${c.ev_nome}" (${c.ev_tipo || 'tipo não informado'}), ${fmtData_(c.ev_data)} em ${c.ev_local}, para ${lc_(c.ev_qtd)} pessoas.`;
    if (dias != null) frase += dias >= 0 ? ` Faltam ${dias} dias.` : ' A data informada já passou.';
    if (dias != null && dias >= 0 && dias < 45) pontos.push(`Evento em ${dias} dias: o pré-evento fica curto. Confirmar o que dá para entregar até lá.`);
    if (/primeira/i.test(c.ev_edicao)) pontos.push('Primeira edição: não há histórico. Perguntar de onde vem o público inicial (base, parceiros, patrocinadores).');
    if (/Nada disso/.test(c.ev_pronto)) pontos.push('Nada estruturado ainda (site, identidade, redes): o escopo começa pela base.');
    else if (c.ev_pronto && !/Verba de mídia/.test(c.ev_pronto)) pontos.push('Não marcou verba de mídia definida: alinhar se mídia paga entra no projeto.');
    if (c.ev_fases && c.ev_fases.split(',').length === 3) pontos.push('Quer as três fases (pré, durante e pós): escopo de ponta a ponta.');
  }
  if (c.interesse === '360') {
    frase += ` Momento: ${lc_(c.m_momento)}. Hoje o marketing fica com: ${lc_(c.m_quem)}. Dor principal: "${c.m_dor}".`;
    if (/Ainda não/.test(c.m_medicao)) pontos.push('Não mede resultados: abrir com diagnóstico e definição de metas e KPIs.');
    if (/Ninguém|correria/.test(c.m_quem)) pontos.push('Sem time de marketing: entender capacidade de execução do lado do cliente.');
    if (/Agência/.test(c.m_quem)) pontos.push('Já tem agência: perguntar o que não está funcionando e se a ideia é substituir ou complementar.');
    if (/poucos compram/.test(c.m_dor)) pontos.push('Leads não convertem: olhar régua de nutrição e passagem para vendas.');
  }
  if (c.interesse === 'campanha') {
    frase += ` Objetivo: ${lc_(c.c_objetivo)}. Ideia: ${lc_(c.c_ideia)}. No ar: ${c.c_periodo}.`;
    if (/conceito/.test(c.c_ideia)) pontos.push('Já tem conceito: escopo tende a ser execução, desdobramentos e mídia.');
    if (/zero/.test(c.c_ideia)) pontos.push('Começa do zero: espaço para conceito e manifesto. Pedir materiais de marca antes da proposta.');
  }
  if (c.interesse === 'gtm') {
    frase += ` Estágio: ${lc_(c.g_estagio)}, venda ${lc_(c.g_mercado)}. Lançamento: ${c.g_data}.`;
    if (/ideia/.test(c.g_estagio)) pontos.push('Produto ainda em ideia: GTM começa por pesquisa e validação com clientes.');
    if (/não decolou/.test(c.g_estagio)) pontos.push('Relançamento: investigar por que não tracionou (mensagem, público, preço, time).');
    if (!/Time comercial/.test(c.g_tem || '')) pontos.push('Sem time comercial definido: incluir sales enablement na conversa.');
  }
  if (c.interesse === 'outro') pontos.push('Pedido fora das trilhas: ler o desafio com atenção e chegar com 2 caminhos possíveis.');

  frase += ` Prazo: ${lc_(c.prazo)}. Investimento: ${lc_(c.investimento)}. Decide: ${lc_(c.decisor)}.`;

  if (/não definimos/.test(c.investimento)) pontos.push('Verba não definida: levar faixas de proposta para ancorar.');
  if (/Até R\$ 5 mil/.test(c.investimento) && c.interesse !== 'outro') pontos.push('Verba baixa para o escopo: pensar num primeiro passo menor (diagnóstico ou sprint).');
  if (/comitê/.test(c.decisor)) pontos.push('Decisão em comitê: perguntar quem mais precisa estar na próxima reunião.');
  if (/O quanto antes/.test(c.prazo)) pontos.push('Urgência alta: responder hoje e propor reunião nos próximos 2 dias.');
  if (c.sucesso) pontos.push(`Sucesso para o cliente: "${c.sucesso}". Usar essa frase na proposta.`);

  return { temperatura, emoji, score, resumo: frase.replace(/\s+/g, ' ').trim(), pontos };
}

/* ============ E-MAILS ============ */
const C = { ink: '#06060b', ink2: '#13131f', cream: '#f3ecdf', dim: '#b9b3a8', lilac: '#c9bcff', violet: '#7a5cf2', line: '#2a2833' };

function emailJami_(id, c, r, respostas, painelUrl) {
  const tempCor = { Quente: '#ff9a7a', Morno: '#f2d48a', Frio: '#9ec3ff' }[r.temperatura];
  const zap = String(c.whatsapp || '').replace(/\D/g, '');
  return `<div style="background:${C.ink};padding:28px 16px;font-family:Arial,Helvetica,sans-serif;color:${C.cream}">
  <div style="max-width:640px;margin:0 auto">
    <p style="font-size:11px;letter-spacing:3px;text-transform:uppercase;color:${C.lilac};margin:0 0 10px">Novo briefing · ${id}</p>
    <h1 style="font-family:Georgia,serif;font-weight:normal;font-size:30px;line-height:1.15;margin:0 0 6px">${h_(c.empresa || c.nome)}</h1>
    <p style="margin:0 0 18px;color:${C.dim}">${h_(c.interesse_nome)}</p>
    <span style="display:inline-block;background:${tempCor};color:${C.ink};font-weight:bold;font-size:12px;letter-spacing:1px;text-transform:uppercase;border-radius:99px;padding:6px 12px">${r.temperatura}</span>
    <div style="background:${C.ink2};border:1px solid ${C.line};border-radius:16px;padding:20px;margin:20px 0">
      <p style="font-size:11px;letter-spacing:3px;text-transform:uppercase;color:${C.lilac};margin:0 0 8px">Resumo</p>
      <p style="font-family:Georgia,serif;font-size:17px;line-height:1.55;margin:0">${h_(r.resumo)}</p>
      ${r.pontos.length ? `<p style="font-size:11px;letter-spacing:3px;text-transform:uppercase;color:${C.lilac};margin:18px 0 8px">Pontos para a reunião</p>
      <ul style="margin:0;padding-left:18px;line-height:1.6">${r.pontos.map(p => `<li style="margin-bottom:6px">${h_(p)}</li>`).join('')}</ul>` : ''}
    </div>
    <p style="margin:0 0 22px">
      <a href="${painelUrl}" style="display:inline-block;background:${C.cream};color:${C.ink};text-decoration:none;font-weight:bold;font-size:12px;letter-spacing:2px;text-transform:uppercase;border-radius:99px;padding:13px 20px;margin:0 8px 8px 0">Abrir no painel</a>
      ${zap ? `<a href="https://wa.me/55${zap.replace(/^55/, '')}" style="display:inline-block;border:1px solid ${C.line};color:${C.cream};text-decoration:none;font-weight:bold;font-size:12px;letter-spacing:2px;text-transform:uppercase;border-radius:99px;padding:12px 20px;margin:0 8px 8px 0">WhatsApp ${h_(c.whatsapp)}</a>` : ''}
    </p>
    ${tabela_(respostas, true)}
    <p style="color:${C.dim};font-size:12px;margin-top:24px">Responder este e-mail fala direto com ${h_(c.nome)} (${h_(c.email)}).</p>
  </div></div>`;
}

function emailCliente_(c, respostas) {
  const nome = String(c.nome || '').trim().split(/\s+/)[0];
  return `<div style="background:#f3ecdf;padding:28px 16px;font-family:Arial,Helvetica,sans-serif;color:#16131f">
  <div style="max-width:600px;margin:0 auto">
    <p style="font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#7a5cf2;margin:0 0 12px">Curiosa House</p>
    <h1 style="font-family:Georgia,serif;font-weight:normal;font-size:30px;line-height:1.15;margin:0 0 16px">Seu briefing chegou na mesa, ${h_(nome)}.</h1>
    <p style="font-size:16px;line-height:1.6;margin:0 0 12px">Obrigada por contar o seu desafio com esse cuidado. Eu leio tudo com calma antes da nossa conversa e, em até 1 dia útil, te chamo para marcar um horário.</p>
    <p style="font-size:16px;line-height:1.6;margin:0 0 22px">Se lembrar de algo importante, é só responder este e-mail. Se for urgente, o WhatsApp é ${WHATS_TXT}.</p>
    <p style="font-family:Georgia,serif;font-style:italic;font-size:18px;margin:0 0 26px">Jami<br><span style="font-family:Arial,sans-serif;font-style:normal;font-size:12px;letter-spacing:2px;text-transform:uppercase;color:#6b6575">Curiosa House</span></p>
    <p style="font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#7a5cf2;margin:0 0 6px">Uma cópia das suas respostas</p>
    ${tabela_(respostas, false)}
  </div></div>`;
}

function tabela_(respostas, dark) {
  const line = dark ? C.line : '#ddd3c2', dim = dark ? C.dim : '#6b6575', sec = dark ? C.lilac : '#7a5cf2';
  let html = '', last = '';
  respostas.forEach(r => {
    if (r.sec !== last) { html += `<tr><td colspan="2" style="padding:18px 0 6px;font-size:11px;letter-spacing:3px;text-transform:uppercase;color:${sec}">${h_(r.sec)}</td></tr>`; last = r.sec; }
    html += `<tr><td style="padding:9px 12px 9px 0;border-top:1px solid ${line};color:${dim};font-size:13px;vertical-align:top;width:42%">${h_(r.q)}</td>
      <td style="padding:9px 0;border-top:1px solid ${line};font-size:14px;vertical-align:top;white-space:pre-wrap">${h_(r.a || '-')}</td></tr>`;
  });
  return `<table role="presentation" style="width:100%;border-collapse:collapse">${html}</table>`;
}

/* ============ UTILS ============ */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#13131f').setFontColor('#f3ecdf');
    sh.setFrozenRows(1);
    sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build());
  }
  return sh;
}
function diasAte_(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || '')) return null;
  const [y, m, d] = iso.split('-').map(Number);
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  return Math.round((new Date(y, m - 1, d) - hoje) / 86400000);
}
function fmtData_(iso) { return /^\d{4}-\d{2}-\d{2}$/.test(iso || '') ? iso.split('-').reverse().join('/') : (iso || 'data não informada'); }
function lc_(s) { s = String(s || ''); return /^[A-ZÀ-Ú](?:[a-zà-ú]|\s)/.test(s) ? s[0].toLowerCase() + s.slice(1) : s; }
function h_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }
function safeParse_(s, d) { try { return JSON.parse(s); } catch (e) { return d; } }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ============ TESTE ============ */
// No editor: selecione "testeManual" e clique em Executar. Autorize. Deve aparecer uma linha e chegar um e-mail.
function testeManual() {
  const payload = {
    campos: { nome: 'Teste Curiosa', email: NOTIFY_EMAIL.split(',')[0], whatsapp: '(19) 99999-9999', empresa: 'Empresa Teste', cargo: 'Diretora',
      interesse: 'evento', interesse_nome: 'Divulgação de ponta a ponta para evento', prazo: 'No próximo mês', investimento: 'De R$ 15 mil a R$ 40 mil',
      decisor: 'Eu', sucesso: 'Casa cheia e 300 leads', ev_nome: 'Expo Teste', ev_data: '2027-03-10', ev_local: 'Piracicaba', ev_tipo: 'Feira ou expo',
      ev_qtd: 'De 1.000 a 5.000', ev_edicao: 'Sim, é a primeira', ev_fases: 'Pré-evento: aquecer e vender', ev_pronto: 'Base de contatos' },
    respostas: [{ sec: 'Sobre você', q: 'Nome', a: 'Teste Curiosa' }, { sec: 'Seu evento', q: 'Qual é o nome do evento?', a: 'Expo Teste' }],
    utm: {}, pagina: 'teste'
  };
  const out = doPost({ parameter: { payload: JSON.stringify(payload) } });
  Logger.log(out.getContent());
}
