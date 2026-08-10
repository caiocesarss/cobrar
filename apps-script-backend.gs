/**
 * BACKEND - Google Apps Script
 * ============================
 * Como instalar:
 * 1. Abra sua planilha no Google Sheets.
 * 2. Extensoes > Apps Script.
 * 3. Apague o conteudo padrao e cole este arquivo inteiro.
 * 4. Ajuste NOME_DA_ABA abaixo se sua aba nao se chamar "Sheet1".
 * 5. Clique em "Implantar" (Deploy) > "Nova implantacao" (New deployment).
 *    - Tipo: "Aplicativo da Web" (Web app)
 *    - Executar como: "Eu" (seu usuario)
 *    - Quem tem acesso: "Qualquer pessoa" (Anyone) - necessario para o
 *      GitHub Pages conseguir chamar essa API.
 * 6. Copie a URL gerada (termina em /exec) - ela vai no arquivo index.html.
 * 7. Toda vez que editar este script, gere uma NOVA implantacao (ou
 *    "Gerenciar implantacoes" > editar > nova versao) para as mudancas
 *    valerem na URL publicada.
 *
 * SENHA DE ACESSO (nao fica no codigo, fica nas Propriedades do Script):
 * 1. No editor do Apps Script, clique no icone de engrenagem
 *    "Configuracoes do projeto" na barra lateral esquerda.
 * 2. Va ate "Propriedades do script" > "Adicionar propriedade do script".
 * 3. Propriedade: PASSWORD   Valor: (escolha uma senha sua)
 * 4. Salve. Essa senha NUNCA aparece no codigo nem no GitHub.
 */

const NOME_DA_ABA = 'Sheet1'; // ajuste para o nome real da sua aba

// Indices de coluna (0 = A, 1 = B, 2 = C ...)
const COL = {
  DATA: 0,        // A - data do mes
  VALOR_PARC: 1,  // B - valor parcela
  INDICE: 2,      // C - indice CUB
  CUB: 3,         // D - valor CUB do mes
  CORRIGIDO: 4,   // E - valor corrigido (formula existente na planilha)
  TRANSFERIDO: 5, // F - valor de cada pagamento/PIX
  SALDO: 6,       // G - saldo acumulado (formula existente na planilha)
  DATA_PIX: 7,    // H - data do pagamento
  NOME: 8,        // I - nome de quem pagou
  TIPO: 9         // J - tipo (PIX, TED...)
};

const MESES_PT = ["Janeiro","Fevereiro","Marco","Abril","Maio","Junho","Julho",
                   "Agosto","Setembro","Outubro","Novembro","Dezembro"];
const MESES_ABREV = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_DA_ABA);
}

/** Compara a senha recebida com a guardada nas Propriedades do Script. */
function senhaValida_(senhaRecebida) {
  const senhaCorreta = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!senhaCorreta) return false; // se voce nao configurou, bloqueia tudo por seguranca
  return senhaRecebida === senhaCorreta;
}

/** Extrai {dia, mes, ano} de um texto tipo "28-Aug-26" ou "10-Aug-26". */
function parseDataExibida_(texto) {
  if (!texto) return null;
  const m = String(texto).match(/(\d{1,2})[\/\-](\w+)[\/\-](\d{2,4})/);
  if (!m) return null;
  const dia = parseInt(m[1], 10);
  const abrev = m[2].toLowerCase().substring(0, 3);
  const mes = MESES_ABREV[abrev] || parseInt(m[2], 10);
  let ano = parseInt(m[3], 10);
  if (ano < 100) ano += 2000;
  return { dia, mes, ano };
}

function pad2_(n) {
  return String(n).padStart(2, '0');
}

/** Considera "linha de mes" tanto uma Date real quanto um texto no padrao dd-Mon-aa. */
function ehLinhaDeMes_(valorCelula, textoExibido) {
  if (valorCelula instanceof Date) return true;
  if (textoExibido && /^\d{1,2}[\/\-]\w+[\/\-]\d{2,4}$/.test(String(textoExibido).trim())) return true;
  return false;
}

/**
 * Localiza o bloco do mes mais recente e retorna os dados necessarios
 * para exibir o status atual e gerar o texto de cobranca.
 */
function calcularStatusAtual_() {
  const sheet = getSheet_();
  const dataRange = sheet.getDataRange();
  const valores = dataRange.getValues();
  const exibidos = dataRange.getDisplayValues();

  let linhaMes = -1;
  for (let i = valores.length - 1; i >= 0; i--) {
    if (ehLinhaDeMes_(valores[i][COL.DATA], exibidos[i][COL.DATA])) {
      linhaMes = i;
      break;
    }
  }
  if (linhaMes === -1) {
    return { erro: 'Nenhuma linha de mes encontrada na planilha.' };
  }

  const dataMesInfo = parseDataExibida_(exibidos[linhaMes][COL.DATA]);
  const idx = valores[linhaMes][COL.INDICE];
  const cub = valores[linhaMes][COL.CUB];
  const saldoAnterior = linhaMes > 0 ? Math.abs(valores[linhaMes - 1][COL.SALDO]) : 0;

  let proximaLinhaMes = valores.length;
  for (let i = linhaMes + 1; i < valores.length; i++) {
    if (ehLinhaDeMes_(valores[i][COL.DATA], exibidos[i][COL.DATA])) {
      proximaLinhaMes = i;
      break;
    }
  }

  const pagamentos = [];
  let totalPago = 0;
  for (let i = linhaMes + 1; i < proximaLinhaMes; i++) {
    const valor = valores[i][COL.TRANSFERIDO];
    if (valor && typeof valor === 'number' && valor > 0) {
      const dInfo = parseDataExibida_(exibidos[i][COL.DATA_PIX]);
      pagamentos.push({
        linha: i + 1,
        valor: valor,
        dataTexto: dInfo ? (pad2_(dInfo.dia) + '/' + pad2_(dInfo.mes)) : '',
        nome: valores[i][COL.NOME] || ''
      });
      totalPago += valor;
    }
  }

  const mult = idx * cub;
  const totalMes = mult + saldoAnterior;
  const aPagar = totalMes - totalPago;

  // Somas de coluna inteira (todas as linhas da planilha, nao so do mes atual)
  let totalCorrigidoGeral = 0;
  let totalTransferidoGeral = 0;
  for (let i = 0; i < valores.length; i++) {
    const corrigido = valores[i][COL.CORRIGIDO];
    if (typeof corrigido === 'number') totalCorrigidoGeral += corrigido;
    const transferido = valores[i][COL.TRANSFERIDO];
    if (typeof transferido === 'number') totalTransferidoGeral += transferido;
  }

  return {
    linhaMes: linhaMes + 1,
    proximaLinhaLivre: proximaLinhaMes + 1,
    mesNome: dataMesInfo ? MESES_PT[dataMesInfo.mes - 1] : '',
    ano: dataMesInfo ? dataMesInfo.ano : '',
    indice: idx,
    cub: cub,
    saldoAnterior: saldoAnterior,
    mult: mult,
    totalMes: totalMes,
    pagamentos: pagamentos,
    totalPago: totalPago,
    aPagar: aPagar,
    totalCorrigidoGeral: totalCorrigidoGeral,
    totalTransferidoGeral: totalTransferidoGeral
  };
}

/**
 * POST /exec -> recebe JSON como texto puro (evita CORS preflight):
 * { action: "getStatus" }
 * { action: "addPagamento", valor: 65.00, data: "2026-08-10", nome: "Fulano" }
 * { action: "addMes", data: "2026-09-28", valorParcela: 10000, indice: 3.6301, cub: 3200.50 }
 *
 * Nota de seguranca: a senha so trafega dentro do corpo (body) do POST,
 * nunca como parametro de URL (?senha=...). Query strings ficam gravadas
 * no historico do navegador e nos logs de execucao do Apps Script, entao
 * o doGet antigo foi removido de proposito - nao reintroduza a senha
 * na URL.
 */
function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (!senhaValida_(body.senha || '')) {
    return ContentService.createTextOutput(JSON.stringify({ erro: 'Senha incorreta.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === 'getStatus') {
    const status = calcularStatusAtual_();
    return ContentService.createTextOutput(JSON.stringify(status))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getSheet_();
  const ultimaLinha = sheet.getLastRow();

  if (body.action === 'addPagamento') {
    const novaLinha = ultimaLinha + 1;
    sheet.getRange(novaLinha, COL.TRANSFERIDO + 1).setValue(Number(body.valor));
    sheet.getRange(novaLinha, COL.DATA_PIX + 1).setValue(new Date(body.data + 'T12:00:00'));
    if (body.nome) sheet.getRange(novaLinha, COL.NOME + 1).setValue(body.nome);
    sheet.getRange(novaLinha, COL.TIPO + 1).setValue('PIX');
  } else if (body.action === 'addMes') {
    const novaLinha = ultimaLinha + 1;
    sheet.getRange(novaLinha, COL.DATA + 1).setValue(new Date(body.data + 'T12:00:00'));
    sheet.getRange(novaLinha, COL.VALOR_PARC + 1).setValue(Number(body.valorParcela));
    sheet.getRange(novaLinha, COL.INDICE + 1).setValue(Number(body.indice));
    sheet.getRange(novaLinha, COL.CUB + 1).setValue(Number(body.cub));
  } else {
    return ContentService.createTextOutput(JSON.stringify({ erro: 'acao invalida' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const statusAtualizado = calcularStatusAtual_();
  return ContentService.createTextOutput(JSON.stringify({ sucesso: true, status: statusAtualizado }))
    .setMimeType(ContentService.MimeType.JSON);
}