/**
 * AutoFollow-PA - Office Script "ExtrairDadosEmail"
 *
 * Roda sobre a planilha triagem.xlsx hospedada no SharePoint, chamado pelo
 * Power Automate na ação "Executar script" (Excel Online Business). É a
 * triagem, e não o Follow-Up: a única tabela que este script abre é a
 * tb_triagem, e apontar a ação para o follow_up.xlsx a faria falhar por não
 * encontrar essa tabela.
 *
 * Existe por um motivo específico: o Power Automate não tem expressões
 * regulares. Toda a leitura de texto do e-mail (códigos de pedido/RMS, prazo
 * de entrega, datas) é feita aqui, em JavaScript, e o fluxo fica só com o que
 * ele faz bem - gatilho, laços, condições e conectores.
 *
 * O script NÃO altera a planilha de Follow-Up: ele só acrescenta uma linha na
 * tabela de triagem (tb_triagem). Quem decide o que entra no Follow-Up é o
 * segundo fluxo, depois da conferência.
 *
 * Parâmetros (o Power Automate preenche, nesta ordem):
 *   idMensagem, assunto, corpo (texto puro), remetente, fornecedor, recebidoEm
 *   (ISO), anexos
 *
 * O idMensagem é a chave da linha na tb_triagem: é por ele que o fluxo volta
 * depois para marcar a linha como "aplicada".
 * Retorno: JSON com o que foi identificado, para o fluxo ramificar.
 */

interface Resultado {
  pedidos: string[];
  rms: string[];
  dataPrevistaSerial: number;
  prazoDias: number;
  prazoUteis: boolean;
  prazoTipo: string; // prazo | imediato | indefinido | nao_encontrado
  prazoTrecho: string;
  dataPrevista: string; // DD/MM/YYYY
  confianca: string; // alta | media | baixa
  linhaTriagem: number;
}

const GATILHOS_FORTES: string[] = [
  "prazo de entrega",
  "entrega prevista",
  "entrega esta prevista",
  "entrega sera",
  "prazo previsto de entrega",
  "previsao de entrega",
  "tempo de entrega",
  "prazo de fabricacao",
];

const GATILHOS_FRACOS: string[] = [
  "entrega em",
  "entrega:",
  "entrega ",
  "prazo:",
  "prazo ",
  "postado",
  "expedicao",
];

const NEGATIVOS: string[] = [
  "pagamento", "pagto", "validade", "valido", "valida", "garantia",
  "faturamento", "faturado", "boleto", "vencimento", "cartorio", "credito",
  "desconto", "devolucao", "devolucoes", "cotacao", "reajuste", "multa",
  "juros", "cancelamento", "atraso", "atrasos",
];

const IMEDIATO: string[] = ["pronta entrega", "pronta-entrega", "imediat"];
const INDEFINIDO: string[] = ["a combinar", "sob consulta", "a definir", "a confirmar", "lead time"];
const PREFIXOS_DE_LUGAR: string[] = ["local de", "local da", "dados para", "dados da", "endereco"];

const JANELA = 120;

function normalizar(texto: string): string {
  return (texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function temNegativo(trecho: string): boolean {
  return NEGATIVOS.some(p => trecho.indexOf(p) >= 0);
}

/** Códigos de pedido e de RMS citados no texto. */
function extrairCodigos(texto: string): { pedidos: string[]; rms: string[] } {
  const plano = (texto || "").toUpperCase();

  const pedidos: string[] = [];
  const rms: string[] = [];

  // ASN2412274/14, PCI2608934, ASI26081748
  const padraoPedido = /\b(AS[NI]|PC[IPN])\d{6,12}(\/\d{1,3})?\b/g;
  let achado = padraoPedido.exec(plano);
  while (achado !== null) {
    if (pedidos.indexOf(achado[0]) < 0) { pedidos.push(achado[0]); }
    achado = padraoPedido.exec(plano);
  }

  // CPI235, REM485, PSP236
  const padraoRms = /\b(CPI|REM|PSP|CP|RM)\d{2,8}\b/g;
  achado = padraoRms.exec(plano);
  while (achado !== null) {
    if (rms.indexOf(achado[0]) < 0) { rms.push(achado[0]); }
    achado = padraoRms.exec(plano);
  }

  // RMS numérica: só quando a palavra "RMS" estiver por perto
  const padraoNumerico = /RMS[^0-9]{0,15}(\d{5,7})/g;
  achado = padraoNumerico.exec(plano);
  while (achado !== null) {
    if (rms.indexOf(achado[1]) < 0) { rms.push(achado[1]); }
    achado = padraoNumerico.exec(plano);
  }

  return { pedidos: pedidos, rms: rms };
}

/**
 * Data prometida no trecho ("entrega prevista para 10/11/2026", "dia 05/12").
 * É o formato típico da resposta do fornecedor a uma cobrança. Devolve a data
 * em DD/MM/AAAA ou "" quando não houver.
 */
function dataNoContexto(contexto: string, hoje: Date): string {
  // a barra é obrigatória quando não há ano: "3-5 dias úteis" é intervalo de
  // prazo, não a data 03/05
  const achado = /\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b|\b(\d{1,2})-(\d{1,2})-(\d{2,4})\b/.exec(contexto);
  if (achado === null) { return ""; }

  const resto = contexto.substring(achado.index + achado[0].length);
  if (/^\s*(dias?|semanas?|mes|meses)/.test(resto)) { return ""; }

  const comBarra = achado[1] !== undefined;
  const dia = parseInt(comBarra ? achado[1] : achado[4], 10);
  const mes = parseInt(comBarra ? achado[2] : achado[5], 10);
  const anoTexto = comBarra ? achado[3] : achado[6];
  let ano = anoTexto ? parseInt(anoTexto, 10) : hoje.getFullYear();
  if (ano < 100) { ano = ano + 2000; }

  if (dia < 1 || dia > 31 || mes < 1 || mes > 12) { return ""; }

  let data = new Date(ano, mes - 1, dia);
  if (data.getDate() !== dia || data.getMonth() !== mes - 1) { return ""; }

  // data sem ano que já passou é do ano que vem
  if (!anoTexto && data.getTime() < hoje.getTime()) {
    data = new Date(ano + 1, mes - 1, dia);
  }

  const dd = ("0" + data.getDate()).slice(-2);
  const mm = ("0" + (data.getMonth() + 1)).slice(-2);
  return dd + "/" + mm + "/" + data.getFullYear();
}

/** Quantidade de tempo do trecho, em dias. Devolve [dias, uteis] ou null. */
function quantidade(trecho: string): [number, boolean] | null {
  const padrao = /(\d{1,3})\s*(?:a|ate|-)\s*(\d{1,3})\s*(dias?|semanas?|mes|meses)|(\d{1,3})\s*(dias?|semanas?|mes|meses)/g;

  let melhor: [number, boolean] | null = null;
  let achado = padrao.exec(trecho);

  while (achado !== null) {
    let valor: number;
    let unidade: string;

    if (achado[3]) {
      valor = Math.max(parseInt(achado[1], 10), parseInt(achado[2], 10));
      unidade = achado[3];
    } else {
      valor = parseInt(achado[4], 10);
      unidade = achado[5];
    }

    let dias = valor;
    if (unidade.indexOf("semana") === 0) { dias = valor * 7; }
    if (unidade.indexOf("mes") === 0) { dias = valor * 30; }

    const depois = trecho.substring(achado.index + achado[0].length, achado.index + achado[0].length + 18);
    const uteis = depois.indexOf("util") >= 0 || depois.indexOf("uteis") >= 0;

    if (melhor === null || dias > melhor[0]) { melhor = [dias, uteis]; }

    achado = padrao.exec(trecho);
  }

  return melhor;
}

/** Procura o prazo de entrega, ancorando a leitura no gatilho. */
function extrairPrazo(texto: string, hoje: Date): { dias: number; uteis: boolean; tipo: string; trecho: string; confianca: string; data: string } {
  const linhas = (texto || "").split("\n");

  let escolhido = { dias: 0, uteis: false, tipo: "nao_encontrado", trecho: "", confianca: "baixa", data: "" };
  let achouForte = false;

  for (let i = 0; i < linhas.length; i++) {
    const base = normalizar(linhas[i]);
    if (!base) { continue; }

    const seguinte = i + 1 < linhas.length ? normalizar(linhas[i + 1]) : "";
    const gatilhos = GATILHOS_FORTES.concat(GATILHOS_FRACOS);

    for (const gatilho of gatilhos) {
      const posicao = base.indexOf(gatilho);
      if (posicao < 0) { continue; }

      const forte = GATILHOS_FORTES.indexOf(gatilho) >= 0;
      if (achouForte && !forte) { continue; }

      const antes = base.substring(Math.max(0, posicao - 40), posicao);
      if (PREFIXOS_DE_LUGAR.some(p => antes.substring(antes.length - 18).indexOf(p) >= 0)) { continue; }
      if (temNegativo(antes)) { continue; }

      const candidatos = [
        base.substring(posicao, posicao + JANELA).split(". ")[0],
        (base + " " + seguinte).substring(posicao, posicao + JANELA).split(". ")[0],
      ];

      for (const contexto of candidatos) {
        if (temNegativo(contexto)) { continue; }

        const trecho = contexto.substring(0, 150);
        const confianca = forte ? "alta" : "media";
        const dataPrometida = dataNoContexto(contexto, hoje);
        const qtd = quantidade(contexto);

        if (dataPrometida !== "") {
          escolhido = { dias: 0, uteis: false, tipo: "data", trecho: trecho, confianca: confianca, data: dataPrometida };
        } else if (qtd !== null) {
          escolhido = { dias: qtd[0], uteis: qtd[1], tipo: "prazo", trecho: trecho, confianca: confianca, data: "" };
        } else if (IMEDIATO.some(p => contexto.indexOf(p) >= 0)) {
          escolhido = { dias: 0, uteis: false, tipo: "imediato", trecho: trecho, confianca: confianca, data: "" };
        } else if (INDEFINIDO.some(p => contexto.indexOf(p) >= 0)) {
          escolhido = { dias: 0, uteis: false, tipo: "indefinido", trecho: trecho, confianca: confianca, data: "" };
        } else {
          continue;
        }

        if (forte) { achouForte = true; }
        break;
      }
    }
  }

  return escolhido;
}

function somarDiasUteis(inicio: Date, dias: number): Date {
  const atual = new Date(inicio.getTime());
  let restantes = dias;
  while (restantes > 0) {
    atual.setDate(atual.getDate() + 1);
    const diaSemana = atual.getDay();
    if (diaSemana !== 0 && diaSemana !== 6) { restantes--; }
  }
  return atual;
}

/**
 * Número de série do Excel para a data (dias desde 30/12/1899).
 *
 * É assim que a data precisa ser gravada em célula: se mandarmos o texto
 * "01/12/2026", o Excel interpreta com a cultura da pasta de trabalho e, em
 * inglês, isso vira 12 de janeiro. Número de série não tem ambiguidade.
 */
function serialExcel(data: Date): number {
  const utc = Date.UTC(data.getFullYear(), data.getMonth(), data.getDate());
  return Math.round((utc - Date.UTC(1899, 11, 30)) / 86400000);
}

function formatar(data: Date): string {
  const dia = ("0" + data.getDate()).slice(-2);
  const mes = ("0" + (data.getMonth() + 1)).slice(-2);
  return dia + "/" + mes + "/" + data.getFullYear();
}

function main(
  workbook: ExcelScript.Workbook,
  idMensagem: string,
  assunto: string,
  corpo: string,
  remetente: string,
  fornecedor: string,
  recebidoEm: string,
  anexos: string
): Resultado {
  const texto = (assunto || "") + "\n" + (corpo || "");

  const base = recebidoEm ? new Date(recebidoEm) : new Date();

  const codigos = extrairCodigos(texto);
  const prazo = extrairPrazo(texto, base);

  let prevista = "";
  let previstaData: Date = null;

  if (prazo.tipo === "data") {
    // "10/12/2026" -> dia, mês, ano
    const partes = prazo.data.split("/");
    previstaData = new Date(parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10));
  } else if (prazo.tipo === "prazo") {
    previstaData = prazo.uteis
      ? somarDiasUteis(base, prazo.dias)
      : new Date(base.getTime() + prazo.dias * 86400000);
  } else if (prazo.tipo === "imediato") {
    previstaData = somarDiasUteis(base, 2);
  }

  if (previstaData !== null) { prevista = formatar(previstaData); }
  const previstaSerial = previstaData !== null ? serialExcel(previstaData) : null;

  // confiança final: só é alta quando há prazo explícito E um código de pedido
  let confianca = prazo.confianca;
  if (prazo.tipo === "nao_encontrado") { confianca = "baixa"; }
  if (codigos.pedidos.length === 0 && codigos.rms.length === 0) { confianca = "baixa"; }
  if (codigos.pedidos.length > 1) { confianca = "baixa"; }

  // a ordem precisa ser exatamente a das 16 colunas da tb_triagem
  const tabela = workbook.getTable("tb_triagem");

  // A coluna de data guarda número de série; o formato abaixo é o que a faz
  // aparecer como data na planilha, e ordenar como data.
  //
  // setNumberFormat (e não setNumberFormatLocal) porque esta função usa sempre
  // os códigos invariantes, em inglês: d de day, m de month, y de year. O Excel
  // os traduz sozinho para o idioma de quem abre a planilha. A versão "Local"
  // espera o código no idioma da interface - em português seria "dd/mm/aaaa" -
  // e por isso quebra assim que alguém abre o arquivo no outro idioma.
  tabela.getColumnByName("Data prevista")
        .getRangeBetweenHeaderAndTotal()
        .setNumberFormat("dd/mm/yyyy");
  tabela.addRow(-1, [
    idMensagem,                                 // 1  ID da mensagem (chave)
    formatar(base),                             // 2  Recebido em
    remetente,                                  // 3  Remetente
    assunto,                                    // 4  Assunto
    fornecedor,                                 // 5  Fornecedor
    codigos.pedidos.join("; "),                 // 6  Pedido (Octa)
    codigos.rms.join("; "),                     // 7  RMS (Octa)
    prazo.tipo,                                 // 8  Tipo de prazo
    prazo.tipo === "prazo" ? prazo.dias : "",   // 9  Prazo (dias)
    prazo.tipo === "prazo" ? (prazo.uteis ? "úteis" : "corridos") : "", // 10 Dias
    previstaSerial,                             // 11 Data prevista (número de série)
    prazo.trecho,                               // 12 Trecho de origem
    confianca,                                  // 13 Confiança
    anexos,                                     // 14 Anexos
    "",                                         // 15 Pedido confirmado (operador)
    "pendente",                                 // 16 Situação
  ]);

  return {
    pedidos: codigos.pedidos,
    rms: codigos.rms,
    dataPrevistaSerial: previstaSerial,
    prazoDias: prazo.dias,
    prazoUteis: prazo.uteis,
    prazoTipo: prazo.tipo,
    prazoTrecho: prazo.trecho,
    dataPrevista: prevista,
    confianca: confianca,
    linhaTriagem: tabela.getRangeBetweenHeaderAndTotal().getRowCount(),
  };
}
