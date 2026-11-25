/**
 * AutoFollow-PA - Office Script "MarcarAplicada"
 *
 * Substitui a ação "Excel Online - Atualizar uma linha" que marcava a linha da
 * triagem como aplicada no fim do Fluxo 1.
 *
 * Por que existe: a ação do conector procura a linha por chave através de um
 * cache (a resposta vem com o cabeçalho x-ms-apihub-cached-response: true).
 * Quando a linha acabou de ser criada pelo script ExtrairDadosEmail, segundos
 * antes, o cache ainda não a contém e a ação devolve 404 - "Nenhuma linha foi
 * encontrada com a ID". A política de repetição não resolve: o Power Automate
 * só repete 408, 429 e 5xx, nunca 404.
 *
 * Este script não passa pelo conector: lê a tabela viva pela API do Excel, a
 * mesma via que criou a linha. Não há cache no meio.
 *
 * Parâmetros (o Power Automate preenche, nesta ordem):
 *   idMensagem - o internetMessageId, chave da linha na tb_triagem
 *   situacao   - o texto a gravar; se vier vazio, grava "aplicada"
 *
 * Se a linha não for encontrada, o script lança erro de propósito, para que o
 * escopo falhe e o tratamento de erro que já existe no fluxo dispare o alerta.
 * Uma linha de Follow-Up atualizada sem a triagem correspondente marcada é
 * justamente o estado que não pode passar em silêncio: ela seria reprocessada.
 */

interface ResultadoMarcacao {
  linha: number;      // número da linha na planilha, para conferência
  situacao: string;   // o que ficou gravado
}

function main(
  workbook: ExcelScript.Workbook,
  idMensagem: string,
  situacao: string
): ResultadoMarcacao {
  const alvo = (idMensagem || "").trim();
  if (alvo === "") {
    throw new Error("MarcarAplicada: o parâmetro idMensagem veio vazio.");
  }

  const valor = (situacao || "").trim() === "" ? "aplicada" : situacao.trim();

  const tabela = workbook.getTable("tb_triagem");
  if (!tabela) {
    throw new Error("MarcarAplicada: a tabela tb_triagem não existe nesta pasta de trabalho.");
  }

  // As colunas são localizadas pelo nome do cabeçalho, nunca por posição fixa:
  // acrescentar ou reordenar uma coluna da tb_triagem não deve quebrar o script.
  const cabecalhos = tabela.getHeaderRowRange().getValues()[0];
  let colId = -1;
  let colSituacao = -1;
  for (let i = 0; i < cabecalhos.length; i++) {
    const nome = String(cabecalhos[i]).trim();
    if (nome === "ID da mensagem") { colId = i; }
    if (nome === "Situação") { colSituacao = i; }
  }
  if (colId < 0) {
    throw new Error("MarcarAplicada: a coluna 'ID da mensagem' não existe na tb_triagem.");
  }
  if (colSituacao < 0) {
    throw new Error("MarcarAplicada: a coluna 'Situação' não existe na tb_triagem.");
  }

  const corpo = tabela.getRangeBetweenHeaderAndTotal();
  if (!corpo) {
    throw new Error("MarcarAplicada: a tb_triagem está vazia - a linha não chegou a ser criada.");
  }
  const valores = corpo.getValues();

  // De baixo para cima: a linha procurada é quase sempre a última acrescentada.
  // Se um dia houver duas linhas com o mesmo ID, a mais recente é a correta.
  for (let linha = valores.length - 1; linha >= 0; linha--) {
    if (String(valores[linha][colId]).trim() === alvo) {
      corpo.getCell(linha, colSituacao).setValue(valor);
      return { linha: linha + 2, situacao: valor };
    }
  }

  throw new Error(
    "MarcarAplicada: nenhuma linha da tb_triagem tem o ID da mensagem " + alvo +
    ". O Follow-Up pode ter sido atualizado sem a triagem ser marcada - confira antes de reenviar."
  );
}
