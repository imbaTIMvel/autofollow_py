# -*- coding: utf-8 -*-
"""
AutoFollow - leitura dos orçamentos anexados (tarefa local diária)

Fecha a lacuna do módulo em nuvem: o Power Automate lê o corpo do e-mail, mas
não consegue abrir o PDF anexado - e é no PDF que a maioria dos fornecedores
escreve o prazo de entrega. Nenhum conector padrão do Power Automate extrai
texto de PDF, e os que existem são premium e mandam o arquivo para servidor de
terceiro, o que não cabe aqui.

Este programa roda na rede do Instituto, lê os PDFs que o fluxo já salvou na
biblioteca do SharePoint (pela pasta sincronizada), extrai o prazo e grava o
resultado em uma planilha de saída que o Fluxo 2 aplica na triagem.

    Fluxo 1 (nuvem)  -> salva o anexo e cria a linha pendente na triagem
    este programa    -> lê o PDF e escreve em prazos_anexos.xlsx
    Fluxo 2 (nuvem)  -> aplica o prazo na linha da triagem

Cada arquivo tem um único programa que escreve nele, de propósito: a triagem é
escrita só pela nuvem, e a planilha de saída só por aqui. Sem isso, duas
escritas simultâneas no mesmo arquivo sincronizado geram cópia conflitada e
perda de dado.

Execução:
    python autofollow_anexos.py --config autofollow_anexos.json
    python autofollow_anexos.py --config ... --simular    (não grava nada)

Dependências: pandas, openpyxl, pdfplumber.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import traceback
from datetime import date, datetime
from urllib.parse import unquote

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import autofollow_core as core
import autofollow_orcamentos as ao

# ==============================
# CONFIGURAÇÃO
# ==============================

ARQUIVO_CONFIG = "autofollow_anexos.json"
ARQUIVO_LOG = "autofollow_anexos_log.txt"

ABA_SAIDA = "Prazos"
TABELA_SAIDA = "tb_prazos_anexos"

CONFIG_PADRAO = {
    "observacao": (
        "Caminhos da pasta sincronizada do SharePoint. Ajuste para a máquina "
        "em que a tarefa agendada roda."
    ),
    "triagem": r"C:\\Instituto Hardware BR\\Compras\\Follow-Up\\triagem.xlsx",
    "orcamentos": r"C:\\Instituto Hardware BR\\Compras\\Orçamentos",
    "saida": r"C:\\Instituto Hardware BR\\Compras\\Follow-Up\\prazos_anexos.xlsx",
    "fornecedores": "",
    "dias_para_desistir": 30,
}

# colunas da tb_triagem que este programa lê (nunca escreve)
COL_ID = "ID da mensagem"
COL_RECEBIDO = "Recebido em"
COL_ANEXOS = "Anexos"
COL_TIPO_PRAZO = "Tipo de prazo"
COL_SITUACAO = "Situação"

COLUNAS_SAIDA = [
    "ID da mensagem",
    "Arquivo",
    "Tipo de prazo",
    "Prazo (dias)",
    "Dias",
    "Data prevista",
    "Trecho de origem",
    "Confiança",
    "Observação",
    "Lido em",
]


def carregar_config(caminho):
    """Lê a configuração; cria o arquivo padrão quando não existir."""
    if not os.path.exists(caminho):
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(CONFIG_PADRAO, f, ensure_ascii=False, indent=2)
        raise core.ErroDeUso(
            "Criei o arquivo de configuração %s com valores de exemplo. Abra, "
            "ajuste os caminhos das pastas e rode de novo." % caminho
        )

    with open(caminho, "r", encoding="utf-8") as f:
        config = json.load(f)

    for chave, valor in CONFIG_PADRAO.items():
        config.setdefault(chave, valor)

    return config


# ==============================
# LEITURA DA TRIAGEM
# ==============================


def ler_triagem(caminho):
    """Lê a tb_triagem em modo somente leitura."""
    if not os.path.exists(caminho):
        raise core.ErroDeUso(
            "Não encontrei a planilha de triagem em %s. Confira se a pasta do "
            "SharePoint está sincronizada nesta máquina e se o caminho no "
            "arquivo de configuração está certo." % caminho
        )

    df = pd.read_excel(caminho, sheet_name=0, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]

    faltando = [c for c in (COL_ID, COL_ANEXOS, COL_TIPO_PRAZO, COL_SITUACAO)
                if c not in df.columns]
    if faltando:
        raise core.ErroDeUso(
            "A planilha de triagem está sem as colunas: %s. Confira se é o "
            "arquivo certo." % ", ".join(faltando)
        )

    return df


def linhas_pendentes(df, dias_para_desistir=30, hoje=None):
    """
    Linhas que valem a pena abrir o anexo: ainda pendentes, com anexo, e em
    que o fluxo não conseguiu identificar o prazo pelo corpo do e-mail.
    """
    hoje = hoje or date.today()
    selecionadas = []

    for _, linha in df.iterrows():
        if str(linha.get(COL_SITUACAO, "")).strip().lower() != "pendente":
            continue
        if not str(linha.get(COL_ANEXOS, "")).strip():
            continue

        tipo = str(linha.get(COL_TIPO_PRAZO, "")).strip().lower()
        if tipo not in ("", "nao_encontrado", "não_encontrado", "indefinido"):
            continue

        # linha antiga demais: o operador já deve ter tratado à mão
        recebido = core.formatar_data(linha.get(COL_RECEBIDO, ""))
        if recebido:
            try:
                dias = (hoje - datetime.strptime(recebido, "%d/%m/%Y").date()).days
                if dias > dias_para_desistir:
                    continue
            except ValueError:
                pass

        selecionadas.append(linha)

    return selecionadas


# ==============================
# LOCALIZAÇÃO DO ANEXO
# ==============================


def nomes_dos_anexos(celula):
    """Extrai os nomes de arquivo da coluna Anexos (que guarda links)."""
    nomes = []

    for parte in str(celula or "").split(";"):
        parte = parte.strip()
        if not parte:
            continue

        # a coluna guarda o link do SharePoint; o nome é o último segmento
        nome = unquote(parte.split("?")[0].rstrip("/").split("/")[-1])
        if nome.lower().endswith(".pdf"):
            nomes.append(nome)

    return nomes


def indexar_orcamentos(pasta):
    """{nome do arquivo em minúsculas: caminho completo} da pasta e subpastas."""
    indice = {}

    if not os.path.isdir(pasta):
        raise core.ErroDeUso(
            "Não encontrei a pasta de orçamentos em %s. Confira se a "
            "biblioteca do SharePoint está sincronizada nesta máquina." % pasta
        )

    for raiz, _, arquivos in os.walk(pasta):
        for nome in arquivos:
            if nome.lower().endswith(".pdf"):
                indice.setdefault(nome.lower(), os.path.join(raiz, nome))

    return indice


def localizar(nome, indice):
    """Acha o PDF no índice, tolerando diferença de acento e de caixa."""
    direto = indice.get(nome.lower())
    if direto:
        return direto

    alvo = ao.normalizar(nome)
    for chave, caminho in indice.items():
        if ao.normalizar(chave) == alvo:
            return caminho

    return None


# ==============================
# PROCESSAMENTO
# ==============================


def processar(config, hoje=None, simular=False):
    """Lê os anexos pendentes e devolve (linhas de saída, ocorrências)."""
    hoje = hoje or date.today()

    df = ler_triagem(config["triagem"])
    indice = indexar_orcamentos(config["orcamentos"])

    fornecedores = []
    if config.get("fornecedores") and os.path.exists(config["fornecedores"]):
        try:
            fornecedores = ao.carregar_fornecedores(config["fornecedores"])
        except Exception:
            fornecedores = []

    pendentes = linhas_pendentes(df, config.get("dias_para_desistir", 30), hoje)

    saida = []
    ocorrencias = []

    for linha in pendentes:
        identificador = str(linha[COL_ID]).strip()
        nomes = nomes_dos_anexos(linha.get(COL_ANEXOS, ""))

        if not nomes:
            continue

        melhor = None

        for nome in nomes:
            caminho = localizar(nome, indice)

            if not caminho:
                ocorrencias.append(
                    "Anexo não encontrado na pasta sincronizada: %s "
                    "(mensagem %s). Pode ser que a sincronização ainda não "
                    "tenha baixado o arquivo." % (nome, identificador[:40])
                )
                continue

            try:
                recebido = core.formatar_data(linha.get(COL_RECEBIDO, ""))
                base = (datetime.strptime(recebido, "%d/%m/%Y").date()
                        if recebido else hoje)
                resultado = ao.analisar_orcamento(
                    caminho, fornecedores, data_recebimento=base
                )
            except Exception as erro:
                ocorrencias.append(
                    "Não foi possível ler o arquivo %s: %s" % (nome, erro)
                )
                continue

            # entre vários anexos, fica o de maior confiança
            ordem = {"alta": 0, "media": 1, "baixa": 2, "nenhuma": 3, "": 4}
            if melhor is None or ordem.get(resultado.confianca_prazo, 4) < ordem.get(
                melhor[1].confianca_prazo, 4
            ):
                melhor = (nome, resultado)

        if melhor is None:
            continue

        nome, resultado = melhor
        prazo = resultado.prazo

        # nada de útil no PDF: não ocupa linha na saída
        if prazo.tipo == "nao_encontrado" and not resultado.observacao:
            continue

        saida.append(
            {
                "ID da mensagem": identificador,
                "Arquivo": nome,
                "Tipo de prazo": prazo.tipo,
                "Prazo (dias)": prazo.dias if prazo.tipo == "prazo" else "",
                "Dias": ("úteis" if prazo.uteis else "corridos")
                        if prazo.tipo == "prazo" else "",
                "Data prevista": resultado.data_prevista,
                "Trecho de origem": prazo.trecho,
                "Confiança": resultado.confianca_prazo,
                "Observação": resultado.observacao,
                "Lido em": hoje,
            }
        )

    return saida, ocorrencias


# ==============================
# GRAVAÇÃO DA SAÍDA
# ==============================

FORMATO_DATA = "dd/mm/yyyy"
FILL_HEADER = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

LARGURAS = {
    "ID da mensagem": 42, "Arquivo": 38, "Tipo de prazo": 15, "Prazo (dias)": 12,
    "Dias": 10, "Data prevista": 14, "Trecho de origem": 55, "Confiança": 11,
    "Observação": 45, "Lido em": 12,
}


def gravar_saida(linhas, caminho):
    """
    Reescreve prazos_anexos.xlsx com o que ainda está pendente.

    O arquivo é a caixa de saída deste programa: o Fluxo 2 lê, aplica na
    triagem, e na execução seguinte a linha some daqui sozinha, porque a
    triagem deixou de estar sem prazo. Por isso é reescrito inteiro, e não
    acrescentado.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = ABA_SAIDA

    ws.append(COLUNAS_SAIDA)

    for linha in linhas:
        ws.append([linha.get(c, "") or "" for c in COLUNAS_SAIDA])

    for celula in ws[1]:
        celula.font = Font(bold=True)
        celula.fill = FILL_HEADER
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for indice, coluna in enumerate(COLUNAS_SAIDA, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = LARGURAS.get(coluna, 18)

    # o limite da tabela é medido ANTES de aplicar o formato nas linhas de
    # folga: tocar numa célula a cria, e a tabela acabaria abrangendo
    # centenas de linhas vazias - que o fluxo depois leria uma a uma
    ultima = max(ws.max_row, 2)
    referencia = "A1:%s%d" % (get_column_letter(len(COLUNAS_SAIDA)), ultima)

    for coluna in ("Data prevista", "Lido em"):
        letra = get_column_letter(COLUNAS_SAIDA.index(coluna) + 1)
        for numero in range(2, ultima + 200):
            ws["%s%d" % (letra, numero)].number_format = FORMATO_DATA

    tabela = Table(displayName=TABELA_SAIDA, ref=referencia)
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(tabela)
    ws.freeze_panes = "A2"

    # grava ao lado e troca no fim, para que o OneDrive nunca encontre o
    # arquivo pela metade
    temporario = caminho + ".tmp"
    wb.save(temporario)
    os.replace(temporario, caminho)

    return caminho


def gravar_log(ocorrencias, resumo, pasta):
    caminho = os.path.join(pasta, ARQUIVO_LOG)

    with open(caminho, "a", encoding="utf-8") as f:
        f.write("\n=== %s ===\n" % datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        f.write(resumo + "\n")
        for ocorrencia in ocorrencias:
            f.write("  - %s\n" % ocorrencia)

    return caminho


# ==============================
# EXECUÇÃO
# ==============================


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Lê os orçamentos anexados e prepara os prazos para o fluxo."
    )
    parser.add_argument("--config", default=ARQUIVO_CONFIG,
                        help="arquivo de configuração (padrão: %s)" % ARQUIVO_CONFIG)
    parser.add_argument("--simular", action="store_true",
                        help="mostra o que faria, sem gravar nada")
    args = parser.parse_args(argv)

    try:
        config = carregar_config(args.config)
        linhas, ocorrencias = processar(config, simular=args.simular)
    except core.ErroDeUso as erro:
        print(erro)
        return 1
    except Exception as erro:
        print(
            "Não foi possível concluir a leitura dos anexos. O detalhe técnico "
            "foi salvo no log, e pode ser enviado à equipe de Processos e "
            "Automação."
        )
        pasta = os.path.dirname(os.path.abspath(args.config))
        with open(os.path.join(pasta, ARQUIVO_LOG), "a", encoding="utf-8") as f:
            f.write("\n=== falha em %s ===\n" % datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            f.write("".join(traceback.format_exception(type(erro), erro, erro.__traceback__)))
        return 1

    com_prazo = sum(1 for l in linhas if l["Tipo de prazo"] == "prazo")
    resumo = ("%d anexo(s) analisado(s): %d com prazo identificado, %d para "
              "conferência manual." % (len(linhas), com_prazo, len(linhas) - com_prazo))

    print(resumo)
    for ocorrencia in ocorrencias:
        print("  -", ocorrencia)

    if args.simular:
        print("\n(simulação: nada foi gravado)")
        for linha in linhas:
            print("  %-28s %-16s %s" % (linha["Arquivo"][:28], linha["Tipo de prazo"],
                                        linha["Data prevista"] or ""))
        return 0

    gravar_saida(linhas, config["saida"])
    gravar_log(ocorrencias, resumo, os.path.dirname(os.path.abspath(args.config)))

    print("Planilha de saída atualizada: %s" % config["saida"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
