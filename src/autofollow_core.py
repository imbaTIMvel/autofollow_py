# -*- coding: utf-8 -*-
"""
AutoFollow - núcleo de processamento (AutoFollow-Py)

Associação automática dos relatórios do Octalink (pedido de compras, produto e
RMS) em uma única planilha de Follow-Up, conforme o documento
"PLANEJAMENTO DE AUTOMAÇÃO - ASSOCIAÇÃO AUTOMÁTICA DE RELATÓRIOS DO OCTALINK".

Este módulo não depende de PySide6: toda a interação com o usuário é feita pelo
autofollow.py. Assim, o processamento pode ser testado isoladamente.

Dependências: pandas, openpyxl (a comparação de textos usa difflib, da
biblioteca padrão, para não acrescentar dependências ao pacote final).
"""

from __future__ import annotations

import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ==============================
# CONSTANTES DA PLANILHA DE SAÍDA
# ==============================

ABA_DADOS = "Follow-Up"
ABA_LISTAS = "aba_listas"
TABELA = "tb_follow_up"
ARQUIVO_LOGS = "error_logs.txt"
ARQUIVO_CONFIG = "tipos_relatorio.json"

COL_RECEBIMENTO = "Data real de recebimento na unidade de destino"
COL_PREV_ATUAL = "Previsão de entrega atualizada"
COL_PREV_INICIAL = "Previsão de entrega inicial"
COL_APROVACAO = "Data de aprovação RMS"
COL_FORNECEDOR = "Fornecedor"
COL_CONTATO = "Contato do fornecedor"
COL_DESCRICAO = "Descrição da RMS (+ itens)"
COL_REQUISITANTE = "Requisitante"
COL_DESTINO = "Destino (Brasília / Campinas / Manaus)"
COL_PROJETO = "Projeto"
COL_COMPRADOR = "Comprador responsável"
COL_PEDIDO = "Pedido (Octa)"
COL_RMS = "RMS (Octa)"
COL_STATUS = "Status"
COL_OBS = "OBS"
COL_FOLLOWUP = "Último follow-up"
COL_EXCECOES = "Exceções do algoritmo de análise"

COLUNAS_SAIDA = [
    COL_RECEBIMENTO,
    COL_PREV_ATUAL,
    COL_PREV_INICIAL,
    COL_APROVACAO,
    COL_FORNECEDOR,
    COL_CONTATO,
    COL_DESCRICAO,
    COL_REQUISITANTE,
    COL_DESTINO,
    COL_PROJETO,
    COL_COMPRADOR,
    COL_PEDIDO,
    COL_RMS,
    COL_STATUS,
    COL_OBS,
    COL_FOLLOWUP,
    COL_EXCECOES,
]

# Colunas de data: normalizadas para DD/MM/YYYY
COLUNAS_DATA = [COL_RECEBIMENTO, COL_APROVACAO, COL_PREV_INICIAL, COL_PREV_ATUAL]

# Colunas preenchidas pelo operador ou pelo AutoFollow-PA: nunca sobrescritas
# quando há mesclagem com uma versão anterior da planilha.
COLUNAS_PRESERVADAS = [
    COL_PREV_ATUAL,
    COL_STATUS,
    COL_OBS,
    COL_FOLLOWUP,
]

# Colunas montadas por concatenação, inclusive entre relatórios diferentes
# (é o caso da descrição, que soma o título da RMS e os itens).
COLUNAS_CONCATENADAS = [COL_DESCRICAO]

STATUS_OPCOES = [
    "1 - Em cotação",
    "2 - Pedido enviado",
    "3 - Pagamento programado",
    "4 - Pagamento realizado",
    "5 - Em produção",
    "6 - Em trânsito",
    "7 - Em desembaraço",
    "8 - Entregue",
]

DESTINOS = ["Brasília", "Campinas", "Manaus"]

# Comentários gravados pelo AutoFollow-PA (fluxo em nuvem) começam com este
# prefixo e são preservados na coluna de exceções a cada nova execução.
PREFIXO_PA = "Verificar manualmente:"

TIPOS_RELATORIO = ["compras", "produtos", "rms", "rms_itens"]
TIPO_FORNECEDORES = "fornecedores"
TIPO_ANTERIOR = "follow_up_anterior"

# ==============================
# CONFIGURAÇÃO (tipos_relatorio.json)
# ==============================

CONFIG_PADRAO = {
    "versao": 2,
    "observacao": (
        "Assinaturas de cabeçalho e associação de colunas do AutoFollow. "
        "Edite este arquivo quando os relatórios do Octalink mudarem de "
        "cabeçalho - não é necessário alterar o programa. Todas as abas de "
        "cada arquivo são lidas: cada aba é classificada pela própria "
        "assinatura."
    ),
    "similaridade_minima_coluna": 0.9,
    "pontuacao_minima_padrao": 0.6,
    "margem_minima": 0.1,
    "janela_busca_linhas": 7,
    "janela_busca_colunas": 7,
    "precedencia": ["produtos", "compras", "rms", "rms_itens"],
    # precedência específica de algumas colunas (quando mais de um relatório
    # traz o mesmo campo preenchido)
    "precedencia_por_coluna": {
        COL_REQUISITANTE: ["rms", "produtos", "compras"],
        # razão social (compras/RMS) antes do nome curto do relatório de
        # produtos, porque é ela que casa com a planilha de contatos
        COL_FORNECEDOR: ["compras", "rms", "produtos"],
    },
    "colunas_verificar_divergencia": [
        COL_PREV_INICIAL,
        COL_APROVACAO,
        COL_DESTINO,
        COL_PROJETO,
    ],
    # datas só são consideradas divergentes quando a diferença passa disso
    "tolerancia_dias_divergencia": 15,
    # limite de itens concatenados na descrição (0 = sem limite)
    "max_itens_descricao": 15,
    "tipos": {
        "compras": {
            "nome_exibido": "Relatório de pedido de compras",
            "colunas_obrigatorias": [
                "COD PEDIDO",
                "COD CONTROLE",
                "NOME",
                "DESCRIÇÃO",
                "DT EMISSÃO",
            ],
            "colunas_opcionais": [
                "OPERAÇÃO",
                "STATUS",
                "REQUISITANTE",
                "CNPJ/CPF",
                "QUANTIDADE",
                "VALOR BRL",
                "TIPO",
                "PRIORIDADE",
            ],
            "chave": {"cod_pedido": ["COD PEDIDO"], "cod_controle": ["COD CONTROLE"]},
            "mapa": {
                COL_APROVACAO: ["DT EMISSÃO"],
                COL_FORNECEDOR: ["NOME"],
                COL_DESCRICAO: ["DESCRIÇÃO"],
                COL_REQUISITANTE: ["REQUISITANTE"],
                COL_COMPRADOR: ["REQUISITANTE"],
                COL_PEDIDO: ["COD PEDIDO"],
                COL_RMS: ["COD CONTROLE"],
            },
        },
        "produtos": {
            "nome_exibido": "Relatório do pedido de compras - produto/serviço",
            "colunas_obrigatorias": [
                "CÓDIGO",
                "CONTROLE",
                "PARTICIPANTE",
                "TIPO",
                "OBJETO",
                "LOCAL DE ENTREGA",
            ],
            "colunas_opcionais": [
                "SAÍDA",
                "EMISSÃO",
                "ENTREGA",
                "COD. ORÇAMENTO",
                "USUÁRIO",
                "CRIADO POR",
                "STATUS",
                "STATUS APROVAÇÃO",
                "VALOR",
                "SALDO",
            ],
            "chave": {"cod_pedido": ["CÓDIGO"], "cod_controle": ["CONTROLE"]},
            "mapa": {
                COL_PREV_INICIAL: ["ENTREGA"],
                COL_APROVACAO: ["SAÍDA"],
                COL_FORNECEDOR: ["PARTICIPANTE"],
                COL_DESCRICAO: ["OBJETO"],
                COL_REQUISITANTE: ["CRIADO POR"],
                COL_DESTINO: ["LOCAL DE ENTREGA"],
                COL_COMPRADOR: ["USUÁRIO"],
                COL_PEDIDO: ["CÓDIGO"],
                COL_RMS: ["CONTROLE"],
            },
        },
        "rms": {
            "nome_exibido": "Relatório de RMS",
            "abas_auxiliares": ["rms_itens"],
            "colunas_obrigatorias": [
                "Chave",
                "Título",
                "Razão social",
                "Requisitante",
                "Criação",
            ],
            "colunas_opcionais": [
                "Código",
                "Vencimento",
                "Destino",
                "CPF/CNPJ",
                "Segmento",
                "Cidade",
                "UF",
                "Rubrica",
                "Categoria",
                "Prioridade",
                "Objeto",
                "Responsável",
                "Aprovador",
                "Status",
                "Entrega",
                "Local de entrega",
                "Valor",
            ],
            "chave": {"cod_pedido": [], "cod_controle": ["Chave", "Código"]},
            # mantém apenas as RMS que viram pedido de compra; as de pagamento
            # (diárias, reembolsos) não têm entrega a acompanhar
            "filtro_linhas": {
                "coluna": "Destino",
                "manter": ["Pedido de Compra"],
                "motivo": "RMS de pagamento (sem entrega a acompanhar)",
            },
            "mapa": {
                COL_PREV_INICIAL: ["Entrega"],
                COL_APROVACAO: ["Criação"],
                COL_FORNECEDOR: ["Razão social"],
                COL_DESCRICAO: ["Título"],
                COL_REQUISITANTE: ["Requisitante"],
                COL_COMPRADOR: ["Responsável"],
                COL_DESTINO: ["Local de entrega"],
                COL_RMS: ["Chave", "Código"],
            },
        },
        "rms_itens": {
            "nome_exibido": "Itens do RMS (aba do relatório de RMS)",
            "auxiliar": True,
            "colunas_obrigatorias": ["Código", "Item", "Projeto", "Centro de Custo"],
            "colunas_opcionais": [
                "Título",
                "Operação",
                "Porcentagem",
                "Quantidade",
                "Valor Unitário Rateado",
            ],
            "chave": {"cod_pedido": [], "cod_controle": ["Código"]},
            "mapa": {
                COL_PROJETO: ["Projeto"],
                COL_DESCRICAO: ["Item"],
            },
        },
        TIPO_FORNECEDORES: {
            "nome_exibido": "Planilha de contatos de fornecedores (opcional)",
            "colunas_obrigatorias": ["Fornecedor", "E-mail de contato"],
            "colunas_opcionais": ["CNPJ", "Razão social", "E-mail", "Email"],
            "mapa_contatos": {
                "razao_social": ["Fornecedor", "Razão social"],
                "email": ["E-mail de contato", "E-mail", "Email"],
            },
        },
        TIPO_ANTERIOR: {
            "nome_exibido": "Planilha de Follow-Up anterior (opcional)",
            "colunas_obrigatorias": [COL_PEDIDO, COL_RMS, COL_FORNECEDOR, COL_STATUS],
            "colunas_opcionais": COLUNAS_SAIDA,
        },
    },
}


class ErroDeUso(Exception):
    """
    Erro cuja mensagem já está escrita em linguagem simples e pode ser exibida
    ao usuário exatamente como está (ao contrário das exceções técnicas).
    """


def resource_path(relative_path):
    """Caminho compatível com execução via PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def caminho_config():
    """A configuração fica ao lado do executável, para poder ser editada."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, ARQUIVO_CONFIG)


def carregar_config(caminho=None):
    """Lê tipos_relatorio.json; cria o arquivo padrão quando não existir."""
    caminho = caminho or caminho_config()

    if not os.path.exists(caminho):
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(CONFIG_PADRAO, f, ensure_ascii=False, indent=2)
        return json.loads(json.dumps(CONFIG_PADRAO))

    with open(caminho, "r", encoding="utf-8") as f:
        config = json.load(f)

    # completa chaves ausentes com o padrão, para tolerar edições parciais
    for chave, valor in CONFIG_PADRAO.items():
        config.setdefault(chave, valor)

    for tipo, definicao in CONFIG_PADRAO["tipos"].items():
        config["tipos"].setdefault(tipo, definicao)

    return config


def tipos_principais(config):
    """Tipos que o operador pode escolher na tela (exclui abas auxiliares)."""
    return [
        (tipo, definicao.get("nome_exibido", tipo))
        for tipo, definicao in config["tipos"].items()
        if not definicao.get("auxiliar")
    ]


# ==============================
# FUNÇÕES AUXILIARES DE TEXTO
# ==============================


def normalizar_texto(texto):
    """minúsculas, sem acentos e sem espaços nas extremidades."""
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = texto.replace("º", "o").replace("°", "o")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def similaridade(a, b):
    """0.0 a 1.0 entre dois textos já normalizados."""
    a, b = normalizar_texto(a), normalizar_texto(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    return SequenceMatcher(None, a, b).ratio()


def normalizar_chave(valor):
    """Códigos comparáveis: sem separadores, sem zeros à esquerda, maiúsculas."""
    texto = normalizar_texto(valor).upper()
    texto = re.sub(r"[^A-Z0-9]", "", texto)
    texto = texto.lstrip("0")
    return texto


def formatar_data(valor):
    """Qualquer data reconhecível vira DD/MM/YYYY; o que não for data volta como texto."""
    if valor is None:
        return ""

    texto = str(valor).strip()
    if texto == "" or normalizar_texto(texto) in ("nan", "nat", "none"):
        return ""

    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")

    for formato in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto[:10], formato).strftime("%d/%m/%Y")
        except ValueError:
            continue

    try:
        convertido = pd.to_datetime(texto, dayfirst=True, errors="raise")
        return convertido.strftime("%d/%m/%Y")
    except Exception:
        return texto


def diferenca_em_dias(data_a, data_b):
    """Diferença absoluta em dias entre duas datas DD/MM/YYYY (None se não forem datas)."""
    try:
        a = datetime.strptime(str(data_a)[:10], "%d/%m/%Y")
        b = datetime.strptime(str(data_b)[:10], "%d/%m/%Y")
    except (ValueError, TypeError):
        return None
    return abs((a - b).days)


def normalizar_para_comparacao(coluna, valor):
    """Normaliza o valor antes de comparar duas fontes (evita falsa divergência)."""
    if coluna in COLUNAS_DATA:
        return formatar_data(valor)

    texto = str(valor).strip()

    # valores numéricos: 1.234,56 e 1234.56 são o mesmo número
    if re.fullmatch(r"-?[\d.,]+", texto or ""):
        numero = texto.replace(" ", "")
        if "," in numero and "." in numero:
            numero = numero.replace(".", "").replace(",", ".")
        elif "," in numero:
            numero = numero.replace(",", ".")
        try:
            return "%.4f" % float(numero)
        except ValueError:
            pass

    return normalizar_texto(texto)


def extrair_destino(valor):
    """
    Procura Brasília / Campinas / Manaus no local de entrega.

    O campo costuma vir como um bloco de endereço de várias linhas, por isso a
    busca é por palavra-chave dentro do texto inteiro.
    """
    texto = normalizar_texto(valor)
    for destino in DESTINOS:
        if normalizar_texto(destino) in texto:
            return destino
    return ""


def sem_codigo_item(texto):
    """Remove o código que antecede a descrição do item ("09.0002 - Passagem")."""
    return re.sub(r"^[\d][\d.\-/]*\s*[-–]\s*", "", str(texto or "").strip())


def juntar_textos(valores, limite_similaridade=0.9, maximo=0):
    """Concatena textos distintos com ' | ', descartando repetições e quase-repetições."""
    partes = []
    comparaveis = []

    for valor in valores:
        texto = str(valor or "").strip()
        if not texto:
            continue

        # a comparação ignora o código do item, para não repetir a mesma
        # descrição duas vezes (uma do pedido, outra do item da RMS)
        chave = sem_codigo_item(texto)
        if any(similaridade(chave, anterior) >= limite_similaridade for anterior in comparaveis):
            continue

        partes.append(texto)
        comparaveis.append(chave)

    if maximo and len(partes) > maximo:
        restantes = len(partes) - maximo
        partes = partes[:maximo] + ["(+%d itens)" % restantes]

    return " | ".join(partes)


# ==============================
# OCORRÊNCIAS (erros e avisos)
# ==============================


@dataclass
class Ocorrencia:
    """Uma incorrência localizável: planilha, aba, linha e coluna."""

    planilha: str
    linha: str
    coluna: str
    mensagem: str
    aba: str = ""
    pedido: str = ""
    controle: str = ""
    gravidade: str = "aviso"  # "aviso" ou "erro"
    momento: str = field(default_factory=lambda: datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

    def origem(self):
        if self.aba:
            return "%s (aba %s)" % (self.planilha, self.aba)
        return self.planilha

    def linha_log(self):
        return " | ".join(
            [
                self.momento,
                self.gravidade.upper(),
                self.origem(),
                "linha %s" % (self.linha or "-"),
                "coluna %s" % (self.coluna or "-"),
                "pedido %s" % (self.pedido or "-"),
                "RMS %s" % (self.controle or "-"),
                self.mensagem,
            ]
        )


@dataclass
class Resultado:
    """Devolvido ao final do processamento, para a interface exibir."""

    df: pd.DataFrame = None
    ocorrencias: list = field(default_factory=list)
    chaves_novas: set = field(default_factory=set)
    chaves_com_excecao: set = field(default_factory=set)
    abas_lidas: list = field(default_factory=list)
    linhas_ignoradas: int = 0
    total_processados: int = 0
    total_novos: int = 0
    total_verificar: int = 0
    contatos_carregados: bool = False


# ==============================
# LEITURA DAS PLANILHAS
# ==============================


def motor_leitura(caminho):
    return "xlrd" if os.path.splitext(caminho)[1].lower() == ".xls" else "openpyxl"


def listar_abas(caminho):
    """Nomes de todas as abas do arquivo, na ordem em que aparecem."""
    return pd.ExcelFile(caminho, engine=motor_leitura(caminho)).sheet_names


def ler_excel_bruto(caminho, aba=0):
    """Lê a aba inteira como texto, sem assumir onde está o cabeçalho."""
    return pd.read_excel(
        caminho,
        sheet_name=aba,
        header=None,
        dtype=str,
        keep_default_na=False,
        engine=motor_leitura(caminho),
    )


def pontuar_linha(valores_linha, obrigatorias, limite):
    """Proporção das colunas obrigatórias encontradas na linha informada."""
    if not obrigatorias:
        return 0.0

    encontradas = 0
    for esperada in obrigatorias:
        for valor in valores_linha:
            if similaridade(valor, esperada) >= limite:
                encontradas += 1
                break

    return encontradas / len(obrigatorias)


def nomes_conhecidos(definicao):
    """Todos os nomes de coluna cadastrados para o tipo (obrigatórios, opcionais e mapeados)."""
    nomes = list(definicao.get("colunas_obrigatorias", []))
    nomes += list(definicao.get("colunas_opcionais", []))

    for lista in definicao.get("mapa", {}).values():
        nomes += list(lista)

    for lista in definicao.get("chave", {}).values():
        nomes += list(lista)

    for lista in definicao.get("mapa_contatos", {}).values():
        nomes += list(lista)

    filtro = definicao.get("filtro_linhas")
    if filtro:
        nomes.append(filtro["coluna"])

    return list(dict.fromkeys(nomes))


def localizar_cabecalho(bruto, obrigatorias, conhecidas, limite, linhas, colunas):
    """
    Varre as primeiras `linhas` x `colunas` células procurando os nomes de
    coluna cadastrados para o tipo.

    A planilha exportada pode começar em qualquer célula (linhas de título,
    logotipo, colunas em branco à esquerda) e a ordem das colunas pode mudar -
    por isso a busca é por âncora: basta encontrar, dentro da janela, uma
    célula cujo texto coincida com um nome de coluna conhecido.

    - `obrigatorias` decide se a linha é mesmo um cabeçalho e dá a pontuação;
    - `conhecidas` (obrigatórias + opcionais) decide onde a tabela começa, para
      não descartar uma coluna opcional que esteja à esquerda da primeira
      obrigatória quando a ordem das colunas mudar.

    Devolve (pontuacao, linha, coluna_inicial) com índices 0-based, ou
    (0.0, None, 0) quando nada é encontrado.
    """
    if not obrigatorias:
        return 0.0, None, 0

    melhor = (0.0, None, 0)

    for indice_linha in range(min(linhas, len(bruto))):
        valores = [str(v) for v in bruto.iloc[indice_linha].tolist()]

        ancoras_obrigatorias = []
        ancoras_conhecidas = []

        for indice_coluna in range(min(colunas, len(valores))):
            celula = valores[indice_coluna].strip()
            if not celula:
                continue

            if any(similaridade(celula, nome) >= limite for nome in obrigatorias):
                ancoras_obrigatorias.append(indice_coluna)
                ancoras_conhecidas.append(indice_coluna)
            elif any(similaridade(celula, nome) >= limite for nome in conhecidas):
                ancoras_conhecidas.append(indice_coluna)

        # sem nenhuma coluna obrigatória na janela, não é o cabeçalho procurado
        if not ancoras_obrigatorias:
            continue

        # a primeira coluna conhecida da linha marca onde a tabela começa
        coluna_inicial = min(ancoras_conhecidas)
        pontuacao = pontuar_linha(valores[coluna_inicial:], obrigatorias, limite)

        if pontuacao > melhor[0]:
            melhor = (pontuacao, indice_linha, coluna_inicial)

    return melhor


def janela_busca(config):
    """Tamanho da janela de varredura (padrão: 7 linhas x 7 colunas)."""
    return (
        config.get("janela_busca_linhas", 7),
        config.get("janela_busca_colunas", 7),
    )


def pontuar_aba(bruto, definicao, config):
    """(pontuacao, linha, coluna) da aba contra a assinatura de um tipo."""
    limite = config.get("similaridade_minima_coluna", 0.9)
    linhas, colunas = janela_busca(config)

    return localizar_cabecalho(
        bruto,
        definicao.get("colunas_obrigatorias", []),
        nomes_conhecidos(definicao),
        limite,
        linhas,
        colunas,
    )


def detectar_tipo(caminho, config):
    """
    Identifica o tipo do arquivo pela assinatura de cabeçalho de suas abas.

    Nunca usa o nome do arquivo. Um arquivo pode ter várias abas (o relatório
    de produtos/serviços do Octalink tem seis); o tipo do arquivo é o da aba
    que melhor pontuar. Abas auxiliares (como "Itens do RMS") não definem o
    tipo do arquivo sozinhas.

    Devolve (tipo, linha_do_cabecalho, coluna_inicial, pontuacao) da melhor
    aba; (None, None, 0, pontuacao) quando não há confiança suficiente.
    """
    margem = config.get("margem_minima", 0.1)
    padrao = config.get("pontuacao_minima_padrao", 0.6)

    resultados = []

    for tipo, definicao in config["tipos"].items():
        if definicao.get("auxiliar"):
            continue

        melhor = (0.0, None, 0)

        for aba in listar_abas(caminho):
            bruto = ler_excel_bruto(caminho, aba)
            pontuacao, linha, coluna = pontuar_aba(bruto, definicao, config)
            if pontuacao > melhor[0]:
                melhor = (pontuacao, linha, coluna)

        resultados.append(
            (tipo, melhor[1], melhor[2], melhor[0], definicao.get("pontuacao_minima", padrao))
        )

    resultados.sort(key=lambda r: r[3], reverse=True)

    primeiro = resultados[0]
    segundo = resultados[1] if len(resultados) > 1 else (None, None, 0, 0.0, 0.0)

    if primeiro[3] < primeiro[4]:
        return None, None, 0, primeiro[3]

    if primeiro[3] - segundo[3] < margem:
        return None, None, 0, primeiro[3]

    return primeiro[0], primeiro[1], primeiro[2], primeiro[3]


def classificar_abas(caminho, tipo, config):
    """
    Classifica cada aba do arquivo entre o tipo escolhido e suas abas
    auxiliares.

    Devolve [(aba, tipo_da_aba, linha_cabecalho, coluna_inicial)], ignorando as
    abas que não correspondem a nenhuma assinatura (capas, resumos etc.).
    """
    definicao = config["tipos"][tipo]
    candidatos = [tipo] + list(definicao.get("abas_auxiliares", []))
    padrao = config.get("pontuacao_minima_padrao", 0.6)

    classificadas = []

    for aba in listar_abas(caminho):
        bruto = ler_excel_bruto(caminho, aba)

        melhor = None
        for candidato in candidatos:
            definicao_candidata = config["tipos"][candidato]
            pontuacao, linha, coluna = pontuar_aba(bruto, definicao_candidata, config)
            minimo = definicao_candidata.get("pontuacao_minima", padrao)

            if pontuacao >= minimo and (melhor is None or pontuacao > melhor[0]):
                melhor = (pontuacao, candidato, linha, coluna)

        if melhor:
            classificadas.append((aba, melhor[1], melhor[2], melhor[3]))

    return classificadas


def montar_tabela(bruto, linha_cabecalho, coluna_inicial):
    """Recorta o DataFrame a partir da célula do cabeçalho."""
    bruto = bruto.iloc[:, coluna_inicial:]

    nomes = [str(c).strip() for c in bruto.iloc[linha_cabecalho].tolist()]
    df = bruto.iloc[linha_cabecalho + 1:].reset_index(drop=True)
    df.columns = nomes

    df = df.loc[:, [c for c in df.columns if c and not c.lower().startswith("unnamed")]]
    df = df[~(df.apply(lambda linha: all(str(v).strip() == "" for v in linha), axis=1))]

    return df


def ler_tabela(caminho, tipo, config, aba=None, linha_cabecalho=None, coluna_inicial=0):
    """
    Lê a tabela de uma aba, a partir do cabeçalho localizado.

    Devolve (DataFrame, número da linha do cabeçalho na planilha, 1-based).
    """
    if aba is None:
        aba = listar_abas(caminho)[0]

    bruto = ler_excel_bruto(caminho, aba)

    if linha_cabecalho is None:
        _, linha_cabecalho, coluna_inicial = pontuar_aba(bruto, config["tipos"][tipo], config)

        if linha_cabecalho is None:
            linhas, colunas = janela_busca(config)
            raise ErroDeUso(
                "Não foi possível localizar o cabeçalho em %s. Verifique se o "
                "arquivo é mesmo o relatório esperado e se a tabela começa nas "
                "primeiras %d linhas e %d colunas da planilha."
                % (os.path.basename(caminho), linhas, colunas)
            )

    return montar_tabela(bruto, linha_cabecalho, coluna_inicial), linha_cabecalho + 1


def mapear_colunas(df, nomes_config, limite):
    """Associa os nomes previstos na configuração às colunas reais do arquivo."""
    encontradas = []

    for esperada in nomes_config:
        for real in df.columns:
            if similaridade(real, esperada) >= limite:
                encontradas.append(real)
                break

    return encontradas


def primeiro_valor(linha, colunas):
    for coluna in colunas:
        valor = str(linha.get(coluna, "")).strip()
        if valor and normalizar_texto(valor) not in ("nan", "none"):
            return valor
    return ""


def valores_da_linha(linha, colunas):
    valores = []
    for coluna in colunas:
        valor = str(linha.get(coluna, "")).strip()
        if valor and normalizar_texto(valor) not in ("nan", "none"):
            valores.append(valor)
    return valores


# ==============================
# CONSOLIDAÇÃO
# ==============================


class Consolidador:
    """Junta os relatórios em registros únicos por (pedido, RMS)."""

    def __init__(self, config):
        self.config = config
        self.limite = config.get("similaridade_minima_coluna", 0.9)
        self.registros = {}      # chave -> {coluna: {tipo: [valores]}}
        self.origens = {}        # chave -> {coluna: {tipo: (origem, linha, coluna_real)}}
        self.ordem = []          # preserva a ordem de entrada
        self.por_controle = {}   # código de controle -> chave
        self.ocorrencias = []
        self.linhas_ignoradas = 0

    # ------------------------------------------------------------------
    def registrar(self, ocorrencia):
        self.ocorrencias.append(ocorrencia)

    def resolver_chave(self, cod_pedido, cod_controle):
        """
        Mesma regra do fluxo de controle: pedido + controle identificam o
        registro; sem pedido (caso do relatório de RMS), vale só o controle.
        """
        if cod_pedido:
            chave = (cod_pedido, cod_controle)
            if cod_controle:
                self.por_controle.setdefault(cod_controle, chave)
            return chave

        if cod_controle and cod_controle in self.por_controle:
            return self.por_controle[cod_controle]

        chave = ("", cod_controle)
        if cod_controle:
            self.por_controle.setdefault(cod_controle, chave)
        return chave

    # ------------------------------------------------------------------
    def adicionar_relatorio(self, tipo, caminho, df, linha_cabecalho, aba=""):
        definicao = self.config["tipos"][tipo]
        nome_arquivo = os.path.basename(caminho)

        cols_pedido = mapear_colunas(df, definicao.get("chave", {}).get("cod_pedido", []), self.limite)
        cols_controle = mapear_colunas(df, definicao.get("chave", {}).get("cod_controle", []), self.limite)

        mapa_real = {}
        for coluna_saida, nomes in definicao.get("mapa", {}).items():
            reais = mapear_colunas(df, nomes, self.limite)
            if reais:
                mapa_real[coluna_saida] = reais

        # filtro de linhas (ex.: manter só as RMS de pedido de compra)
        filtro = definicao.get("filtro_linhas")
        col_filtro = ""
        if filtro:
            encontradas = mapear_colunas(df, [filtro["coluna"]], self.limite)
            col_filtro = encontradas[0] if encontradas else ""

        ignoradas_aqui = 0

        for posicao, (_, linha) in enumerate(df.iterrows()):
            numero_linha = linha_cabecalho + posicao + 1

            if col_filtro:
                valor_filtro = normalizar_texto(linha.get(col_filtro, ""))
                permitidos = [normalizar_texto(v) for v in filtro.get("manter", [])]
                if permitidos and valor_filtro not in permitidos:
                    ignoradas_aqui += 1
                    continue

            bruto_pedido = primeiro_valor(linha, cols_pedido)
            bruto_controle = primeiro_valor(linha, cols_controle)

            cod_pedido = normalizar_chave(bruto_pedido)
            cod_controle = normalizar_chave(bruto_controle)

            if not cod_pedido and not cod_controle:
                self.registrar(
                    Ocorrencia(
                        planilha=nome_arquivo,
                        aba=aba,
                        linha=str(numero_linha),
                        coluna=" / ".join(cols_pedido + cols_controle) or "-",
                        mensagem=(
                            "Esta linha está sem o código do pedido e sem o código de "
                            "controle, então não foi possível saber a qual pedido ela "
                            "pertence. A linha foi ignorada. Confira essa linha no "
                            "Octalink e gere o relatório novamente."
                        ),
                        gravidade="erro",
                    )
                )
                continue

            chave = self.resolver_chave(cod_pedido, cod_controle)

            if chave not in self.registros:
                self.registros[chave] = {}
                self.origens[chave] = {}
                self.ordem.append(chave)

            for coluna_saida, colunas_reais in mapa_real.items():
                if coluna_saida in COLUNAS_CONCATENADAS:
                    valores = valores_da_linha(linha, colunas_reais)
                else:
                    valor = primeiro_valor(linha, colunas_reais)
                    valores = [valor] if valor else []

                if coluna_saida == COL_DESTINO:
                    convertidos = []
                    for valor in valores:
                        destino = extrair_destino(valor)
                        if destino:
                            convertidos.append(destino)
                        else:
                            self.registrar(
                                Ocorrencia(
                                    planilha=nome_arquivo,
                                    aba=aba,
                                    linha=str(numero_linha),
                                    coluna=colunas_reais[0],
                                    mensagem=(
                                        "Não foi possível identificar a unidade de destino "
                                        "(Brasília, Campinas ou Manaus) no local de entrega. "
                                        "O campo ficou vazio."
                                    ),
                                    pedido=bruto_pedido,
                                    controle=bruto_controle,
                                )
                            )
                    valores = convertidos

                if coluna_saida in COLUNAS_DATA:
                    valores = [formatar_data(v) for v in valores]

                valores = [v for v in valores if v]
                if not valores:
                    continue

                destino_lista = self.registros[chave].setdefault(coluna_saida, {}).setdefault(tipo, [])

                for valor in valores:
                    if coluna_saida in COLUNAS_CONCATENADAS:
                        # várias linhas do mesmo pedido = vários itens
                        if valor not in destino_lista:
                            destino_lista.append(valor)
                    elif not destino_lista:
                        # demais colunas: vale o primeiro valor não vazio
                        destino_lista.append(valor)

                self.origens[chave].setdefault(coluna_saida, {}).setdefault(
                    tipo, (nome_arquivo if not aba else "%s (aba %s)" % (nome_arquivo, aba),
                           numero_linha, colunas_reais[0])
                )

            # o código apresentado ao usuário é o do relatório, não o normalizado
            if bruto_pedido:
                self.registros[chave].setdefault(COL_PEDIDO, {}).setdefault(tipo, [bruto_pedido])
            if bruto_controle:
                self.registros[chave].setdefault(COL_RMS, {}).setdefault(tipo, [bruto_controle])

        if ignoradas_aqui:
            self.linhas_ignoradas += ignoradas_aqui
            self.registrar(
                Ocorrencia(
                    planilha=nome_arquivo,
                    aba=aba,
                    linha="-",
                    coluna=col_filtro,
                    mensagem=(
                        "%d linha(s) não entraram na planilha de Follow-Up: %s."
                        % (ignoradas_aqui, filtro.get("motivo", "fora do filtro configurado"))
                    ),
                )
            )

    # ------------------------------------------------------------------
    def consolidar(self):
        """Aplica a precedência e registra as divergências entre relatórios."""
        precedencia_geral = self.config.get("precedencia", TIPOS_RELATORIO)
        precedencia_coluna = self.config.get("precedencia_por_coluna", {})
        verificar = self.config.get(
            "colunas_verificar_divergencia", CONFIG_PADRAO["colunas_verificar_divergencia"]
        )
        maximo_itens = self.config.get("max_itens_descricao", 0)
        tolerancia = self.config.get("tolerancia_dias_divergencia", 0)
        limite = self.limite

        linhas = []
        chaves_com_excecao = set()

        for chave in self.ordem:
            registro = self.registros[chave]
            saida = {coluna: "" for coluna in COLUNAS_SAIDA}
            excecoes = []

            def ordem_fontes(coluna):
                fontes = list(precedencia_coluna.get(coluna, precedencia_geral))
                # tipos fora da lista entram ao final, para nunca se perder nada
                for tipo in registro.get(coluna, {}):
                    if tipo not in fontes:
                        fontes.append(tipo)
                return fontes

            def escolher(coluna):
                candidatos = registro.get(coluna, {})
                if not candidatos:
                    return ""

                if coluna in COLUNAS_CONCATENADAS:
                    valores = []
                    for tipo in ordem_fontes(coluna):
                        valores.extend(candidatos.get(tipo, []))
                    return juntar_textos(valores, limite, maximo_itens)

                for tipo in ordem_fontes(coluna):
                    if candidatos.get(tipo):
                        return candidatos[tipo][0]
                return ""

            pedido_exibido = escolher(COL_PEDIDO)
            rms_exibido = escolher(COL_RMS)

            for coluna in COLUNAS_SAIDA:
                candidatos = registro.get(coluna, {})
                if not candidatos:
                    continue

                saida[coluna] = escolher(coluna)

                if coluna not in verificar or len(candidatos) < 2:
                    continue

                # agrupa os valores equivalentes: datas normalizadas e textos
                # com grau de similaridade acima do limite contam como iguais
                grupos = []  # [(valor_normalizado, tipo, valor_original)]
                for tipo, valores in candidatos.items():
                    if not valores:
                        continue
                    valor = valores[0]
                    normalizado = normalizar_para_comparacao(coluna, valor)
                    igual = False
                    for existente in grupos:
                        if coluna in COLUNAS_DATA:
                            dias = diferenca_em_dias(existente[0], normalizado)
                            # datas próximas são a mesma previsão em estágios
                            # diferentes do processo, não uma divergência
                            igual = existente[0] == normalizado or (
                                dias is not None and dias <= tolerancia
                            )
                        else:
                            igual = similaridade(existente[0], normalizado) >= limite
                        if igual:
                            break
                    if not igual:
                        grupos.append((normalizado, tipo, valor))

                if len(grupos) < 2:
                    continue

                detalhes = []
                for _, tipo, valor in grupos:
                    origem = self.origens[chave].get(coluna, {}).get(tipo)
                    nome = origem[0] if origem else tipo
                    detalhes.append("%s = %s" % (nome, valor))

                sufixo = ""
                if coluna in COLUNAS_DATA and len(grupos) == 2:
                    dias = diferenca_em_dias(grupos[0][0], grupos[1][0])
                    if dias is not None:
                        sufixo = " Diferença de %d dias." % dias

                excecoes.append(
                    "Verificar: %s está diferente entre os relatórios (%s).%s"
                    % (coluna, "; ".join(detalhes), sufixo)
                )

                for _, tipo, valor in grupos:
                    origem = self.origens[chave].get(coluna, {}).get(tipo)
                    self.registrar(
                        Ocorrencia(
                            planilha=origem[0] if origem else "-",
                            linha=str(origem[1]) if origem else "-",
                            coluna=origem[2] if origem else coluna,
                            mensagem=(
                                'Informação divergente na coluna "%s": este relatório '
                                'traz "%s". Verifique qual valor está correto.'
                                % (coluna, valor)
                            ),
                            pedido=pedido_exibido,
                            controle=rms_exibido,
                        )
                    )

            if excecoes:
                chaves_com_excecao.add(chave)
                saida[COL_EXCECOES] = "\n".join(excecoes)

            saida["_chave"] = chave
            linhas.append(saida)

        df = pd.DataFrame(linhas, columns=COLUNAS_SAIDA + ["_chave"])
        return df, chaves_com_excecao


# ==============================
# CONTATOS DOS FORNECEDORES
# ==============================


def carregar_contatos(caminho, config):
    """Lê fornecedores.xlsx e devolve [(razão social, e-mail), ...]."""
    definicao = config["tipos"][TIPO_FORNECEDORES]
    limite = config.get("similaridade_minima_coluna", 0.9)

    contatos = []

    for aba, tipo, linha, coluna in classificar_abas(caminho, TIPO_FORNECEDORES, config):
        df, _ = ler_tabela(caminho, tipo, config, aba, linha, coluna)

        mapa = definicao.get("mapa_contatos", {})
        cols_nome = mapear_colunas(df, mapa.get("razao_social", []), limite)
        cols_email = mapear_colunas(df, mapa.get("email", []), limite)

        for _, registro in df.iterrows():
            nome = primeiro_valor(registro, cols_nome)
            email = primeiro_valor(registro, cols_email)
            if nome and email:
                contatos.append((nome, email))

    return contatos


def aplicar_contatos(df, contatos, limite=0.9):
    """
    Preenche "Contato do fornecedor" cruzando a razão social.

    Havendo mais de um contato para o mesmo fornecedor, os endereços são
    gravados na mesma célula, separados por ponto e vírgula.
    """
    sem_contato = []

    for indice, linha in df.iterrows():
        fornecedor = str(linha.get(COL_FORNECEDOR, "")).strip()
        if not fornecedor:
            continue

        encontrados = [
            email for nome, email in contatos if similaridade(nome, fornecedor) >= limite
        ]

        unicos = list(dict.fromkeys(encontrados))

        if unicos:
            df.at[indice, COL_CONTATO] = "; ".join(unicos)
        else:
            sem_contato.append(indice)

    return sem_contato


# ==============================
# MESCLAGEM COM A VERSÃO ANTERIOR
# ==============================


def ler_follow_up_anterior(caminho, config):
    """Lê a planilha de Follow-Up de uma execução anterior."""
    classificadas = classificar_abas(caminho, TIPO_ANTERIOR, config)

    if not classificadas:
        raise ErroDeUso(
            "Não foi possível ler a planilha de Follow-Up anterior (%s). "
            "Verifique se o arquivo tem as colunas do Follow-Up no cabeçalho."
            % os.path.basename(caminho)
        )

    aba, tipo, linha, coluna = classificadas[0]
    df, _ = ler_tabela(caminho, tipo, config, aba, linha, coluna)

    limite = config.get("similaridade_minima_coluna", 0.9)
    renomear = {}
    for coluna_padrao in COLUNAS_SAIDA:
        for real in df.columns:
            if similaridade(real, coluna_padrao) >= limite:
                renomear[real] = coluna_padrao
                break

    df = df.rename(columns=renomear)

    for coluna_saida in COLUNAS_SAIDA:
        if coluna_saida not in df.columns:
            df[coluna_saida] = ""

    # o Excel devolve as datas como "2026-02-10 00:00:00"; normaliza para
    # DD/MM/YYYY antes de qualquer comparação ou regravação
    for coluna_data in COLUNAS_DATA:
        df[coluna_data] = df[coluna_data].apply(formatar_data)

    df["_chave"] = df.apply(
        lambda registro: (
            normalizar_chave(registro.get(COL_PEDIDO, "")),
            normalizar_chave(registro.get(COL_RMS, "")),
        ),
        axis=1,
    )

    return df[COLUNAS_SAIDA + ["_chave"]]


def comentarios_do_pa(texto):
    """Mantém apenas os comentários gravados pelo fluxo em nuvem."""
    linhas = [l.strip() for l in str(texto or "").split("\n") if l.strip()]
    return [l for l in linhas if l.startswith(PREFIXO_PA)]


def mesclar(df_novo, df_anterior):
    """
    Mescla preservando o trabalho manual da equipe.

    Devolve (df_final, chaves_novas, ocorrencias).
    """
    ocorrencias = []
    anterior_por_chave = {linha["_chave"]: linha for _, linha in df_anterior.iterrows()}
    chaves_novas = set()

    for indice, linha in df_novo.iterrows():
        chave = linha["_chave"]
        antiga = anterior_por_chave.get(chave)

        if antiga is None:
            chaves_novas.add(chave)
            continue

        for coluna in COLUNAS_PRESERVADAS:
            valor_antigo = str(antiga.get(coluna, "")).strip()
            if valor_antigo:
                df_novo.at[indice, coluna] = valor_antigo

        # previsão inicial: o valor do relatório só entra se ainda não havia um
        prev_antiga = str(antiga.get(COL_PREV_INICIAL, "")).strip()
        if prev_antiga:
            df_novo.at[indice, COL_PREV_INICIAL] = prev_antiga

        # data real de recebimento: preservada sempre que já preenchida
        recebimento_antigo = str(antiga.get(COL_RECEBIMENTO, "")).strip()
        if recebimento_antigo:
            df_novo.at[indice, COL_RECEBIMENTO] = recebimento_antigo

        # contato: preservado quando a planilha de contatos não trouxe endereço
        contato_novo = str(linha.get(COL_CONTATO, "")).strip()
        contato_antigo = str(antiga.get(COL_CONTATO, "")).strip()
        if not contato_novo and contato_antigo:
            df_novo.at[indice, COL_CONTATO] = contato_antigo

        # exceções: divergências são reescritas, comentários do fluxo ficam
        preservados = comentarios_do_pa(antiga.get(COL_EXCECOES, ""))
        atuais = [
            l.strip()
            for l in str(linha.get(COL_EXCECOES, "") or "").split("\n")
            if l.strip()
        ]
        todos = list(dict.fromkeys(atuais + preservados))
        df_novo.at[indice, COL_EXCECOES] = "\n".join(todos)

    # linhas que sumiram dos relatórios não são excluídas
    chaves_atuais = set(df_novo["_chave"])
    remanescentes = [
        linha for chave, linha in anterior_por_chave.items() if chave not in chaves_atuais
    ]

    for linha in remanescentes:
        ocorrencias.append(
            Ocorrencia(
                planilha="planilha de Follow-Up anterior",
                linha="-",
                coluna="-",
                mensagem=(
                    "Este pedido estava na planilha anterior, mas não apareceu nos "
                    "relatórios desta execução. A linha foi mantida como estava."
                ),
                pedido=str(linha.get(COL_PEDIDO, "")),
                controle=str(linha.get(COL_RMS, "")),
            )
        )

    if remanescentes:
        df_final = pd.concat(
            [df_novo, pd.DataFrame(remanescentes, columns=df_novo.columns)],
            ignore_index=True,
        )
    else:
        df_final = df_novo

    return df_final, chaves_novas, ocorrencias


# ==============================
# PROCESSAMENTO COMPLETO
# ==============================


def processar(arquivos, config, caminho_anterior=None):
    """
    arquivos: dicionário {tipo: caminho} já confirmado pelo operador na tela.

    Todas as abas de cada arquivo são lidas e classificadas individualmente.
    Devolve um Resultado com o DataFrame pronto e as ocorrências encontradas.
    """
    consolidador = Consolidador(config)
    resultado = Resultado()

    # ordem fixa: o relatório de RMS entra por último, porque é o único sem
    # código de pedido e precisa se apoiar nos códigos de controle já vistos
    for tipo in ("compras", "produtos", "rms"):
        caminho = arquivos.get(tipo)
        if not caminho:
            continue

        classificadas = classificar_abas(caminho, tipo, config)

        if not classificadas:
            raise ErroDeUso(
                "Nenhuma aba de %s corresponde ao relatório escolhido. Confira "
                "se o tipo selecionado para esse arquivo está correto."
                % os.path.basename(caminho)
            )

        # dentro do arquivo, as abas auxiliares entram depois da principal
        classificadas.sort(key=lambda c: 1 if config["tipos"][c[1]].get("auxiliar") else 0)

        for aba, tipo_aba, linha, coluna in classificadas:
            df, linha_cabecalho = ler_tabela(caminho, tipo_aba, config, aba, linha, coluna)
            consolidador.adicionar_relatorio(tipo_aba, caminho, df, linha_cabecalho, aba)
            resultado.abas_lidas.append(
                "%s / %s (%d linhas)" % (os.path.basename(caminho), aba, len(df))
            )

    df, chaves_com_excecao = consolidador.consolidar()
    ocorrencias = list(consolidador.ocorrencias)
    resultado.linhas_ignoradas = consolidador.linhas_ignoradas

    # contatos dos fornecedores (opcional)
    caminho_contatos = arquivos.get(TIPO_FORNECEDORES)
    if caminho_contatos:
        resultado.contatos_carregados = True
        contatos = carregar_contatos(caminho_contatos, config)
        sem_contato = aplicar_contatos(df, contatos, config.get("similaridade_minima_coluna", 0.9))

        for indice in sem_contato:
            comentario = "%s fornecedor sem e-mail de contato cadastrado" % PREFIXO_PA
            atuais = [l for l in str(df.at[indice, COL_EXCECOES] or "").split("\n") if l]
            if comentario not in atuais:
                atuais.append(comentario)
            df.at[indice, COL_EXCECOES] = "\n".join(atuais)
            chaves_com_excecao.add(df.at[indice, "_chave"])

            ocorrencias.append(
                Ocorrencia(
                    planilha=os.path.basename(caminho_contatos),
                    linha="-",
                    coluna=COL_CONTATO,
                    mensagem=(
                        'Não há e-mail cadastrado para o fornecedor "%s". A cobrança '
                        "automática não será enviada para este pedido."
                        % df.at[indice, COL_FORNECEDOR]
                    ),
                    pedido=str(df.at[indice, COL_PEDIDO]),
                    controle=str(df.at[indice, COL_RMS]),
                )
            )

    # mesclagem com a versão anterior (opcional)
    chaves_novas = set()
    if caminho_anterior:
        df_anterior = ler_follow_up_anterior(caminho_anterior, config)
        df, chaves_novas, ocorrencias_merge = mesclar(df, df_anterior)
        ocorrencias.extend(ocorrencias_merge)

    resultado.df = df
    resultado.ocorrencias = ocorrencias
    resultado.chaves_novas = chaves_novas
    resultado.chaves_com_excecao = chaves_com_excecao
    resultado.total_processados = len(df)
    resultado.total_novos = len(chaves_novas) if caminho_anterior else len(df)
    resultado.total_verificar = int((df[COL_EXCECOES].astype(str).str.strip() != "").sum())

    return resultado


# ==============================
# GRAVAÇÃO DA PLANILHA
# ==============================

# formato das colunas de data e estilo da tabela do Excel
FORMATO_DATA = "DD/MM/YYYY"
ESTILO_TABELA = "TableStyleMedium2"

FILL_NOVO = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_EXCECAO = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
FILL_HEADER = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

LARGURAS = {
    COL_RECEBIMENTO: 18,
    COL_PREV_ATUAL: 16,
    COL_PREV_INICIAL: 16,
    COL_APROVACAO: 16,
    COL_FORNECEDOR: 32,
    COL_CONTATO: 30,
    COL_DESCRICAO: 55,
    COL_REQUISITANTE: 22,
    COL_DESTINO: 16,
    COL_PROJETO: 26,
    COL_COMPRADOR: 22,
    COL_PEDIDO: 14,
    COL_RMS: 14,
    COL_STATUS: 22,
    COL_OBS: 40,
    COL_FOLLOWUP: 16,
    COL_EXCECOES: 50,
}


def para_data(valor):
    """
    Converte o texto DD/MM/YYYY em uma data de verdade, para o Excel.

    Devolve None quando a célula está vazia e o próprio texto quando não for
    uma data reconhecível (nesse caso a célula continua como texto, para não
    perder a informação).
    """
    texto = formatar_data(valor)
    if not texto:
        return None

    try:
        return datetime.strptime(texto[:10], "%d/%m/%Y").date()
    except ValueError:
        return texto


def gravar_planilha(resultado, caminho):
    """
    Grava follow_up.xlsx como tabela do Excel.

    As colunas de data são gravadas como data de verdade (não como texto), de
    modo que a ordenação e os filtros de data do Excel funcionem; o restante é
    texto. A faixa toda é uma tabela nomeada, com filtro em cada coluna.
    """
    df = resultado.df
    wb = Workbook()
    ws = wb.active
    ws.title = ABA_DADOS

    ws.append(COLUNAS_SAIDA)

    indices_data = [COLUNAS_SAIDA.index(coluna) + 1 for coluna in COLUNAS_DATA]
    indices_texto_longo = [
        COLUNAS_SAIDA.index(coluna) + 1 for coluna in (COL_DESCRICAO, COL_OBS, COL_EXCECOES)
    ]

    for _, linha in df.iterrows():
        valores = []
        for coluna in COLUNAS_SAIDA:
            bruto = linha.get(coluna, "")
            if coluna in COLUNAS_DATA:
                valores.append(para_data(bruto))
            else:
                texto = str(bruto or "").strip()
                # célula realmente vazia (e não texto vazio), para que os
                # filtros do Excel reconheçam o "(Vazias)"
                valores.append(texto or None)
        ws.append(valores)

    # ---------- cabeçalho ----------
    for celula in ws[1]:
        celula.font = Font(bold=True)
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for indice, coluna in enumerate(COLUNAS_SAIDA, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = LARGURAS.get(coluna, 18)

    # ---------- corpo ----------
    col_excecoes = COLUNAS_SAIDA.index(COL_EXCECOES) + 1

    for posicao, (_, linha) in enumerate(df.iterrows()):
        numero_linha = posicao + 2
        tem_excecao = str(linha.get(COL_EXCECOES, "") or "").strip() != ""
        eh_nova = linha.get("_chave") in resultado.chaves_novas

        for indice in range(1, len(COLUNAS_SAIDA) + 1):
            celula = ws.cell(row=numero_linha, column=indice)

            if indice in indices_data:
                celula.number_format = FORMATO_DATA
                celula.alignment = Alignment(horizontal="center", vertical="top")
            elif indice in indices_texto_longo:
                celula.alignment = Alignment(vertical="top", wrap_text=True)
            else:
                celula.alignment = Alignment(vertical="top")

            if tem_excecao and indice == col_excecoes:
                celula.fill = FILL_EXCECAO
            elif eh_nova:
                celula.fill = FILL_NOVO

    ws.freeze_panes = "A2"

    # ---------- tabela ----------
    ultima_linha = max(ws.max_row, 2)
    referencia = "A1:%s%d" % (get_column_letter(len(COLUNAS_SAIDA)), ultima_linha)

    tabela = Table(displayName=TABELA, ref=referencia)
    tabela.tableStyleInfo = TableStyleInfo(
        name=ESTILO_TABELA,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    # filtro em cada coluna (o botão de seta no cabeçalho)
    tabela.autoFilter = AutoFilter(ref=referencia)
    ws.add_table(tabela)

    # ---------- aba de listas + validação de dados ----------
    listas = wb.create_sheet(title=ABA_LISTAS)
    listas["A1"] = "Status do pedido"
    listas["A1"].font = Font(bold=True)

    for posicao, opcao in enumerate(STATUS_OPCOES, start=2):
        listas.cell(row=posicao, column=1, value=opcao)

    listas.column_dimensions["A"].width = 28

    intervalo = "%s!$A$2:$A$%d" % (ABA_LISTAS, len(STATUS_OPCOES) + 1)
    validacao = DataValidation(type="list", formula1="=%s" % intervalo, allow_blank=True)
    validacao.error = "Escolha uma das opções da lista."
    validacao.errorTitle = "Status inválido"
    validacao.prompt = "Escolha a etapa em que o pedido se encontra."
    validacao.promptTitle = "Status do pedido"
    ws.add_data_validation(validacao)

    letra_status = get_column_letter(COLUNAS_SAIDA.index(COL_STATUS) + 1)
    validacao.add("%s2:%s%d" % (letra_status, letra_status, max(ultima_linha, 2) + 500))

    # as linhas acrescentadas no futuro (pelo operador ou pelo fluxo em nuvem)
    # já nascem com o formato de data nas colunas de data
    for indice in indices_data:
        letra = get_column_letter(indice)
        for numero_linha in range(ultima_linha + 1, ultima_linha + 501):
            ws["%s%d" % (letra, numero_linha)].number_format = FORMATO_DATA

    wb.save(caminho)
    return caminho


def gravar_logs(ocorrencias, caminho):
    """Exporta error_logs.txt agrupado por planilha e ordenado pela linha."""
    if not ocorrencias:
        return None

    def ordenar(ocorrencia):
        try:
            numero = int(ocorrencia.linha)
        except (TypeError, ValueError):
            numero = 10 ** 9
        return (ocorrencia.origem(), numero)

    with open(caminho, "w", encoding="utf-8") as f:
        f.write("AutoFollow - registro de ocorrências\n")
        f.write("Gerado em %s\n" % datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        f.write("Formato: data e hora | gravidade | planilha (aba) | linha | coluna | "
                "pedido | RMS | descrição\n")
        f.write("=" * 110 + "\n\n")

        origem_atual = None
        for ocorrencia in sorted(ocorrencias, key=ordenar):
            if ocorrencia.origem() != origem_atual:
                origem_atual = ocorrencia.origem()
                f.write("\n--- %s ---\n" % origem_atual)
            f.write(ocorrencia.linha_log() + "\n")

    return caminho


def resumo_para_usuario(resultado, caminho_saida, caminho_logs):
    """Texto final exibido na tela, em linguagem simples."""
    linhas = [
        "Planilha de Follow-Up criada com sucesso.",
        "",
        "Pedidos processados: %d" % resultado.total_processados,
        "Pedidos novos nesta execução: %d" % resultado.total_novos,
        "Pedidos que precisam de verificação: %d" % resultado.total_verificar,
    ]

    if resultado.abas_lidas:
        linhas.append("")
        linhas.append("Abas lidas:")
        for descricao in resultado.abas_lidas:
            linhas.append("  - %s" % descricao)

    if resultado.linhas_ignoradas:
        linhas.append("")
        linhas.append(
            "%d linha(s) ficaram de fora por não se referirem a pedidos de compra "
            "(por exemplo, RMS de pagamento)." % resultado.linhas_ignoradas
        )

    if not resultado.contatos_carregados:
        linhas.append("")
        linhas.append(
            "A planilha de contatos não foi carregada, portanto a coluna de "
            "contato do fornecedor ficou vazia."
        )

    if caminho_logs:
        linhas.append("")
        linhas.append(
            "Algumas linhas precisam de conferência. O arquivo %s, salvo na mesma "
            "pasta, indica a planilha, a linha e a coluna de cada caso."
            % os.path.basename(caminho_logs)
        )

    linhas.append("")
    linhas.append("Arquivo salvo em: %s" % caminho_saida)

    return "\n".join(linhas)
