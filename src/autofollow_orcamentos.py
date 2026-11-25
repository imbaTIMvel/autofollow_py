# -*- coding: utf-8 -*-
"""
AutoFollow - identificação de orçamentos (apoio ao AutoFollow-PA)

Lê os orçamentos recebidos (PDF) e extrai o que interessa ao follow-up:
fornecedor, número da proposta, data de emissão, validade e, principalmente,
o prazo de entrega prometido - convertido em data prevista.

A saída é uma tabela de triagem (tb_triagem) que o fluxo em nuvem usa para
atualizar a planilha de Follow-Up, e que o operador consegue conferir linha a
linha.

Este módulo não depende de PySide6 nem do Power Automate: roda sozinho sobre
uma pasta de PDFs.

Dependências: pandas, openpyxl, pdfplumber.
"""

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import pandas as pd

try:
    import pdfplumber
except ImportError:  # pragma: no cover
    pdfplumber = None


# ==============================
# PALAVRAS-CHAVE DO MODELO
# ==============================

# a linha precisa falar de entrega (ou de prazo, que é o mesmo campo em boa
# parte dos modelos de orçamento) para ser considerada
GATILHOS_ENTREGA = [
    "prazo de entrega",
    "prazo previsto de entrega",
    "previsao de entrega",
    "tempo de entrega",
    "prazo de fabricacao",
    "prazo de envio",
    "entrega em",
    "entrega:",
    "entrega ",
    "prazo:",
    "prazo ",
    "postado",
    "expedicao",
]

# gatilhos que, sozinhos, já identificam o campo com segurança
GATILHOS_FORTES = [
    "prazo de entrega",
    "entrega prevista",
    "entrega esta prevista",
    "entrega sera",
    "prazo previsto de entrega",
    "previsao de entrega",
    "tempo de entrega",
    "prazo de fabricacao",
    "descricao do prazo de entrega",
]

# se a frase fala de uma destas coisas, o número não é prazo de entrega
NEGATIVOS = [
    "pagamento",
    "pagto",
    "validade",
    "valido",
    "valida",
    "garantia",
    "faturamento",
    "faturado",
    "boleto",
    "vencimento",
    "cartorio",
    "credito",
    "desconto",
    "devolucao",
    "devolucoes",
    "proposta valida",
    "cotacao",
    "reajuste",
    "multa",
    "juros",
    "cancelamento",
    "antes do prazo",
    "atraso",
    "atrasos",
    "em caso de",
]

# trecho, a partir do gatilho, em que o prazo é procurado
JANELA_CONTEXTO = 120

# páginas varridas na busca por tabelas
PAGINAS_TABELA = 4

# o que vem antes do gatilho e indica que a frase fala de lugar, não de prazo
PREFIXOS_DE_LUGAR = ["local de", "local da", "dados para", "dados da", "endereco", "end.",
                     "cidade de", "ponto de"]

IMEDIATO = ["pronta entrega", "pronta-entrega", "prontaentrega", "imediat", "disponivel em estoque"]
INDEFINIDO = ["a combinar", "sob consulta", "a definir", "conforme disponibilidade",
              "estimativa de lead time", "lead time", "sujeito a variacao", "a confirmar"]

# CNPJ do próprio Instituto: nunca é o fornecedor
CNPJ_PROPRIO = "09429074000112"

# provedores gratuitos: o domínio não identifica o fornecedor
DOMINIOS_GENERICOS = {
    "gmail.com", "hotmail.com", "outlook.com", "outlook.com.br", "yahoo.com",
    "yahoo.com.br", "bol.com.br", "uol.com.br", "terra.com.br", "live.com",
    "icloud.com", "msn.com", "ig.com.br", "globo.com",
}

# palavras que sozinhas não identificam ninguém
PALAVRAS_GENERICAS = {
    "similar", "comercio", "comercial", "industria", "industrial", "materiais",
    "material", "nacional", "brasil", "produtos", "servicos", "solucoes",
    "tecnologia", "distribuidora", "suprimentos", "ferramentas", "eletronica",
    "importacao", "representacoes", "equipamentos", "automacao", "central",
}

SUFIXOS_EMPRESA = {"ltda", "me", "epp", "eireli", "sa", "s/a", "s.a", "cia", "e", "de", "da", "do"}

MESES = {
    "janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}


# ==============================
# TEXTO
# ==============================


def normalizar(texto):
    """minúsculas, sem acentos, espaços colapsados."""
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ASCII", "ignore").decode()
    return re.sub(r"\s+", " ", texto).strip().lower()


def so_digitos(texto):
    return re.sub(r"\D", "", str(texto or ""))


def extrair_texto(caminho):
    """Texto do PDF. Devolve '' quando o arquivo é digitalizado (só imagem)."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber não está instalado (pip install pdfplumber).")

    partes = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            partes.append(pagina.extract_text() or "")

    return "\n".join(partes)


# ==============================
# CAMPOS SIMPLES
# ==============================


def extrair_cnpjs(texto):
    """CNPJs do documento, sem o do próprio Instituto, na ordem de aparição."""
    encontrados = []
    for bruto in re.findall(r"\d{2}[.\s]?\d{3}[.\s]?\d{3}[/\s]?\d{4}[-\s]?\d{2}", texto):
        numero = so_digitos(bruto)
        if len(numero) == 14 and numero != CNPJ_PROPRIO and numero not in encontrados:
            encontrados.append(numero)
    return encontrados


def extrair_codigos(texto):
    """
    Códigos de pedido e de RMS citados no texto.

    Mesma lógica do Office Script ExtrairDadosEmail.ts - os dois precisam
    concordar, porque o script roda em produção e este módulo é o que se testa.
    """
    plano = str(texto or "").upper()

    pedidos = []
    for achado in re.findall(r"\b(?:AS[NI]|PC[IPN])\d{6,12}(?:/\d{1,3})?\b", plano):
        if achado not in pedidos:
            pedidos.append(achado)

    rms = []
    for achado in re.findall(r"\b(?:CPI|REM|PSP|CP|RM)\d{2,8}\b", plano):
        if achado not in rms:
            rms.append(achado)

    # RMS numérica só vale com a palavra "RMS" por perto
    for achado in re.findall(r"RMS[^0-9]{0,15}(\d{5,7})", plano):
        if achado not in rms:
            rms.append(achado)

    return {"pedidos": pedidos, "rms": rms}


def classificar_email(assunto, corpo, remetente="", recebido_em=None, fornecedores=None):
    """
    Avalia um e-mail como o fluxo em nuvem avalia: códigos, prazo e confiança.

    É o espelho em Python do Office Script, usado para testar o modelo sem
    depender do Power Automate.
    """
    texto = (assunto or "") + "\n" + (corpo or "")
    recebido_em = recebido_em or date.today()

    codigos = extrair_codigos(texto)
    prazo = extrair_prazo(texto)
    prevista = prever_entrega(prazo, recebido_em)

    confianca = prazo.confianca
    if prazo.tipo == "nao_encontrado":
        confianca = "baixa"
    if not codigos["pedidos"] and not codigos["rms"]:
        confianca = "baixa"
    if len(codigos["pedidos"]) > 1:
        confianca = "baixa"

    fornecedor = ""
    if fornecedores:
        fornecedor = identificar_fornecedor(texto, fornecedores, remetente)[0]

    return {
        "pedidos": codigos["pedidos"],
        "rms": codigos["rms"],
        "fornecedor": fornecedor,
        "tipo_prazo": prazo.tipo,
        "dias": prazo.dias,
        "uteis": prazo.uteis,
        "data_prevista": prevista,
        "trecho": prazo.trecho,
        "confianca": confianca,
        "aplica_automaticamente": confianca == "alta" and len(codigos["pedidos"]) == 1,
    }


def extrair_numero_orcamento(texto):
    """Número da proposta/orçamento, quando o documento traz."""
    padroes = [
        r"or[çc]amento\s*(?:n[ºo°.]?|:|n[uú]mero)?\s*[:#]?\s*(\d{3,12})",
        r"proposta\s*(?:comercial\s*)?(?:n[ºo°.]?|:)?\s*[:#]?\s*(\d{3,12})",
        r"cota[çc][aã]o\s*(?:n[ºo°.]?|:)?\s*[:#]?\s*(\d{3,12})",
        r"n[uú]mero\s*:?\s*(\d{4,12})",
        r"pedido\s*(?:n[ºo°.]?|:)?\s*[:#]?\s*(\d{4,12})",
    ]

    plano = normalizar(texto)
    for padrao in padroes:
        achado = re.search(padrao, plano)
        if achado:
            return achado.group(1)
    return ""


def extrair_datas(texto):
    """(data de emissão, data de validade) quando identificáveis."""
    plano = normalizar(texto)

    def buscar(rotulos):
        for rotulo in rotulos:
            achado = re.search(
                rotulo + r"[^0-9]{0,20}(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", plano
            )
            if achado:
                return converter_data(achado.group(1))
        return None

    emissao = buscar([r"data de emiss[aã]o", r"emiss[aã]o", r"\bdata\b"])
    validade = buscar([r"validade", r"v[aá]lid[oa] at[eé]"])

    return emissao, validade


def converter_data(texto):
    texto = str(texto or "").strip()
    for formato in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    return None


def extrair_valor_total(texto):
    """Maior valor em reais do documento - serve para casar com o pedido."""
    valores = []
    for bruto in re.findall(r"R\$\s*([\d.]+,\d{2})", texto):
        try:
            valores.append(float(bruto.replace(".", "").replace(",", ".")))
        except ValueError:
            continue
    return max(valores) if valores else None


# ==============================
# PRAZO DE ENTREGA
# ==============================


@dataclass
class Prazo:
    """O prazo prometido, já interpretado."""

    dias: int = None            # em dias corridos ou úteis, conforme `uteis`
    uteis: bool = False
    data: date = None           # quando o fornecedor informa a data, não o prazo
    tipo: str = "nao_encontrado"  # prazo | data | imediato | indefinido | nao_encontrado
    trecho: str = ""
    confianca: str = "baixa"      # alta | media | baixa
    alternativas: list = field(default_factory=list)

    def descricao(self):
        if self.tipo == "data":
            return "entrega em %s" % self.data.strftime("%d/%m/%Y")
        if self.tipo == "imediato":
            return "pronta entrega"
        if self.tipo == "indefinido":
            return "a combinar / sob consulta"
        if self.tipo == "prazo":
            return "%d dias %s" % (self.dias, "úteis" if self.uteis else "corridos")
        return "não encontrado"


def _tem_negativo(frase):
    return any(palavra in frase for palavra in NEGATIVOS)


def _quantidade(frase):
    """
    Converte a quantidade de tempo da frase em (dias, úteis?).

    Intervalos ("3 a 5 dias", "02 A 03 dias") devolvem o maior valor, que é o
    compromisso real do fornecedor.
    """
    unidades = {"dia": 1, "semana": 7, "mes": 30, "meses": 30}

    padrao = (
        r"(\d{1,3})\s*(?:a|à|-|ate|até)\s*(\d{1,3})\s*(dias?|semanas?|m[eê]s|meses)"
        r"|(\d{1,3})\s*(dias?|semanas?|m[eê]s|meses)"
    )

    melhor = None
    for achado in re.finditer(padrao, frase):
        if achado.group(3):
            valor = max(int(achado.group(1)), int(achado.group(2)))
            unidade = achado.group(3)
        else:
            valor = int(achado.group(4))
            unidade = achado.group(5)

        raiz = unidade.rstrip("s").replace("ê", "e")
        dias = valor * unidades.get(raiz, 1)

        # "úteis" costuma vir logo depois da unidade
        trecho_depois = frase[achado.end(): achado.end() + 18]
        uteis = "util" in trecho_depois or "uteis" in trecho_depois

        if melhor is None or dias > melhor[0]:
            melhor = (dias, uteis)

    return melhor


def _gatilhos(frase):
    """Posições dos gatilhos de entrega na frase, sem sobreposição."""
    achados = []

    for gatilho in GATILHOS_FORTES + GATILHOS_ENTREGA:
        posicao = frase.find(gatilho)
        while posicao >= 0:
            achados.append((posicao, gatilho, gatilho in GATILHOS_FORTES))
            posicao = frase.find(gatilho, posicao + 1)

    # na mesma posição, vale o gatilho mais longo ("prazo de entrega" antes de "prazo ")
    achados.sort(key=lambda a: (a[0], -len(a[1])))

    selecionados = []
    fim_anterior = -1
    for posicao, gatilho, forte in achados:
        if posicao <= fim_anterior:
            continue
        selecionados.append((posicao, gatilho, forte))
        fim_anterior = posicao + len(gatilho) - 1

    return selecionados


def _clausula(texto):
    """Só a oração: corta no primeiro ponto final ou ponto e vírgula."""
    return re.split(r"[;.]\s|\.$", texto[:JANELA_CONTEXTO])[0]


def _data_no_contexto(contexto, hoje=None):
    """
    Data prometida no trecho ("entrega prevista para 10/11/2026", "dia 05/12").

    É o formato típico da resposta do fornecedor a um e-mail de cobrança, em
    que ele informa a data e não o prazo. Sem o ano, assume o próximo 10/11 a
    partir de hoje.
    """
    hoje = hoje or date.today()

    # a barra é obrigatória quando não há ano: "3-5 dias úteis" é intervalo de
    # prazo, não a data 03/05
    achado = re.search(
        r"\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b|\b(\d{1,2})-(\d{1,2})-(\d{2,4})\b",
        contexto,
    )
    if not achado:
        return None

    # "10/11 dias" não é data
    if re.match(r"\s*(dias?|semanas?|mes|meses)", contexto[achado.end():]):
        return None

    grupos = achado.groups()
    if grupos[0] is not None:
        dia, mes, ano = int(grupos[0]), int(grupos[1]), grupos[2]
    else:
        dia, mes, ano = int(grupos[3]), int(grupos[4]), grupos[5]
    com_ano = ano is not None

    if com_ano:
        ano = int(ano)
        if ano < 100:
            ano += 2000
    else:
        ano = hoje.year

    try:
        encontrada = date(ano, mes, dia)
    except ValueError:
        return None

    # data sem ano que já passou é do ano que vem
    if not com_ano and encontrada < hoje:
        try:
            encontrada = date(ano + 1, mes, dia)
        except ValueError:
            return None

    return encontrada


def _avaliar_contexto(contexto, forte, hoje=None):
    """Interpreta o trecho que segue o gatilho. Devolve um Prazo ou None."""
    if _tem_negativo(contexto):
        return None

    trecho = re.sub(r"\s+", " ", contexto).strip()[:150]
    confianca = "alta" if forte else "media"

    data_prometida = _data_no_contexto(contexto, hoje)
    if data_prometida:
        return Prazo(data=data_prometida, tipo="data", trecho=trecho, confianca=confianca)

    quantidade = _quantidade(contexto)
    if quantidade:
        return Prazo(dias=quantidade[0], uteis=quantidade[1], tipo="prazo",
                     trecho=trecho, confianca=confianca)

    if any(termo in contexto for termo in IMEDIATO):
        return Prazo(dias=0, tipo="imediato", trecho=trecho, confianca=confianca)

    if any(termo in contexto for termo in INDEFINIDO):
        return Prazo(tipo="indefinido", trecho=trecho, confianca=confianca)

    return None


def extrair_prazo(texto):
    """
    Procura o prazo de entrega no texto do orçamento.

    A busca é ancorada no gatilho: encontrada uma expressão como "prazo de
    entrega", só o trecho imediatamente seguinte é interpretado. É isso que
    evita confundir o prazo com o de pagamento, a validade da proposta ou a
    garantia, que usam as mesmas unidades e aparecem na mesma página.
    """
    candidatos = []
    linhas = str(texto or "").split("\n")

    for indice, linha_bruta in enumerate(linhas):
        base = normalizar(linha_bruta)
        if not base:
            continue

        # o prazo às vezes quebra no fim da linha ("prazo de 90 a" / "120 dias")
        seguinte = normalizar(linhas[indice + 1]) if indice + 1 < len(linhas) else ""
        completa = (base + " " + seguinte).strip()

        for posicao, gatilho, forte in _gatilhos(base):
            antes = base[max(0, posicao - 40):posicao]

            # "local de entrega", "dados para entrega": é lugar, não prazo
            if any(prefixo in antes[-18:] for prefixo in PREFIXOS_DE_LUGAR):
                continue

            # o que vem imediatamente antes do gatilho também desqualifica
            # ("devoluções serão aceitas no prazo de 30 dias")
            if _tem_negativo(antes):
                continue

            # só a oração iniciada pelo gatilho é interpretada
            candidato = _avaliar_contexto(_clausula(base[posicao:]), forte)

            # nada na própria linha: o prazo pode ter quebrado para a seguinte
            if candidato is None and seguinte:
                candidato = _avaliar_contexto(_clausula(completa[posicao:]), forte)

            if candidato is not None:
                candidatos.append(candidato)

    if not candidatos:
        return Prazo()

    ordem = {"alta": 0, "media": 1, "baixa": 2}
    # um prazo com número vale mais que "pronta entrega" ou "a combinar", mesmo
    # que estes apareçam em uma frase mais explícita; entre dois números, o
    # maior, que é o compromisso que o fornecedor realmente assumiu
    peso_tipo = {"data": 0, "prazo": 1, "imediato": 2, "indefinido": 3, "nao_encontrado": 4}
    candidatos.sort(key=lambda p: (peso_tipo[p.tipo], ordem[p.confianca], -(p.dias or 0)))

    escolhido = candidatos[0]
    assinatura = (escolhido.tipo, escolhido.dias, escolhido.uteis)

    escolhido.alternativas = []
    vistos = {assinatura}
    for candidato in candidatos[1:]:
        atual = (candidato.tipo, candidato.dias, candidato.uteis)
        if atual in vistos:
            continue
        vistos.add(atual)
        escolhido.alternativas.append(candidato.descricao() + " (" + candidato.trecho[:60] + ")")

    # conflito entre candidatos do mesmo nível derruba a confiança
    conflito = {
        (c.tipo, c.dias, c.uteis) for c in candidatos if c.confianca == escolhido.confianca
    }
    if len(conflito) > 1:
        escolhido.confianca = "media" if escolhido.confianca == "alta" else "baixa"

    escolhido.alternativas = escolhido.alternativas[:3]

    return escolhido


# ==============================
# PRAZO EM TABELA
# ==============================

# Cabeçalhos de coluna que indicam prazo de ENTREGA. "Prazo", sozinho, fica de
# fora de propósito: em vários orçamentos é o cabeçalho da tabela de parcelas
# do pagamento, e aceitá-lo produziria uma data de entrega errada com cara de
# certa.
CABECALHOS_ENTREGA = [
    "prazo de entrega",
    "prazo entrega",
    "previsao de entrega",
    "prazo de fabricacao",
    "disponibilidade",
    "disponib",
    "entrega",
    "embarque",
    "expedicao",
]

# Se o cabeçalho (ou a legenda logo acima) da tabela fala disto, a tabela
# inteira é descartada: são os prazos que não são de entrega.
TABELA_NEGATIVA = [
    "pagto", "pagamento", "parcela", "parcelamento", "faturamento",
    "validade", "garantia", "vencimento", "boleto", "duplicata",
]

# "8 dias", "3 d.u", "5 du", "2 a 3 dias uteis"
VALOR_PRAZO = re.compile(
    r"(\d{1,3})\s*(?:a|ate|-)\s*(\d{1,3})\s*(d\.?\s*u\.?|dias?\s*uteis?|dias?|semanas?)"
    r"|(\d{1,3})\s*(d\.?\s*u\.?|dias?\s*uteis?|dias?|semanas?)",
    re.IGNORECASE,
)


def _valor_de_prazo(celula):
    """Lê o prazo de uma célula de tabela. Devolve (dias, uteis) ou None."""
    texto = normalizar(celula)
    if not texto or _tem_negativo(texto):
        return None

    achado = VALOR_PRAZO.search(texto)
    if not achado:
        return None

    if achado.group(3):
        valor = max(int(achado.group(1)), int(achado.group(2)))
        unidade = achado.group(3)
    else:
        valor = int(achado.group(4))
        unidade = achado.group(5)

    unidade = unidade.replace(".", "").replace(" ", "")
    uteis = unidade.startswith("du") or "uteis" in unidade or "util" in unidade
    dias = valor * 7 if unidade.startswith("semana") else valor

    if dias <= 0 or dias > 400:
        return None

    return dias, uteis


def extrair_prazo_tabela(caminho):
    """
    Procura o prazo de entrega nas tabelas do PDF.

    Serve para os orçamentos que guardam o prazo numa coluna por item, em vez
    de escrevê-lo em frase. É um caminho mais fraco que o texto corrido - e por
    isso a confiança nunca passa de média -, porque o cabeçalho da coluna é a
    única pista de que aquele número é de entrega.

    Quando a mesma coluna traz valores diferentes (orçamentos que oferecem
    preço por prazo: 3, 5 ou 10 dias úteis), nenhum é escolhido: a decisão é
    comercial, não de extração. O caso vai para conferência com as opções
    registradas.
    """
    if pdfplumber is None:
        return Prazo()

    encontrados = []      # [(dias, uteis, cabecalho)]
    tabelas_vistas = []   # [(tabela, texto do topo)] para a segunda passada

    try:
        with pdfplumber.open(caminho) as pdf:
            for pagina in pdf.pages[:PAGINAS_TABELA]:
                for tabela in pagina.extract_tables() or []:
                    if not tabela or len(tabela) < 2:
                        continue

                    # a tabela inteira é descartada se fala de pagamento,
                    # garantia ou validade em qualquer célula do cabeçalho
                    topo = " ".join(normalizar(c) for c in tabela[0] if c)
                    tabelas_vistas.append((tabela, topo))
                    if any(termo in topo for termo in TABELA_NEGATIVA):
                        continue

                    for indice, titulo in enumerate(tabela[0]):
                        rotulo = normalizar(titulo)
                        if not rotulo or len(rotulo) > 40:
                            continue
                        if not any(p in rotulo for p in CABECALHOS_ENTREGA):
                            continue
                        if _tem_negativo(rotulo):
                            continue

                        for linha in tabela[1:]:
                            if indice >= len(linha):
                                continue
                            valor = _valor_de_prazo(linha[indice])
                            if valor:
                                encontrados.append((valor[0], valor[1], rotulo))
    except Exception:
        return Prazo()

    # Tabela de preço por prazo: duas ou mais células da MESMA linha são
    # prazos ("3 d.u", "5 d.u", "10 d.u" como cabeçalhos de colunas de preço).
    # O fornecedor está oferecendo opções, e qual foi contratada é decisão
    # comercial - não se escolhe nenhuma aqui.
    if not encontrados:
        for tabela, topo in tabelas_vistas:
            if any(termo in topo for termo in TABELA_NEGATIVA):
                continue
            for linha in tabela:
                na_linha = []
                for celula in linha:
                    valor = _valor_de_prazo(celula)
                    if valor and valor not in na_linha:
                        na_linha.append(valor)
                if len(na_linha) >= 2:
                    opcoes = ", ".join(
                        "%d dias %s" % (d, "úteis" if u else "corridos")
                        for d, u in sorted(na_linha)
                    )
                    return Prazo(
                        tipo="indefinido",
                        trecho="o orçamento oferece preço por prazo: %s" % opcoes,
                        confianca="baixa",
                        alternativas=[opcoes],
                    )

    if not encontrados:
        return Prazo()

    distintos = sorted({(d, u) for d, u, _ in encontrados})
    rotulo = encontrados[0][2]

    if len(distintos) == 1:
        dias, uteis = distintos[0]
        return Prazo(
            dias=dias,
            uteis=uteis,
            tipo="prazo",
            trecho="coluna \"%s\" da tabela de itens: %d dias %s"
                   % (rotulo, dias, "úteis" if uteis else "corridos"),
            confianca="media",
        )

    # várias opções de prazo: não escolhe nenhuma
    opcoes = ", ".join("%d dias %s" % (d, "úteis" if u else "corridos") for d, u in distintos)
    return Prazo(
        tipo="indefinido",
        trecho="coluna \"%s\" oferece mais de um prazo: %s" % (rotulo, opcoes),
        confianca="baixa",
        alternativas=[opcoes],
    )


def somar_dias_uteis(inicio, dias):
    """Soma dias úteis (sem feriados - o calendário do Instituto pode entrar depois)."""
    atual = inicio
    restantes = dias
    while restantes > 0:
        atual += timedelta(days=1)
        if atual.weekday() < 5:
            restantes -= 1
    return atual


def prever_entrega(prazo, data_base):
    """Data prevista de entrega a partir do prazo e da data de referência."""
    if prazo.tipo == "data":
        return prazo.data
    if prazo.tipo == "imediato":
        return somar_dias_uteis(data_base, 2)  # postagem + trânsito mínimo
    if prazo.tipo != "prazo" or prazo.dias is None:
        return None
    if prazo.uteis:
        return somar_dias_uteis(data_base, prazo.dias)
    return data_base + timedelta(days=prazo.dias)


# ==============================
# FORNECEDOR
# ==============================


def carregar_fornecedores(caminho):
    """Lê a lista oficial de fornecedores da equipe de Compras."""
    bruto = pd.read_excel(caminho, header=None, dtype=str, keep_default_na=False)

    linha_cabecalho = 0
    for i in range(min(10, len(bruto))):
        valores = normalizar(" ".join(str(v) for v in bruto.iloc[i].tolist()))
        if "razao social" in valores and "mail" in valores:
            linha_cabecalho = i
            break

    df = bruto.iloc[linha_cabecalho + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in bruto.iloc[linha_cabecalho].tolist()]
    df = df.loc[:, [c for c in df.columns if c]]

    def coluna(*nomes):
        for nome in nomes:
            for real in df.columns:
                if normalizar(real) == normalizar(nome):
                    return real
        return None

    col_nome = coluna("RAZÃO SOCIAL", "Fornecedor")
    col_email = coluna("E-MAIL", "E-mail de contato")
    col_site = coluna("SITE")
    col_cnpj = coluna("CNPJ")

    fornecedores = []
    for _, linha in df.iterrows():
        nome = str(linha.get(col_nome, "")).strip()
        if not nome:
            continue
        email = str(linha.get(col_email, "")).strip() if col_email else ""
        site = str(linha.get(col_site, "")).strip() if col_site else ""
        dominios = {
            d
            for d in [
                email.split("@")[-1].lower() if "@" in email else "",
                re.sub(r"^(https?://)?(www\.)?", "", site).split("/")[0].lower(),
            ]
            if d and d not in DOMINIOS_GENERICOS
        }
        fornecedores.append(
            {
                "nome": nome,
                "email": email,
                "dominios": dominios,
                "cnpj": so_digitos(linha.get(col_cnpj, "")) if col_cnpj else "",
            }
        )

    return fornecedores


def identificar_fornecedor(texto, fornecedores, remetente=""):
    """
    Descobre de quem é o orçamento, por três caminhos independentes.

    Devolve (nome_cadastrado, nome_no_documento, como_foi_identificado,
    confianca).
    """
    plano = normalizar(texto)
    dominio_remetente = remetente.split("@")[-1].lower().strip() if "@" in remetente else ""

    # 1) CNPJ do documento contra o cadastro
    cnpjs = extrair_cnpjs(texto)
    for cnpj in cnpjs:
        for fornecedor in fornecedores:
            if fornecedor["cnpj"] and fornecedor["cnpj"] == cnpj:
                return fornecedor["nome"], fornecedor["nome"], "CNPJ", "alta"

    # 2) domínio do remetente (ou e-mail que apareça no documento)
    dominios_no_texto = set(re.findall(r"[\w.+-]+@([\w-]+\.[\w.-]+)", plano))
    for fornecedor in fornecedores:
        if dominio_remetente and dominio_remetente in fornecedor["dominios"]:
            return fornecedor["nome"], fornecedor["nome"], "domínio do remetente", "alta"
        if fornecedor["dominios"] & dominios_no_texto:
            return fornecedor["nome"], fornecedor["nome"], "e-mail no documento", "alta"

    # 3) razão social cadastrada aparecendo no texto
    melhor = None
    for fornecedor in fornecedores:
        alvo = normalizar(fornecedor["nome"])
        palavras = [p for p in alvo.split() if p not in SUFIXOS_EMPRESA]
        if not palavras:
            continue

        nucleo = " ".join(palavras[:2])

        # um nome de uma palavra só precisa ser distintivo o bastante
        if len(palavras) == 1 and (len(nucleo) < 9 or nucleo in PALAVRAS_GENERICAS):
            continue
        if all(p in PALAVRAS_GENERICAS for p in nucleo.split()):
            continue

        if re.search(r"\b" + re.escape(nucleo) + r"\b", plano):
            if melhor is None or len(nucleo) > len(melhor[1]):
                melhor = (fornecedor["nome"], nucleo)

    if melhor:
        return melhor[0], melhor[0], "razão social no documento", "média"

    # 4) não cadastrado: devolve o nome que aparece no topo do documento
    nome_documento = ""
    for linha in str(texto or "").split("\n")[:8]:
        limpa = re.sub(r"\s+", " ", linha).strip()
        if len(limpa) >= 6 and not re.fullmatch(r"[\d\s.\-/]+", limpa):
            nome_documento = limpa[:60]
            break

    return "", nome_documento, "não cadastrado", "baixa"


# ==============================
# ANÁLISE DE UM ORÇAMENTO
# ==============================


@dataclass
class Orcamento:
    arquivo: str = ""
    fornecedor_cadastrado: str = ""
    fornecedor_documento: str = ""
    identificado_por: str = ""
    confianca_fornecedor: str = ""
    numero: str = ""
    emissao: date = None
    validade: date = None
    valor_total: float = None
    prazo: Prazo = field(default_factory=Prazo)
    data_prevista: date = None
    confianca_prazo: str = ""
    observacao: str = ""

    def como_linha(self):
        return {
            "Arquivo": self.arquivo,
            "Fornecedor (cadastro)": self.fornecedor_cadastrado,
            "Fornecedor (documento)": self.fornecedor_documento,
            "Identificado por": self.identificado_por,
            "Nº do orçamento": self.numero,
            "Emissão": self.emissao.strftime("%d/%m/%Y") if self.emissao else "",
            "Validade": self.validade.strftime("%d/%m/%Y") if self.validade else "",
            "Valor total": self.valor_total,
            "Prazo informado": self.prazo.descricao(),
            "Trecho de origem": self.prazo.trecho,
            "Data prevista de entrega": (
                self.data_prevista.strftime("%d/%m/%Y") if self.data_prevista else ""
            ),
            "Confiança": self.confianca_prazo,
            "Observação": self.observacao,
        }


def analisar_orcamento(caminho, fornecedores, remetente="", data_recebimento=None):
    """Lê um PDF de orçamento e devolve tudo o que foi possível identificar."""
    resultado = Orcamento(arquivo=os.path.basename(caminho))

    texto = extrair_texto(caminho)

    if len(texto.strip()) < 40:
        resultado.observacao = (
            "O arquivo parece ser uma imagem digitalizada: não há texto para ler. "
            "Confira o prazo de entrega manualmente."
        )
        resultado.confianca_prazo = "nenhuma"
        return resultado

    (
        resultado.fornecedor_cadastrado,
        resultado.fornecedor_documento,
        resultado.identificado_por,
        resultado.confianca_fornecedor,
    ) = identificar_fornecedor(texto, fornecedores, remetente)

    resultado.numero = extrair_numero_orcamento(texto)
    resultado.emissao, resultado.validade = extrair_datas(texto)
    resultado.valor_total = extrair_valor_total(texto)
    resultado.prazo = extrair_prazo(texto)

    # o texto corrido tem prioridade; a tabela só entra quando ele não achou
    # nada, porque o cabeçalho de coluna é uma pista mais fraca que a frase
    if resultado.prazo.tipo == "nao_encontrado":
        da_tabela = extrair_prazo_tabela(caminho)
        if da_tabela.tipo != "nao_encontrado":
            resultado.prazo = da_tabela

    resultado.confianca_prazo = resultado.prazo.confianca

    base = resultado.emissao or data_recebimento or date.today()
    resultado.data_prevista = prever_entrega(resultado.prazo, base)

    avisos = []
    if resultado.prazo.tipo == "nao_encontrado":
        avisos.append("Prazo de entrega não encontrado no documento.")
    if resultado.prazo.tipo == "indefinido":
        avisos.append("O fornecedor não comprometeu prazo (a combinar / sob consulta).")
    if resultado.prazo.alternativas:
        avisos.append("Há mais de um prazo no documento: " + "; ".join(resultado.prazo.alternativas))
    if not resultado.fornecedor_cadastrado:
        avisos.append("Fornecedor não está na lista oficial.")

    resultado.observacao = " ".join(avisos)
    return resultado


def analisar_pasta(pasta, caminho_fornecedores):
    """Analisa todos os PDFs de uma pasta e devolve um DataFrame de triagem."""
    fornecedores = carregar_fornecedores(caminho_fornecedores)

    linhas = []
    for nome in sorted(os.listdir(pasta)):
        if not nome.lower().endswith(".pdf"):
            continue
        linhas.append(analisar_orcamento(os.path.join(pasta, nome), fornecedores).como_linha())

    return pd.DataFrame(linhas)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("uso: python autofollow_orcamentos.py <pasta_de_pdfs> <fornecedores.xlsx> [saida.xlsx]")
        raise SystemExit(1)

    tabela = analisar_pasta(sys.argv[1], sys.argv[2])

    if len(sys.argv) > 3:
        tabela.to_excel(sys.argv[3], index=False, sheet_name="tb_triagem")
        print("tabela de triagem gravada em", sys.argv[3])
    else:
        print(tabela.to_string(index=False))
