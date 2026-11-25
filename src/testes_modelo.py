# -*- coding: utf-8 -*-
"""
AutoFollow - bateria de testes do modelo de identificação

Roda os casos de teste contra o módulo autofollow_orcamentos (o espelho em
Python do Office Script) e informa quais passaram. Use antes de mexer nas
listas de gatilhos e palavras negativas - e de novo depois, para ter certeza
de que o ajuste não quebrou outro caso.

    python testes_modelo.py              # executa e mostra o resultado
    python testes_modelo.py casos.xlsx   # grava também a planilha de execução

A planilha gerada é o roteiro que a equipe usa para testar o fluxo de verdade:
cada linha é um e-mail a enviar para a caixa de compras, com o resultado que
se espera ver na tabela de triagem.
"""

from datetime import date

import pandas as pd

import autofollow_orcamentos as ao

# data fixa, para que o resultado dos testes não mude com o calendário
HOJE = date(2026, 11, 16)  # segunda-feira

PEDIDO = "PCI2608934"
RMS = "CPI639"

# cada caso: (id, o que se verifica, assunto, corpo, esperado)
CASOS = [
    (
        "T01", "Prazo em dias úteis, com pedido no assunto",
        "Orçamento 4821 - Pedido %s" % PEDIDO,
        "Prezados,\nSegue nossa proposta.\nPrazo de entrega: 5 dias úteis após aprovação.\nAtt.",
        {"tipo_prazo": "prazo", "dias": 5, "uteis": True, "data_prevista": date(2026, 11, 23),
         "confianca": "alta", "aplica_automaticamente": True},
    ),
    (
        "T02", "Prazo em dias corridos",
        "Proposta comercial - %s" % PEDIDO,
        "Prazo de entrega: 15 dias.",
        {"tipo_prazo": "prazo", "dias": 15, "uteis": False, "data_prevista": date(2026, 12, 1),
         "confianca": "alta", "aplica_automaticamente": True},
    ),
    (
        "T03", "Intervalo: vale o maior, que é o compromisso real",
        "Cotação - pedido %s" % PEDIDO,
        "Previsão de entrega: 3 a 5 dias úteis.",
        {"tipo_prazo": "prazo", "dias": 5, "uteis": True, "confianca": "alta"},
    ),
    (
        "T04", "Prazo em semanas",
        "Orçamento - %s" % PEDIDO,
        "PRAZO DE ENTREGA: 10 SEMANAS",
        {"tipo_prazo": "prazo", "dias": 70, "uteis": False, "confianca": "alta"},
    ),
    (
        "T05", "Pronta entrega",
        "Orçamento - %s" % PEDIDO,
        "Prazo de entrega: pronta entrega, salvo venda prévia.",
        {"tipo_prazo": "imediato", "confianca": "alta", "aplica_automaticamente": True},
    ),
    (
        "T06", "Prazo a combinar: não compromete data",
        "Orçamento - %s" % PEDIDO,
        "Prazo de entrega: a combinar.",
        {"tipo_prazo": "indefinido", "data_prevista": None, "confianca": "alta"},
    ),
    (
        "T07", "Armadilha: prazo de pagamento na mesma mensagem",
        "Orçamento - %s" % PEDIDO,
        "Condições de pagamento: boleto 30 dias após envio.\n"
        "Prazo de entrega: 4 dias úteis.\nFaturamento mínimo: R$ 300,00.",
        {"tipo_prazo": "prazo", "dias": 4, "uteis": True, "confianca": "alta"},
    ),
    (
        "T08", "Armadilha: só validade da proposta, sem prazo de entrega",
        "Orçamento - %s" % PEDIDO,
        "Validade da proposta: 7 dias.\nGarantia: 12 meses.",
        {"tipo_prazo": "nao_encontrado", "confianca": "baixa", "aplica_automaticamente": False},
    ),
    (
        "T09", "Armadilha: 'local de entrega' é lugar, não prazo",
        "Orçamento - %s" % PEDIDO,
        "Local de entrega: a combinar.\nPrazo de entrega: 8 dias úteis.",
        {"tipo_prazo": "prazo", "dias": 8, "uteis": True, "confianca": "alta"},
    ),
    (
        "T10", "Armadilha: cláusula de cancelamento cita prazo de entrega",
        "Orçamento - %s" % PEDIDO,
        "Cancelamentos podem ser feitos até 30 dias antes do prazo de entrega acordado.",
        {"tipo_prazo": "nao_encontrado", "confianca": "baixa"},
    ),
    (
        "T11", "Resposta a cobrança: fornecedor informa a data, não o prazo",
        "RE: Follow-up de entrega - Pedido %s / RMS %s" % (PEDIDO, RMS),
        "Bom dia,\nA entrega está prevista para 10/12/2026.\nAtenciosamente,",
        {"tipo_prazo": "data", "data_prevista": date(2026, 12, 10), "confianca": "alta",
         "aplica_automaticamente": True},
    ),
    (
        "T12", "Data sem o ano: assume o próximo",
        "RE: Follow-up de entrega - Pedido %s" % PEDIDO,
        "Entrega prevista para 05/01.",
        {"tipo_prazo": "data", "data_prevista": date(2027, 1, 5), "confianca": "alta"},
    ),
    (
        "T13", "Sem código de pedido: vai para conferência",
        "Orçamento de peças",
        "Prazo de entrega: 6 dias úteis.",
        {"tipo_prazo": "prazo", "dias": 6, "confianca": "baixa",
         "aplica_automaticamente": False, "pedidos": []},
    ),
    (
        "T14", "Dois pedidos citados: não adivinha qual é",
        "Orçamentos dos pedidos %s e ASI26081748" % PEDIDO,
        "Prazo de entrega: 4 dias úteis para ambos.",
        {"confianca": "baixa", "aplica_automaticamente": False},
    ),
    (
        "T15", "Só a RMS com letras é citada",
        "Orçamento referente à RMS %s" % RMS,
        "Prazo de entrega: 9 dias úteis.",
        {"rms": [RMS], "pedidos": [], "confianca": "alta", "aplica_automaticamente": False},
    ),
    (
        "T16", "RMS numérica só vale com a palavra RMS por perto",
        "Orçamento RMS 112534",
        "Prazo de entrega: 3 dias úteis. Valor total 112999 reais.",
        {"rms": ["112534"], "confianca": "alta"},
    ),
    (
        "T17", "E-mail interno, sem assunto de orçamento e sem prazo",
        "Reunião de alinhamento de compras",
        "Pessoal, vamos remarcar para quinta?",
        {"tipo_prazo": "nao_encontrado", "confianca": "baixa", "aplica_automaticamente": False},
    ),
    (
        "T18", "Pedido identificado, mas sem prazo no texto",
        "Orçamento - Pedido %s" % PEDIDO,
        "Segue o orçamento em anexo conforme solicitado.",
        {"tipo_prazo": "nao_encontrado", "confianca": "baixa", "aplica_automaticamente": False,
         "pedidos": [PEDIDO]},
    ),
    (
        "T19", "Prazo quebrado em duas linhas",
        "Orçamento - %s" % PEDIDO,
        "Itens disponíveis em estoque parcial.\nNa falta, prazo de 90 a\n120 dias.",
        {"tipo_prazo": "prazo", "dias": 120, "uteis": False},
    ),
    (
        "T20", "Prazo escrito em caixa alta e sem acento",
        "ORCAMENTO - %s" % PEDIDO,
        "DESCRICAO DO PRAZO DE ENTREGA 20 DIAS UTEIS",
        {"tipo_prazo": "prazo", "dias": 20, "uteis": True, "confianca": "alta"},
    ),
]


def executar():
    resultados = []
    passou_total = 0

    for identificador, descricao, assunto, corpo, esperado in CASOS:
        obtido = ao.classificar_email(assunto, corpo, recebido_em=HOJE)

        divergencias = []
        for campo, valor_esperado in esperado.items():
            valor_obtido = obtido.get(campo)
            if valor_obtido != valor_esperado:
                divergencias.append("%s: esperado %r, obtido %r" % (campo, valor_esperado, valor_obtido))

        passou = not divergencias
        passou_total += 1 if passou else 0

        resultados.append(
            {
                "Caso": identificador,
                "O que verifica": descricao,
                "Assunto": assunto,
                "Corpo": corpo,
                "Prazo esperado": esperado.get("tipo_prazo", "-"),
                "Data prevista esperada": (
                    esperado["data_prevista"].strftime("%d/%m/%Y")
                    if esperado.get("data_prevista") else ""
                ),
                "Confiança esperada": esperado.get("confianca", "-"),
                "Aplica sozinho?": "sim" if esperado.get("aplica_automaticamente") else "não",
                "Resultado do modelo": "OK" if passou else "DIVERGENTE",
                "Divergências": "; ".join(divergencias),
                "Resultado no fluxo": "",
                "Conferido por": "",
            }
        )

        marca = "ok  " if passou else "FALHA"
        print("%s %s  %s" % (marca, identificador, descricao))
        for divergencia in divergencias:
            print("        %s" % divergencia)

    print("\n%d de %d casos conforme o esperado." % (passou_total, len(CASOS)))
    return pd.DataFrame(resultados)


if __name__ == "__main__":
    import sys

    tabela = executar()

    if len(sys.argv) > 1:
        tabela.to_excel(sys.argv[1], index=False, sheet_name="casos_de_teste")
        print("planilha de execução gravada em", sys.argv[1])
