# -*- coding: utf-8 -*-
"""Teste rápido: gera planilhas fictícias, roda o processamento e imprime o resultado.

Uso: python teste_rapido.py   (os arquivos vão para /tmp/af_teste ou %TEMP%\\af_teste)
"""
import os
import shutil

import pandas as pd
from openpyxl import load_workbook

import autofollow_core as core

import tempfile
PASTA = os.path.join(tempfile.gettempdir(), "af_teste")
shutil.rmtree(PASTA, ignore_errors=True)
os.makedirs(PASTA)


def salvar(df, nome, linhas_lixo=0, colunas_lixo=0, embaralhar=False):
    """Grava a planilha deslocada e, se pedido, com as colunas fora de ordem."""
    if embaralhar:
        df = df[list(reversed(list(df.columns)))]

    caminho = os.path.join(PASTA, nome)
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        df.to_excel(
            w,
            index=False,
            startrow=linhas_lixo,
            startcol=colunas_lixo,
            sheet_name="Planilha1",
        )

    if linhas_lixo or colunas_lixo:
        wb = load_workbook(caminho)
        wb["Planilha1"]["A1"] = "Relatório exportado do Octalink em 24/09/2026"
        wb.save(caminho)

    return caminho


compras = pd.DataFrame({
    "COD PEDIDO": ["1001", "1002", "1003"],
    "COD CONTROLE": ["RMS-50", "RMS-51", ""],
    "NOME": ["ACME Componentes LTDA", "Beta Instrumentos SA", "Gama Ltda"],
    "DESCRIÇÃO": ["Osciloscópio 200 MHz", "Fonte de bancada", "Cabos diversos"],
    "OPERAÇÃO": ["COMPRA NACIONAL", "COMPRA NACIONAL", "COMPRA NACIONAL"],
    "DT EMISSÃO": ["05/08/2026", "12/08/2026", "20/08/2026"],
    "REQUISITANTE": ["Coord. Laboratório", "Coord. Laboratório", "Coord. Projetos"],
    "STATUS": ["Enviado", "Enviado", "Em cotação"],
})

produtos = pd.DataFrame({
    "CÓDIGO": ["1001", "1002"],
    "CONTROLE": ["RMS-50", "RMS-51"],
    "PARTICIPANTE": ["ACME Componentes LTDA", "Beta Instrumentos S/A"],
    "TIPO": ["Equipamento", "Equipamento"],
    "OBJETO": ["Osciloscópio 200 MHz 4 canais", "Fonte simétrica 30V"],
    "LOCAL DE ENTREGA": ["Sede - Campinas/SP", "Escritório Manaus/AM"],
    "ENTREGA": ["", ""],
    "SAÍDA": ["05/08/2026", "15/08/2026"],   # divergência proposital no 1002
    "CRIADO POR": ["ass.estoque", "ass.estoque"],
    "USUÁRIO": ["comprador.1", "comprador.2"],
    "STATUS": ["Em produção", "Enviado"],
})

rms = pd.DataFrame({
    "Chave": ["RMS-50", "RMS-52"],
    "Código": ["50", "52"],
    "Título": ["Aquisição de osciloscópio", "Serviço de calibração"],
    "Item": ["1 un - osciloscópio", "1 serviço"],
    "Operação": ["COMPRA", "SERVIÇO"],
    "Razão social": ["ACME Componentes LTDA", "Delta Metrologia ME"],
    "Projeto": ["LASC 2026", "LASC 2026"],
    "Requisitante": ["Coord. Laboratório", "Coord. Qualidade"],
    "Responsável": ["comprador.1", "comprador.3"],
    "Criação": ["05/08/2026", "22/08/2026"],
    "Entrega": ["", ""],
    "Status": ["Em produção", "Em cotação"],
})

fornecedores = pd.DataFrame({
    "Fornecedor": ["ACME Componentes LTDA", "ACME Componentes LTDA", "Delta Metrologia ME"],
    "E-mail de contato": ["vendas@acme.example", "pos.venda@acme.example", "contato@delta.example"],
})

# cabeçalho deslocado 3 linhas para baixo e 2 colunas para a direita
p_compras = salvar(compras, "relatorio_a.xlsx", linhas_lixo=3, colunas_lixo=2)
# mesmas colunas, ordem invertida
p_produtos = salvar(produtos, "relatorio_b.xlsx", embaralhar=True)
p_rms = salvar(rms, "relatorio_c.xlsx")
p_forn = salvar(fornecedores, "contatos.xlsx")

config = core.carregar_config(os.path.join(PASTA, core.ARQUIVO_CONFIG))

print("=== DETECÇÃO DE TIPO (sem usar o nome do arquivo) ===")
for caminho in (p_compras, p_produtos, p_rms, p_forn):
    tipo, linha, coluna, pontuacao = core.detectar_tipo(caminho, config)
    print("%-18s -> %-20s cabeçalho na linha %s, coluna %s  (%.2f)"
          % (os.path.basename(caminho), tipo,
             "-" if linha is None else linha + 1,
             "-" if linha is None else coluna + 1,
             pontuacao))

arquivos = {
    "compras": p_compras,
    "produtos": p_produtos,
    "rms": p_rms,
    core.TIPO_FORNECEDORES: p_forn,
}

resultado = core.processar(arquivos, config)

print("\n=== PLANILHA CONSOLIDADA ===")
colunas = [core.COL_PEDIDO, core.COL_RMS, core.COL_FORNECEDOR, core.COL_CONTATO,
           core.COL_DESCRICAO, core.COL_DESTINO, core.COL_APROVACAO, core.COL_STATUS]
print(resultado.df[colunas].to_string(index=False))

print("\n=== EXCEÇÕES ===")
for _, l in resultado.df.iterrows():
    if str(l[core.COL_EXCECOES]).strip():
        print("-", l[core.COL_PEDIDO], "/", l[core.COL_RMS], "->", l[core.COL_EXCECOES].replace("\n", " // "))

saida = os.path.join(PASTA, "follow_up.xlsx")
core.gravar_planilha(resultado, saida)
logs = core.gravar_logs(resultado.ocorrencias, os.path.join(PASTA, core.ARQUIVO_LOGS))

print("\n=== LOGS ===")
print(open(logs, encoding="utf-8").read())

print("=== RESUMO ===")
print(core.resumo_para_usuario(resultado, saida, logs))

# ---------------------------------------------------------------- 2ª execução
print("\n\n############ SEGUNDA EXECUÇÃO (mesclagem) ############")
wb = load_workbook(saida)
ws = wb[core.ABA_DADOS]
cab = [c.value for c in ws[1]]
ws.cell(row=2, column=cab.index(core.COL_STATUS) + 1, value="5 - Em produção")
ws.cell(row=2, column=cab.index(core.COL_PREV_INICIAL) + 1, value="30/09/2026")
ws.cell(row=2, column=cab.index(core.COL_OBS) + 1, value="fornecedor confirmou por telefone")
ws.cell(row=2, column=cab.index(core.COL_EXCECOES) + 1,
        value="Verificar manualmente: nenhum e-mail encontrado para este pedido.")
anterior = os.path.join(PASTA, "follow_up_anterior.xlsx")
wb.save(anterior)

compras2 = compras.copy()
compras2.loc[len(compras2)] = ["1004", "RMS-60", "Épsilon Cabos", "Cabos coaxiais",
                               "COMPRA NACIONAL", "01/09/2026", "Coord. Projetos", "Enviado"]
p_compras2 = salvar(compras2, "relatorio_a2.xlsx", linhas_lixo=3, colunas_lixo=2)

r2 = core.processar({"compras": p_compras2, "produtos": p_produtos, "rms": p_rms},
                    config, caminho_anterior=anterior)

print(r2.df[[core.COL_PEDIDO, core.COL_RMS, core.COL_STATUS, core.COL_PREV_INICIAL,
             core.COL_OBS, core.COL_CONTATO]].to_string(index=False))
print("\nchaves novas:", r2.chaves_novas)
print("exceções preservadas do fluxo:",
      [l for l in str(r2.df.iloc[0][core.COL_EXCECOES]).split("\n") if l])
print(core.resumo_para_usuario(r2, saida, None))

core.gravar_planilha(r2, os.path.join(PASTA, "follow_up_2.xlsx"))
wb2 = load_workbook(os.path.join(PASTA, "follow_up_2.xlsx"))
print("\nabas:", wb2.sheetnames, "| tabelas:", list(wb2[core.ABA_DADOS].tables))
print("validações:", len(wb2[core.ABA_DADOS].data_validations.dataValidation))
