# -*- coding: utf-8 -*-
"""
AutoFollow - interface (AutoFollow-Py)

Associa os relatórios do Octalink (pedido de compras, produto e RMS) em uma
única planilha de Follow-Up, opcionalmente preenchendo o contato dos
fornecedores e preservando o trabalho manual da planilha anterior.

O processamento fica em autofollow_core.py. Aqui está apenas a interface e a
conversa com o usuário - sempre em linguagem simples, sem termos técnicos.

Execução:  python autofollow.py
Empacotamento (exemplo):
    pyinstaller --onefile --windowed --add-data "bg_hbr.png;." autofollow.py
"""

import os
import subprocess
import sys
import traceback
from datetime import datetime

from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QBrush, QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

import autofollow_core as core

VERSAO = "0.1.0"

# ==============================
# ESTADO
# ==============================

# cada item: {"caminho": str, "tipo": str|None, "linha": QWidget, "combo": QComboBox}
arquivos = []

config = core.carregar_config()

# opções do menu de tipo, na ordem em que aparecem para o operador
OPCOES_TIPO = [
    ("compras", config["tipos"]["compras"]["nome_exibido"]),
    ("produtos", config["tipos"]["produtos"]["nome_exibido"]),
    ("rms", config["tipos"]["rms"]["nome_exibido"]),
    (core.TIPO_FORNECEDORES, config["tipos"][core.TIPO_FORNECEDORES]["nome_exibido"]),
    (core.TIPO_ANTERIOR, config["tipos"][core.TIPO_ANTERIOR]["nome_exibido"]),
    (None, "Não utilizar este arquivo"),
]

NAO_IDENTIFICADO = "Não identificado - escolha o tipo"


# ==============================
# MENSAGENS EM LINGUAGEM SIMPLES
# ==============================


def mensagem_amigavel(erro):
    """Converte a exceção em uma frase compreensível para quem vai operar."""
    if isinstance(erro, core.ErroDeUso):
        # mensagem já escrita para o usuário
        return str(erro)

    if isinstance(erro, PermissionError):
        return (
            "Não foi possível salvar a planilha porque o arquivo está aberto. "
            "Feche a planilha no Excel e execute novamente."
        )

    if isinstance(erro, FileNotFoundError):
        return (
            "Um dos arquivos selecionados não foi encontrado. Verifique se ele "
            "não foi movido ou renomeado e selecione o arquivo novamente."
        )

    texto = str(erro)

    if "Excel file format" in texto or "not supported" in texto:
        return (
            "Um dos arquivos não está no formato de planilha do Excel (.xlsx). "
            "Exporte o relatório novamente pelo Octalink e tente de novo."
        )

    return (
        "Não foi possível concluir o processamento. Confira os arquivos "
        "selecionados e tente novamente. O detalhe técnico foi salvo no arquivo "
        "%s, que pode ser enviado à equipe de Processos e Automação."
        % core.ARQUIVO_LOGS
    )


def registrar_falha_tecnica(pasta, erro):
    """Grava o detalhe técnico no log, longe da tela do usuário."""
    try:
        caminho = os.path.join(pasta or os.path.expanduser("~"), core.ARQUIVO_LOGS)
        with open(caminho, "a", encoding="utf-8") as f:
            f.write("\n=== Falha em %s ===\n" % datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            f.write("".join(traceback.format_exception(type(erro), erro, erro.__traceback__)))
        return caminho
    except Exception:
        return None


# ==============================
# SELEÇÃO DE ARQUIVOS
# ==============================


def selecionar_arquivos():
    caminhos, _ = QFileDialog.getOpenFileNames(
        window,
        "Selecionar arquivos de entrada",
        "",
        "Planilhas do Excel (*.xlsx *.xls)",
    )

    if not caminhos:
        return

    QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

    nao_identificados = []

    try:
        for caminho in caminhos:
            if any(a["caminho"] == caminho for a in arquivos):
                continue

            try:
                tipo, _, _, _ = core.detectar_tipo(caminho, config)
            except Exception as erro:
                registrar_falha_tecnica(os.path.dirname(caminho), erro)
                tipo = None

            if tipo is None:
                nao_identificados.append(os.path.basename(caminho))

            adicionar_linha(caminho, tipo)

    finally:
        QApplication.restoreOverrideCursor()

    if nao_identificados:
        QMessageBox.information(
            window,
            "Confira o tipo do arquivo",
            "Não foi possível reconhecer o tipo destes arquivos:\n\n%s\n\n"
            "Escolha o tipo na lista ao lado de cada um ou remova o arquivo."
            % "\n".join("- " + nome for nome in nao_identificados),
        )

    atualizar_status()


def adicionar_linha(caminho, tipo):
    """Cria a linha da lista: nome do arquivo + menu colapsável com o tipo."""
    linha = QFrame()
    linha.setStyleSheet(
        """
        QFrame {
            background-color: rgba(255,255,255,10);
            border-radius: 8px;
        }
        """
    )

    layout = QHBoxLayout(linha)
    layout.setContentsMargins(10, 6, 10, 6)
    layout.setSpacing(8)

    nome = QLabel(os.path.basename(caminho))
    nome.setToolTip(caminho)
    nome.setWordWrap(True)
    nome.setMinimumWidth(180)
    nome.setStyleSheet("background: transparent; color: #dddddd; font-size: 11px;")

    combo = QComboBox()
    combo.setStyleSheet(combo_style)
    combo.setFixedHeight(30)
    combo.setMinimumWidth(240)

    if tipo is None:
        combo.addItem(NAO_IDENTIFICADO, None)

    for valor, rotulo in OPCOES_TIPO:
        combo.addItem(rotulo, valor)

    if tipo is not None:
        indice = combo.findData(tipo)
        combo.setCurrentIndex(max(indice, 0))

    combo.currentIndexChanged.connect(atualizar_status)

    btn_remover = QPushButton("x")
    btn_remover.setFixedSize(30, 30)
    btn_remover.setStyleSheet(button_style)
    btn_remover.clicked.connect(lambda: remover_arquivo(caminho))

    layout.addWidget(nome, 1)
    layout.addWidget(combo)
    layout.addWidget(btn_remover)

    # mantém o espaçador sempre ao final da lista
    lista_layout.insertWidget(lista_layout.count() - 1, linha)

    arquivos.append({"caminho": caminho, "linha": linha, "combo": combo})


def remover_arquivo(caminho):
    for item in list(arquivos):
        if item["caminho"] == caminho:
            item["linha"].setParent(None)
            arquivos.remove(item)
    atualizar_status()


def remover_todos():
    for item in list(arquivos):
        item["linha"].setParent(None)
    arquivos.clear()
    atualizar_status()


def tipos_selecionados():
    """{tipo: caminho} apenas dos arquivos com tipo definido."""
    selecionados = {}
    for item in arquivos:
        tipo = item["combo"].currentData()
        if tipo:
            selecionados.setdefault(tipo, []).append(item["caminho"])
    return selecionados


def atualizar_status():
    """Libera o botão Executar e explica, em uma frase, o que falta."""
    selecionados = tipos_selecionados()

    duplicados = [t for t, caminhos in selecionados.items() if len(caminhos) > 1]
    sem_tipo = [
        item for item in arquivos if item["combo"].currentData() is None
        and item["combo"].currentText() == NAO_IDENTIFICADO
    ]
    relatorios = [t for t in selecionados if t in core.TIPOS_RELATORIO]

    if not arquivos:
        aviso.setText("Selecione os relatórios exportados do Octalink para começar.")
        btn_executar.setEnabled(False)
        return

    if sem_tipo:
        aviso.setText(
            "Escolha o tipo de %d arquivo(s) na lista ao lado do nome."
            % len(sem_tipo)
        )
        btn_executar.setEnabled(False)
        return

    if duplicados:
        nomes = [dict(OPCOES_TIPO).get(t, t) for t in duplicados]
        aviso.setText(
            "Há mais de um arquivo com o mesmo tipo: %s. Remova um deles."
            % ", ".join(nomes)
        )
        btn_executar.setEnabled(False)
        return

    if not relatorios:
        aviso.setText("Selecione ao menos um relatório do Octalink (pedido, produto ou RMS).")
        btn_executar.setEnabled(False)
        return

    partes = ["%d relatório(s) do Octalink" % len(relatorios)]
    if core.TIPO_FORNECEDORES in selecionados:
        partes.append("contatos dos fornecedores")
    else:
        partes.append("sem planilha de contatos (a coluna de contato ficará vazia)")
    if core.TIPO_ANTERIOR in selecionados:
        partes.append("planilha anterior para mesclagem")

    aviso.setText("Pronto para executar: " + "; ".join(partes) + ".")
    btn_executar.setEnabled(True)


# ==============================
# EXECUÇÃO
# ==============================


def executar():
    selecionados = {t: caminhos[0] for t, caminhos in tipos_selecionados().items()}

    caminho_saida, _ = QFileDialog.getSaveFileName(
        window,
        "Escolha onde salvar a planilha de Follow-Up",
        "follow_up.xlsx",
        "Planilhas do Excel (*.xlsx)",
    )

    if not caminho_saida:
        return

    if not caminho_saida.lower().endswith(".xlsx"):
        caminho_saida += ".xlsx"

    pasta = os.path.dirname(caminho_saida)
    caminho_anterior = selecionados.pop(core.TIPO_ANTERIOR, None)

    QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
    btn_executar.setEnabled(False)

    try:
        resultado = core.processar(selecionados, config, caminho_anterior=caminho_anterior)
        core.gravar_planilha(resultado, caminho_saida)

        caminho_logs = core.gravar_logs(
            resultado.ocorrencias, os.path.join(pasta, core.ARQUIVO_LOGS)
        )

    except Exception as erro:
        QApplication.restoreOverrideCursor()
        btn_executar.setEnabled(True)
        registrar_falha_tecnica(pasta, erro)

        QMessageBox.critical(window, "Não foi possível concluir", mensagem_amigavel(erro))
        return

    QApplication.restoreOverrideCursor()
    btn_executar.setEnabled(True)

    QMessageBox.information(
        window,
        "Pronto",
        core.resumo_para_usuario(resultado, caminho_saida, caminho_logs),
    )

    if abrir_checkbox.isChecked():
        abrir_arquivo(caminho_saida)


def abrir_arquivo(caminho):
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.call(["open", caminho])
        else:
            subprocess.call(["xdg-open", caminho])
    except Exception:
        QMessageBox.information(
            window,
            "Planilha salva",
            "A planilha foi salva, mas não foi possível abri-la automaticamente. "
            "Abra o arquivo pela pasta onde ele foi gravado.",
        )


# ==============================
# APLICAÇÃO
# ==============================

app = QApplication(sys.argv)

# ==============================
# WINDOW

window = QWidget()

window.setWindowTitle("AutoFollow")

caminho_icone = core.resource_path("autofollow.ico")
if os.path.exists(caminho_icone):
    window.setWindowIcon(QIcon(caminho_icone))

window.resize(960, 540)
window.setMinimumSize(960, 540)

# ==============================
# ESTILO GERAL

window.setStyleSheet("""
QWidget {
    background-color: #1e1e1e;
    font-family: "Segoe UI";
    color: white;
}

QScrollArea {
    border: none;
    background: transparent;
}

QScrollBar:vertical {
    background: #2b2b2b;
    width: 10px;
    border-radius: 5px;
}

QScrollBar::handle:vertical {
    background: #f9b02e;
    border-radius: 5px;
}

QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    height: 0px;
}

QCheckBox {
    color: #dddddd;
    font-size: 12px;
    spacing: 12px;
    background: transparent;
}
""")

# ==============================
# BACKGROUND IMAGE

bg_label = QLabel(window)

bg_pixmap = QPixmap(core.resource_path("bg_hbr.png"))

bg_label.setPixmap(bg_pixmap)
bg_label.setScaledContents(True)

# ==============================
# BOTÕES

button_style = """
QPushButton {
    background-color: #f9b02e;
    color: black;
    border: none;
    border-radius: 8px;
    padding: 10px;
    font-weight: bold;
    font-size: 14px;
}

QPushButton:hover {
    background-color: #ffd166;
}

QPushButton:pressed {
    background-color: #e69500;
}

QPushButton:disabled {
    background-color: #666666;
    color: #aaaaaa;
}
"""

combo_style = """
QComboBox {
    background-color: #2b2b2b;
    color: white;
    border: 1px solid #3a3a3a;
    border-radius: 8px;
    padding: 4px 8px;
    font-size: 11px;
}

QComboBox:hover {
    border: 1px solid #f9b02e;
}

QComboBox::drop-down {
    border: none;
    width: 18px;
}

QComboBox QAbstractItemView {
    background-color: #2b2b2b;
    color: white;
    selection-background-color: #f9b02e;
    selection-color: black;
    border: 1px solid #3a3a3a;
}
"""

# ==============================
# SCROLL AREA

scroll = QScrollArea(window)

scroll.setWidgetResizable(True)
scroll.setStyleSheet("background: transparent;")

container = QWidget()
container.setStyleSheet("background: transparent;")

scroll.setWidget(container)

# ==============================
# CARD CENTRAL

card = QFrame()

card.setObjectName("card")

card.setStyleSheet("""
#card {
    background-color: rgba(30, 30, 30, 220);
    border-radius: 18px;
    border: 1px solid rgba(255,255,255,25);
}
""")

main_layout = QVBoxLayout(container)

main_layout.addStretch()

main_layout.addWidget(card, alignment=Qt.AlignmentFlag.AlignCenter)

main_layout.addStretch()

main_layout.setContentsMargins(30, 30, 30, 30)

card_layout = QVBoxLayout(card)

card_layout.setContentsMargins(24, 24, 24, 24)
card_layout.setSpacing(14)

# ==============================
# TITLE

title = QLabel("AutoFollow")

title.setAlignment(Qt.AlignmentFlag.AlignCenter)

title.setStyleSheet("""
font-family: "Bahnschrift Condensed";
font-size: 38px;
font-weight: bold;
color: white;
padding-bottom: 2px;
background: transparent;
""")

card_layout.addWidget(title)

subtitulo = QLabel(
    "Associação dos relatórios do Octalink na planilha de Follow-Up"
)

subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)

subtitulo.setStyleSheet("""
color: rgba(255,255,255,150);
font-size: 12px;
background: transparent;
padding-bottom: 6px;
""")

card_layout.addWidget(subtitulo)

# ==============================
# BLOCO: ARQUIVOS DE ENTRADA

bloco_arquivos = QFrame()

bloco_arquivos.setStyleSheet("""
QFrame {
    background-color: rgba(255,255,255,8);
    border-radius: 12px;
}
""")

bloco_layout = QVBoxLayout(bloco_arquivos)
bloco_layout.setContentsMargins(14, 14, 14, 14)
bloco_layout.setSpacing(10)

titulo_bloco = QLabel("Arquivos de entrada (.xlsx)")

titulo_bloco.setStyleSheet("""
    color: white;
    font-size: 13px;
    font-weight: bold;
    background: transparent;
""")

bloco_layout.addWidget(titulo_bloco)

explicacao = QLabel(
    "Selecione os relatórios exportados do Octalink. O tipo de cada arquivo é "
    "reconhecido pelo cabeçalho - confira e corrija no menu ao lado, se preciso. "
    "A planilha de contatos e a planilha de Follow-Up anterior são opcionais."
)

explicacao.setWordWrap(True)

explicacao.setStyleSheet("""
    color: #aaaaaa;
    font-size: 11px;
    background: transparent;
""")

bloco_layout.addWidget(explicacao)

botoes_layout = QHBoxLayout()

btn_selecionar = QPushButton("Selecionar")
btn_selecionar.setFixedSize(120, 34)
btn_selecionar.setStyleSheet(button_style)
btn_selecionar.clicked.connect(selecionar_arquivos)

btn_remover_todos = QPushButton("Remover")
btn_remover_todos.setFixedSize(120, 34)
btn_remover_todos.setStyleSheet(button_style)
btn_remover_todos.clicked.connect(remover_todos)

botoes_layout.addWidget(btn_selecionar)
botoes_layout.addWidget(btn_remover_todos)
botoes_layout.addStretch()

bloco_layout.addLayout(botoes_layout)

# lista de arquivos carregados
lista_area = QScrollArea()
lista_area.setWidgetResizable(True)
lista_area.setMinimumHeight(140)
lista_area.setMaximumHeight(220)
lista_area.setStyleSheet("""
QScrollArea {
    background-color: #2b2b2b;
    border: 1px solid #3a3a3a;
    border-radius: 8px;
}
""")

lista_container = QWidget()
lista_container.setStyleSheet("background: transparent;")

lista_layout = QVBoxLayout(lista_container)
lista_layout.setContentsMargins(8, 8, 8, 8)
lista_layout.setSpacing(6)
lista_layout.addStretch()

lista_area.setWidget(lista_container)

bloco_layout.addWidget(lista_area)

card_layout.addWidget(bloco_arquivos)

# ==============================
# AVISO DE STATUS

aviso = QLabel("Selecione os relatórios exportados do Octalink para começar.")

aviso.setWordWrap(True)

aviso.setStyleSheet("""
    background-color: #2b2b2b;
    border: 1px solid #3a3a3a;
    border-radius: 8px;
    padding: 8px;
    color: #cccccc;
    font-size: 11px;
""")

card_layout.addWidget(aviso)


# ==============================
# CHECKBOX


class ToggleSwitch(QCheckBox):

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)

        self.setFixedHeight(28)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def paintEvent(self, event):

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # TRACK
        track_rect = QRectF(0, 4, 44, 20)

        if self.isChecked():
            painter.setBrush(QBrush(QColor("#f9b02e")))
        else:
            painter.setBrush(QBrush(QColor("#3a3a3a")))

        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(track_rect, 10, 10)

        # KNOB
        knob_x = 24 if self.isChecked() else 2

        knob_rect = QRectF(knob_x, 6, 16, 16)

        painter.setBrush(QBrush(QColor("white")))
        painter.drawEllipse(knob_rect)

        # TEXT
        painter.setPen(QColor("#dddddd"))

        painter.drawText(54, 19, self.text())


abrir_checkbox = ToggleSwitch("Abrir planilha quando estiver pronta")

card_layout.addWidget(abrir_checkbox)

# ==============================
# BOTÃO FINAL

btn_executar = QPushButton("Executar")
btn_executar.setEnabled(False)
btn_executar.setFixedSize(128, 42)
btn_executar.setStyleSheet(button_style)
btn_executar.clicked.connect(executar)

executar_layout = QHBoxLayout()

executar_layout.addStretch()
executar_layout.addWidget(btn_executar)
executar_layout.addStretch()

card_layout.addLayout(executar_layout)

# ==============================
# FOOTER

github = QLabel()
github.setAlignment(Qt.AlignmentFlag.AlignCenter)
github.setText(
        '<a href = "https://github.com/imbaTIMvel/autofollow_py">'
        'FaceList v%s - GitHub'
        '</a>' % VERSAO
    )
github.setOpenExternalLinks(True)
github.setStyleSheet("""
    QLabel {
        background-color: transparent;
        color: rgba(255,255,255,120);
        font-size: 11px;
    }
    QLabel:hover {
        color: #f9b02e;
    }
    """)
main_layout.addWidget(github)

footer = QLabel(
    "Desenvolvido por: Diretoria Administrativa Financeira — DAF"
)

footer.setAlignment(Qt.AlignmentFlag.AlignCenter)

footer.setStyleSheet("""
QLabel {
    background-color: transparent;
    color: rgba(255,255,255,120);
    font-size: 10px;
    padding-bottom: 4px;
}
""")

# ==============================
# WINDOW LAYOUT

window_layout = QVBoxLayout(window)

window_layout.setContentsMargins(0, 0, 0, 0)

window_layout.addWidget(scroll)

window.setLayout(window_layout)

# ==============================
# RESPONSIVE BACKGROUND


def resize_event(event):

    bg_label.resize(window.size())

    scroll.resize(window.size())

    card.setMaximumWidth(680)


window.resizeEvent = resize_event

# ==============================
# SHOW

main_layout.addWidget(github, alignment=Qt.AlignmentFlag.AlignBottom)
main_layout.addWidget(footer, alignment=Qt.AlignmentFlag.AlignBottom)

window.show()

sys.exit(app.exec())
