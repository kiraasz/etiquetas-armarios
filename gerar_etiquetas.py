import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

# ================= CONFIGURAÇÕES =================
EXCEL = "dados/Relação Armários novos.xlsx"
OUT_DIR = "output"
LOGO = "assets/logo.png"

os.makedirs(OUT_DIR, exist_ok=True)

# Tamanho RG (horizontal)
ETIQUETA_LARG = 8.8 * 28.35  # cm -> pt
ETIQUETA_ALT = 5.7 * 28.35

MARGEM_X = 20
MARGEM_Y = 20
COLUNAS = 2
ESPACO_X = 10
ESPACO_Y = 10

# ================= ESTILOS =================
style_nome = ParagraphStyle(
    name="Nome",
    fontName="Helvetica-Bold",
    fontSize=11,
    alignment=1,
    leading=12
)

style_setor = ParagraphStyle(
    name="Setor",
    fontName="Helvetica-Bold",
    fontSize=8,
    alignment=1
)

style_info = ParagraphStyle(
    name="Info",
    fontName="Helvetica",
    fontSize=7,
    alignment=0
)

style_gestor = ParagraphStyle(
    name="Gestor",
    fontName="Helvetica-Oblique",
    fontSize=7,
    alignment=0
)

# ================= FUNÇÃO PDF =================
def gerar_pdf(df, arquivo, titulo):
    c = canvas.Canvas(arquivo, pagesize=A4)
    largura_pagina, altura_pagina = A4

    x = MARGEM_X
    y = altura_pagina - MARGEM_Y - ETIQUETA_ALT
    coluna = 0
    armario = 1

    for _, row in df.iterrows():
        if y < MARGEM_Y:
            c.showPage()
            y = altura_pagina - MARGEM_Y - ETIQUETA_ALT
            x = MARGEM_X
            coluna = 0

        # Borda
        c.setStrokeColor(colors.black)
        c.rect(x, y, ETIQUETA_LARG, ETIQUETA_ALT, stroke=1, fill=0)

        # Logo
        if os.path.exists(LOGO):
            c.drawImage(LOGO, x + 5, y + ETIQUETA_ALT - 25, width=50, height=20, preserveAspectRatio=True)

        # Número do armário
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(x + ETIQUETA_LARG - 6, y + ETIQUETA_ALT - 18, f"{armario:03d}")

        # Faixa do setor (cinza)
        c.setFillColor(colors.lightgrey)
        c.rect(x + 4, y + ETIQUETA_ALT - 42, ETIQUETA_LARG - 8, 16, stroke=0, fill=1)
        c.setFillColor(colors.black)

        p_setor = Paragraph(row["Centro de Custo"], style_setor)
        w, h = p_setor.wrap(ETIQUETA_LARG - 12, 16)
        p_setor.drawOn(c, x + 6, y + ETIQUETA_ALT - 40)

        # Nome (grande, centralizado, quebra automática)
        p_nome = Paragraph(row["Nome Completo"], style_nome)
        w, h = p_nome.wrap(ETIQUETA_LARG - 12, 40)
        p_nome.drawOn(c, x + 6, y + ETIQUETA_ALT - 85)

        # Matrícula
        p_mat = Paragraph(f"Matrícula: {row['Matricula']}", style_info)
        p_mat.wrapOn(c, ETIQUETA_LARG - 12, 10)
        p_mat.drawOn(c, x + 6, y + 28)

        # Gestor (itálico)
        p_gestor = Paragraph(f"Gestor: {row['Gestor']}", style_gestor)
        p_gestor.wrapOn(c, ETIQUETA_LARG - 12, 10)
        p_gestor.drawOn(c, x + 6, y + 14)

        # Próxima posição
        armario += 1
        coluna += 1

        if coluna == COLUNAS:
            coluna = 0
            x = MARGEM_X
            y -= ETIQUETA_ALT + ESPACO_Y
        else:
            x += ETIQUETA_LARG + ESPACO_X

    c.save()

# ================= LEITURA E ORGANIZAÇÃO =================
df = pd.read_excel(EXCEL)

df["Centro de Custo"] = df["Centro de Custo"].str.upper().str.strip()
df["Nome Completo"] = df["Nome Completo"].str.upper().str.strip()

# Masculino
df_m = df[df["Sigla Sexo"] == "M"].copy()
df_m = df_m.sort_values(by=["Centro de Custo", "Nome Completo"]).reset_index(drop=True)

# Feminino
df_f = df[df["Sigla Sexo"] == "F"].copy()
df_f = df_f.sort_values(by=["Centro de Custo", "Nome Completo"]).reset_index(drop=True)

# Geração
gerar_pdf(df_m, f"{OUT_DIR}/etiquetas_masculino.pdf", "Masculino")
gerar_pdf(df_f, f"{OUT_DIR}/etiquetas_feminino.pdf", "Feminino")

print("PDFs gerados com sucesso")



