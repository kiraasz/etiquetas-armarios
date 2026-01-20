import pandas as pd
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.colors import black, lightgrey

# ======================
# CONFIGURAÇÕES
# ======================
ARQUIVO_EXCEL = "dados/Relação Armários novos.xlsx"  # ajuste se o nome for diferente
OUT_DIR = "output"

LARGURA_RG = 85.6 * mm
ALTURA_RG = 54 * mm

MARGEM = 6 * mm

# ======================
# FUNÇÃO DE QUEBRA DE TEXTO
# ======================
def desenhar_paragraph(c, texto, estilo, x, y, largura):
    p = Paragraph(texto, estilo)
    w, h = p.wrap(largura, ALTURA_RG)
    p.drawOn(c, x, y - h)
    return h + 2

# ======================
# FUNÇÃO PRINCIPAL PDF
# ======================
def gerar_pdf(df, nome_pdf, titulo):
    c = canvas.Canvas(nome_pdf, pagesize=(LARGURA_RG, ALTURA_RG))

    # ESTILOS
    style_nome = ParagraphStyle(
        "nome",
        fontSize=11,
        leading=13,
        alignment=1,  # centralizado
        spaceAfter=4,
        fontName="Helvetica-Bold"
    )

    style_setor = ParagraphStyle(
        "setor",
        fontSize=9,
        leading=11,
        alignment=1,
        backColor=lightgrey
    )

    style_gestor = ParagraphStyle(
        "gestor",
        fontSize=8,
        leading=10,
        alignment=1,
        fontName="Helvetica-Oblique"
    )

    numero_armario = 1

    for _, row in df.iterrows():
        # BORDA
        c.setStrokeColor(black)
        c.rect(1, 1, LARGURA_RG - 2, ALTURA_RG - 2)

        y = ALTURA_RG - MARGEM

        # NÚMERO DO ARMÁRIO
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(LARGURA_RG / 2, y, f"Armário {numero_armario}")
        y -= 12

        # NOME
        y -= desenhar_paragraph(
            c,
            row["Nome"],
            style_nome,
            MARGEM,
            y,
            LARGURA_RG - 2 * MARGEM
        )

        # SETOR
        y -= desenhar_paragraph(
            c,
            row["Setor"],
            style_setor,
            MARGEM,
            y,
            LARGURA_RG - 2 * MARGEM
        )

        # GESTOR
        if pd.notna(row["Gestor"]):
            desenhar_paragraph(
                c,
                f"Gestor: {row['Gestor']}",
                style_gestor,
                MARGEM,
                y,
                LARGURA_RG - 2 * MARGEM
            )

        c.showPage()
        numero_armario += 1

    c.save()

# ======================
# LEITURA E PROCESSAMENTO
# ======================
df = pd.read_excel(ARQUIVO_EXCEL)

# NORMALIZA COLUNAS
df.columns = [c.strip().capitalize() for c in df.columns]

# ORDENA: SETOR -> NOME
df = df.sort_values(by=["Setor", "Nome"])

# MASCULINO
df_m = df[df["Genero"].str.lower() == "masculino"]
gerar_pdf(df_m, f"{OUT_DIR}/etiquetas_masculino.pdf", "Masculino")

# FEMININO
df_f = df[df["Genero"].str.lower() == "feminino"]
gerar_pdf(df_f, f"{OUT_DIR}/etiquetas_feminino.pdf", "Feminino")

print("PDFs gerados com sucesso!")


