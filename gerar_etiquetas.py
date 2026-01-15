import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

# ================= CONFIGURAÇÕES =================
EXCEL = "dados/Relacao_Armarios.xlsx"
LOGO = "assets/logo.png"
OUT_DIR = "output"

os.makedirs(OUT_DIR, exist_ok=True)

# Tamanho RG (horizontal)
LARGURA = 85.6 * mm
ALTURA = 54 * mm

# ================= ESTILOS =================
style_nome = ParagraphStyle(
    name="Nome",
    fontName="Helvetica-Bold",
    fontSize=9,
    alignment=1,   # centralizado
    leading=10
)

style_info = ParagraphStyle(
    name="Info",
    fontName="Helvetica",
    fontSize=7.8,
    leading=9
)

# ================= FUNÇÃO PDF =================
def gerar_pdf(df, arquivo_saida, titulo):
    c = canvas.Canvas(arquivo_saida, pagesize=A4)

    margem_x = 10 * mm
    margem_y = 15 * mm
    espacamento = 6 * mm

    x = margem_x
    y = A4[1] - ALTURA - margem_y

    armario = 1

    for _, row in df.iterrows():

        if y < 20 * mm:
            c.showPage()
            x = margem_x
            y = A4[1] - ALTURA - margem_y

        # Título do vestiário
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margem_x, A4[1] - 10 * mm, titulo.upper())

        # Borda da etiqueta
        c.setLineWidth(1.5)
        c.rect(x, y, LARGURA, ALTURA)

        # Logo
        c.drawImage(
            LOGO,
            x + 4 * mm,
            y + ALTURA - 26 * mm,
            22 * mm,
            22 * mm,
            preserveAspectRatio=True,
            mask="auto"
        )

        # ===== NOME (com quebra automática) =====
        nome = str(row["Nome Completo"]).upper()
        p_nome = Paragraph(nome, style_nome)

        largura_nome = LARGURA - 34 * mm
        altura_nome = 18 * mm

        w, h = p_nome.wrap(largura_nome, altura_nome)
        p_nome.drawOn(
            c,
            x + 30 * mm,
            y + ALTURA - h - 6 * mm
        )

        # Linha separadora
        c.setLineWidth(0.8)
        c.line(
            x + 30 * mm,
            y + ALTURA - 26 * mm,
            x + LARGURA - 4 * mm,
            y + ALTURA - 26 * mm
        )

        # ===== INFORMAÇÕES =====
        infos = (
            f"<b>Matrícula:</b> {row['Id Contratado']}<br/>"
            f"<b>Cargo:</b> {row['Cargo']}<br/>"
            f"<b>Setor:</b> {row['Centro de Custo']}<br/>"
            f"<b>Gestor:</b> {row['Gestor']}"
        )

        p_info = Paragraph(infos, style_info)
        p_info.wrapOn(c, largura_nome, 22 * mm)
        p_info.drawOn(
            c,
            x + 30 * mm,
            y + 8 * mm
        )

        # ===== ARMÁRIO =====
        c.setFont("Helvetica-Bold", 9)
        c.drawRightString(
            x + LARGURA - 4 * mm,
            y + 4 * mm,
            f"Armário {armario:03d}"
        )

        armario += 1

        # Próxima etiqueta
        x += LARGURA + espacamento
        if x + LARGURA > A4[0]:
            x = margem_x
            y -= ALTURA + espacamento

    c.save()

# ================= EXECUÇÃO =================
df = pd.read_excel(EXCEL)

df_m = df[df["Sigla Sexo"] == "M"]
df_f = df[df["Sigla Sexo"] == "F"]

gerar_pdf(df_m, f"{OUT_DIR}/etiquetas_masculino.pdf", "Vestiário Masculino")
gerar_pdf(df_f, f"{OUT_DIR}/etiquetas_feminino.pdf", "Vestiário Feminino")

print("PDFs gerados com sucesso")

