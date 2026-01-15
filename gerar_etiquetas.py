import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import os

EXCEL = "dados/Relacao_Armarios.xlsx"
LOGO = "assets/logo.png"
OUT_DIR = "output"

os.makedirs(OUT_DIR, exist_ok=True)

LARGURA = 85.6 * mm
ALTURA = 54 * mm

def gerar_pdf(df, nome_pdf, titulo):
    c = canvas.Canvas(nome_pdf, pagesize=A4)

    x = 10 * mm
    y = A4[1] - ALTURA - 15 * mm
    armario = 1

    for _, row in df.iterrows():
        if y < 20 * mm:
            c.showPage()
            y = A4[1] - ALTURA - 15 * mm
            x = 10 * mm

        c.setFont("Helvetica-Bold", 10)
        c.drawString(10*mm, A4[1] - 10*mm, titulo.upper())

        c.setLineWidth(1.5)
        c.rect(x, y, LARGURA, ALTURA)

        c.drawImage(
            LOGO,
            x + 4*mm,
            y + 16*mm,
            22*mm,
            22*mm,
            preserveAspectRatio=True,
            mask="auto"
        )

        tx = x + 30*mm
        ty = y + ALTURA - 9*mm

        c.setFont("Helvetica-Bold", 11)
        c.drawString(tx, ty, str(row["Nome Completo"]).upper())

        c.setLineWidth(1)
        c.line(tx, ty - 3*mm, x + LARGURA - 4*mm, ty - 3*mm)

        c.setFont("Helvetica", 8.2)
        c.drawString(tx, ty - 8*mm, f"Matrícula: {row['Id Contratado']}")
        c.drawString(tx, ty - 13*mm, f"Cargo: {row['Cargo']}")
        c.drawString(tx, ty - 18*mm, f"Setor: {row['Centro de Custo']}")
        c.drawString(tx, ty - 23*mm, f"Gestor: {row['Gestor']}")

        c.setFont("Helvetica-Bold", 9)
        c.drawString(tx, ty - 29*mm, f"Armário: {armario:03d}")
        armario += 1

        x += LARGURA + 6*mm
        if x + LARGURA > A4[0]:
            x = 10 * mm
            y -= ALTURA + 6 * mm

    c.save()

df = pd.read_excel(EXCEL)

df_m = df[df["Sigla Sexo"] == "M"]
df_f = df[df["Sigla Sexo"] == "F"]

gerar_pdf(df_m, f"{OUT_DIR}/etiquetas_masculino.pdf", "Vestiário Masculino")
gerar_pdf(df_f, f"{OUT_DIR}/etiquetas_feminino.pdf", "Vestiário Feminino")

print("PDFs gerados com sucesso")
