from pathlib import Path
import pandas as pd
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import lightgrey

# ================== CAMINHOS ==================
BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_EXCEL = BASE_DIR / "dados" / "base.xlsx"
ASSETS_DIR = BASE_DIR / "assets"
OUT_DIR = Path("output")
OUT_DIR.mkdir(exist_ok=True)

# ================== CONFIG ETIQUETA ==================
ETIQUETA_LARGURA = 85.6 * mm
ETIQUETA_ALTURA = 54 * mm

MARGEM_X = 6 * mm
MARGEM_Y = 6 * mm

# ================== ESTILOS ==================
style_nome = ParagraphStyle(
    "Nome",
    fontSize=11,
    leading=13,
    alignment=1,  # centralizado
    spaceAfter=4
)

style_setor = ParagraphStyle(
    "Setor",
    fontSize=8,
    leading=10,
    alignment=1,
    backColor=lightgrey,
    spaceAfter=4
)

style_gestor = ParagraphStyle(
    "Gestor",
    fontSize=7,
    leading=9,
    alignment=1,
    italic=True
)

style_armario = ParagraphStyle(
    "Armario",
    fontSize=16,
    leading=18,
    alignment=1,
    spaceAfter=6
)

# ================== FUNÇÃO PDF ==================
def gerar_pdf(df, caminho_pdf, titulo):
    c = canvas.Canvas(str(caminho_pdf), pagesize=(ETIQUETA_LARGURA, ETIQUETA_ALTURA))

    logo_path = ASSETS_DIR / "logo.png"

    for _, row in df.iterrows():
        x = MARGEM_X
        y = MARGEM_Y
        largura = ETIQUETA_LARGURA - 2 * MARGEM_X
        altura = ETIQUETA_ALTURA - 2 * MARGEM_Y

        y_atual = ETIQUETA_ALTURA - MARGEM_Y

        # LOGO (se existir)
        if logo_path.exists():
            c.drawImage(
                str(logo_path),
                x,
                y_atual - 14 * mm,
                width=14 * mm,
                height=14 * mm,
                preserveAspectRatio=True
            )

        # NÚMERO DO ARMÁRIO
        p_arm = Paragraph(f"<b>{row['armario']}</b>", style_armario)
        w, h = p_arm.wrap(largura, altura)
        p_arm.drawOn(c, x, y_atual - 22 * mm)

        # NOME
        p_nome = Paragraph(row["nome"], style_nome)
        w, h = p_nome.wrap(largura, altura)
        p_nome.drawOn(c, x, y_atual - 34 * mm)

        # SETOR
        p_setor = Paragraph(row["setor"], style_setor)
        w, h = p_setor.wrap(largura, altura)
        p_setor.drawOn(c, x, y_atual - 42 * mm)

        # GESTOR
        if pd.notna(row["gestor"]):
            p_gestor = Paragraph(f"Gestor: {row['gestor']}", style_gestor)
            w, h = p_gestor.wrap(largura, altura)
            p_gestor.drawOn(c, x, y_atual - 50 * mm)

        c.showPage()

    c.save()

# ================== LEITURA EXCEL ==================
if not ARQUIVO_EXCEL.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_EXCEL}")

df = pd.read_excel(ARQUIVO_EXCEL)

# normaliza nomes de colunas
df.columns = [c.strip().lower() for c in df.columns]

# valida colunas
colunas_obrigatorias = {"nome", "setor", "gestor", "sexo"}
if not colunas_obrigatorias.issubset(df.columns):
    raise ValueError(f"O Excel deve conter as colunas: {colunas_obrigatorias}")

# ================== ORDENAÇÃO ==================
df = df.sort_values(by=["setor", "nome"])

# ================== DIVISÃO POR SEXO ==================
df_m = df[df["sexo"].str.lower() == "masculino"].reset_index(drop=True)
df_f = df[df["sexo"].str.lower() == "feminino"].reset_index(drop=True)

df_m["armario"] = range(1, len(df_m) + 1)
df_f["armario"] = range(1, len(df_f) + 1)

# ================== GERAR PDFS ==================
if len(df_m) > 0:
    gerar_pdf(df_m, OUT_DIR / "etiquetas_masculino.pdf", "Masculino")

if len(df_f) > 0:
    gerar_pdf(df_f, OUT_DIR / "etiquetas_feminino.pdf", "Feminino")

print("PDFs gerados com sucesso!")
