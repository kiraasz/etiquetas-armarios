from pathlib import Path
import pandas as pd
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import lightgrey

# ==================================================
# CAMINHOS
# ==================================================
BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_EXCEL = BASE_DIR / "dados" / "base.xlsx"
ASSETS_DIR = BASE_DIR / "assets"
OUT_DIR = BASE_DIR / "output"
OUT_DIR.mkdir(exist_ok=True)

# ==================================================
# CONFIGURAÇÃO – TAMANHO EXATO DO NOVO RG (ID-1)
# ==================================================
ETIQUETA_LARGURA = 85.6 * mm
ETIQUETA_ALTURA = 53.98 * mm

# Margens mínimas (capinha justa)
MARGEM_X = 3 * mm
MARGEM_Y = 3 * mm

# ==================================================
# ESTILOS (OTIMIZADOS PARA LEITURA E ESPAÇO)
# ==================================================
style_armario = ParagraphStyle(
    "Armario",
    fontSize=19,
    leading=21,
    alignment=1
)

style_nome = ParagraphStyle(
    "Nome",
    fontSize=12,
    leading=14,
    alignment=1
)

style_setor = ParagraphStyle(
    "Setor",
    fontSize=9,
    leading=11,
    alignment=1,
    backColor=lightgrey
)

style_gestor = ParagraphStyle(
    "Gestor",
    fontSize=8,
    leading=10,
    alignment=1,
    italic=True
)

# ==================================================
# FUNÇÃO PARA GERAR PDF
# ==================================================
def gerar_pdf(df, caminho_pdf):
    c = canvas.Canvas(
        str(caminho_pdf),
        pagesize=(ETIQUETA_LARGURA, ETIQUETA_ALTURA)
    )

    logo_path = ASSETS_DIR / "logo.png"
    largura_util = ETIQUETA_LARGURA - 2 * MARGEM_X

    for _, row in df.iterrows():
        y_topo = ETIQUETA_ALTURA - MARGEM_Y
        x = MARGEM_X

        # LOGO
        if logo_path.exists():
            c.drawImage(
                str(logo_path),
                x,
                y_topo - 11 * mm,
                width=11 * mm,
                height=11 * mm,
                preserveAspectRatio=True
            )

        # NÚMERO DO ARMÁRIO
        p_arm = Paragraph(f"<b>{row['armario']}</b>", style_armario)
        p_arm.wrap(largura_util, ETIQUETA_ALTURA)
        p_arm.drawOn(c, x, y_topo - 19 * mm)

        # NOME
        p_nome = Paragraph(row["nome"], style_nome)
        p_nome.wrap(largura_util, ETIQUETA_ALTURA)
        p_nome.drawOn(c, x, y_topo - 31 * mm)

        # SETOR
        p_setor = Paragraph(row["setor"], style_setor)
        p_setor.wrap(largura_util, ETIQUETA_ALTURA)
        p_setor.drawOn(c, x, y_topo - 40 * mm)

        # GESTOR
        if pd.notna(row["gestor"]) and str(row["gestor"]).strip():
            p_gestor = Paragraph(f"Gestor: {row['gestor']}", style_gestor)
            p_gestor.wrap(largura_util, ETIQUETA_ALTURA)
            p_gestor.drawOn(c, x, y_topo - 49 * mm)

        c.showPage()

    c.save()

# ==================================================
# LEITURA DO EXCEL
# ==================================================
df = pd.read_excel(ARQUIVO_EXCEL)
df.columns = [c.strip().lower() for c in df.columns]

# Normalização do campo sexo (robusto)
df["sexo"] = (
    df["sexo"]
    .astype(str)
    .str.strip()
    .str.lower()
    .replace({
        "m": "masculino",
        "masc": "masculino",
        "f": "feminino",
        "fem": "feminino"
    })
)

# ==================================================
# ORDENAÇÃO
# ==================================================
df = df.sort_values(by=["setor", "nome"])

# ==================================================
# DIVISÃO POR SEXO
# ==================================================
df_m = df[df["sexo"] == "masculino"].reset_index(drop=True)
df_f = df[df["sexo"] == "feminino"].reset_index(drop=True)

df_m["armario"] = range(1, len(df_m) + 1)
df_f["armario"] = range(1, len(df_f) + 1)

print("Total registros:", len(df))
print("Masculino:", len(df_m))
print("Feminino:", len(df_f))

# ==================================================
# GERAÇÃO DOS PDFs
# ==================================================
if not df_m.empty:
    gerar_pdf(df_m, OUT_DIR / "etiquetas_masculino.pdf")

if not df_f.empty:
    gerar_pdf(df_f, OUT_DIR / "etiquetas_feminino.pdf")

print("PDFs gerados:")
for f in OUT_DIR.glob("*.pdf"):
    print(f)

