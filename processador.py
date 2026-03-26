import fitz  # PyMuPDF
import pandas as pd
import re
import os


# =========================
# 🔹 FILTRO ISA
# =========================
def eh_instrumento(tag):
    prefixos = [
        "TI","PI","FI","LI",
        "PT","TT","FT","LT",
        "FC","LC","HC","HS","LS",
        "PC","TC"
    ]
    return any(tag.startswith(p) for p in prefixos)


# =========================
# 🔹 EXTRAÇÃO PRINCIPAL (SIMPLES E ESTÁVEL)
# =========================
def extrair_tags_basico(texto):
    return re.findall(r'[A-Z]{1,4}-?\d{1,4}', texto)


# =========================
# 🔹 RECUPERAÇÃO INTELIGENTE (AQUI ESTÁ O GANHO)
# =========================
def recuperar_tags(texto_total):

    texto = texto_total.upper()

    # remove sujeira leve
    texto = re.sub(r'[^A-Z0-9\- ]', ' ', texto)

    # corrige hífen quebrado
    texto = texto.replace("- ", "-")
    texto = texto.replace(" -", "-")

    # junta TI 101 → TI101
    texto = re.sub(r'([A-Z]{2,4})\s+(\d{2,4})', r'\1\2', texto)

    # remove espaços duplicados
    texto = re.sub(r'\s+', ' ', texto)

    tags = re.findall(r'[A-Z]{1,4}-?\d{1,4}', texto)

    return tags


# =========================
# 🔹 PROCESSAMENTO PRINCIPAL
# =========================
def processar_pdf(pdf_path):

    print(f"\n📂 Processando: {pdf_path}")

    doc = fitz.open(pdf_path)
    resultados = []

    for pagina_num, pagina in enumerate(doc):
        print(f"\n📄 Página {pagina_num + 1}")

        texto = pagina.get_text()

        # 🔹 EXTRAÇÃO BASE (já funciona bem)
        tags_base = extrair_tags_basico(texto)

        # 🔹 RECUPERAÇÃO (corrige os que faltam)
        tags_extra = recuperar_tags(texto)

        # 🔹 JUNTA TUDO
        todas_tags = tags_base + tags_extra

        print(f"Total bruto encontrado: {len(todas_tags)}")

        # 🔹 FILTRA
        for tag in todas_tags:
            if eh_instrumento(tag):
                tipo = re.match(r'[A-Z]+', tag).group()

                resultados.append({
                    "Tipo": tipo,
                    "Tag": tag
                })

    if not resultados:
        print("❌ Nenhum instrumento encontrado")
        return None

    df = pd.DataFrame(resultados)

    # 🔹 remove duplicados
    df = df.drop_duplicates()

    df = df.sort_values(by=["Tipo", "Tag"])

    print("\n📊 Resumo:")
    print(df["Tipo"].value_counts())

    os.makedirs("temp", exist_ok=True)

    output = os.path.join(
        "temp",
        os.path.basename(pdf_path).replace(".pdf", "_instrumentos.xlsx")
    )

    # 🔹 escrita segura
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    print(f"\n✅ Gerado: {output}")

    return output
