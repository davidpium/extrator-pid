import fitz  # PyMuPDF
import pandas as pd
import re
import os
from sklearn.cluster import DBSCAN
import numpy as np


# =========================
# 🔹 NORMALIZAÇÃO
# =========================
def limpar_texto(texto):
    texto = texto.upper()
    texto = re.sub(r'[^A-Z0-9\-]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto


# =========================
# 🔹 DETECÇÃO FLEXÍVEL DE TAG
# =========================
def extrair_candidatos(texto):
    padrao = r'\b[A-Z]{1,4}-?\d{1,4}\b'
    return re.findall(padrao, texto)


# =========================
# 🔹 FILTRO INTELIGENTE (ISA-like)
# =========================
def eh_instrumento(tag):
    prefixos = [
        "TI","PI","FI","LI","AI","DI",
        "PT","TT","FT","LT",
        "FC","LC","HC","HS","LS",
        "PC","TC","SC"
    ]
    return any(tag.startswith(p) for p in prefixos)


def filtro_final(tag):
    return (
        any(c.isdigit() for c in tag) and
        2 <= len(tag) <= 10
    )


# =========================
# 🔹 CLUSTER (REMOVE DUPLICADOS ESPACIAIS)
# =========================
def clusterizar(palavras):
    if len(palavras) == 0:
        return []

    coords = np.array([[p["x"], p["y"]] for p in palavras])

    clustering = DBSCAN(eps=20, min_samples=1).fit(coords)
    labels = clustering.labels_

    resultado = []
    usados = set()

    for i, label in enumerate(labels):
        if label not in usados:
            resultado.append(palavras[i])
            usados.add(label)

    return resultado


# =========================
# 🔹 EXTRAÇÃO PRINCIPAL
# =========================
def processar_pdf(pdf_path):

    print(f"\n📂 Processando: {pdf_path}")

    doc = fitz.open(pdf_path)
    resultados = []

    for pagina_num, pagina in enumerate(doc):
        print(f"\n📄 Página {pagina_num + 1}")

        palavras = pagina.get_text("words")

        candidatos = []

        for w in palavras:
            texto = limpar_texto(w[4])

            encontrados = extrair_candidatos(texto)

            for tag in encontrados:
                if eh_instrumento(tag) and filtro_final(tag):
                    candidatos.append({
                        "tag": tag,
                        "x": w[0],
                        "y": w[1]
                    })

        print(f"Candidatos brutos: {len(candidatos)}")

        candidatos_cluster = clusterizar(candidatos)

        print(f"Após cluster: {len(candidatos_cluster)}")

        for c in candidatos_cluster:
            tipo = re.match(r'[A-Z]+', c["tag"]).group()
            resultados.append({
                "Tipo": tipo,
                "Tag": c["tag"]
            })

    if not resultados:
        print("❌ Nenhum instrumento encontrado")
        return None

    df = pd.DataFrame(resultados)

    df = df.drop_duplicates()
    df = df.sort_values(by=["Tipo", "Tag"])

    print("\n📊 Resumo:")
    print(df["Tipo"].value_counts())

    os.makedirs("temp", exist_ok=True)

    output = os.path.join(
        "temp",
        os.path.basename(pdf_path).replace(".pdf", "_instrumentos.xlsx")
    )

    # Escrita segura (evita corrupção)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    print(f"\n✅ Gerado: {output}")

    return output
