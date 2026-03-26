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
# 🔹 DETECÇÃO HÍBRIDA
# =========================
def extrair_tags_inteligente(texto):

    # 🔒 modo preciso (preserva PDF bom)
    padrao_estrito = r'\b[A-Z]{2,3}-\d{2,4}\b'
    estritos = re.findall(padrao_estrito, texto)

    # 🔓 modo flexível (salva PDFs ruins)
    padrao_flex = r'\b[A-Z]{1,4}-?\d{1,4}\b'
    flex = re.findall(padrao_flex, texto)

    # 🎯 decisão inteligente
    if len(estritos) >= 10:
        return estritos
    else:
        return flex


# =========================
# 🔹 FILTRO INTELIGENTE
# =========================
def eh_instrumento(tag):

    prefixos_fortes = [
        "TI","PI","FI","LI",
        "PT","TT","FT","LT"
    ]

    prefixos_medios = [
        "FC","LC","HC","HS","LS",
        "PC","TC"
    ]

    if any(tag.startswith(p) for p in prefixos_fortes):
        return True

    if any(tag.startswith(p) for p in prefixos_medios):
        return len(tag) >= 4

    return False


def filtro_final(tag):
    return (
        any(c.isdigit() for c in tag) and
        2 <= len(tag) <= 10
    )


# =========================
# 🔹 CLUSTER (SUAVE)
# =========================
def clusterizar(palavras):

    if len(palavras) == 0:
        return []

    coords = np.array([[p["x"], p["y"]] for p in palavras])

    clustering = DBSCAN(eps=10, min_samples=1).fit(coords)
    labels = clustering.labels_

    resultado = []
    usados = set()

    for i, label in enumerate(labels):
        if label not in usados:
            resultado.append(palavras[i])
            usados.add(label)

    return resultado


# =========================
# 🔹 PROCESSAMENTO PRINCIPAL
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

            tags = extrair_tags_inteligente(texto)

            for tag in tags:
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

    # escrita segura
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    print(f"\n✅ Gerado: {output}")

    return output
