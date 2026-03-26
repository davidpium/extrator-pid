import fitz  # PyMuPDF
import pandas as pd
import re
import os
import numpy as np
from sklearn.cluster import DBSCAN


# =========================
# 🔹 NORMALIZAÇÃO
# =========================
def limpar_texto(texto):
    texto = texto.upper()
    texto = re.sub(r'[^A-Z0-9\-]', '', texto)
    return texto


# =========================
# 🔹 DETECTA TAG SIMPLES
# =========================
def extrair_tag(texto):
    return re.findall(r'[A-Z]{1,4}-?\d{1,4}', texto)


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
# 🔹 RECONSTRUÇÃO POR PROXIMIDADE
# =========================
def reconstruir_tags(palavras):

    candidatos = []

    for i, w1 in enumerate(palavras):

        t1 = limpar_texto(w1[4])

        if not t1:
            continue

        # 🔹 sozinho
        for tag in extrair_tag(t1):
            candidatos.append({
                "tag": tag,
                "x": w1[0],
                "y": w1[1]
            })

        # 🔹 combina com próximas palavras (janela)
        for j in range(1, 3):  # olha até 2 palavras à frente
            if i + j >= len(palavras):
                continue

            w2 = palavras[i + j]

            # distância espacial
            dx = abs(w1[0] - w2[0])
            dy = abs(w1[1] - w2[1])

            # só combina se estiver próximo
            if dx < 50 and dy < 10:

                t2 = limpar_texto(w2[4])

                combinado = t1 + t2

                for tag in extrair_tag(combinado):
                    candidatos.append({
                        "tag": tag,
                        "x": w1[0],
                        "y": w1[1]
                    })

    return candidatos


# =========================
# 🔹 CLUSTER SUAVE
# =========================
def clusterizar(pontos):

    if not pontos:
        return []

    coords = np.array([[p["x"], p["y"]] for p in pontos])

    clustering = DBSCAN(eps=12, min_samples=1).fit(coords)
    labels = clustering.labels_

    resultado = []
    usados = set()

    for i, label in enumerate(labels):
        if label not in usados:
            resultado.append(pontos[i])
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

        # 🔥 reconstrução inteligente
        candidatos = reconstruir_tags(palavras)

        print(f"Candidatos brutos: {len(candidatos)}")

        # 🔹 filtra instrumentos
        candidatos = [c for c in candidatos if eh_instrumento(c["tag"])]

        # 🔹 cluster
        candidatos = clusterizar(candidatos)

        print(f"Após cluster: {len(candidatos)}")

        for c in candidatos:
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

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    print(f"\n✅ Gerado: {output}")

    return output
