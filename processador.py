import fitz
import pandas as pd
import re
import numpy as np
from sklearn.cluster import DBSCAN
import os

EPS = 60
MIN_SAMPLES = 2


def processar_pdf(pdf_path):
    print(f"\n📂 Processando: {pdf_path}")

    doc = fitz.open(pdf_path)
    instrumentos = []

    for page_num, page in enumerate(doc):
        print(f"\n📄 Página {page_num+1}")

        words = page.get_text("words")

        tokens = []
        for w in words:
            texto = re.sub(r'[^A-Z0-9\-]', '', w[4].upper())
            tokens.append((texto, w[0], w[1]))

        candidatos = []

        # =========================
        # 🔹 1. SEU MÉTODO ORIGINAL (mantido)
        # =========================
        for i in range(len(tokens) - 1):
            t1, x1, y1 = tokens[i]
            t2, x2, y2 = tokens[i + 1]

            if re.match(r'^[A-Z]{1,3}$', t1) and re.match(r'^\d{3,4}$', t2):
                if re.match(r'^[TPFALH][A-Z]?$', t1):
                    candidatos.append({
                        "Tipo": t1,
                        "Tag": t2,
                        "x": x1,
                        "y": y1,
                        "Pagina": page_num + 1
                    })

        # =========================
        # 🔹 2. NOVO: detectar TI101 direto
        # =========================
        for t, x, y in tokens:
            match = re.match(r'^([A-Z]{1,3})-?(\d{3,4})$', t)
            if match:
                tipo, tag = match.groups()

                if re.match(r'^[TPFALH][A-Z]?$', tipo):
                    candidatos.append({
                        "Tipo": tipo,
                        "Tag": tag,
                        "x": x,
                        "y": y,
                        "Pagina": page_num + 1
                    })

        # =========================
        # 🔹 3. NOVO: juntar tokens próximos (TI + -101)
        # =========================
        for i in range(len(tokens) - 1):
            t1, x1, y1 = tokens[i]
            t2, x2, y2 = tokens[i + 1]

            combinado = t1 + t2

            match = re.match(r'^([A-Z]{1,3})-?(\d{3,4})$', combinado)
            if match:
                tipo, tag = match.groups()

                if re.match(r'^[TPFALH][A-Z]?$', tipo):
                    candidatos.append({
                        "Tipo": tipo,
                        "Tag": tag,
                        "x": x1,
                        "y": y1,
                        "Pagina": page_num + 1
                    })

        print(f"Candidatos brutos: {len(candidatos)}")

        if not candidatos:
            continue

        # =========================
        # CLUSTER (mantido)
        # =========================
        coords = np.array([[c["x"], c["y"]] for c in candidatos])

        clustering = DBSCAN(eps=EPS, min_samples=MIN_SAMPLES).fit(coords)
        labels = clustering.labels_

        clusters_validos = []

        for label in set(labels):
            grupo = [candidatos[i] for i in range(len(labels)) if labels[i] == label]

            if label == -1:
                for g in grupo:
                    if re.match(r'^[TPFALH][A-Z]?$', g["Tipo"]):
                        clusters_validos.append(g)
                continue

            if len(grupo) >= MIN_SAMPLES:
                clusters_validos.extend(grupo)

        print(f"Após cluster: {len(clusters_validos)}")

        instrumentos.extend(clusters_validos)

    if not instrumentos:
        print("❌ Nenhum instrumento encontrado")
        return None

    df = pd.DataFrame(instrumentos)

    df["Instrumento"] = df["Tipo"] + df["Tag"]
    df = df.drop_duplicates(subset=["Instrumento"])
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
