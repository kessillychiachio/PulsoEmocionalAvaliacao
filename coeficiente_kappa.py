import pandas as pd
from sklearn.metrics import cohen_kappa_score

def _load_table(xlsx: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(xlsx)
    except FileNotFoundError:
        return None

def _normalize_label(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.upper()

def interpretar_kappa(kappa_score: float) -> str:
    if kappa_score >= 0.81:
        return "Quase Perfeita"
    if kappa_score >= 0.61:
        return "Substancial"
    if kappa_score >= 0.41:
        return "Moderada"
    if kappa_score >= 0.21:
        return "Razoável"
    if kappa_score >= 0.01:
        return "Leve"
    return "Ruim"

def consolidar_e_calcular_kappa():
    base = _load_table("resultados_da_ia.xlsx")
    avaliador1 = _load_table("avaliacoes_1.xlsx")
    avaliador2 = _load_table("avaliacoes_2.xlsx")

    if base is None or avaliador1 is None or avaliador2 is None:
        print("ERRO: Arquivos necessários não encontrados. Gere as planilhas e preencha as avaliações.")
        return
    
    base.columns = base.columns.str.strip()
    avaliador1.columns = avaliador1.columns.str.strip()
    avaliador2.columns = avaliador2.columns.str.strip()

    if "Texto" not in base.columns:
        print("ERRO: A planilha base precisa conter a coluna 'Texto'.")
        return

    if "Polaridade" not in base.columns:
        print("ERRO: A planilha base precisa conter a coluna 'Polaridade'.")
        return

    if "Texto" not in avaliador1.columns or "Avaliação Humana (avaliador1)" not in avaliador1.columns:
        print("ERRO: A planilha do avaliador1 deve conter as colunas 'Texto' e 'Avaliação Humana (avaliador1)'.")
        return
    if "Texto" not in avaliador2.columns or "Avaliação Humana (avaliador2)" not in avaliador2.columns:
        print("ERRO: A planilha do avaliador2 deve conter as colunas 'Texto' e 'Avaliação Humana (avaliador2)'.")
        return
    
    avaliador1 = avaliador1.rename(columns={"Avaliação Humana (avaliador1)": "Avaliacao_1"})
    avaliador2 = avaliador2.rename(columns={"Avaliação Humana (avaliador2)": "Avaliacao_2"})

    base["Texto"] = base["Texto"].astype(str)
    avaliador1["Texto"] = avaliador1["Texto"].astype(str)
    avaliador2["Texto"] = avaliador2["Texto"].astype(str)

    df = base.merge(avaliador1[["Texto", "Avaliacao_1"]], on="Texto", how="left")
    df = df.merge(avaliador2[["Texto", "Avaliacao_2"]], on="Texto", how="left")

    df["Polaridade"] = _normalize_label(df["Polaridade"])
    df["Avaliacao_1"] = _normalize_label(df["Avaliacao_1"])
    df["Avaliacao_2"] = _normalize_label(df["Avaliacao_2"]) 

    print("Amostra consolidada:")
    print(df.head())
    print("-" * 50)

    df_k = df[(df["Avaliacao_1"] != "") & (df["Avaliacao_2"] != "")]
    if df_k.empty:
        print("ERRO: Nenhuma avaliação humana preenchida nas duas planilhas.")
        return

    kappa_ia_avaliador1 = cohen_kappa_score(df_k["Polaridade"], df_k["Avaliacao_1"])
    kappa_ia_avaliador2 = cohen_kappa_score(df_k["Polaridade"], df_k["Avaliacao_2"])
    kappa_avaliador1_avaliador2 = cohen_kappa_score(df_k["Avaliacao_1"], df_k["Avaliacao_2"])

    print(f"Coeficiente de Kappa (IA vs. avaliador1): {kappa_ia_avaliador1:.4f} -> {interpretar_kappa(kappa_ia_avaliador1)}")
    print(f"Coeficiente de Kappa (IA vs. avaliador2): {kappa_ia_avaliador2:.4f} -> {interpretar_kappa(kappa_ia_avaliador2)}")
    print(f"Coeficiente de Kappa (avaliador1 vs. avaliador2): {kappa_avaliador1_avaliador2:.4f} -> {interpretar_kappa(kappa_avaliador1_avaliador2)}")

    try:
        df.to_excel("resultados_consolidados_finais.xlsx", index=False, engine="openpyxl")
        print("Planilha 'resultados_consolidados_finais.xlsx' criada.")
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")


if __name__ == "__main__":
    consolidar_e_calcular_kappa()
