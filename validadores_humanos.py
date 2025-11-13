import pandas as pd
import os

def _carregar_base():
    try:
        return pd.read_excel('resultados_da_ia.xlsx')
    except FileNotFoundError:
        try:
            return pd.read_csv('resultados_da_ia.csv')
        except FileNotFoundError:
            print("Erro: Nenhum arquivo de resultados encontrado ('resultados_da_ia.xlsx' ou 'resultados_da_ia.csv').")
            print("Gere os resultados primeiro executando: python classificador_youtube.py --video-id SEU_ID")
            return None

def preparar_planilhas_para_avaliacao():
    df_base = _carregar_base()
    if df_base is None:
        return

    if 'Texto' not in df_base.columns:
        if 'Mensagem Coletada' in df_base.columns:
            df_base['Texto'] = df_base['Mensagem Coletada']
        else:
            print("Erro: A coluna 'Texto' não foi encontrada na base de resultados.")
            print("Verifique se a planilha contém a coluna 'Texto' ou 'Mensagem Coletada'.")
            return

    df_para_avaliadores = df_base[['Texto']].copy()
    df_para_avaliadores.drop_duplicates(inplace=True)
    df_para_avaliadores.reset_index(drop=True, inplace=True)

    df_avaliador1 = df_para_avaliadores.copy()
    df_avaliador1 ['Avaliação Humana (avaliador1)'] = ''
    nome_arquivo_avaliador1 = 'avaliacoes_1.xlsx'
    try:
        df_avaliador1.to_excel(nome_arquivo_avaliador1, index=False, engine='openpyxl')
    except ImportError:
        nome_arquivo_avaliador1 = 'avaliacoes_1.csv'
        df_avaliador1.to_csv(nome_arquivo_avaliador1, index=False, encoding='utf-8')
    print(f"Planilha '{nome_arquivo_avaliador1}' criada com sucesso para avaliador1!")

    df_avaliador2 = df_para_avaliadores.copy()
    df_avaliador2['Avaliação Humana (avaliador2)'] = ''
    nome_arquivo_avaliador2 = 'avaliacoes_2.xlsx'
    try:
        df_avaliador2.to_excel(nome_arquivo_avaliador2, index=False, engine='openpyxl')
    except ImportError:
        nome_arquivo_avaliador2 = 'avaliacoes_2.csv'
        df_avaliador2.to_csv(nome_arquivo_avaliador2, index=False, encoding='utf-8')
    print(f"Planilha '{nome_arquivo_avaliador2}' criada com sucesso para avaliador2!")

    print("\nInstruções: Envie as planilhas para cada avaliador. Após o preenchimento, use o script 'coeficiente_kappa.py'.")

if __name__ == "__main__":
    preparar_planilhas_para_avaliacao()
