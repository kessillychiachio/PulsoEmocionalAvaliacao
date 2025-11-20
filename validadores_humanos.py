import pandas as pd
import os

def _carregar_base():
    try:
        return pd.read_excel('resultados_da_ia.xlsx')
    except FileNotFoundError:
        print("Erro: O arquivo 'resultados_da_ia.xlsx' não foi encontrado.")
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
            return

    df_para_avaliadores = df_base[['Texto']].copy()
    df_para_avaliadores.drop_duplicates(inplace=True)
    df_para_avaliadores.reset_index(drop=True, inplace=True)

    df_avaliador1 = df_para_avaliadores.copy()
    df_avaliador1['Avaliação Humana (avaliador1)'] = ''
    
    try:
        df_avaliador1.to_excel('avaliacoes_1.xlsx', index=False, engine='openpyxl')
        print("Planilha 'avaliacoes_1.xlsx' criada com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar 'avaliacoes_1.xlsx': {e}")
    df_avaliador2 = df_para_avaliadores.copy()
    df_avaliador2['Avaliação Humana (avaliador2)'] = ''
    
    try:
        df_avaliador2.to_excel('avaliacoes_2.xlsx', index=False, engine='openpyxl')
        print("Planilha 'avaliacoes_2.xlsx' criada com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar 'avaliacoes_2.xlsx': {e}")

    print("\nInstruções: Envie as planilhas para cada avaliador. Após o preenchimento, use o script 'coeficiente_kappa.py'.")

if __name__ == "__main__":
    preparar_planilhas_para_avaliacao()
