import pandas as pd
from datetime import datetime
import win32com.client as win32
import os

## Programa para ler base de dados excel e separar os arquivos conforme a regra de negócio

# Proporções de cada grupo
proporcoes_grupo = {
    "Backoffice": 0.50,
    "Intimações": 0.30,
    "Consolidações": 0.20
}

# Proporções dentro de cada grupo
proporcoes_advogados = {
    "Backoffice": {
        "Matheus Cezar Dias": 0.20,
        "Ana Carolina Bressan da Silva": 0.30,
        "Erick Damin Bitencourt": 0.30,
        "Victor Lopes Machado Gonçalves": 0.20
    },
    "Intimações": {
        "Sarah Raquel Lopes Gonçalves": 0.28,
        "João Lucas Martins Falcão": 0.38,
        "Darlei Jacoby Kayser": 0.33
    },
    "Consolidações": {
        "Rafaella Rodrigues dos Santos Marques": 0.33,
        "Felipe Machado da Luz": 0.32,
        "Danielle Lais da Silva Lutkemeyer": 0.35
    }
}



def tratamento_dados(dados):
    # Filtrando registros que não são de interesse
    if 'Subase' in dados.columns:
        dados_filtrados = dados[dados['Subase'] != "AF 12 - Pagamento"]
    else:
        print("A coluna 'Subase' não foi encontrada.")
        return None

    # Removendo registros do mês atual
    if 'Data atualização Benner' in dados_filtrados.columns:
        dados_filtrados['Data atualização Benner'] = pd.to_datetime(dados_filtrados['Data atualização Benner'])
        mes_atual = datetime.now().month    
        dados_filtrados = dados_filtrados.loc[dados_filtrados['Data atualização Benner'].dt.month != mes_atual]
    else:
        print("A coluna 'Data atualização Benner' não foi encontrada.")
        return None
    
    # Mapeando núcleo por advogado
    nucleo_por_advogado = {
        "Ana Carolina Bressan da Silva": "Backoffice",
        "Bruno Gonçalves Barrios": "Intimações",
        "Darlei Jacoby Kayser": "Intimações",
        "Erick Damin Bitencourt": "Backoffice",
        "Felipe Machado da Luz": "Consolidações",
        "Rafaella Rodrigues dos Santos Marques": "Consolidações",
        "Danielle Lais da Silva Lutkemeyer": "Consolidações",
        "Gustavo Araujo Tavares": "Intimações",
        "Sarah Raquel Lopes Gonçalves": "Intimações",
        "Matheus Cezar Dias": "Backoffice",
        "João Lucas Martins Falcão" : "Intimação",
        "Victor Lopes Machado Gonçalves" : "Backoffice"
    }
    
    dados_filtrados['Núcleo'] = dados_filtrados['Advogado interno'].map(nucleo_por_advogado)

    return dados_filtrados

def dividir_base_por_grupos(df, proporcoes_grupo):
    base_dividida = {}
    total_registros = len(df)
    print(total_registros)

    for grupo, proporcao in proporcoes_grupo.items():
        tamanho_grupo = int(total_registros * proporcao)
        base_dividida[grupo] = df.sample(n=tamanho_grupo, random_state=42)
        df = df.drop(base_dividida[grupo].index)

    return base_dividida

def dividir_base_por_advogados(base_dividida, proporcoes_advogados):
    base_final = {}

    for grupo, df_grupo in base_dividida.items():
        total_registros_grupo = len(df_grupo)
        base_final[grupo] = {}

        for advogado, proporcao in proporcoes_advogados[grupo].items():
            tamanho_advogado = int(total_registros_grupo * proporcao)
            base_final[grupo][advogado] = df_grupo.sample(n=tamanho_advogado, random_state=42)
            df_grupo = df_grupo.drop(base_final[grupo][advogado].index)

    return base_final

def salvar_e_enviar_arquivos(base_final):
    for grupo, advogados in base_final.items():
        for advogado, df_advogado in advogados.items():
            # Salvar o arquivo
            file_name = f"{advogado}_{grupo}.xlsx"
            df_advogado.to_excel(file_name, index=False)
            


if __name__ == "__main__":

    inputar = input('Inserir caminho para o arquivo excel: ')
    inputar = inputar.replace('"', '')

    # Carregue a base de dados
    df = pd.read_excel(inputar, header=0)
    df = tratamento_dados(df)
    
    # Divida a base por grupos e depois por advogados
    base_dividida = dividir_base_por_grupos(df, proporcoes_grupo)
    base_final = dividir_base_por_advogados(base_dividida, proporcoes_advogados)
    
    # Salve os arquivos e envie por e-mail
    salvar_e_enviar_arquivos(base_final)
