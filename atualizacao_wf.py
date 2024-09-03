import pandas as pd

inputar = input('Inserir caminho para o arquivo excel: ')
inputar = inputar.replace('"', '')
# Ler o arquivo Excel
dados = pd.read_excel(inputar)

# Dividir o DataFrame com base na coluna 'Advogado interno'
grupos = dados.groupby('Advogado interno')

# Iterar sobre os grupos e criar arquivos Excel para cada advogado
for advogado, grupo in grupos:
    # Criar um arquivo Excel com o nome do advogado
    nome_arquivo = f'{advogado}.xlsx'
    
    # Salvar o grupo do advogado no arquivo Excel
    grupo.to_excel(nome_arquivo, index=False)

    print(f"Arquivo {nome_arquivo} criado para o advogado {advogado}")

print("Processo concluído.")