import pandas as pd
import os
from glob import glob

# Input do usuário
codigo_assessor_anterior = input("Digite o código do assessor que deseja verificar: ")
codigo_assessor_atual = "4062851.0"  # Substitua pelo código real de "Wert Digital"

# Diretório contendo os arquivos históricos
diretorio_historico = r"C:\Users\gabriel.martins\Downloads\2021.11.30_BaseBTG"
arquivo_atual = r"C:\Users\gabriel.martins\Downloads\Base BTG.new.xlsx"  # Arquivo mais recente

# Lista de arquivos Excel no diretório histórico
arquivos_historicos = glob(os.path.join(diretorio_historico, "*.xlsx"))
print(f"Arquivos encontrados no diretório histórico: {arquivos_historicos}")

# Consolidar os dados históricos
dados_historicos = pd.DataFrame()

# Carregar os arquivos históricos
for arquivo in arquivos_historicos:
    try:
        print(f"Carregando arquivo histórico: {arquivo}")
        df = pd.read_excel(arquivo, sheet_name=0)  # Lê a primeira aba de cada arquivo
        df['Fonte'] = os.path.basename(arquivo)  # Adiciona a origem do arquivo
        dados_historicos = pd.concat([dados_historicos, df], ignore_index=True)
    except Exception as e:
        print(f"Erro ao carregar o arquivo histórico {arquivo}: {e}")

# Deduplicar dados históricos
dados_historicos = dados_historicos.drop_duplicates(subset=["Conta", "Código do Assessor"])

# Carregar o arquivo atual
try:
    print(f"Carregando arquivo atual: {arquivo_atual}")
    dados_atuais = pd.read_excel(arquivo_atual, sheet_name=0)
except Exception as e:
    raise ValueError(f"Erro ao carregar o arquivo atual {arquivo_atual}: {e}")

# Diagnóstico inicial: Exportar os dados carregados para análise
diretorio_downloads = os.path.expanduser("~/Downloads")

# Garantir que os códigos de assessor sejam tratados como strings
dados_historicos["Código do Assessor"] = dados_historicos["Código do Assessor"].astype(str)
dados_atuais["Código do Assessor"] = dados_atuais["Código do Assessor"].astype(str)
codigo_assessor_anterior = str(codigo_assessor_anterior)
codigo_assessor_atual = str(codigo_assessor_atual)

# Verificar se as colunas necessárias existem em ambos os conjuntos de dados
colunas_necessarias = {"Código do Assessor", "Conta"}
for df, tipo in [(dados_historicos, "históricos"), (dados_atuais, "atuais")]:
    if not colunas_necessarias.issubset(df.columns):
        raise ValueError(f"As colunas necessárias {colunas_necessarias} não estão presentes nos dados {tipo}.")

# Filtrar clientes que atualmente pertencem ao código do assessor atual
clientes_atual = dados_atuais[dados_atuais["Código do Assessor"] == codigo_assessor_atual]
numero_clientes_atual = len(clientes_atual)
print(f"Número de clientes atuais com assessor {codigo_assessor_atual}: {numero_clientes_atual}")

# Identificar clientes que estavam com o código do assessor anterior nos históricos
clientes_historicos = dados_historicos[dados_historicos["Código do Assessor"] == codigo_assessor_anterior]
print(f"Número de clientes históricos com assessor {codigo_assessor_anterior}: {len(clientes_historicos)}")

# Salvar o histórico de clientes do assessor
historico_saida = os.path.join(diretorio_downloads, f"historico_clientes_{codigo_assessor_anterior}.xlsx")
clientes_historicos.to_excel(historico_saida, index=False)
print(f"Histórico de clientes salvo em '{historico_saida}'.")

# Comparar os clientes entre os dois DataFrames usando o identificador único (ex.: Conta)
if not clientes_atual.empty and not clientes_historicos.empty:
    clientes_filtrados = pd.merge(
        clientes_atual,
        clientes_historicos,
        on="Conta",
        suffixes=("_Atual", "_Historico")
    )
    print(f"Número de clientes filtrados: {len(clientes_filtrados)}")

    # Construir o caminho completo para salvar o arquivo no diretório de downloads
    arquivo_saida = os.path.join(diretorio_downloads, f"clientes_filtrados_{codigo_assessor_anterior}.xlsx")

    # Exportar os resultados para Excel
    clientes_filtrados.to_excel(arquivo_saida, index=False)
    print(f"Resultado salvo em '{arquivo_saida}'.")
else:
    print("Nenhum cliente encontrado que atenda aos critérios especificados.")
    