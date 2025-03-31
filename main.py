import pandas as pd
import os
from glob import glob


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

# Garantir que os códigos de assessor sejam tratados como strings
dados_historicos["Código do Assessor"] = dados_historicos["Código do Assessor"].astype(str)
dados_atuais["Código do Assessor"] = dados_atuais["Código do Assessor"].astype(str)

# Verificar se as colunas necessárias existem em ambos os conjuntos de dados
colunas_necessarias = {"Código do Assessor", "Conta"}
for df, tipo in [(dados_historicos, "históricos"), (dados_atuais, "atuais")]:
    if not colunas_necessarias.issubset(df.columns):
        raise ValueError(f"As colunas necessárias {colunas_necessarias} não estão presentes nos dados {tipo}.")

# Início do loop
while True:
    codigo_assessor_anterior = input("Digite o código do assessor anterior (ou 'sair' para encerrar): ").strip()
    if codigo_assessor_anterior.lower() == "sair":
        print("Encerrando o programa.")
        break

    codigo_assessor_atual = "4062851.0"  # Código do assessor atual (Wert Digital)

    # Filtrar clientes que atualmente pertencem ao código do assessor atual
    clientes_atual = dados_atuais[dados_atuais["Código do Assessor"] == codigo_assessor_atual]

    # Identificar clientes que estavam com o código do assessor anterior nos históricos
    clientes_historicos = dados_historicos[dados_historicos["Código do Assessor"] == codigo_assessor_anterior]

    # Salvar o histórico de clientes do assessor anterior
    if not clientes_historicos.empty:
        diretorio_downloads = os.path.expanduser("~/Downloads")
        historico_saida = os.path.join(diretorio_downloads, f"historico_clientes_{codigo_assessor_anterior}.xlsx")
        clientes_historicos.to_excel(historico_saida, index=False)
        print(f"Histórico de clientes do assessor {codigo_assessor_anterior} salvo em '{historico_saida}'.")
    else:
        print(f"Nenhum cliente encontrado para o assessor anterior {codigo_assessor_anterior}.")

    # Comparar os clientes entre os dois DataFrames usando o identificador único (ex.: Conta)
    if not clientes_atual.empty and not clientes_historicos.empty:
        clientes_filtrados = pd.merge(
            clientes_atual,
            clientes_historicos,
            on="Conta",
            suffixes=("_Atual", "_Historico")
        )

        if not clientes_filtrados.empty:
            # Construir o caminho completo para salvar o arquivo no diretório de downloads
            arquivo_saida = os.path.join(diretorio_downloads, f"clientes_filtrados_{codigo_assessor_anterior}.xlsx")
            clientes_filtrados.to_excel(arquivo_saida, index=False)
            print(f"Resultado salvo em '{arquivo_saida}'.")
        else:
            print("Nenhum cliente encontrado que atenda aos critérios especificados.")
    else:
        print("Nenhum cliente encontrado que atenda aos critérios especificados.")
   
    # Caminho para tirar as duplicadas
    # Caminho do arquivo Excel

    arquivo_excel = rf"C:\Users\gabriel.martins\Downloads\historico_clientes_{codigo_assessor_anterior}.xlsx"

    # Carregar o arquivo Excel
    df = pd.read_excel(arquivo_excel)

    # Verificar se a coluna "Conta" existe
    if 'Conta' in df.columns:
        # Remover duplicatas com base na coluna "Conta"
        df_sem_duplicatas = df.drop_duplicates(subset='Conta', keep='first')
        
        # Salvar o DataFrame atualizado no mesmo arquivo
        with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
            df_sem_duplicatas.to_excel(writer, index=False, sheet_name='Dados Limpos')
        
        print("Duplicatas removidas e arquivo atualizado com sucesso.")
    else:
        print("A coluna 'Conta' não existe no arquivo.")

    # Comparar os dados
    # Caminhos dos arquivos
    diretorio_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    arquivo_historico = os.path.join(diretorio_downloads, f"historico_clientes_{codigo_assessor_anterior}.xlsx")
    arquivo_atual = os.path.join(diretorio_downloads, "Base BTG ApenasWertDigital.xlsx")
    arquivo_saida = os.path.join(diretorio_downloads, f"clientes_interseccao{codigo_assessor_anterior}.xlsx")

    # Carregar as planilhas
    try:
        print("Carregando a planilha histórica do assessor...")
        df_historico = pd.read_excel(arquivo_historico)
        print("Carregando a planilha atual de outro assessor...")
        df_atual = pd.read_excel(arquivo_atual)
    except Exception as e:
        raise ValueError(f"Erro ao carregar as planilhas: {e}")

    # Garantir que a coluna 'Conta' esteja presente em ambas as planilhas
    if "Conta" not in df_historico.columns or "Conta" not in df_atual.columns:
        raise ValueError("As planilhas devem conter a coluna 'Conta' para realizar a verificação.")

    # Garantir que a coluna 'Conta' seja tratada como string
    df_historico["Conta"] = df_historico["Conta"].astype(str)
    df_atual["Conta"] = df_atual["Conta"].astype(str)

    # Realizar o merge para encontrar interseção com base na coluna 'Conta'
    clientes_interseccao = pd.merge(df_historico, df_atual, on="Conta", suffixes=("_Historico", "_Atual"))

    # Verificar se há resultados
    if clientes_interseccao.empty:
        print("Nenhuma correspondência encontrada entre as planilhas.")
    else:
        # Salvar os resultados em uma nova planilha
        clientes_interseccao.to_excel(arquivo_saida, index=False)
        print(f"Clientes correspondentes salvos em '{arquivo_saida}'.")
