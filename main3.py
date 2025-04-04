import pandas as pd
import os


# Caminhos dos 
#arquivos
diretorio_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
arquivo_historico = os.path.join(diretorio_downloads, "historico_clientes_1393034.0.xlsx")
arquivo_atual = os.path.join(diretorio_downloads, "Base BTG ApenasWertDigital.xlsx")
arquivo_saida = os.path.join(diretorio_downloads, "clientes_interseccao.xlsx")

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
