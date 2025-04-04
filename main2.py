import pandas as pd

# Caminho do arquivo 
# Excel
arquivo_excel = r"C:\Users\gabriel.martins\Downloads\historico_clientes_1393034.0.xlsx"

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
