import pandas as pd

# Caminho do arquivo Excel
caminho_arquivo = 'C:\\Users\\daniel\\Desktop\\data_fgv\\db_sce.xlsx'

# Lista de abas (planilhas) para processar
abas = ['ITEM 1', 'ITEM 2', 'ITEM 3', 'ITEM 4']

# Dicionário para armazenar os DataFrames filtrados
df_filtrados = {}

# Ler cada planilha e filtrar as linhas com ID_FUNCIONAL = 43554075
for aba in abas:
    try:
        # Ler a planilha
        df = pd.read_excel(caminho_arquivo, sheet_name=aba, skiprows=2)
        
        # Verificar se a coluna ID_FUNCIONAL existe
        if 'ID_FUNCIONAL' not in df.columns:
            print(f"A coluna 'ID_FUNCIONAL' não foi encontrada na planilha {aba}. Verifique o nome da coluna.")
            continue  # Pula para a próxima planilha
        
        # Filtrar as linhas com ID_FUNCIONAL = 43554075
        df_filtrado = df[df['ID_FUNCIONAL'] == 43554075]
        
        # Armazenar o DataFrame filtrado no dicionário
        df_filtrados[aba] = df_filtrado
    except Exception as e:
        print(f"Erro ao processar a planilha {aba}: {e}")

# Exportar os DataFrames filtrados para um único arquivo Excel, com cada planilha em uma aba diferente
if df_filtrados:  # Verifica se há dados filtrados
    caminho_saida = 'C:\\Users\\daniel\\Desktop\\data_fgv\\resultados_filtrados02.xlsx'
    with pd.ExcelWriter(caminho_saida) as writer:
        for aba, df in df_filtrados.items():
            df.to_excel(writer, sheet_name=aba, index=False)
    print(f"Resultados filtrados exportados com sucesso para: {caminho_saida}")
else:
    print("Nenhum dado foi filtrado. Verifique os nomes das colunas e os dados das planilhas.")


