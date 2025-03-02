import pandas as pd
from openpyxl import load_workbook

# Caminhos dos arquivos de entrada e saída
file_path = "data/db_sce.xlsx"
output_file = "data/db_sce_formatado_ID_LOCACAO.xlsx"  # Novo nome

# Ler a aba "ITEM 3"
df = pd.read_excel(file_path, sheet_name="ITEM 3")

# Concatenar todas as disciplinas por IDFUNCIONAL (limitando o tamanho)
df["DISCIPLINA DE ALOCAÇÃO"] = df.groupby("IDFUNCIONAL")["DISCIPLINA DE ALOCAÇÃO"].transform(lambda x: ', '.join(x)[:100])  # Limita a 100 caracteres

# Remover duplicatas mantendo apenas uma linha por IDFUNCIONAL
df_grouped = df.drop_duplicates(subset="IDFUNCIONAL")

# Salvar o DataFrame em um arquivo Excel
df_grouped.to_excel(output_file, index=False, engine="openpyxl")

# Carregar o arquivo Excel para ajustar as larguras das colunas
wb = load_workbook(output_file)
ws = wb.active

# Ajustar a largura das colunas
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter  # Obtém a letra da coluna

    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))

    # Define uma largura máxima para evitar colunas enormes
    ws.column_dimensions[col_letter].width = min(max_length, 30)  # Máximo de 30 caracteres de largura

# Salvar o arquivo com as larguras ajustadas
wb.save(output_file)

print(f"Arquivo '{output_file}' gerado com sucesso!")
