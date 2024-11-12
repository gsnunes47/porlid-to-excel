import pandas as pd
import re

file_path = 'samples/PORLID_8casosdevento.txt'
with open(file_path, 'r') as file:
    lines = file.readlines()

# Processamento das linhas
headers = []
data = []


header_line_1 = lines[0].strip().split('\t')

for i, v in enumerate(header_line_1):
    if v == '-':
        header_line_1[i] = header_line_1[i-1]

# Combinar headers para criar nomes únicos de coluna
for i in range(len(header_line_1)):
    case = header_line_1[i].strip()
    headers.append(f"{case}")


# Processar cada linha de dados dos elementos
for line in lines[1:]:
    # Remover espaços e dividir a linha
    line = re.sub(r'\s+', ' ', line.strip())
    parts = line.split()
    
    # Primeiro elemento é o nome do elemento (e.g., P2A, P3A)
    element_name = parts[0]
    values = parts[1:]
    
    # Converter para float e organizar em uma linha de dados
    row = [element_name] + [value.replace(',', '.') for value in values]
    data.append(row)


# # Criar DataFrame e exportar para Excel
df = pd.DataFrame(data, columns=headers)
df.iloc[0] = [None] + df.iloc[0].tolist()[:-1]
output_path = 'dados_convertidos.xlsx'
df.to_excel(output_path, index=False)

print(f"Arquivo Excel criado: {output_path}")
