import pandas as pd
import string

#section-function
def letra_para_numero(letra):
    """Converte uma letra de coluna do Excel para um índice numérico (0-based)."""
    alfabeto = string.ascii_uppercase
    numero = 0
    for char in letra:
        numero = numero * 26 + (alfabeto.index(char) + 1)
    return numero - 1  # Ajuste para índice baseado em zero

def listarEmultiplicar(file_path, output_path, colunas):

    #section-tratamento
    colunas_numericas = {
        chave: [letra_para_numero(col) for col in valor if col]  # Ignora strings vazias
        for chave, valor in colunas.items()
    }

    #section-main
    df = pd.read_excel(file_path)
    segunda_linha = df.iloc[0].copy()

    for nome_coluna, indice_coluna in colunas_numericas.items():
        
        if nome_coluna == "Elem" or nome_coluna == "Todas permanentes e acidentais dos pavimentos":
            continue
        
        print('  Coluna: ', nome_coluna)

        #input majorador    
        while True:

            try:
                majorador = str(input('  Majorador: ')).strip()
                majorador = float(majorador.replace(',', '.'))
                break

            except ValueError:
                print()
                print("  Erro: Digite um número válido.")
                print()
        
        print()
        indice_coluna[1] += 1

        df_recorte = df.iloc[:, indice_coluna[0]:indice_coluna[1]]

        #section-tratamento-corte
        df_recorte = df_recorte.replace(',', '.', regex=True)  # Troca vírgulas por pontos
        df_recorte = df_recorte.apply(lambda x: x.str.strip() if x.dtype == "object" else x)  # Remove espaços extras

        df_recorte = df_recorte.apply(pd.to_numeric, errors='coerce')

        #aplicando majorador
        df_recorte = df_recorte.multiply(majorador)

        #arredondando para duas casas decimais
        df_recorte = df_recorte.round(2)    

        df.iloc[:, indice_coluna[0]:indice_coluna[1]] = df_recorte

    segunda_linha = segunda_linha.apply(lambda x: x if pd.notna(x) else "")
    df.iloc[0] = segunda_linha
    df.to_excel(output_path, index=False)

# listarEmultiplicar(
#     file_path = r"C:\Users\gusta\OneDrive\Documentos\GitHub\porlid-to-excel\samples\dados-mesclados.xlsx", 
#     output_path = r"C:\Users\gusta\OneDrive\Documentos\GitHub\porlid-to-excel\samples\dados-banana.xlsx", 
#     colunas = {'Elem': ['A', ''], 'Todas permanentes e acidentais dos pavimentos': ['B', 'D'], 'Empuxo': ['E', 'I'], 'Vento (1) 78°': ['J', 'N'], 'Vento (2) 258°': ['O', 'S'], 'Vento (3) 348°': ['T', 'X'], 'Vento (4) 168°': ['Y', 'AC'], 'Vento (5) 33°': ['AD', 'AH'], 'Vento (6) 123°': ['AI', 'AM'], 'Vento (7) 213°': ['AN', 'AR'], 'Vento (8) 303°': ['AS', 'AW']}
# )
