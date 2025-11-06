import pandas as pd

ARQUIVO_EXCEL = r"Seu_Arquivo.xlsx"  # Caminho da planilha

# Lê a planilha inteira
df = pd.read_excel(ARQUIVO_EXCEL)

# Garante que estamos pegando a coluna C (índice 2, pois começa em 0) e converte para string
coluna_c = df.iloc[:, 2].astype(str)

# --- ALTERAÇÃO AQUI ---
# Função lambda modificada para simular quebras de linha do Windows (CRLF: \r\n)
# Substitui o '\n' (Linux/Excel) por '\r\n' (Windows) ANTES de contar o comprimento.
def contar_caracteres_windows(texto):
    # 1. Simula CRLF: Substitui cada quebra de linha ('\n') por '\r\n'.
    # Isso garante que cada nova linha conte como 2 caracteres (o cenário mais restritivo)
    texto_crlf = str(texto).replace('\n', '\r\n')
    
    # 2. Retorna a contagem exata (garantindo o suporte Unicode)
    return len(texto_crlf.encode('utf-8').decode('utf-8'))
# Aplica a nova função de contagem à coluna
contagem_caracteres = coluna_c.apply(contar_caracteres_windows)
# -----------------------

# Filtra linhas que ultrapassam 5000 caracteres
linhas_excedentes = contagem_caracteres[contagem_caracteres > 5000]

if not linhas_excedentes.empty:
    print("Linhas com mais de 5000 caracteres na coluna C:")
    # Nota: idx + 2 é para compensar o índice 0 e o cabeçalho do Excel.
    for idx, qtd in linhas_excedentes.items():
        print(f"➡ Linha {idx + 2} (Excel): {qtd} caracteres")
else:
    print("Nenhuma célula da coluna C ultrapassa 5000 caracteres")