import pandas as pd

# Carregar o arquivo Excel
df = pd.read_excel('Op Padrão - Planos para carga F2 (version 1) - Copia.xlsm', sheet_name='BASE')

# Extrair os dados necessários
parametro = df.loc[69:139, 'J']  # J70:J140, considerando que o índice é 0-based
coluna_f = df.loc[69:139, 'F']
coluna_d = df.loc[69:139, 'D']
operacoes = df.loc[69:139, 'U']

# Criar um DataFrame com os dados extraídos
resultado = pd.DataFrame({
    'Titulo': parametro,
    'Coluna F': coluna_f,
    'Coluna D': coluna_d,
    'Operações (Coluna U)': operacoes
})

# Salvar o resultado em uma nova planilha Excel
with pd.ExcelWriter('resultado.xlsx', engine='openpyxl') as writer:
    resultado.to_excel(writer, sheet_name='Resultado', index=False)

