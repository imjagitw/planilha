import pandas as pd
import numpy as np

print("-> Iniciando a leitura do arquivo...")

# 1. Carregar a folha de cálculo pulando o cabeçalho desconfigurado (as 2 primeiras linhas)
df = pd.read_excel('Execução da despesa-1.xlsx', sheet_name=0, skiprows=2)

print(f"-> Dados carregados! A planilha possui {df.shape[0]} linhas e {df.shape[1]} colunas.")

# 2. Remover a última linha se for o "Total Geral" da exportação
if str(df.iloc[-1, 0]).strip().lower() == 'total':
    df = df.iloc[:-1]

# 3. Renomear as colunas para um formato amigável para o Looker Studio
df.columns = [
    'UG_Responsavel_Cod', 'UG_Responsavel_Nome',
    'Favorecido_CNPJ', 'Favorecido_Nome',
    'PI_Cod', 'PI_Nome',
    'Natureza_Despesa_Cod', 'Natureza_Despesa_Nome',
    'NE_CCor_Empenho', 'NE_CCor_Info_Complementar', 'Ano_Emissao_Empenho',
    'Ano_Lancamento',
    'Despesas_Empenhadas',            # Coluna 29 do sistema original
    'Despesas_Liquidadas',            # Coluna 31 
    'Despesas_Pagas',                 # Coluna 34 
    'RAP_Processados_Pagos',          # Coluna 38 
    'RAP_Nao_Processados_Liquidados', # Coluna 44 
    'RAP_Nao_Processados_Pagos',      # Coluna 46 
    'Total_Geral'
]

# 4. Tratar valores nulos (transformar NaN em 0 nas colunas numéricas)
colunas_valores = [
    'Despesas_Empenhadas', 'Despesas_Liquidadas', 'Despesas_Pagas', 
    'RAP_Processados_Pagos', 'RAP_Nao_Processados_Liquidados', 
    'RAP_Nao_Processados_Pagos', 'Total_Geral'
]
df[colunas_valores] = df[colunas_valores].fillna(0)

# 5. Criar as regras de negócio e categorizações do painel

# 5.1 Tag: Com contrato / Sem contrato baseada na coluna PI
df['Status_Contrato'] = np.where(
    df['PI_Nome'].astype(str).str.contains('CONTRATO', case=False, na=False) | 
    df['PI_Cod'].astype(str).str.contains('CTN', case=False, na=False), 
    'Com contrato', 
    'Sem contrato'
)

# -----------------------------------------------------------------------------
# 5.2 Agrupamento da Natureza da Despesa e Adequação dos Nomes (Atualizado)
# -----------------------------------------------------------------------------

# Garantir que o código da natureza de despesa é interpretado como string e sem espaços
df['Natureza_Despesa_Cod'] = df['Natureza_Despesa_Cod'].astype(str).str.strip()

# Regra 1: Categorização Principal
df['Categoria_Natureza'] = np.where(
    df['Natureza_Despesa_Cod'].str.startswith('3'),
    'Material de consumo e serviço',
    'Investimentos'
)

# Regra 2: Padronização dos Nomes conforme o Manual (Quadro 1 e 2)
mapa_naturezas_manual = {
    '339014': 'DIÁRIAS PESSOAL CIVIL',
    '339018': 'AUXÍLIO FINANCEIRO A ESTUDANTES',
    '339020': 'AUXÍLIO FINANCEIRO A PESQUISADORES',
    '339030': 'MATERIAL DE CONSUMO',
    '339031': 'PREMIAÇÕES CULTURAIS, ARTÍSTICAS, CIENTÍFICAS, DESPORTIVAS E OUTROS',
    '339032': 'MATERIAL, BEM OU SERVIÇO PARA DISTRIBUIÇÃO GRATUITA',
    '339033': 'PASSAGENS E DESPESAS COM LOCOMOÇÃO',
    '339035': 'SERVICOS DE CONSULTORIA',
    '339036': 'OUTROS SERVIÇOS DE TERCEIROS PESSOA FÍSICA',
    '339037': 'LOCAÇÃO DE MÃO DE OBRA',
    '339039': 'OUTROS SERVIÇOS DE TERCEIROS - PESSOA JURÍDICA',
    '339040': 'SERVIÇOS DE TECNOLOGIA DA INFORMAÇÃO E COMUNICAÇÃO',
    '339047': 'OBRIGAÇÕES TRIBUTÁRIAS E CONTRIBUTIVAS',
    '339048': 'OUTROS AUXÍLIOS FINANCEIROS A PESSOAS FÍSICAS',
    '339092': 'DESPESAS DE EXERCÍCIOS ANTERIORES',
    '339093': 'INDENIZAÇÕES E RESTITUIÇÕES',
    '449051': 'OBRAS E INSTALAÇÕES',
    '449052': 'EQUIPAMENTOS E MATERIAL PERMANENTE'
}

df['Codigo_6_digitos'] = df['Natureza_Despesa_Cod'].str.replace('.', '', regex=False).str[:6]
df['Natureza_Despesa_Nome'] = df['Codigo_6_digitos'].map(mapa_naturezas_manual).fillna(df['Natureza_Despesa_Nome'])
df = df.drop(columns=['Codigo_6_digitos'])

# -----------------------------------------------------------------------------

# 5.3 Campos Calculados Totais
df['Total_RAP_Pagos'] = df['RAP_Processados_Pagos'] + df['RAP_Nao_Processados_Pagos']
df['Total_Pago_Geral'] = df['Despesas_Pagas'] + df['Total_RAP_Pagos']

# 6. Exportar a base final tratada
print(f"-> Tratamento concluído. Exportando {df.shape[0]} linhas...")

# Salva em Excel
df.to_excel('Base_Tratada_Painel_PRAD.xlsx', index=False)

# Salva em CSV (Plano B para subir no painel)
df.to_csv('Base_Tratada_Painel_PRAD.csv', index=False, sep=';', encoding='utf-8-sig')

print("-> Arquivos gerados com sucesso! Verifique a pasta.")