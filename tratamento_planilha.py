from pathlib import Path
import pandas as pd
import numpy as np

print("-> Iniciando a leitura do arquivo...")
planilhas_path = Path('./planilhas')

# 1. Carregar a folha de cálculo pulando o cabeçalho desconfigurado (as 2 primeiras linhas)
planilha_original = planilhas_path / 'Execução da despesa-1.xlsx'
df = pd.read_excel(planilha_original, sheet_name=0, skiprows=2)

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
    'Despesas Correntes',
    'Despesas de Capital'
)

# Regra 2: Padronização dos Nomes conforme o Manual (Quadro 1 e 2)
mapa_naturezas_manual = {
    '339014': 'Diárias - Pessoal Civil',
    '339018': 'Auxílio Financeiro a Estudantes',
    '339020': 'Auxílio Financeiro a Pesquisadores',
    '339030': 'Material de Consumo',
    '339031': 'Premiações Culturais, Artísticas, Científicas, Desportivas e Outros',
    '339032': 'Material, Bem ou Serviço para Distribuição Gratuita',
    '339033': 'Passagens e Despesas com Locomoção',
    '339035': 'Serviços de Consultoria',
    '339036': 'Outros Serviços de Terceiros - Pessoa Física',
    '339037': 'Locação de Mão de Obra',
    '339039': 'Outros Serviços de Terceiros - Pessoa Jurídica',
    '339040': 'Serviços de Tecnologia da Informação e Comunicação - pessoa jurídica',
    '339047': 'Obrigações Tributárias e Contributivas',
    '339048': 'Outros Auxílios Financeiros a Pessoas Físicas',
    '339092': 'Despesas de Exercícios Anteriores',
    '339093': 'Indenizações e Restituições',
    '339147': 'Obrigações Tributárias e Contributivas em Operações Intraorçamentárias',
    '449051': 'Obras e Instalações',
    '449052': 'Equipamentos e Material Permanente',
    '449039': 'Serviços de terceiros - Pessoa Jurídica (Capital)'
}

df['Codigo_6_digitos'] = df['Natureza_Despesa_Cod'].str.replace('.', '', regex=False).str[:6]
df['Natureza_Despesa_Nome'] = df['Codigo_6_digitos'].map(mapa_naturezas_manual).fillna(df['Natureza_Despesa_Nome'])
df = df.drop(columns=['Codigo_6_digitos'])

# -----------------------------------------------------------------------------

# 5.3 Campos Calculados Totais
df['Total_RAP_Pagos'] = df['RAP_Processados_Pagos'] + df['RAP_Nao_Processados_Pagos']
df['Total_Pago_Geral'] = df['Despesas_Pagas'] + df['Total_RAP_Pagos']

# 5.4 Variável Anos_RAP (Contém todos os anos de emissão, exceto o ano vigente)
df['Ano_Emissao_Empenho'] = df['Ano_Emissao_Empenho'].astype('Int64')
df['Ano_Lancamento'] = df['Ano_Lancamento'].astype('Int64')

ano_vigente = df['Ano_Lancamento'].max()
df['Anos_RAP'] = np.where(
    df['Ano_Emissao_Empenho'] != ano_vigente,
    df['Ano_Emissao_Empenho'],
    np.nan
)
df['Anos_RAP'] = df['Anos_RAP'].astype('Int64')

# 6. Exportar a base final tratada
print(f"-> Tratamento concluído. Exportando {df.shape[0]} linhas...")

# Salva em Excel
df.to_excel(planilhas_path / 'Base_Tratada_Painel_PRAD.xlsx', index=False)

# Salva em CSV (Plano B para subir no painel)
df.to_csv(planilhas_path / 'Base_Tratada_Painel_PRAD.csv', index=False, sep=';', encoding='utf-8-sig')

print("-> Arquivos gerados com sucesso! Verifique a pasta.")