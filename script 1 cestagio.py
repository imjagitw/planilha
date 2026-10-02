from pathlib import Path
import pandas as pd
import numpy as np

print("-> Iniciando a leitura do arquivo...")
planilhas_path = Path('./planilhas')
planilha_original = planilhas_path / 'Execução da despesa-1.xlsx'

if not planilha_original.exists():
    print("Arquivo não encontrado:", planilha_original.resolve())
    print("Arquivos na pasta:", [p.name for p in planilhas_path.glob('*')])
    raise SystemExit

# 1. Ler SEM cabeçalho e pular as 3 linhas de título (código do item, nome do item e rótulo).
#    Assim as colunas são identificadas por posição, que é estável nessa exportação.
df = pd.read_excel(planilha_original, sheet_name=0, header=None, skiprows=3)
print(f"-> Dados carregados! {df.shape[0]} linhas e {df.shape[1]} colunas.")

# 2. Remover a linha "Total" do final da exportação
df = df[df[0].astype(str).str.strip().str.lower() != 'total'].copy()

# 3. Manter só as 20 colunas com significado e renomear.
#    As colunas 20 a 23 da exportação não têm cabeçalho e foram descartadas.
df = df.iloc[:, :20]
df.columns = [
    'UG_Responsavel_Cod', 'UG_Responsavel_Nome',
    'Favorecido_CNPJ', 'Favorecido_Nome',
    'PI_Cod', 'PI_Nome',
    'Natureza_Despesa_Cod', 'Natureza_Despesa_Nome',
    'NE_CCor_Empenho', 'NE_CCor_Info_Complementar', 'Ano_Emissao_Empenho',
    'Ano_Lancamento',
    'Despesas_Empenhadas',            # item 29
    'Despesas_Liquidadas',            # item 31
    'Despesas_Pagas',                 # item 34
    'RAP_NP_Inscritos',               # item 40
    'RAP_NP_Reinscritos',             # item 41
    'RAP_NP_Liquidados',              # item 44
    'RAP_NP_Pagos',                   # item 46
    'Total_Geral'
]

# 4. Valores nulos viram 0 nas colunas numéricas
colunas_valores = [
    'Despesas_Empenhadas', 'Despesas_Liquidadas', 'Despesas_Pagas',
    'RAP_NP_Inscritos', 'RAP_NP_Reinscritos', 'RAP_NP_Liquidados',
    'RAP_NP_Pagos', 'Total_Geral'
]
for c in colunas_valores:
    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

# 5. Regras de negócio

# 5.1 Com contrato / Sem contrato
df['Status_Contrato'] = np.where(
    df['PI_Nome'].astype(str).str.contains('CONTRATO', case=False, na=False) |
    df['PI_Cod'].astype(str).str.contains('CTN', case=False, na=False),
    'Com contrato',
    'Sem contrato'
)

# 5.2 Natureza da despesa
# O código vem como número (339037.0); converter para inteiro antes de virar texto
df['Natureza_Despesa_Cod'] = (
    pd.to_numeric(df['Natureza_Despesa_Cod'], errors='coerce')
    .astype('Int64').astype(str).str.strip()
)

df['Categoria_Natureza'] = np.where(
    df['Natureza_Despesa_Cod'].str.startswith('3'),
    'Despesas Correntes',
    'Despesas de Capital'
)

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
df['Natureza_Despesa_Nome'] = (
    df['Natureza_Despesa_Cod'].str[:6].map(mapa_naturezas_manual)
    .fillna(df['Natureza_Despesa_Nome'])
)

# 5.3 MÉTRICAS DA EXECUÇÃO ANUAL
# RAP não processado "a pagar" = inscritos + reinscritos
df['RAP_NP_Total'] = df['RAP_NP_Inscritos'] + df['RAP_NP_Reinscritos']

# (1) Montante = Empenho + Restos a pagar não processados
df['Montante'] = df['Despesas_Empenhadas'] + df['RAP_NP_Total']

# (2) Execução (liquidação) = Empenhos liquidados + RAP NP liquidados
df['Execucao_Liquidacao'] = df['Despesas_Liquidadas'] + df['RAP_NP_Liquidados']

# Saldo = (1) - (2)  -> "a liquidar"
df['Saldo_a_Liquidar'] = df['Montante'] - df['Execucao_Liquidacao']

# RAP Pago (Tudo de RAP Pago)
df['Total_RAP_Pagos'] = df['RAP_NP_Pagos']
df['Total_Pago_RAP'] = df['Total_RAP_Pagos']

# (3) Total Geral Pago = Soma de Despesas Pagas + Total RAP Pago
df['Total_Pago_Geral'] = df['Despesas_Pagas'] + df['Total_RAP_Pagos']
df['Pagamento'] = df['Total_Pago_Geral']

# 5.4 Anos_RAP: apenas os anos de emissão do empenho (sem valores nulos)
df['Ano_Emissao_Empenho'] = pd.to_numeric(df['Ano_Emissao_Empenho'], errors='coerce').astype('Int64')
df['Ano_Lancamento'] = pd.to_numeric(df['Ano_Lancamento'], errors='coerce').astype('Int64')
ano_vigente = df['Ano_Lancamento'].max()
df['Ano_Vigente'] = ano_vigente
df['Anos_RAP'] = np.where(
    df['Ano_Emissao_Empenho'] != ano_vigente,
    df['Ano_Emissao_Empenho'],
    np.nan
)
df['Anos_RAP'] = df['Anos_RAP'].astype('Int64')

# 5.5 Colunas de Filtro com Código e Nome (para o Looker Studio)
def limpa_codigo(val):
    if pd.isna(val):
        return ''
    s = str(val).strip()
    return s[:-2] if s.endswith('.0') else s

df['UG_Responsavel_Filtro'] = (
    df['UG_Responsavel_Cod'].map(limpa_codigo).replace('', 'Sem código') +
    ' - ' + df['UG_Responsavel_Nome'].fillna('').astype(str).str.strip()
)
df['Favorecido_Filtro'] = (
    df['Favorecido_CNPJ'].map(limpa_codigo).replace('', 'Sem código') +
    ' - ' + df['Favorecido_Nome'].fillna('').astype(str).str.strip()
)
df['Natureza_Filtro'] = (
    df['Natureza_Despesa_Cod'].map(limpa_codigo).replace('', 'Sem código') +
    ' - ' + df['Natureza_Despesa_Nome'].fillna('').astype(str).str.strip()
)

# 6. Conferência rápida dos totais
print("\n--- CONFERÊNCIA (R$) ---")
print(f"Empenhado ............. {df['Despesas_Empenhadas'].sum():>18,.2f}")
print(f"RAP NP (insc.+reinsc.)  {df['RAP_NP_Total'].sum():>18,.2f}")
print(f"(1) Montante .......... {df['Montante'].sum():>18,.2f}")
print(f"(2) Execução/Liquidação {df['Execucao_Liquidacao'].sum():>18,.2f}")
print(f"    Saldo a liquidar .. {df['Saldo_a_Liquidar'].sum():>18,.2f}")
print(f"(3) Pagamento ......... {df['Pagamento'].sum():>18,.2f}")
print(f"Linhas com saldo negativo: {(df['Saldo_a_Liquidar'] < -0.01).sum()}")
print(f"Linhas com pagamento > liquidação: {(df['Pagamento'] > df['Execucao_Liquidacao'] + 0.01).sum()}")

# 7. Exportar
print(f"\n-> Exportando {df.shape[0]} linhas...")
df.to_excel(planilhas_path / 'Base_Tratada_Painel_PRAD.xlsx', index=False)
df.to_csv(planilhas_path / 'Base_Tratada_Painel_PRAD.csv', index=False, sep=';', encoding='utf-8-sig')
print("-> Arquivos gerados com sucesso! Verifique a pasta.")