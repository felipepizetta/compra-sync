import pandas as pd

# RC
requisicoes_df = pd.read_excel('Requisitions/Rec.xlsx')
print("RC: ")
print(requisicoes_df.columns)

# OC
ordens_df = pd.read_excel('Purchasing/Ord.xlsx')
print("OC: ")
print(ordens_df.columns)

requisicoes_df['ID do projeto'] = requisicoes_df['ID do projeto'].astype(object)
ordens_df['ID do projeto'] = ordens_df['ID do projeto'].astype(object)

merged_df = pd.merge(
    requisicoes_df, 
    ordens_df, 
    left_on=['Site', 'Preparador', 'Data de emissão', 'Nome2', 'ID do projeto'], 
    right_on=['Site', 'Solicitante', 'Data de recebimento solicitada', 'Autor da ordem', 'ID do projeto'],  # note o espaço
    how='inner'
)

final_df = merged_df[['Site', 'Conta de fornecedor', 'Ordem de compra', 'Requisição de compra', 'Nome', 'Preparador', 'Status da ordem de compra', 'ID do projeto', 'Comprador_x', 'Classificação do contrato de compra', 'Tipo de operação']]

final_df.to_excel('Tabela.xlsx', index=False)