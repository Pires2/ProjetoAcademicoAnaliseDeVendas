import pandas as pd

# Caminho para o arquivo original e o arquivo de exportação
file_path = 'Dados das Vendas.xlsx'  # Caminho do arquivo original
output_path = 'Resumo das vendas.xlsx'  # Caminho do arquivo de exportação

# Carregar a planilha original com os dados de vendas
df = pd.read_excel(file_path)

# 1. Investimento, Lucro e Total de Vendas por Marca
investimento_lucro_vendas = df[['Marcas', 'Investimento', 'Lucro', 'Total de Vendas']]

# 2. Ranking das Peças de Roupa Mais Vendidas
quantidade_pecas = df[['Camisas', 'Calças', 'Sapatos', 'Casacos', 'Acessórios', 'Bijuterias', 'Roupas Íntimas']].sum()
ranking_pecas = quantidade_pecas.sort_values()

# 3. Vendas Totais por Marca com total de cada peça
pecas_por_marca = df.groupby('Marcas')[['Camisas', 'Calças', 'Sapatos', 'Casacos', 'Acessórios', 'Bijuterias', 'Roupas Íntimas']].sum()
vendas_por_marca = df.groupby('Marcas')['Total de Vendas'].sum()
vendas_detalhadas_por_marca = pd.concat([vendas_por_marca, pecas_por_marca], axis=1)

# Salvar os dados processados em um único arquivo Excel com abas separadas
with pd.ExcelWriter(output_path) as writer:
    investimento_lucro_vendas.to_excel(writer, sheet_name='Investimento_Lucro_Vendas', index=False)
    ranking_pecas.to_frame(name='Quantidade Vendida').to_excel(writer, sheet_name='Ranking_Pecas', index=True)
    vendas_detalhadas_por_marca.to_excel(writer, sheet_name='Vendas_Por_Marca', index=True)

print(f"Dados atualizados e exportados para: {output_path}")
