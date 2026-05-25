# Boas Práticas - Estudo de Cobertura

## Antes de gerar

- Use Sell-in mensal. Não use MAT, YTD, média ou total como se fossem valores mensais.
- Escolha a mesma métrica entre Sell-in e Sell-out: volume com volume, quantidade com quantidade.
- Em arquivos grandes do FTP, filtre o Fabricante antes ou dentro do app para melhorar a performance.
- Use Congelado somente quando precisar apoiar mapeamento por SKU/EAN, Marca, Fabricante ou Est Mer 7.
- Se precisar excluir marcas, faça isso antes de carregar os arquivos no app.

## Durante a geração

- Confira a prévia das primeiras linhas para validar se o arquivo correto foi carregado.
- Mantenha as saídas principais ligadas, principalmente Parâmetros, Avisos e Descrição Cálculos.
- Ative Base Contribuição Sell-out apenas quando realmente precisar dessa aba técnica.

## Depois de gerar

- Leia a aba Avisos antes de interpretar o resultado.
- Confira se os meses do Sell-in e Sell-out estão alinhados.
- Em comparação 2.0 x 3.0, confirme se cada arquivo pertence à versão correta.
- No modo Dash, valide se a versão do SKU/Congelado, quando usado, corresponde à versão avaliada do SM.
