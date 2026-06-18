# Estudo de Cobertura - Código modificado

## Alterações principais

- Adicionada a opção `Gráficos Cobertura` na interface e no motor.
- Criada aba `Gráficos Cobertura` com imagens estáticas geradas por Matplotlib e tabelas editáveis.
- Em comparação 2.0 x 3.0, o gráfico usa Sell-in, Sell-out 2.0, Sell-out 3.0, Cobertura 2.0 e Cobertura 3.0.
- Coberturas aparecem apenas no último mês disponível, em eixo secundário percentual.
- Ajuste de volumetria passou a ser por categoria/PROD, não mais global.
- Aba Parâmetros inclui o divisor aplicado na volumetria por categoria.
- Cálculos de SKU em comum e Base SKUs usam janela de 12 meses por categoria.
- MAT e tendências principais só são calculados com 24 meses completos.
- Tendência 6M só é calculada quando há pelo menos 12 meses completos.
- Removidas colunas técnicas excedentes do Resumo Categorias.
- Nomenclatura alterada de Variação para Tendência.

## Validação executada

Foram gerados arquivos de teste com dados sintéticos para validar abertura do XLSX, existência dos gráficos, ausência de erros textuais/fórmulas comuns e ausência de coluna `Cobertura SO Total`.

## Observação

O XLSX de teste não representa os dados reais do estudo. Ele serve apenas para validar a estrutura do arquivo e o funcionamento do código.
