# Notas da modificação v2

## Ajuste feito no SKU em comum

- O bloco **SKU em Comum** deixou de forçar cálculo somente em janela fixa de 12 meses.
- Agora ele replica o mesmo período definido na tabela superior da aba:
  - mesmo período anterior;
  - mesmo período atual;
  - mesma lógica FY/YTD/total disponível.
- A única diferença é que as somas são filtradas para SKUs/EANs que existem simultaneamente no Sell-in e no Sell-out dentro de cada período.
- Foi mantido um detalhe mensal de apoio apenas para visualização do período atual, sem recalcular 12M.

## Ajuste feito no divisor por categoria

- A aba **Parâmetros** não acumula mais todos os divisores em uma única célula.
- Agora existe uma tabela separada chamada **Divisor aplicado na volumetria por categoria**.
- Essa tabela tem uma linha por categoria/PROD, com:
  - Categoria;
  - Coluna ajustada;
  - Divisor aplicado;
  - Ratio antes;
  - Ratio depois;
  - Status.

## Correção adicional

- A aba **Gráficos Cobertura** foi adicionada à lista de abas de sistema da comparação 2.0 x 3.0, para não ser interpretada como categoria/PROD.

## Validação

- O motor foi validado com geração individual e comparação 2.0 x 3.0 usando os arquivos de teste.
- Os arquivos abriram sem corrupção.
- Não foram encontrados erros de fórmula como #REF!, #DIV/0!, #VALUE!, #NAME? ou #N/A nos testes gerados.
