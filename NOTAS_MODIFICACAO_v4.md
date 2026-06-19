# Notas de modificação v4

## Correção solicitada

### Base SKUs
Foram adicionadas colunas específicas para análise por período, sempre mantendo Sell-in e Sell-out lado a lado no mesmo recorte:

- Período Último Ano Fechado
- Volume Sell-in Último Ano Fechado
- Volume Sell-out Último Ano Fechado
- Período MAT 12M
- Volume Sell-in MAT 12M
- Volume Sell-out MAT 12M

## Regras aplicadas

### Último Ano Fechado
- Em bases mensais, usa o último ano civil completo dentro do intervalo comum entre Sell-in e Sell-out.
- Exemplo: se houver dados de jan/2024 a abr/2026, o último ano fechado será FY 2025.
- Em bases sem mês, usa o maior ano comum disponível entre Sell-in e Sell-out.

### MAT 12M
- Usa os últimos 12 meses encerrados no último mês comum entre Sell-in e Sell-out.
- Exemplo: se o último mês comum for abr/2026, o MAT será mai/2025 a abr/2026.
- Quando não há 12 meses mensais disponíveis, o período MAT fica sem aplicação.

## Observação sobre UF
A tabela por UF continua usando a janela de 12 meses móveis no modo mensal, alinhada à cobertura por UF. Essa regra não foi alterada nesta versão.

## Validação
Foram gerados arquivos de teste individual e comparação 2.0 x 3.0. A validação verificou:

- XLSX abre corretamente.
- Base SKUs contém as novas colunas de Último Ano Fechado e MAT 12M.
- Não há erros de fórmula como #REF!, #DIV/0!, #VALUE!, #NAME? ou #N/A.
