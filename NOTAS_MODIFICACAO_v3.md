# Notas da modificação v3

Correções aplicadas nesta versão:

1. Tabela e gráfico mensal voltaram à regra visual original:
   - se houver 2 anos fechados consecutivos, exibe os 2 anos fechados + os meses posteriores disponíveis;
   - se não houver 2 anos fechados, exibe no máximo 24 meses ao todo.

2. A aba Gráficos Cobertura usa a mesma regra acima para a tabela editável e para a imagem do gráfico.

3. SKU em comum não força mais janela de 12 meses para volume mensal:
   - usa o mesmo período exibido/calculado na tabela;
   - apenas filtra os SKUs presentes nos dois lados.

4. Base SKUs voltou a usar a base completa do estudo, sem filtrar previamente em 12 meses.

5. A aba Parâmetros passou a exibir a volumetria por categoria em tabela própria:
   - Categoria;
   - Ajuste do divisor;
   - Divisor aplicado;
   - Ratio antes;
   - Ratio depois;
   - Status.

6. A aba Gráficos Cobertura não é mais interpretada como categoria no modo comparação 2.0 x 3.0.

7. Validação feita com teste de 28 meses: jan/2024 a abr/2026, confirmando que a tabela e o gráfico mantêm 2024 completo + 2025 completo + meses de 2026.
