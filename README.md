# Estudo de Cobertura - Streamlit

Pacote completo para rodar a interface profissional do Estudo de Cobertura.

## Arquivos principais

- `app.py`: interface Streamlit profissional.
- `gerar_estudo_cobertura_anexo_corrigido.py`: motor/base atual do estudo.
- `requirements.txt`: dependências necessárias.
- `instalar_dependencias.bat`: instala as bibliotecas no Windows.
- `run_streamlit.bat`: abre o app localmente.
- `verificar_instalacao.bat`: testa se o motor está sendo carregado corretamente.

## Como rodar no Windows

1. Extraia o `.zip` em uma pasta simples, por exemplo:
   `C:\Users\Usuario\Documents\Estudo_Cobertura_Streamlit`

2. Dê duplo clique em:
   `instalar_dependencias.bat`

3. Depois dê duplo clique em:
   `run_streamlit.bat`

4. O navegador deve abrir com o painel do Streamlit.

## Como rodar pelo terminal

Abra o CMD ou PowerShell dentro da pasta extraída e execute:

```bat
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Correção do erro de motor não carregado

Esta versão força o app a carregar o motor pela própria pasta onde o `app.py` está salvo.

Mesmo que o Streamlit seja iniciado de outro diretório, o app procura automaticamente:

```text
gerar_estudo_cobertura_anexo_corrigido.py
```

na mesma pasta do `app.py`.

## Sobre volume, quantidade e escala

O motor considera as opções:

- `Volume`
- `Quantidade`
- `Volume variável`

E possui ajuste automático de escala entre Sell-in e Sell-out quando identifica diferenças compatíveis com dezenas, centenas, milhares, milhões, bilhões ou trilhões.

Exemplo esperado:

```text
Sell-in: 340.123
Sell-out: 320.150.850.321
```

Nesse cenário, o Sell-out pode ser ajustado por divisor compatível com milhão, resultando numa escala comparável ao Sell-in, sem alterar a base original carregada.

## Observação importante

A prévia de linhas na tela é apenas uma leitura inicial para conferência visual. A validação real continua sendo feita pelo motor do estudo durante o processamento.
