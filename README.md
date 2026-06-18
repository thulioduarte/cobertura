# Estudo de Cobertura - Streamlit

## Como rodar localmente

```bat
cd "C:\caminho\da\pasta"
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Também é possível executar `instalar_dependencias.bat` e depois `run_streamlit.bat`.

## Arquivos principais

- `app.py`: interface Streamlit profissional.
- `gerar_estudo_cobertura_anexo_corrigido.py`: motor do Estudo de Cobertura.
- `Manual de Uso - Estudo de Cobertura.docx`: manual baixável no painel lateral.
- `Boas Praticas - Estudo de Cobertura.md`: boas práticas baixáveis no painel lateral.

## Atualizações desta versão

- Removido o card visual “Motor carregado”.
- Removido o card “Motor ativo”.
- `Base Contribuição Sell-out` desativada por padrão.
- `Congelado` exibido de forma discreta: só aparece após marcar `Usar Congelado opcional`.
- Manual de Uso e Boas Práticas adicionados no painel lateral, com botão de download.
- Manual atualizado com imagens novas da interface Streamlit.


## Dependências extras desta versão

- `matplotlib`: usado para gerar a aba `Gráficos Cobertura` como imagem estática no mesmo padrão visual do PPT.
