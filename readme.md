# Obtem as informações de uma planilha e gera uma arquivo CSV
Este script gera um arquivo CSV no formato exigido pelo Google Agenda, permitindo a importação automática de eventos a partir de uma planilha.

## Requisitos
- pandas
- os
- re
- datetime

## Instalação

Instale as dependências do projeto:
```bash
pip install -r requirements.txt
```
Execute o comando abaixo para criar o arquivo de configuração:
```bash
cp config-example.py config.py
```

Acesse o arquivo `config.py` e informe os parâmetros necessários.

## Uso

1. Execute o arquivo `createCsv.py`:
```bash
python createCsv.py
```