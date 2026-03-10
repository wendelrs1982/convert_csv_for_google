# Obter dados de um arquivo xlxs e converter para CSV
A função desse código é realizar a criação de um arquivo csv para ser importado no google agenda.

## Requisitos
- pandas
- os
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