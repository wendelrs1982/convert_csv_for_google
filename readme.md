# Gerador de CSV para Google Agenda

Este projeto tem como objetivo gerar automaticamente um arquivo **CSV compatível com o Google Agenda**, permitindo a importação de eventos a partir de uma planilha Excel.

O script lê uma planilha `.xlsx`, processa as informações de eventos e gera um arquivo `.csv` pronto para importação no Google Agenda.

---

## 📌 Funcionalidades

- Leitura de planilha Excel (`.xlsx`)
- Conversão automática de datas
- Conversão de horários para formato **AM/PM**
- Tratamento de intervalos de dias (`15`, `15 e 16`, `13 a 15`)
- Definição de valor padrão para campos vazios
- Geração de arquivo **CSV compatível com Google Agenda**
- Geração de arquivo **XLSX com registros não exportados**

---

## ⚙️ Requisitos

- Python 3.10+
- Bibliotecas Python:
  - pandas
  - os
  - re
  - datetime

---

## ⚙️ Instalação:

Instale as dependências do projeto:
```bash
pip install -r requirements.txt
```
Execute o comando abaixo para criar o arquivo de configuração:
```bash
cp config-example.py config.py
```
Acesse o arquivo `config.py` e informe os parâmetros necessários.

---

## ▶️ Como usar

1. Coloque o arquivo Excel na pasta:
```bash
xlsx.in/
```
2. Execute o script:
```bash
python createCsv.py
```
3. Os arquivos gerados serão gravados nos seguintes diretórios:

- Arquivo CSV
```bash
csv_out/
```
- Arquivo XLSX com os eventos que não foram inseridos no CSV
```bash
nao_processados/
```
---

## 📝 Observações

- Apenas eventos que possuem `HORA INICIAL` e `HORA FINAL` são exportados.
- Eventos sem horário serão gravados no arquivo `nao_exportados.xlsx`.
- Caso o campo `LOCAL` venha vazio, será utilizado um valor padrão (definido no config.py).

---

## 👤 Autor

Projeto desenvolvido por Wendel Ribeiro Silva.