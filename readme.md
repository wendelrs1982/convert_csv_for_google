# Gerador de CSV para Google Agenda

![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

Este projeto gera automaticamente um **arquivo CSV compatível com o Google Agenda** a partir de uma planilha Excel contendo informações de eventos.

O script lê uma planilha `.xlsx`, processa os dados (datas, horários e eventos) e gera um arquivo `.csv` pronto para importação no Google Agenda.

---

## 📌 Funcionalidades

* Leitura de planilha Excel (`.xlsx`)
* Conversão automática de datas
* Conversão de horários para formato **AM/PM**
* Tratamento de intervalos de dias (`15`, `15 e 16`, `13 a 15`)
* Definição de valor padrão para campos vazios
* Exportação apenas de eventos que possuem horário
* Geração de arquivo **CSV compatível com Google Agenda**
* Geração de arquivo **XLSX contendo registros não exportados**

---

## 📂 Estrutura do projeto

```
project/
│
├── createCsv.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
├── config-example.py
├── config.py
├── in/
│   └── agenda.xlsx
│
└── out/
    ├── google_agenda.csv
    └── eventos_nao_importados.xlsx
```

---

## ⚙️ Requisitos

* Python **3.10 ou superior**
* Bibliotecas Python:

```
pandas
openpyxl
```

Instale as dependências com:

```
pip install -r requirements.txt
```

Crie o arquivo `config.py`:

```
cp config-example.py config.py
```
Em seguida, abra o arquivo `config.py` e informe os parâmetros necessários.

---

## ▶️ Como usar

1️⃣ Coloque o arquivo Excel na pasta:

```
in/
```

2️⃣ Execute o script:

```
python createCsv.py
```

3️⃣ O resultado será gerado na pasta:

```
out/
```

Arquivos gerados:

* `google_agenda.csv`
* `eventos_nao_importados.xlsx`

---

## 📅 Layout de planilha contendo os compromissos

|  MÊS  | DIA       | DIA DA SEMANA | HORA INICIAL | HORA FINAL |   EVENTO  |   LOCAL      |  RESPONSAVEL |
| ----- | --------- | ------------- | ------------ | ---------- | --------- | ------------ | ------------ |
| MARÇO | 19        |   Quarta      | 20h          | 22:00      | Evento 01 | Restaurante  | Equipe 01    |
| ABRIL | 12 e 13   |   Quinta      | 20           | 22h        | Evento 02 | Pousada      | Equipe 02    |

---

## 📄 Layout do arquivo CSV gerado

| Subject   | Start Date  |  End Date  | Start Time | End Time | All Day Event | Description | Location     |
| --------- | ----------- | ---------- | ---------- | -------- | ------------- | ----------- | ------------ |
| Evento 01 | 19/03/2026  | 19/03/2026 | 08:00 PM   | 10:00 PM |     False     |  Equipe 01  | Restaurante  |
| Evento 02 | 12/04/2026  | 12/04/2026 | 08:00 PM   | 10:00 PM |     False     |  Equipe 02  | Pousada      |
| Evento 02 | 13/04/2026  | 12/04/2026 | 08:00 PM   | 10:00 PM |     False     |  Equipe 02  | Pousada      |

---

## 📅 Importando no Google Agenda

1. Acesse o **Google Agenda**
2. Clique em **Configurações**
3. Vá em **Importar e exportar**
4. Selecione o arquivo `agenda_convertida.csv`
5. Escolha o calendário desejado
6. Clique em **Importar**

---

## 📝 Regras de processamento

* Apenas linhas que possuem **HORA INICIAL ou HORA FINAL** são exportadas.
* Linhas sem horário são armazenadas no arquivo `nao_exportados.xlsx`.
* Caso o campo **LOCAL** esteja vazio, um valor padrão será utilizado.
* Intervalos de datas são automaticamente convertidos em múltiplos eventos.

Exemplo:

```
DIA: 15 e 16
```

gera:

```
15/03/2026
16/03/2026
```

---

## 📜 Licença

Este projeto está licenciado sob a licença **MIT**.

---

## 👤 Autor

Projeto desenvolvido por **Wendel Ribeiro**.