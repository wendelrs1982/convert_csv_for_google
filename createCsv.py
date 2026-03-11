import config
import os
import pandas as pd
import re
from datetime import datetime

ANO = config.ANO_AGENDA

PASTA_ENTRADA = config.PASTA_IN_XLSX
PASTA_SAIDA = config.PASTA_OUT_CSV
PASTA_NAO_PROCESSADOS = config.PASTA_XLSX_NAO_PROCESSADOS

ARQUIVO_ENTRADA = os.path.join(PASTA_ENTRADA, config.NOME_ARQUIVO_XLSX)
ARQUIVO_CSV = os.path.join(PASTA_SAIDA, config.NOME_ARQUIVO_CSV)
ARQUIVO_NAO_EXPORTADOS = os.path.join(PASTA_NAO_PROCESSADOS, config.ARQUIVO_XLSX_NAO_PROCESSADOS)

os.makedirs(PASTA_SAIDA, exist_ok=True)

meses = {
    "JANEIRO": "01",
    "FEVEREIRO": "02",
    "MARÇO": "03",
    "MARCO": "03",
    "ABRIL": "04",
    "MAIO": "05",
    "JUNHO": "06",
    "JULHO": "07",
    "AGOSTO": "08",
    "SETEMBRO": "09",
    "OUTUBRO": "10",
    "NOVEMBRO": "11",
    "DEZEMBRO": "12"
}

def extrair_dias(texto):

    texto = str(texto).lower()

    if " a " in texto:
        inicio, fim = map(int, re.findall(r'\d+', texto))
        return list(range(inicio, fim + 1))

    if " e " in texto:
        return list(map(int, re.findall(r'\d+', texto)))

    return [int(re.findall(r'\d+', texto)[0])]

def converter_hora(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor).lower().replace("h", "").strip()

    # se vier no formato 20:00:00
    if len(texto.split(":")) == 3:
        texto = ":".join(texto.split(":")[:2])

    # se vier apenas hora (20 ou 07)
    if ":" not in texto:
        texto = f"{texto}:00"

    hora = datetime.strptime(texto, "%H:%M")

    return hora.strftime("%I:%M %p")

print("\nIniciando o Processamento...")
print("Lendo o arquivo:", ARQUIVO_ENTRADA)

df = pd.read_excel(ARQUIVO_ENTRADA)

df = df.dropna(how="all")

# NOVA REGRA
# exporta somente linhas com algum horário

df_validos = df[
    df["HORA INICIAL"].notna() |
    df["HORA FINAL"].notna()
]

df_invalidos = df[
    df["HORA INICIAL"].isna() &
    df["HORA FINAL"].isna()
]

linhas_saida = []

for _, row in df_validos.iterrows():

    mes_nome = str(row["MÊS"]).upper().strip()

    if mes_nome not in meses:
        print("Mês inválido:", mes_nome)
        continue

    mes_num = meses[mes_nome]

    dias = extrair_dias(row["DIA"])

    hora_inicio = converter_hora(row["HORA INICIAL"])
    hora_fim = converter_hora(row["HORA FINAL"])

    local = row["LOCAL"] 

    if pd.isna(local) or str(local).strip() == "":
        local = config.LOCAL_EVENTO

    for dia in dias:

        data = f"{str(dia).zfill(2)}/{mes_num}/{ANO}"

        linhas_saida.append({
            "Subject": row["EVENTO"],
            "Start Date": data,
            "End Date": data,
            "Start Time": hora_inicio,
            "End Time": hora_fim,
            "All Day Event": "False",
            "Description": row["RESPONSAVEL"],
            "Location": local
        })


df_saida = pd.DataFrame(linhas_saida)

df_saida.to_csv(ARQUIVO_CSV, index=False, encoding="utf-8")

# Gerando XLSX com linhas não exportadas para o CSV
df_invalidos.to_excel(ARQUIVO_NAO_EXPORTADOS, index=False)

print("CSV gerado:", ARQUIVO_CSV)
print("Eventos exportados:", len(df_saida))
print("Eventos ignorados:", len(df_invalidos))
print("Arquivo contendoos eventos que não foram exportados:", ARQUIVO_NAO_EXPORTADOS)  
print("Processamento Concluído\n")