import os
import pandas as pd
import re
from datetime import datetime

ANO = "2026"

PASTA_ENTRADA = "xlsx_in"
PASTA_SAIDA = "csv_out"

ARQUIVO_ENTRADA = os.path.join(PASTA_ENTRADA, "agenda.xlsx")
ARQUIVO_SAIDA = os.path.join(PASTA_SAIDA, "agenda_convertida.csv")

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

    if ":" not in texto:
        texto = f"{texto}:00"

    hora = datetime.strptime(texto, "%H:%M")

    return hora.strftime("%I:%M %p")


print("Lendo arquivo:", ARQUIVO_ENTRADA)

df = pd.read_excel(ARQUIVO_ENTRADA)

df = df.dropna(how="all")

df = df[df["EXPORTAR"] == 1]

linhas_saida = []

for _, row in df.iterrows():

    mes_nome = str(row["MÊS"]).upper().strip()

    if mes_nome not in meses:
        print("Mês inválido:", mes_nome)
        continue

    mes_num = meses[mes_nome]

    dias = extrair_dias(row["DIA"])

    hora_inicio = converter_hora(row["HORA INICIAL"])
    hora_fim = converter_hora(row["HORA FINAL"])

    for dia in dias:

        data = f"{str(dia).zfill(2)}/{mes_num}/{ANO}"

        # linhas_saida.append({
        #     "DATA": data,
        #     "HORA_INICIO": hora_inicio,
        #     "HORA_FIM": hora_fim,
        #     "EVENTO": row["EVENTO"],
        #     "LOCAL": row["LOCAL"],
        #     "RESPONSAVEL": row["RESPONSAVEL"]
        # })

        linhas_saida.append({
            "Subject": row["EVENTO"],
            "Start Date": data,
            "End Date": data,
            "Start Time": hora_inicio,
            "End Time": hora_fim,
            "All Day Event": "False",
            "Description": row["RESPONSAVEL"],
            "Location": row["LOCAL"]
        })


df_saida = pd.DataFrame(linhas_saida)

df_saida.to_csv(ARQUIVO_SAIDA, index=False, encoding="utf-8")

print("CSV gerado com sucesso!")
print("Arquivo:", ARQUIVO_SAIDA)
print("Total de eventos exportados:", len(df_saida))