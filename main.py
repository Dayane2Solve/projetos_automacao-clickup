# DOCS:
# https://clickup.com/api/clickupreference/operation/Gettimeentryhistory/

# Antes de rodar esse código, acesse esse link e pegue o código:
# https://app.clickup.com/api?client_id=9BQNB7IFOMPHHLCGVVE3R2R36H59AOYK&redirect_uri=https://www.2solve.com/

import requests
import json
from openpyxl import Workbook
from datetime import datetime

team_id = "3075996"
header = [
    "SETOR",
    "TAREFA",
    "RESPONSÁVEL",
    "TEMPO ESTIMADO (h)",
    "CUSTO ESTIMADO",
    "TEMPO RASTREADO (h)",
    "CUSTO REAL",
    "STATUS",
    "% CONCLUÍDO",
]
custo_funcionario = {
    "Ana Luiza": 21.93,
    "Anderson Ferreira": 14.41,
    "Bernardo Zampirole Brandão": 29.39,
    "Bianca Aguiar": 20.37,
    "Dayane Erlacher": 47.64,
    "Fernando Bisi Vieira": 29.50,
    "Franco Louzada": 49.75,
    "Gabriel da Silva Biancardi": 21.79,
    "Gabriel Farias": 1,
    "GUSTAVO FRINHANI": 76.45,
    "Henrique Loss Lopes": 21.79,
    "Henrique Puppim": 32.04,
    "João Fernando Rangel Guimarães": 24.39,
    "João Pedro Souza Rocha": 29.59,
    "Mariana Gomes Calheiros": 22.23,
    "Mateus Ruy Soares Gaudio": 1,
    "Mauricio Calheiros": 83.57,
    "Melquisedeque Shaloon Bento da Silva Gomes": 1,
    "Menno": 75.68,
    "Michel Carvalho": 34.09,
    "Pedro Vieira Lopes": 21.23,
    "Phelipe Augusto": 36.83,
    "Rafael Antunes Costa": 1,
    "Rodrigo Merigueti": 33.51,
    "Waldelicio Junior": 93.43,
    "Willian Pacheco Silva": 33.46,
}
data_hoje = datetime.now()
data_formatada = data_hoje.strftime("%Y-%m-%d")
space_exclude = ["90131034562", "90131103749", "90131210570", "90131669969"]
# id: 90131034562 nome: Marketing
# id: 90131103749 nome: Comercial
# id: 90131210570 nome: Gestão do Conhecimento
# id: 90131669969 nome: Processos e Qualidade

folder_exclude = ["90132771240"]
# id: 90132771240 nome: Brain Storm

# code = "EWOVWCI2APY52Z57XDP79O88R0JS96SY"

# reqUrl = "https://app.clickup.com/api/v2/oauth/token"

# headersList = {
#  "Accept":.*/*,
#  "User-Agent": Thunder Client (https://www.thunderclient.co.),
#  "Content-Type": pplication.jso"
# }

# payload = json.dumps({
#   "client_id":.9BQNB7IFOMPHHLCGVVE3R2R36H59AOYK,
#   "client_secret":.YBR6BQJEACHDHELWHT4OZ16H49LWVJ2B51SH1DREN03BWP7RVMK6O73C1CI6SBPR,
#   "code": code,
#   "redirect_uri": https://www.2solve.com"
# })

# response = requests.request("POST", reqUrl, data=payload,  headers=headersList)

# response = response.json()
# if "access_token" in response:
#   print(response["access_token"])
# else:
#   print(response)

response = {}
response["access_token"] = (
    "3138343_6fada438d889343a074632a34b7c61afcff6e35c1bc63ba13127367756e4d2b2"
)

url = f"https://api.clickup.com/api/v2/team/{team_id}/space?archived=false"
headers = {"Authorization": f'Bearer {response["access_token"]}'}

try:
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        spaces = response.json()["spaces"]
        for space in spaces:
            workbook = Workbook()
            space_id = space["id"]
            if space_id in space_exclude:
                continue
            space_name = space["name"]
            url = (
                f"https://api.clickup.com/api/v2/space/{space_id}/folder?archived=false"
            )
            response = requests.get(url, headers=headers)

            if response.status_code == 200:
                folders = response.json()["folders"]
                for folder in folders:
                    folder_id = folder["id"]
                    if folder_id in folder_exclude:
                        continue

                    url = f"https://api.clickup.com/api/v2/folder/{folder_id}/list?archived=false"
                    response = requests.get(url, headers=headers)
                    if response.status_code == 200:
                        lists = response.json()["lists"]
                        if not lists:
                            continue
                        aba = workbook.create_sheet(
                            title=f'{lists[0]["folder"]["name"]}'
                        )
                        aba.append(header)
                        for list in lists:
                            list_id = list["id"]
                            print(list["name"])
                            url = f"https://api.clickup.com/api/v2/list/{list_id}/task?archived=false&include_closed=true"
                            response = requests.get(url, headers=headers)
                            if response.status_code == 200:
                                tasks = response.json()["tasks"]
                                for task in tasks:
                                    responsaveis = ";".join(
                                        assignee["username"]
                                        for assignee in task["assignees"]
                                    )
                                    nomes_lista = responsaveis.split(";")
                                    custos_presentes = [
                                        custo_funcionario[nome]
                                        for nome in nomes_lista
                                        if nome in custo_funcionario
                                    ]
                                    custo = (
                                        max(custos_presentes) if custos_presentes else 0
                                    )
                                    time_estimate = (
                                        task.get("time_estimate") or 0
                                    ) / 3600000
                                    cost_estimate = time_estimate * custo
                                    time_spent = task.get("time_spent", 0) / 3600000
                                    cost_spent = time_spent * custo
                                    percent_concluido = (
                                        1
                                        if task["status"]["status"].lower()
                                        == "concluído" or task["status"]["status"].lower()
                                        == "fechado" 
                                        else 0
                                    )

                                    # Adiciona os dados da tarefa à planilha
                                    aba.append(
                                        [
                                            task["list"]["name"],
                                            task["name"],
                                            responsaveis,
                                            time_estimate,
                                            cost_estimate,
                                            time_spent,
                                            cost_spent,
                                            task["status"]["status"],
                                            percent_concluido,
                                        ]
                                    )

                            else:
                                print(
                                    f"Erro ao obter listas da tarefa: código {response.status_code}"
                                )

                    else:
                        print(
                            f"Erro ao obter listas de pastas: código {response.status_code}"
                        )
                    
            else:
                print(f"Erro ao obter pastas do espaço: código {response.status_code}")
            
            workbook.remove(workbook["Sheet"])
            workbook.save(f"{data_formatada}_{space_name}.xlsx")
            


    else:
        print(f"Erro ao obter espaços: código {response.status_code}")

except requests.RequestException as e:
    print(f"Erro de requisição: {e}")
