# DOCS:
# https://clickup.com/api/clickupreference/operation/Gettimeentryhistory/

# Antes de rodar esse código, acesse esse link e pegue o código:
# https://app.clickup.com/api?client_id=9BQNB7IFOMPHHLCGVVE3R2R36H59AOYK&redirect_uri=https://www.2solve.com/

import requests
import json
from openpyxl import Workbook
from openpyxl import load_workbook
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
    "Gabriel Farias": 21.79, #mudar
    "GUSTAVO FRINHANI": 76.45,
    "Henrique Loss Lopes": 21.79,
    "Henrique Puppim": 32.04,
    "João Fernando Rangel Guimarães": 24.39,
    "João Pedro Souza Rocha": 29.59,
    "Mariana Gomes Calheiros": 22.23,
    "Mateus Ruy Soares Gaudio": 29.39, #mudar
    "Mauricio Calheiros": 83.57,
    "Melquisedeque Shaloon Bento da Silva Gomes": 47.64, #mudar
    "Menno": 75.68,
    "Michel Carvalho": 34.09,
    "Pedro Vieira Lopes": 21.23,
    "Phelipe Augusto": 36.83,
    "Rafael Antunes Costa": 29.39, #mudar
    "Rodrigo Merigueti": 33.51,
    "Waldelicio Junior": 93.43,
    "Willian Pacheco Silva": 33.46,
    "Ricardo Calheiros": 109.10
}
data_hoje = datetime.now().strftime("%Y-%m-%d")
space_exclude = ["90131034562", "90131103749", "90131210570", "90131669969", "90131064862", "90130523790", "90131679449"]
# id: 90131034562 nome: Marketing
# id: 90131103749 nome: Comercial
# id: 90131210570 nome: Gestão do Conhecimento
# id: 90131669969 nome: Processos e Qualidade
# id: 90131064862 nome: Gerente Gustavo
# id: 90130523790 nome: Gerente Maurício
# id: 90131679449 nome: Gerente Júnior


folder_exclude = ["90132771240", "90132869998", "90132004295", "90132676002", "90131094720", "90130575812", "90130575784"]
# id: 90132771240 nome: Brain Storm
# id: 90132869998 nome: Consultoria GSMAS
# id: 90132004295 nome: Gestão de projetos
# id: 90132676002 nome: TI
# id: 90131094720 nome: Tarefas Gerais
# id: 90130575812 nome: Interno Software
# id: 90130575784 nome: Suporte Software

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

headers = {"Authorization": f'Bearer {response["access_token"]}'}

def obter_dados_api(url):
    """Função para realizar requisição à API e retornar dados em JSON."""
    response = requests.get(url, headers=headers)
    return response.json() if response.status_code == 200 else None

try:
    spaces = obter_dados_api(f"https://api.clickup.com/api/v2/team/{team_id}/space?archived=false")["spaces"]

    for space in spaces:
        if space["id"] in space_exclude:
            continue
        space_name = space["name"]

        folders = obter_dados_api(f"https://api.clickup.com/api/v2/space/{space['id']}/folder?archived=false")["folders"]

        for folder in folders:
            if folder["id"] in folder_exclude:
                continue
            # print(folder["id"])
            # print(f'{space["name"]}_{folder["name"]}')
            # continue
            caminho_arquivo = f"./Projetos/{space["name"]}_{folder["name"]}.xlsx"
            print(caminho_arquivo)
            workbook = load_workbook(caminho_arquivo)
            # Seleciona a aba "Tarefas Gerais"
            if "Tarefas Gerais" in workbook.sheetnames:
                aba = workbook["Tarefas Gerais"]
            else:
                raise ValueError("A aba 'Tarefas Gerais' não existe na planilha.")

            aba.delete_rows(1, aba.max_row)
            aba.append(header)

            
            lists = obter_dados_api(f"https://api.clickup.com/api/v2/folder/{folder['id']}/list?archived=false")["lists"]
            if not lists:
                print("lists vazio")
                continue

            for lista in lists:
                tasks = obter_dados_api(f"https://api.clickup.com/api/v2/list/{lista['id']}/task?archived=false&include_closed=true")["tasks"]
                if not tasks:
                    print("tasks vazio")
                    continue
                
                for task in tasks:
                    responsaveis = ";".join(assignee["username"] for assignee in task["assignees"])
                    custos_presentes = [custo_funcionario.get(nome, 0) for nome in responsaveis.split(";")]
                    custo = max(custos_presentes, default=0)
                    # custo = 0
                    # custo += sum(custo_funcionario.get(nome, 0) for nome in responsaveis.split(";"))
                    if custo == 0:
                        print(f"custo 0: {folder['name']} - {responsaveis}")
                    time_estimate = (task.get("time_estimate") or 0) / 3600000
                    cost_estimate = time_estimate * custo
                    time_spent = task.get("time_spent", 0) / 3600000
                    cost_spent = time_spent * custo
                    percent_concluido = 1 if task["status"]["status"].lower() in {"concluídas", "fechadas"} else 0

                    # Adiciona os dados da tarefa à planilha
                    aba.append([
                        task["list"]["name"], task["name"], responsaveis,
                        time_estimate, cost_estimate, time_spent,
                        cost_spent, task["status"]["status"], percent_concluido
                    ])
            workbook.save(caminho_arquivo)

except requests.RequestException as e:
    print(f"Erro de requisição: {e}")