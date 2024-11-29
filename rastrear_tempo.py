import requests
import json
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime

workbook = Workbook()
sheet = workbook.active

team_id = "3075996"
response = {}
response["access_token"] = (
    "3138343_6fada438d889343a074632a34b7c61afcff6e35c1bc63ba13127367756e4d2b2"
)

headers = {"Authorization": f'Bearer {response["access_token"]}', 'Content-Type': 'application/json'}

ids_usuarios = [
    {"id": 82184428, "username": "Thiago Ribeiro Pompermayer"},
{"id": 82174345, "username": "Gabriel Farias"},
{"id": 82173689, "username": "Mateus Ruy Soares Gaudio"},
{"id": 82164056, "username": "João Pedro Souza Rocha"},
{"id": 82155499, "username": "Pedro Vieira Lopes"},
{"id": 82150797, "username": "Gustavo Stein"},
{"id": 82150464, "username": "Mariana Gomes Calheiros"},
{"id": 82145319, "username": "Henrique Loss Lopes"},
{"id": 3138354, "username": "Waldelicio Junior"},
{"id": 82083217, "username": "João Fernando Rangel Guimarães"},
{"id": 82074328, "username": "Michel Carvalho"},
{"id": 89143754, "username": "Anderson Ferreira"},
{"id": 60976772, "username": "Ana Luiza"},
# {"id": 3138012, "username": "Ricardo Calheiros"},
# {"id": 84164888, "username": "Bianca Aguiar"},
{"id": 82042179, "username": "Gabriel da Silva Biancardi"},
{"id": 84636582, "username": "Willian Pacheco Silva"},
{"id": 84636590, "username": "Bernardo Zampirole Brandão"},
{"id": 81917609, "username": "Rafael Antunes Costa"},
# {"id": 3138355, "username": "GUSTAVO FRINHANI"},
{"id": 3138350, "username": "Melquisedeque Shaloon Bento da Silva Gomes"},
{"id": 60932018, "username": "Fernando Bisi Vieira"},
{"id": 3269111, "username": "Henrique Puppim"},
{"id": 3248906, "username": "Rodrigo Merigueti"},
# {"id": 3219862, "username": "Mauricio Calheiros"},
{"id": 3164413, "username": "Phelipe Augusto"},
# {"id": 3138343, "username": "Dayane Erlacher"},
# {"id": 3138065, "username": "Franco Louzada"},
{"id": 3137926, "username": "Menno"}
]

# Funções Auxiliares
def obter_dados_api(url):
    response = requests.get(url, headers=headers)
    return response.json() if response.status_code == 200 else []

try:
    strdata = "24/11/2024"
    data = datetime.strptime(strdata, "%d/%m/%Y")
    dias_uteis = 4
    start_date = int(data.timestamp() * 1000)
    
    sheet.append(["Colaborador", "Tempo", "%"])
    for id_user in ids_usuarios:
        detaild_user = obter_dados_api(f"https://api.clickup.com/api/v2/team/{team_id}/time_entries?start_date={start_date}&assignee={id_user["id"]}")["data"]
        time_count = 0
        for user in detaild_user:
            time_count += int(user["duration"])
        # Considerando 80% do tempo
        sheet.append([id_user["username"], time_count/3600000,(time_count/3600000)/(dias_uteis*7)])
    workbook.save("usuarios.xlsx")
except requests.RequestException as e:
    print(f"Erro de requisição: {e}")