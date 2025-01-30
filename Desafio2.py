import subprocess
import sys
import os
import pandas as pd
from datetime import datetime


# Função para instalar os módulos necessários
def instalar_modulo(modulo):
    subprocess.check_call([sys.executable, "-m", "pip", "install", modulo])


# Verifica e instala os módulos necessários
for modulo in ["pandas", "openpyxl"]:
    try:
        __import__(modulo)
    except ImportError:
        print(f"O módulo {modulo} não está instalado. Instalando agora...")
        instalar_modulo(modulo)


# Função para ocultar 'S24/' do número da amostra
def ocultar_s24(sample_number):
    return sample_number.replace('S24/', '') if isinstance(sample_number, str) else sample_number


# Função para formatar os valores
def formatar_valor(value):
    """Formata o valor com vírgula no lugar do ponto decimal e retorna '0' para valores zerados."""
    try:
        if pd.isna(value) or isinstance(value, str):
            return "0"
        formatted_value = f"{float(value):.2f}".replace('.', ',')
        return "0" if formatted_value == "0,00" else formatted_value
    except (ValueError, TypeError):
        return "0"


# Função principal para processar os arquivos Excel e gerar TXT
def process_excel_to_txt():
    # Pergunta ao usuário a pasta onde estão os arquivos
    folder_path = input("Digite o caminho da pasta onde buscar os arquivos do Excel e aperte ENTER, Ex: Dir -> ")

    if not os.path.isdir(folder_path):
        print("❌ Caminho inválido. Verifique e tente novamente.")
        return

    # Obtém a data e hora atual no formato DD-MM-AAAA_HH-MM-SS
    timestamp = datetime.now().strftime('%d-%m-%Y_%H_%M_%S')

    for file_name in os.listdir(folder_path):
        if file_name.startswith("RE_Química") and file_name.endswith(('.xls', '.xlsx')):
            file_path = os.path.join(folder_path, file_name)

            # Carrega o arquivo Excel
            df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')

            # Cabeçalho fixo
            header_client = "p\t\t\tDesafio Técnico"

            # Filtra linhas que contenham "FAZENDA"
            farms = df[df.iloc[:, 1].astype(str).str.contains("FAZENDA", na=False, case=False)]

            if farms.empty:
                print(f"⚠️ Nenhuma fazenda encontrada no arquivo: {file_name}")
                continue

            results = [header_client]

            for _, row in farms.iterrows():
                farm_name = row.iloc[1]  # Nome da Fazenda
                sample_number = ocultar_s24(row.iloc[0])  # Número da Amostra sem 'S24/'

                # Data formatada
                today = datetime.today().strftime('%d-%m-%Y')

                # Adiciona cabeçalhos
                results.append("")
                results.append(f"f\t{farm_name}")
                results.append(f"a\t{sample_number}\t{datetime.today().year}\t{today}")

                # Dicionário de valores
                values = {
                    "01MO": row.iloc[6],
                    "01PHCACl2": row.iloc[5],
                    "01PRES": row.iloc[7],
                    "01K": row.iloc[12],
                    "01CA": row.iloc[9],
                    "01MG": row.iloc[10],
                    "01AL": row.iloc[13],
                    "01H+Al": row.iloc[14],
                    "01S": row.iloc[8],
                    "01H": row.iloc[14],
                    "01SB": row.iloc[15],
                    "01CTC": row.iloc[16],
                    "01V": row.iloc[17],
                    "01M": row.iloc[18],
                    "01KCTC": (row.iloc[12] / row.iloc[16]) * 100 if row.iloc[16] != 0 else 0,
                    "01CACTC": (row.iloc[9] / row.iloc[16]) * 100 if row.iloc[16] != 0 else 0,
                    "01MGCTC": (row.iloc[10] / row.iloc[16]) * 100 if row.iloc[16] != 0 else 0,
                    "01CAMG": ((row.iloc[9] + row.iloc[10]) / row.iloc[16]) * 100 if row.iloc[16] != 0 else 0,
                }

                # Unidades de medida
                unit_map = {
                    "01MO": "g.dm-3", "01S": "mg.dm-3", "01PRES": "mg.dm-3", "01K": "mmolc.dm-3",
                    "01CA": "mmolc.dm-3", "01MG": "mmolc.dm-3", "01AL": "mmolc.dm-3", "01H+Al": "mmolc.dm-3",
                    "01H": "mmolc.dm-3", "01SB": "mmolc.dm-3", "01CTC": "mmolc.dm-3", "01V": "%",
                    "01M": "%", "01KCTC": "%", "01CACTC": "%", "01MGCTC": "%", "01CAMG": "%", "01PHCACl2": ""
                }

                # Adiciona os resultados formatados
                for key, value in values.items():
                    results.append(f"r\t{key}\t{unit_map.get(key, '')}\t{formatar_valor(value)}")

            # Cria uma pasta para armazenar os arquivos TXT
            output_folder = os.path.join(folder_path, "Resultados_TXT")
            os.makedirs(output_folder, exist_ok=True)

            # Salva cada execução com nome único para evitar sobrescrita
            txt_file_name = f"{os.path.splitext(file_name)[0]}_{timestamp}.txt"
            txt_file_path = os.path.join(output_folder, txt_file_name)

            with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write("\n".join(results))

            print(f"✅ Arquivo salvo: {txt_file_path}")


if __name__ == "__main__":
    process_excel_to_txt()