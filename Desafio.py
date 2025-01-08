import subprocess
import sys

# Função para instalar os módulos necessários
def instalar_modulo(modulo):
    subprocess.check_call([sys.executable, "-m", "pip", "install", modulo])

## Verifica e instala os módulos necessários
try:
    import os
    import pandas as pd
    from datetime import datetime
except ImportError as e:
    modulo_faltante = str(e).split()[-1][1:-1]  # Extrai o nome do módulo faltante
    print(f"O módulo {modulo_faltante} não está instalado. Instalando agora...")
    instalar_modulo(modulo_faltante)
    # Re-importa os módulos após a instalação
    import os
    import pandas as pd
    from datetime import datetime

def ocultar_s24(sample_number):
    # Substitui 'S24' por asteriscos, mantendo o restante do número intacto
    return sample_number.replace('S24/', '')

def formatar_valor(value):
    """Formata o valor com vírgula no lugar do ponto decimal, retornando '0' se for 0.00."""
    try:
        if pd.isna(value) or isinstance(value, str):
            return "0"  # Retorna vazio se for texto ou NaN
        formatted_value = f"{float(value):.2f}"
        if formatted_value == "0.00":
            return "0"  # Retorna apenas 0 para valores 0.00
        return formatted_value.replace('.', ',')
    except (ValueError, TypeError):
        return ""

def process_excel_to_txt():
    # Pergunta ao usuário a pasta de busca
    folder_path = input("Digite o caminho da pasta onde buscar os arquivos: ")

    # Verifica se o caminho é válido
    if not os.path.isdir(folder_path):
        print("Caminho inválido. Certifique-se de que o caminho está correto.")
        return

    # Percorre a pasta buscando arquivos que começam com "RE_Química"
    for file_name in os.listdir(folder_path):
        if file_name.startswith("RE_Química") and file_name.endswith(('.xls', '.xlsx')):
            file_path = os.path.join(folder_path, file_name)

            # Carrega o arquivo Excel
            df = pd.read_excel(file_path, sheet_name=0)

            # Extrai o cabeçalho
            header_client = "p\t\t\tDesafio Técnico"  # Cabeçalho

            # Identifica as fazendas e processa os dados
            farms = df[df.iloc[:, 1].str.contains("FAZENDA", na=False)]

            results = [header_client]

            for _, row in farms.iterrows():
                farm_name = row.iloc[1]  # Nome da Fazenda
                sample_number = row.iloc[0]  # Número da Amostra
                sample_number_ocultado = ocultar_s24(sample_number)  # Oculta o 'S24'

                # Data no formato DDMMAAAA
                year = datetime.today().strftime('%Y')
                today = datetime.today().strftime('%d%m%Y')

                # Adiciona cabeçalhos e informações da fazenda
                results.append("")
                results.append(f"f\t{farm_name}")
                results.append(f"a\t{sample_number_ocultado}\t{year}\t{today}")

                # Extrai os resultados necessários
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

                # Adiciona os resultados formatados
                for key, value in values.items():
                    formatted_value = formatar_valor(value)

                    # Adiciona as linhas no formato correto
                    unit = {
                        "01MO": "g.dm-3",
                        "01S": "mg.dm-3",
                        "01PRES": "mg.dm-3",
                        "01K": "mmolc.dm-3",
                        "01CA": "mmolc.dm-3",
                        "01MG": "mmolc.dm-3",
                        "01AL": "mmolc.dm-3",
                        "01H+Al": "mmolc.dm-3",
                        "01H": "mmolc.dm-3",
                        "01SB": "mmolc.dm-3",
                        "01CTC": "mmolc.dm-3",
                        "01V": "%",
                        "01M": "%",
                        "01KCTC": "%",
                        "01CACTC": "%",
                        "01MGCTC": "%",
                        "01CAMG": "%",
                        "01PHCACl2": "",
                    }.get(key, "")

                    results.append(f"r\t{key}\t{unit}\t{formatted_value}")

            # Salva o arquivo TXT com o mesmo nome do Excel e a data atual
            txt_file_name = f"{os.path.splitext(file_name)[0]}_{today}.txt"
            txt_file_path = os.path.join(folder_path, txt_file_name)

            with open(txt_file_path, 'w') as txt_file:
                txt_file.write("\n".join(results))

            print(f"Arquivo processado e salvo em: {txt_file_path}")


if __name__ == "__main__":
    process_excel_to_txt()
