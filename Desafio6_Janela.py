import subprocess
import sys
import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def instalar_modulo(modulo):
    subprocess.check_call([sys.executable, "-m", "pip", "install", modulo])

def ocultar_s24(sample_number):
    return sample_number.replace('S24/', '')

def formatar_valor(value):
    try:
        if pd.isna(value) or isinstance(value, str):
            return "0"
        formatted_value = f"{float(value):.2f}"
        return "0" if formatted_value == "0.00" else formatted_value.replace('.', ',')
    except (ValueError, TypeError):
        return ""

def process_excel(folder_path, save_format):
    results_folder = os.path.join(folder_path, "Resultados_TXT" if save_format == 'txt' else "Resultados_EXCEL")
    os.makedirs(results_folder, exist_ok=True)

    for file_name in os.listdir(folder_path):
        if file_name.startswith("RE_Química") and file_name.endswith(('.xls', '.xlsx')):
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_excel(file_path, sheet_name=0)
            farms = df[df.iloc[:, 1].str.contains("FAZENDA", na=False)]

            today = datetime.today().strftime('%d_%m_%Y_%H_%M')
            results = [["Fazenda", "Amostra", "Ano", "Data"]]

            for _, row in farms.iterrows():
                farm_name = row.iloc[1]
                sample_number = ocultar_s24(row.iloc[0])
                year = datetime.today().strftime('%Y')
                results.append([farm_name, sample_number, year, today])

                values = {
                    "MO": row.iloc[6], "PH": row.iloc[5], "PRES": row.iloc[7],
                    "K": row.iloc[12], "CA": row.iloc[9], "MG": row.iloc[10],
                    "AL": row.iloc[13], "H+Al": row.iloc[14], "S": row.iloc[8],
                    "SB": row.iloc[15], "CTC": row.iloc[16], "V": row.iloc[17],
                    "M": row.iloc[18]
                }
                for key, value in values.items():
                    results.append([key, formatar_valor(value)])

            file_base_name = f"{os.path.splitext(file_name)[0]}_{today}"

            if save_format == 'txt':
                txt_file_path = os.path.join(results_folder, f"{file_base_name}.txt")
                with open(txt_file_path, 'w') as txt_file:
                    for row in results:
                        txt_file.write("\t".join(map(str, row)) + "\n")
                print(f"Arquivo TXT salvo em: {txt_file_path}")
            else:
                excel_file_path = os.path.join(results_folder, f"{file_base_name}.xlsx")
                df_results = pd.DataFrame(results[1:], columns=results[0])
                df_results.to_excel(excel_file_path, index=False)
                print(f"Arquivo Excel salvo em: {excel_file_path}")

    messagebox.showinfo("Sucesso", "Processamento concluído!")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Tela 1 - ATHENAS CONSULTORIA AGRÍCOLA")
        self.root.geometry("500x300")

        self.folder_path = tk.StringVar()
        self.save_format = tk.StringVar(value="txt")

        # Frame principal
        frame = ttk.Frame(root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # Seleção de pasta
        ttk.Label(frame, text="Selecione a pasta com os arquivos Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.folder_path, width=40).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(frame, text="Procurar", command=self.selecionar_pasta).grid(row=1, column=1, padx=5, pady=5)

        # Seleção de formato
        ttk.Label(frame, text="Escolha o formato de saída:").grid(row=2, column=0, sticky=tk.W)
        ttk.Radiobutton(frame, text="TXT", variable=self.save_format, value="txt").grid(row=3, column=0, sticky=tk.W)
        ttk.Radiobutton(frame, text="Excel", variable=self.save_format, value="excel").grid(row=4, column=0, sticky=tk.W)

        # Botão de processamento
        ttk.Button(frame, text="Processar", command=self.processar).grid(row=5, column=0, columnspan=2, pady=10)

    def selecionar_pasta(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def processar(self):
        folder_path = self.folder_path.get()
        save_format = self.save_format.get()

        if not folder_path:
            messagebox.showerror("Erro", "Selecione uma pasta válida!")
            return

        try:
            process_excel(folder_path, save_format)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()