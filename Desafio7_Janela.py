import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph


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


def salvar_pdf(results, file_path):
    pdf = SimpleDocTemplate(file_path, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    # Título do PDF
    title = Paragraph("Relatório de Análise de Solo", styles['Title'])
    elements.append(title)

    # Tabela de dados
    table_data = [results[0]] + results[1:]
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)

    # Gera o PDF
    pdf.build(elements)


def process_excel(folder_path, save_format):
    results_folder = os.path.join(folder_path, f"Resultados_{save_format.upper()}")
    os.makedirs(results_folder, exist_ok=True)

    # Verifica se há arquivos que começam com "RE_Química" no diretório
    arquivos_encontrados = [f for f in os.listdir(folder_path) if
                            f.startswith("RE_Química") and f.endswith(('.xls', '.xlsx'))]

    if not arquivos_encontrados:
        messagebox.showwarning("Aviso", "Nenhum arquivo 'RE_Química' encontrado no diretório selecionado.")
        return

    for file_name in arquivos_encontrados:
        file_path = os.path.join(folder_path, file_name)
        try:
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
            elif save_format == 'excel':
                excel_file_path = os.path.join(results_folder, f"{file_base_name}.xlsx")
                df_results = pd.DataFrame(results[1:], columns=results[0])
                df_results.to_excel(excel_file_path, index=False)
                print(f"Arquivo Excel salvo em: {excel_file_path}")
            elif save_format == 'pdf':
                pdf_file_path = os.path.join(results_folder, f"{file_base_name}.pdf")
                salvar_pdf(results, pdf_file_path)
                print(f"Arquivo PDF salvo em: {pdf_file_path}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo {file_name}: {str(e)}")
            continue

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
        ttk.Radiobutton(frame, text="Excel", variable=self.save_format, value="excel").grid(row=4, column=0,
                                                                                            sticky=tk.W)
        ttk.Radiobutton(frame, text="PDF", variable=self.save_format, value="pdf").grid(row=5, column=0, sticky=tk.W)

        # Botão de processamento
        ttk.Button(frame, text="Processar", command=self.processar).grid(row=6, column=0, columnspan=2, pady=10)

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