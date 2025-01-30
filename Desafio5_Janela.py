import tkinter as tk

def on_button_click():
    print("Botão clicado!")

# Cria a janela principal
janela = tk.Tk()
janela.title("TELA 1 - ATHENAS CONSULTORIA AGRÍCOLA")
janela.geometry("800x600")
janela.resizable(True, True)  # Permite redimensionar largura e altura
janela.state('zoomed')  # Maximiza a janela

# Cria um botão
botao = tk.Button(janela, text="Clique aqui", command=on_button_click)
botao.pack(pady=20)

# Inicia o loop principal da aplicação
janela.mainloop()