import math
import pandas as pd
from openpyxl import load_workbook
from tkinter import Tk, Button, Label, filedialog, messagebox
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk

import pandas as pd

# Comentários de instrução
# Comando para gerar o executável: 
# pyinstaller --onefile --windowed --icon=icon.ico app.py
# Certifique-se de colocar todos os arquivos necessários no mesmo diretório

# ---------- Funções de Cálculo e Manipulação de Dados ---------- #

def marred(valor, multiplo):
    """Arredonda o valor para o múltiplo mais próximo de multiplo."""
    return round(valor / multiplo) * multiplo

def calcular_quantidade_envio(produtos):
    """Calcula a quantidade a ser enviada com base na lógica especificada."""
    resultados = []

    for _, produto in produtos.iterrows():
        SALDO_CD = produto["SALDO_CD"]
        VENDAS = produto["VENDAS"]
        EXPOSICAO = produto["EXPOSICAO"]
        UNITIZACAO = produto["UNITIZACAO"]
        SALDO_PDV = produto["SALDO_PDV"]

        try:
            valor_calculado = SALDO_PDV - (VENDAS * 7) - EXPOSICAO
            valor_positivo = abs(valor_calculado)
            valor_arredondado = marred(valor_positivo, UNITIZACAO)

            quantidade_a_enviar = min(valor_arredondado, SALDO_CD)
        except Exception as e:
            print(f"Erro: {e}")
            quantidade_a_enviar = 0

        resultados.append({
            "COD": produto['COD'],
            "DESCRICAO": produto['DESCRICAO'],
            "ENVIAR": quantidade_a_enviar
        })

    return resultados

def escrever_resultados_no_excel(nome_arquivo, resultados):
    """Escreve os resultados em uma nova aba do Excel chamada 'Demanda'."""
    df_resultados = pd.DataFrame(resultados)
    book = load_workbook(nome_arquivo)

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl', mode='a') as writer:
        df_resultados.to_excel(writer, sheet_name='Demanda', index=False)

    print(f"Os resultados foram adicionados na aba 'Demanda' da planilha '{nome_arquivo}'")

# ---------- Funções de Interface Gráfica ---------- #

def carregar_e_processar_arquivo():
    """Carrega o arquivo Excel, processa os dados e escreve os resultados."""
    nome_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if nome_arquivo:
        df_produtos = pd.read_excel(nome_arquivo)
        resultados = calcular_quantidade_envio(df_produtos)
        escrever_resultados_no_excel(nome_arquivo, resultados)
        messagebox.showinfo("SUCESSO!", "DEMANDA GERADA COM SUCESSO!")

def redimensionar_imagem(caminho_imagem, largura, altura):
    """Redimensiona a imagem para as dimensões especificadas."""
    img = Image.open(caminho_imagem)
    img = img.resize((largura, altura), Image.ANTIALIAS)
    return ImageTk.PhotoImage(img)

def criar_tooltip(widget, texto):
    """Cria um tooltip para o widget fornecido."""
    tooltip = tk.Toplevel(root)
    tooltip.wm_overrideredirect(True)
    tooltip.wm_geometry(f"+{widget.winfo_pointerx()+10}+{widget.winfo_pointery()+10}")
    tooltip.configure(bg='white')

    label = tk.Label(tooltip, text=texto, bg='white', fg='black', padx=5, pady=5, relief='solid')
    label.pack()

    def mostrar_tooltip(event):
        tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        tooltip.deiconify()

    def esconder_tooltip(event):
        tooltip.withdraw()

    widget.bind("<Enter>", mostrar_tooltip)
    widget.bind("<Leave>", esconder_tooltip)
    tooltip.withdraw()

def mostrar_tela_demanda():
    root.destroy()

    # ---------- Tela Demanda ------------ #

    root_Demanda = Tk()
    root_Demanda.title("DEMANDA")
    root_Demanda.iconbitmap("icon.ico")
    root_Demanda.geometry('1200x600')
    root_Demanda.resizable(False, False)
    root_Demanda.configure(bg='#ccc')

    # ------- Itens Necessários --------
    # => Produtos e verificar quais estão ativos
    # => e carregar seus respectivos fornecedores

    # -- Frame Topo
    header = tk.Frame(root_Demanda, bg='#fff', relief='solid', width=1200, height=50)
    header.propagate(False)
    header.grid(row=0, column=0)

    # -- ComboBox --
    lbl_Fornecedor = Label(header, width=15, height=1, bg="#fff", text="FORNECEDORES", anchor="w", font=("Ivy", 10))
    lbl_Fornecedor.place(x=10, y=2)

    data = ["CSSR EDITORA SANTUARIO","DEVOTUM"]

    CB_Fornecedores = ttk.Combobox(header, values = data, width = 40)
    CB_Fornecedores.place(x= 10, y=28)

    # Botão Consultar Fornecedor
    img_lupa = redimensionar_imagem('lupa.png', 20, 20)

    BTN_Consultar_Fornecedor = tk.Button(header, image= img_lupa, relief="solid", border=0, cursor='hand2', bg="#fff")
    BTN_Consultar_Fornecedor.place(x=130, y=2)

    # -- BTN FILTRAR ---
    BTN_filtrar = tk.Button(header, text="FILTRAR", width=10, bg="#06074f", height=1, fg="#fff", cursor="hand2")
    BTN_filtrar.place(x=300, y=25)

    BTN_Buscar = tk.Button(header, text="BUSCAR", width=10, bg="#06074f", height=1, fg="#fff", cursor="hand2")
    BTN_Buscar.place(x=390, y=25)

    # Botão Infos
    img_Info = redimensionar_imagem('informacoes.png', 50, 50)

    BTN_Info = tk.Button(header, image= img_Info, relief="solid", border=0, cursor='hand2', bg="#fff")
    BTN_Info.place(x=1120, y=0)

    # -- ComboBox Operador--
    lbl_Operador = Label(header, width=15, height=1, bg="#fff", text="RESPONSÁVEL", anchor="w", font=("Ivy", 10))
    lbl_Operador.place(x=500, y=2)

    data_op = ["OPERADOR_1","OPERADOR_2","OPERADOR_3","OPERADOR_4"]

    CB_Operadores = ttk.Combobox(header, values = data_op, width = 20)
    CB_Operadores.place(x=500, y=28)



    # ---- TABELA PRODUTOS -----


    
    root_Demanda.mainloop()


def mostrar_tela_pedidos():
    root.destroy()

    # ---------- Tela Demanda ------------ #

    root_Pedidos = Tk()
    root_Pedidos.title("PEDIDOS")
    root_Pedidos.iconbitmap("icon.ico")
    root_Pedidos.geometry('1000x600')
    root_Pedidos.resizable(False, False)
    root_Pedidos.configure(bg='#ccc')
    
    root_Pedidos.mainloop()

# ---------- Configuração da Interface Tkinter ---------- #

root = Tk()
root.title("DEMANDA AUTOMÁTICA")
root.iconbitmap("icon.ico")
root.geometry('1000x600')
root.resizable(False, False)
root.configure(bg='#ccc')

# Header
header = tk.Frame(root, bg='#06074f', relief='solid', width=1000, height=80)
header.propagate(False)
header.grid(row=0, column=0)

# Frame principal
img_logo = redimensionar_imagem('logo.png', 200, 200)

main = tk.Frame(root, bg='#ccc', relief='solid', width=1000, height=720)
main.propagate(False)
main.grid(row=1, column=0)

lbl_ImgLogo = Label(main, image=img_logo, width=200, height=200, bg="#ccc")
lbl_ImgLogo.pack(pady=120)

# Frame de Processamento
Frame_Process = tk.Frame(root, bg='#ccc', relief='solid', width=1000, height=720)
Frame_Process.propagate(False)

# Botão Empurrar
img_empurrar = redimensionar_imagem('distribuicao.png', 70, 60)

btnEmpurrar = tk.Button(header, image=img_empurrar, relief="solid", border=0, cursor='hand2', command=mostrar_tela_demanda)
btnEmpurrar.place(x=900, y=8)
criar_tooltip(btnEmpurrar, "Empurrar Demanda")

# Botão Pedidos
img_Comprar = redimensionar_imagem('shopping-cart.png', 70, 60)

btnComprar = tk.Button(header, image=img_Comprar, relief="solid", border=0, cursor='hand2', command=mostrar_tela_pedidos)
btnComprar.place(x=800, y=8)
criar_tooltip(btnComprar, "Realizar Pedidos de Compra")

# Botão Carregar
img_upload = redimensionar_imagem('upload.png', 40, 40)

btn_carregar = Button(Frame_Process, image=img_upload, text="CARREGAR", width=140, height=40, compound="left", relief="flat", cursor="hand2", command=carregar_e_processar_arquivo)
btn_carregar.place(x=10, y=20)
criar_tooltip(btn_carregar, "CARREGAR ARQUIVO DEMANDA")

# Iniciar a interface
root.mainloop()
