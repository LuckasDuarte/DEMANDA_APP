import math
import pandas as pd
from openpyxl import load_workbook
from tkinter import Tk, Button, Label, filedialog
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox


# gerar executavel => pyinstaller --onefile --windowed --icon=icon.ico app.py
# COLOCAR OS ARQUIVOS NO MESMO DIRETORIO

# Função para arredondar o valor para o múltiplo mais próximo de multiplo
def marred(valor, multiplo):
    return round(valor / multiplo) * multiplo

# Função para calcular a quantidade a ser enviada com base na lógica especificada
def calcular_quantidade_envio(produtos):

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

            if valor_arredondado >= SALDO_CD:
                quantidade_a_enviar = SALDO_CD
            else:
                quantidade_a_enviar = valor_arredondado

        except Exception as e:
            print(f"Erro: {e}")
            quantidade_a_enviar = 0

        # Adicionar resultados à lista
        resultados.append({
            "COD": produto['COD'],
            "DESCRICAO": produto['DESCRICAO'],
            "ENVIAR": quantidade_a_enviar
        })

    return resultados

# Função para escrever os resultados em uma nova aba do Excel
def escrever_resultados_no_excel(nome_arquivo, resultados):
    # Converter os resultados para um DataFrame
    df_resultados = pd.DataFrame(resultados)

    # Carregar a planilha existente
    book = load_workbook(nome_arquivo)

    # Escrever os resultados em uma nova aba chamada 'Demanda'
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl', mode='a') as writer:
        df_resultados.to_excel(writer, sheet_name='Demanda', index=False)

    print(f"Os resultados foram adicionados na aba 'Demanda' da planilha '{nome_arquivo}'")

# Função para carregar o arquivo Excel e processar os dados
def carregar_e_processar_arquivo():
    nome_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if nome_arquivo:
        df_produtos = pd.read_excel(nome_arquivo)
        resultados = calcular_quantidade_envio(df_produtos)
        escrever_resultados_no_excel(nome_arquivo, resultados)
        messagebox.showinfo("SUCESSO!","DEMANDA GERADA COM SUCESSO!")

# Função para redimensionar a imagem
def redimensionar_imagem(caminho_imagem, largura, altura):
    img = Image.open(caminho_imagem)
    img = img.resize((largura, altura), Image.ANTIALIAS)
    return ImageTk.PhotoImage(img)

# Função para criar um tooltip
def criar_tooltip(widget, texto):
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

# Função para alternar entre os frames
def mostrar_frame():
    main.destroy()
    Frame_Process.grid(row=1, column= 0)

# Configuração da interface Tkinter
root = Tk()
root.title("DEMANDA AUTOMÁTICA")
root.iconbitmap("icon.ico")
root.geometry('1000x600')
root.resizable(False, False)
root.configure(bg='#ccc')

# ---- TELA HOME ---- #
header =  tk.Frame(root, bg='#06074f', relief='solid', width=1000, height=80)
header.propagate(False)
header.grid(row=0, column= 0)

# FRAME LOGO
img_logo = redimensionar_imagem('logo.png', 200, 200)

main =  tk.Frame(root,bg='#ccc', relief='solid', width=1000, height=720)
main.propagate(False)
main.grid(row=1, column= 0)

lbl_ImgLogo = Label(main, image= img_logo, width=200, height=200, bg="#ccc")
lbl_ImgLogo.pack(pady=120)

# Frame Processar
Frame_Process =  tk.Frame(root,bg='#ccc', relief='solid', width=1000, height=720)
Frame_Process.propagate(False)

#BTN EMPURRAR
img_empurrar = redimensionar_imagem('distribuicao.png', 70, 60)

# --- BTN EMPURRAR --- #
btnEmpurrar = tk.Button(header, image=img_empurrar, relief="solid", border=0, cursor='hand2', command=mostrar_frame)
btnEmpurrar.place(x=860, y=8)

# Criar o tooltip para o botão
criar_tooltip(btnEmpurrar, "Clique aqui para processar o arquivo Excel")


img_upload = redimensionar_imagem('upload.png', 40 , 40)

btn_carregar = Button(Frame_Process,image= img_upload ,text="CARREGAR",width=140,height=40,compound="left", relief="flat", cursor="hand2", command=carregar_e_processar_arquivo)
btn_carregar.place(x=10, y=20)

criar_tooltip(btn_carregar, "CARREGAR ARQUIVO DEMANDA")

root.mainloop()
