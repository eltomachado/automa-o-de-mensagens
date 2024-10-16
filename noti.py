import pandas as pd
import pywhatkit as kit
from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox

# Função para enviar mensagem no WhatsApp
def enviar_mensagem_whatsapp(numero_cliente, mensagem):
    try:
        kit.sendwhatmsg_instantly(f"+{numero_cliente}", mensagem)
        print(f"Mensagem enviada para {numero_cliente}: {mensagem}")
    except Exception as e:
        print(f"Erro ao enviar mensagem para {numero_cliente}: {e}")

# Função para gerar a mensagem personalizada
def gerar_mensagem(nome, data_entrega, numero_nota):
    return (f" *Prezado Cliente {nome},*\n\n"
            f"Gostaríamos de informá-lo que sua carga foi entregue com sucesso no endereço indicado na data *{data_entrega}*.\n\n"
            f"Número da nota fiscal: *{numero_nota}*.\n\n"
            f"⚠️ Caso tenha qualquer problema ou sugestão, não hesite em entrar em contato com o nosso SAC pelo número 📞 *(11) 97087-9123*.\n\n"
            f"✨ *A equipe da 123 Util está à disposição para atendê-lo.*\n\n"
            f"👍 *Gostaríamos de saber como foi sua experiência com a entrega!*\n"
            f"Por favor, avalie seu atendimento respondendo com um número:\n\n"
            f"1️⃣ - Ótimo\n"
            f"2️⃣ - Bom\n"
            f"3️⃣ - Regular\n"
            f"4️⃣ - Ruim\n\n"
            f"Atenciosamente, \n*123 Útil*")

# Função para carregar e processar a planilha
def processar_planilha(caminho_planilha):
    # Ler a planilha de clientes
    planilha = pd.read_excel(caminho_planilha)

    # Loop para verificar cada cliente
    for index, cliente in planilha.iterrows():
        nome = cliente['Nome']
        telefone = cliente['Telefone']
        data_entrega = cliente['Data'].strftime('%d/%m/%Y')  # Converter a data para o formato desejado
        numero_nota = cliente['Nota']

        # Mensagem personalizada
        mensagem = gerar_mensagem(nome, data_entrega, numero_nota)

        # Enviar a mensagem
        enviar_mensagem_whatsapp(telefone, mensagem)
        print(f"Mensagem enviada para {nome} no número {telefone}")

# Função para iniciar o processo com os dados da interface
def iniciar_processamento():
    caminho_planilha = os.path.join(os.path.dirname(__file__), 'clientes.xlsx')
    processar_planilha(caminho_planilha)
    messagebox.showinfo("Info", "Mensagens enviadas com sucesso!")

# Configurar a interface gráfica
root = tk.Tk()
root.title("Configuração de Notificação")

# Configurar a largura e a altura da janela
root.geometry("400x200")

# Botão para iniciar o processamento
processar_button = tk.Button(root, text="Enviar Mensagens", command=iniciar_processamento)
processar_button.pack(pady=20)

# Iniciar o loop da interface gráfica
root.mainloop()
