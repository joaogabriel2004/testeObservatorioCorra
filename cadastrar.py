import tkinter as tk
from tkinter import messagebox
from docx import Document

import pyrebase

import sys

import os

config = {
  "apiKey": "AIzaSyDIAZ_CJbJUh8WRjS9B5GXf_y_OMhbEvhk",
  "authDomain": "testeobserv.firebaseapp.com",
  "databaseURL": "https://testeobserv-default-rtdb.firebaseio.com",
  "projectId": "testeobserv",
  "storageBucket": "testeobserv.appspot.com",
  "messagingSenderId": "77277924546",
  "appId": "1:77277924546:web:60edcbe989fe25ff81bf5b"
}

firebase = pyrebase.initialize_app(config)

def replace_text(doc, tag, value):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            inline = paragraph.runs
            for i in range(len(inline)):
                if tag in inline[i].text:
                    text = inline[i].text.replace(tag, value)
                    inline[i].text = text

def create_document(name, value):
    doc = Document('../template.docx')
    replace_text(doc, '<NOME>', name)
    replace_text(doc, '<VALOR>', value)

    downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    nome_arquivo = 'formulario_preenchido.docx'
    caminho_arquivo = os.path.join(downloads_path, nome_arquivo)


    doc.save(caminho_arquivo)
    messagebox.showinfo("Informação", f"Documento criads com sucesso.")
    
def post_firebase(name, value):
    db = firebase.database()
    data = {"nome": name,
            "valor": value}
    db.child("app").set(data)

def submit():
    name = name_entry.get()
    value = value_entry.get()

    if name and value:
        create_document(name, value)
        post_firebase(name, value)


        # Encerra a aplicação
        sys.exit()
        
    else:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos.")

# Criando a janela
root = tk.Tk()
root.title("Formulário de Nome e Valor")
root.iconbitmap("../icon.ico")

# Criando os widgets
name_label = tk.Label(root, text="Nome:")
name_entry = tk.Entry(root)

value_label = tk.Label(root, text="Valor:")
value_entry = tk.Entry(root)

submit_button = tk.Button(root, text="Enviar", command=submit)

# Posicionando os widgets na janela usando o gerenciador de layout grid
name_label.grid(row=0, column=0, padx=10, pady=5)
name_entry.grid(row=0, column=1, padx=10, pady=5)

value_label.grid(row=1, column=0, padx=10, pady=5)
value_entry.grid(row=1, column=1, padx=10, pady=5)

submit_button.grid(row=2, columnspan=2, padx=10, pady=10)

# Iniciando o loop principal da interface
root.mainloop()