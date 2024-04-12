import customtkinter as ctk
from tkinter import messagebox, filedialog, Tk, Toplevel, Entry, Label, Button
import mysql.connector
import pandas as pd
import os

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

base_file_path = r"C:\Users\pop\OneDrive - CADS Consultoria\Área de Trabalho\sistema de frequencia\DATA\planilha_base.xlsx"

def open_file():
    file_path = filedialog.askopenfilename(filetypes=(("Arquivos Excel", "*.xls;*.xlsx"), ("Todos os arquivos", "*.*")))
    if file_path:
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip().str.lower()  # Remove espaços extras e converte para minúsculas
            print(f"Colunas do DataFrame: {df.columns}")  # Imprime as colunas do DataFrame
            if 'tabela base' in file_path:
                df.to_excel(base_file_path, engine='xlsxwriter', index=False)
                messagebox.showinfo("Sucesso", "Tabela base salva com sucesso!")
            elif 'tabela comparativa' in file_path:
                if 'nome' not in df.columns:
                    messagebox.showerror("Erro", "A coluna 'nome' não existe na tabela comparativa.")
                    return
                base_df = pd.read_excel(base_file_path)
                base_df.columns = base_df.columns.str.strip().str.lower()  # Remove espaços extras e converte para minúsculas
                print(f"Colunas da tabela base: {base_df.columns}")  # Imprime as colunas da tabela base
                base_df['frequência'] = base_df['nome'].apply(lambda x: 'X' if x in df['nome'].values else '')
                base_df.to_excel(base_file_path, engine='xlsxwriter', index=False)
                messagebox.showinfo("Sucesso", "Tabela comparativa processada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo: {e}")





def download_base_file():
    try:
        df = pd.read_excel(base_file_path)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save_path:  # Verifica se o usuário escolheu um local para salvar o arquivo
            df.to_excel(save_path, engine='xlsxwriter', index=False)
            messagebox.showinfo("Download concluído", "A tabela base foi baixada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao baixar a tabela base: {e}")


def download_base_file():
    try:
        df = pd.read_excel(base_file_path)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save_path:  # Verifica se o usuário escolheu um local para salvar o arquivo
            df.to_excel(save_path, engine='xlsxwriter', index=False)
            messagebox.showinfo("Download concluído", "A planilha base foi baixada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao baixar a planilha base: {e}")

def open_new_window():
    new_window = Toplevel()  
    new_window.title("Nova Janela")
    new_window.geometry("500x500")
    new_window.configure(bg='black')  
    label = ctk.CTkLabel(new_window, text="Bem-vindo ao sistema!")
    label.pack()
    upload_button = ctk.CTkButton(new_window, text="Enviar arquivo XLS", command=open_file)
    upload_button.pack()
    download_button = ctk.CTkButton(new_window, text="Baixar arquivo XLS", command=download_base_file)
    download_button.pack()

def check_login():
    mydb = mysql.connector.connect(
        host="localhost", 
        user="root",
        password="181199", 
        database="login"
    )
    mycursor = mydb.cursor()
    # Consulta SQL para verificar o id e a password
    mycursor.execute("SELECT * FROM tbl_usuario WHERE id = %s AND password = %s",
                     (id_entry.get(), password_entry.get()))
    myresult = mycursor.fetchone()
    if myresult:
        messagebox.showinfo("Login info", "Login bem sucedido!")
        root.withdraw()  
        open_new_window()  
    else:
        messagebox.showerror("Login info", "ID ou password incorretos!")

def register_user():
    mydb = mysql.connector.connect(
        host="localhost", 
        user="root",
        password="181199",  
        database="login"  
    )
    mycursor = mydb.cursor()
    # Consulta SQL para inserir o novo usuário
    sql = "INSERT INTO tbl_usuario (id, password) VALUES (%s, %s)"
    val = (id_entry.get(), password_entry.get())
    mycursor.execute(sql, val)
    mydb.commit()
    messagebox.showinfo("Registro", "Usuário registrado com sucesso!")

root = Tk()
root.title("Tela de Login")
root.geometry("500x500")
root.configure(bg='black')
id_label = ctk.CTkLabel(root, text="ID:")
id_label.pack()
id_entry = ctk.CTkEntry(root)
id_entry.pack()
password_label = ctk.CTkLabel(root, text="Senha:")
password_label.pack()
password_entry = ctk.CTkEntry(root, show="*")
password_entry.pack()
login_button = ctk.CTkButton(root, text="Login", command=check_login)
login_button.pack(padx=10, pady=10)
register_button = ctk.CTkButton(root, text="Cadastrar", command=register_user)
register_button.pack(padx=10, pady=10)
root.mainloop()
