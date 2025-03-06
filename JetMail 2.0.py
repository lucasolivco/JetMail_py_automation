import os
import pandas as pd
import yagmail
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import threading
import queue
import keyring
import time

# Configuração do tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Fila para envio de mensagens de log
log_queue = queue.Queue()

# Variáveis globais para controle da barra de progresso
sending_in_progress = False
progress_value = 0.0
total_emails = 0
current_email = 0

def select_excel_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*"))
    )
    if file_path:
        excel_entry.delete(0, ctk.END)
        excel_entry.insert(0, file_path)
        try:
            # Tenta carregar o arquivo para verificar se é válido
            df = pd.read_excel(file_path)
            log_queue.put(f"Arquivo Excel carregado com {len(df)} registros.")
        except Exception as e:
            log_queue.put(f"Erro ao validar arquivo Excel: {e}")

def select_pdf_folder():
    folder_path = filedialog.askdirectory(title="Selecione a pasta dos PDFs")
    if folder_path:
        pdf_entry.delete(0, ctk.END)
        pdf_entry.insert(0, folder_path)
        # Conta quantos PDFs existem na pasta
        pdf_count = len([f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')])
        log_queue.put(f"Pasta selecionada: {pdf_count} arquivos PDF encontrados.")

def update_log_text():
    """Atualiza a caixa de log com mensagens vindas da fila."""
    while not log_queue.empty():
        log_message = log_queue.get()
        current_time = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{current_time}] {log_message}"
        log_text.insert(ctk.END, formatted_message + "\n")
        log_text.see(ctk.END)
    app.after(100, update_log_text)

def update_progress():
    """Atualiza a barra de progresso com base no progresso real."""
    global sending_in_progress, current_email, total_emails
    if sending_in_progress:
        if total_emails > 0:
            progress_bar.set(current_email / total_emails)
            progress_text.configure(text=f"Progresso: {current_email}/{total_emails}")
        app.after(100, update_progress)

def finish_ui_updates():
    """Finaliza a animação e reabilita o botão de envio."""
    global sending_in_progress
    sending_in_progress = False
    progress_bar.set(1.0)
    progress_text.configure(text="Concluído!")
    send_button.configure(state="normal")
    messagebox.showinfo("Concluído", "Processo de envio concluído!")

def save_credentials():
    """Salva as credenciais no sistema (opcional)."""
    if remember_var.get():
        email = sender_entry.get()
        password = password_entry.get()
        if email and password:
            try:
                keyring.set_password("email_pdf_sender", email, password)
                log_queue.put("Credenciais salvas com segurança.")
            except Exception as e:
                log_queue.put(f"Erro ao salvar credenciais: {e}")
    else:
        # Se a opção for desmarcada, tenta remover as credenciais salvas
        email = sender_entry.get()
        if email:
            try:
                keyring.delete_password("email_pdf_sender", email)
                log_queue.put("Credenciais removidas.")
            except:
                pass

def load_credentials():
    """Carrega credenciais salvas, se existirem."""
    email = sender_entry.get()
    if email:
        try:
            password = keyring.get_password("email_pdf_sender", email)
            if password:
                password_entry.delete(0, ctk.END)
                password_entry.insert(0, password)
                remember_var.set(True)
        except:
            pass

def send_emails_thread():
    """Função que roda em thread e realiza o envio dos e-mails."""
    global total_emails, current_email
    
    excel_path = excel_entry.get()
    pdf_folder = pdf_entry.get()
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    delay_seconds = float(delay_entry.get()) if delay_entry.get() else 0

    # Validação dos dados informados
    if not excel_path or not os.path.exists(excel_path):
        log_queue.put("Erro: Selecione um arquivo Excel válido.")
        app.after(0, finish_ui_updates)
        return
    if not pdf_folder or not os.path.exists(pdf_folder):
        log_queue.put("Erro: Selecione uma pasta válida para os PDFs.")
        app.after(0, finish_ui_updates)
        return
    if not sender_email or not sender_password:
        log_queue.put("Erro: Insira o email e a senha do remetente.")
        app.after(0, finish_ui_updates)
        return

    try:
        log_queue.put(f"Conectando ao servidor de e-mail...")
        yag = yagmail.SMTP(sender_email, sender_password)
        log_queue.put(f"Conexão estabelecida com sucesso.")
    except Exception as e:
        log_queue.put(f"Erro ao conectar com o servidor: {e}")
        app.after(0, finish_ui_updates)
        return

    try:
        log_queue.put(f"Carregando dados do Excel...")
        df = pd.read_excel(excel_path)
        total_emails = len(df)
        log_queue.put(f"Dados carregados: {total_emails} e-mails para enviar.")
    except Exception as e:
        log_queue.put(f"Erro ao ler o arquivo Excel: {e}")
        app.after(0, finish_ui_updates)
        return

    # Tenta salvar as credenciais
    save_credentials()

    # Abre (ou cria) o arquivo de log para registrar as operações
    log_filename = f"log_envio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(log_filename, "a", encoding="utf-8") as log_file:
        log_file.write("\n---------- Início do envio: " +
                       datetime.now().strftime("%Y-%m-%d %H:%M:%S") +
                       " ----------\n")

        current_email = 0
        for index, row in df.iterrows():
            # Verifica se o usuário cancelou o processo
            if not sending_in_progress:
                log_queue.put("Processo interrompido pelo usuário.")
                break

            nome = str(row.get('RESPONSAVEL', 'Cliente'))
            destinatario = str(row.get('Contato', ''))
            
            # Verifica se há dados válidos
            if not destinatario or '@' not in destinatario:
                log_message = f"Linha {index+2}: E-mail inválido ou não fornecido: '{destinatario}'"
                log_queue.put(log_message)
                log_file.write(log_message + "\n")
                continue
                
            # Determina o nome do arquivo PDF
            pdf_filename = f"{nome}.pdf"
            pdf_path = os.path.join(pdf_folder, pdf_filename)

            if not os.path.exists(pdf_path):
                log_message = (f"Linha {index+2}: {destinatario} - Arquivo {pdf_filename} não encontrado")
                log_queue.put(log_message)
                log_file.write(log_message + "\n")
                continue

            try:
                corpo = template_entry.get("1.0", ctk.END).strip()
                if not corpo:
                    corpo = f"Olá {nome}, segue em anexo o seu arquivo PDF."
                else:
                    corpo = corpo.replace("{nome}", nome)
                
                assunto = subject_entry.get()
                if not assunto:
                    assunto = "Seu arquivo PDF"
                
                log_queue.put(f"Enviando para: {destinatario}...")
                yag.send(to=destinatario, subject=assunto, contents=[corpo, pdf_path])
                
                log_message = f"Linha {index+2}: {destinatario} - E-mail enviado com sucesso."
                log_queue.put(log_message)
                log_file.write(log_message + "\n")
                
                # Adiciona o delay configurado entre envios
                if delay_seconds > 0 and index < len(df) - 1:
                    log_queue.put(f"Aguardando {delay_seconds} segundos antes do próximo envio...")
                    time.sleep(delay_seconds)
                    
            except Exception as e:
                log_message = f"Linha {index+2}: {destinatario} - Erro ao enviar: {e}"
                log_queue.put(log_message)
                log_file.write(log_message + "\n")
            
            current_email += 1

        log_file.write("---------- Fim do envio: " + 
                       datetime.now().strftime("%Y-%m-%d %H:%M:%S") +
                       " ----------\n")
    
    # Resuma as operações no final
    log_queue.put(f"Processo de envio concluído. Log salvo em: {log_filename}")
    app.after(0, finish_ui_updates)

def start_sending():
    """Inicia o processo de envio em uma thread separada e anima a barra de progresso."""
    global sending_in_progress, progress_value, total_emails, current_email
    
    # Confirma com o usuário
    if not messagebox.askyesno("Confirmar", "Iniciar o envio de e-mails?"):
        return
        
    send_button.configure(state="disabled")
    stop_button.configure(state="normal")
    sending_in_progress = True
    progress_value = 0.0
    total_emails = 0
    current_email = 0
    update_progress()
    threading.Thread(target=send_emails_thread, daemon=True).start()

def stop_sending():
    """Para o processo de envio em andamento."""
    global sending_in_progress
    if messagebox.askyesno("Confirmar", "Deseja interromper o envio de e-mails?"):
        sending_in_progress = False
        stop_button.configure(state="disabled")
        log_queue.put("Finalizando processo... Aguarde a conclusão do e-mail atual.")

def on_email_changed(event):
    """Carrega automaticamente a senha salva quando o e-mail for alterado."""
    app.after(100, load_credentials)

def show_about():
    """Mostra informações sobre o aplicativo."""
    messagebox.showinfo(
        "Sobre o Aplicativo", 
        "Automação de Envio de PDFs por Email\n\n"
        "Versão 1.1\n"
        "© 2025 - Todos os direitos reservados\n\n"
        "Este aplicativo automatiza o envio de PDFs por e-mail com base em uma planilha Excel."
    )

def show_help():
    """Mostra instruções de uso."""
    messagebox.showinfo(
        "Ajuda", 
        "Como usar o aplicativo:\n\n"
        "1. Selecione um arquivo Excel contendo colunas 'RESPONSAVEL' e 'Contato'\n"
        "2. Selecione a pasta que contém os arquivos PDF (nome igual ao RESPONSAVEL)\n"
        "3. Insira seu e-mail e senha\n"
        "4. Personalize o assunto e o corpo do e-mail (opcional)\n"
        "5. Defina um intervalo entre os envios (opcional)\n"
        "6. Clique em 'Enviar E-mails'\n\n"
        "O progresso será exibido na barra inferior e os detalhes no log."
    )

def toggle_password():
    """Alterna a visibilidade da senha."""
    current = password_entry.cget("show")
    if current == "*":
        password_entry.configure(show="")
        show_password_btn.configure(text="Ocultar")
    else:
        password_entry.configure(show="*")
        show_password_btn.configure(text="Mostrar")

# Cria a interface com customtkinter
app = ctk.CTk()
app.geometry("700x800")
app.title("Automação de Envio de PDFs por Email")

# Menu principal
menu_frame = ctk.CTkFrame(app, height=40)
menu_frame.pack(fill="x", padx=0, pady=0)

help_button = ctk.CTkButton(menu_frame, text="Ajuda", width=80, command=show_help)
help_button.pack(side="left", padx=10, pady=5)

about_button = ctk.CTkButton(menu_frame, text="Sobre", width=80, command=show_about)
about_button.pack(side="right", padx=10, pady=5)

frame = ctk.CTkScrollableFrame(app)
frame.pack(pady=5, padx=20, fill="both", expand=True)

# Seleção do arquivo Excel
excel_label = ctk.CTkLabel(frame, text="Arquivo Excel (com colunas 'RESPONSAVEL' e 'Contato'):")
excel_label.pack(pady=(10, 0), anchor="w")
excel_entry = ctk.CTkEntry(frame, width=600)
excel_entry.pack(pady=(5, 5), fill="x")
excel_button = ctk.CTkButton(frame, text="Selecionar Excel", command=select_excel_file)
excel_button.pack(pady=(0, 10))

# Seleção da pasta dos PDFs
pdf_label = ctk.CTkLabel(frame, text="Pasta dos PDFs (nomeados como 'RESPONSAVEL.pdf'):")
pdf_label.pack(pady=(10, 0), anchor="w")
pdf_entry = ctk.CTkEntry(frame, width=600)
pdf_entry.pack(pady=(5, 5), fill="x")
pdf_button = ctk.CTkButton(frame, text="Selecionar Pasta", command=select_pdf_folder)
pdf_button.pack(pady=(0, 10))

# Dados do remetente
sender_frame = ctk.CTkFrame(frame)
sender_frame.pack(pady=10, fill="x")

sender_label = ctk.CTkLabel(sender_frame, text="Email do Remetente:")
sender_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)
sender_entry = ctk.CTkEntry(sender_frame, width=400)
sender_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
sender_entry.bind("<FocusOut>", on_email_changed)

password_label = ctk.CTkLabel(sender_frame, text="Senha:")
password_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)
password_entry = ctk.CTkEntry(sender_frame, width=400, show="*")
password_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
show_password_btn = ctk.CTkButton(sender_frame, text="Mostrar", width=80, command=toggle_password)
show_password_btn.grid(row=1, column=2, padx=5, pady=5)

# Opção para lembrar credenciais
remember_var = ctk.IntVar(value=0)
remember_check = ctk.CTkCheckBox(sender_frame, text="Lembrar credenciais", variable=remember_var)
remember_check.grid(row=2, column=1, padx=10, pady=5, sticky="w")

sender_frame.columnconfigure(1, weight=1)

# Configurações de e-mail
email_config_frame = ctk.CTkFrame(frame)
email_config_frame.pack(pady=10, fill="x")

subject_label = ctk.CTkLabel(email_config_frame, text="Assunto do e-mail:")
subject_label.pack(pady=(10, 0), padx=10, anchor="w")
subject_entry = ctk.CTkEntry(email_config_frame, width=600)
subject_entry.pack(pady=(5, 10), padx=10, fill="x")
subject_entry.insert(0, "Seu arquivo PDF")

template_label = ctk.CTkLabel(email_config_frame, text="Modelo de mensagem: (use {nome} para inserir o nome do destinatário)")
template_label.pack(pady=(10, 0), padx=10, anchor="w")
template_entry = ctk.CTkTextbox(email_config_frame, width=600, height=100)
template_entry.pack(pady=(5, 10), padx=10, fill="x")
template_entry.insert("1.0", "Olá {nome},\n\nSegue em anexo o seu arquivo PDF conforme solicitado.\n\nAtenciosamente,\nEquipe de Suporte")

# Configuração de intervalo
delay_label = ctk.CTkLabel(email_config_frame, text="Intervalo entre envios (segundos):")
delay_label.pack(pady=(10, 0), padx=10, anchor="w")
delay_entry = ctk.CTkEntry(email_config_frame, width=100)
delay_entry.pack(pady=(5, 10), padx=10, anchor="w")
delay_entry.insert(0, "2")

# Log de envio
log_label = ctk.CTkLabel(frame, text="Log de Envio:")
log_label.pack(pady=(10, 0), anchor="w")
log_text = ctk.CTkTextbox(frame, width=650, height=200)
log_text.pack(pady=(5, 10), fill="both", expand=True)

# Frame para barra de progresso e textos
progress_frame = ctk.CTkFrame(app, height=80)
progress_frame.pack(fill="x", padx=20, pady=10)

# Barra de progresso
progress_bar = ctk.CTkProgressBar(progress_frame, width=500)
progress_bar.set(0)
progress_bar.pack(pady=(10, 5))

# Texto de progresso
progress_text = ctk.CTkLabel(progress_frame, text="Pronto para iniciar")
progress_text.pack(pady=(0, 5))

# Frame para botões de ação
button_frame = ctk.CTkFrame(app)
button_frame.pack(fill="x", padx=20, pady=10)

# Botões para controlar o envio
send_button = ctk.CTkButton(button_frame, text="Enviar E-mails", command=start_sending, width=200)
send_button.pack(side="left", padx=20, pady=10)

stop_button = ctk.CTkButton(button_frame, text="Interromper", command=stop_sending, width=200, state="disabled")
stop_button.pack(side="right", padx=20, pady=10)

# Inicia a atualização contínua da caixa de log
app.after(100, update_log_text)

# Mensagem inicial
log_queue.put("Aplicativo iniciado. Preencha os campos e clique em 'Enviar E-mails'.")

app.mainloop()