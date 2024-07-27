from playwright.sync_api import sync_playwright
import datetime
import tkinter as tk
from tkinter import messagebox
from ttkbootstrap import Style, ttk
import json
import os

class BatidaDePontoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ADPoint - Automação de Ponto ADP")
        self.root.resizable(False, False)  # Bloquear redimensionamento da janela
        
        self.style = Style(theme='darkly')  # Usando o tema 'darkly' que é semelhante ao estilo da Netflix
        self.style.configure('TLabel', font=('Helvetica', 12))
        self.style.configure('TEntry', font=('Helvetica', 12))
        self.style.configure('TButton', font=('Helvetica', 12))
        self.style.configure('TCheckbutton', font=('Helvetica', 12))
        
        self.create_widgets()
        
        self.schedule = []
        self.email = ""
        self.senha = ""
        self.credentials_file = "credentials.json"
        self.automation_running = False
        
        self.load_credentials()
        
    def create_widgets(self):
        self.label_email = ttk.Label(self.root, text="Email:")
        self.label_email.grid(row=0, column=0, padx=10, pady=10)
        self.entry_email = ttk.Entry(self.root)
        self.entry_email.grid(row=0, column=1, padx=10, pady=10)

        self.label_senha = ttk.Label(self.root, text="Senha:")
        self.label_senha.grid(row=1, column=0, padx=10, pady=10)
        self.entry_senha = ttk.Entry(self.root, show="*")
        self.entry_senha.grid(row=1, column=1, padx=10, pady=10)

        self.save_credentials_var = tk.IntVar()
        self.checkbox_save_credentials = ttk.Checkbutton(self.root, text="Salvar credenciais", variable=self.save_credentials_var)
        self.checkbox_save_credentials.grid(row=2, columnspan=2, padx=10, pady=10)

        self.label_entrada = ttk.Label(self.root, text="Horário de Entrada (HH:MM):")
        self.label_entrada.grid(row=3, column=0, padx=10, pady=10)
        self.entry_entrada = ttk.Entry(self.root)
        self.entry_entrada.grid(row=3, column=1, padx=10, pady=10)
        self.entry_entrada.bind("<KeyRelease>", self.mask_time)
        
        self.label_saida_almoco = ttk.Label(self.root, text="Saída para Almoço (HH:MM):")
        self.label_saida_almoco.grid(row=4, column=0, padx=10, pady=10)
        self.entry_saida_almoco = ttk.Entry(self.root)
        self.entry_saida_almoco.grid(row=4, column=1, padx=10, pady=10)
        self.entry_saida_almoco.bind("<KeyRelease>", self.mask_time)
        
        self.label_volta_almoco = ttk.Label(self.root, text="Volta do Almoço (HH:MM):")
        self.label_volta_almoco.grid(row=5, column=0, padx=10, pady=10)
        self.entry_volta_almoco = ttk.Entry(self.root)
        self.entry_volta_almoco.grid(row=5, column=1, padx=10, pady=10)
        self.entry_volta_almoco.bind("<KeyRelease>", self.mask_time)
        
        self.label_saida_trabalho = ttk.Label(self.root, text="Saída do Trabalho (HH:MM):")
        self.label_saida_trabalho.grid(row=6, column=0, padx=10, pady=10)
        self.entry_saida_trabalho = ttk.Entry(self.root)
        self.entry_saida_trabalho.grid(row=6, column=1, padx=10, pady=10)
        self.entry_saida_trabalho.bind("<KeyRelease>", self.mask_time)
        
        # Configurando o botão para verde inicialmente
        self.style.configure('Green.TButton', background='green', foreground='white')

        self.button_iniciar = ttk.Button(self.root, text="Iniciar", command=self.toggle_automation, style='Green.TButton')
        self.button_iniciar.grid(row=7, columnspan=2, padx=10, pady=10)
        self.button_iniciar.bind("<Enter>", self.on_mouse_enter)
        self.button_iniciar.bind("<Leave>", self.on_mouse_leave)

        # TextBox para contagem regressiva
        self.textbox_contagem = tk.Text(self.root, height=2, width=50)
        self.textbox_contagem.grid(row=8, columnspan=2, padx=10, pady=10)
        self.textbox_contagem.config(state=tk.DISABLED)

        # TextBox para status de batida de ponto
        self.textbox_status = tk.Text(self.root, height=2, width=50)
        self.textbox_status.grid(row=9, columnspan=2, padx=10, pady=10)
        self.textbox_status.config(state=tk.DISABLED)

    def mask_time(self, event):
        entry = event.widget
        value = entry.get()
        
        if len(value) == 2 and value[-1] != ':':
            entry.insert(tk.END, ":")
        elif len(value) > 5:
            entry.delete(5, tk.END)

    def validate_time(self, value):
        try:
            if len(value) == 5:
                hours, minutes = value.split(':')
                if not (0 <= int(hours) < 24 and 0 <= int(minutes) < 60):
                    return False
                datetime.datetime.strptime(value, '%H:%M')
                return True
            return False
        except ValueError:
            return False
    
    def toggle_automation(self):
        if self.automation_running:
            self.stop_automation()
        else:
            self.start_automation()
    
    def start_automation(self):
        self.email = self.entry_email.get()
        self.senha = self.entry_senha.get()
        
        if not self.email or not self.senha:
            messagebox.showerror("Erro", "Email e senha são obrigatórios.")
            return

        if self.save_credentials_var.get():
            self.save_credentials()
        else:
            self.delete_credentials()

        horarios = [
            self.entry_entrada.get(),
            self.entry_saida_almoco.get(),
            self.entry_volta_almoco.get(),
            self.entry_saida_trabalho.get()
        ]
        
        self.schedule.clear()
        for horario in horarios:
            if horario:
                if not self.validate_time(horario):
                    self.update_status("Erro", "Formato de hora inválido: {}".format(horario))
                    return
                hora, minuto = map(int, horario.split(':'))
                self.schedule.append(datetime.time(hora, minuto))
        
        self.automation_running = True
        self.style.configure('Running.TButton', background='#3b5998', foreground='white')
        self.button_iniciar.config(text="Executando", style='Running.TButton')
        self.clear_status()
        self.verificar_batida()

    def stop_automation(self):
        self.automation_running = False
        self.style.configure('Green.TButton', background='green', foreground='white')
        self.button_iniciar.config(text="Iniciar", style='Green.TButton')
        self.clear_status()
        self.update_status("Informação", "Automação parada.")

    def verificar_batida(self):
        if not self.automation_running:
            return

        agora = datetime.datetime.now().time()
        if self.schedule:
            proximo_horario = min(self.schedule)
            delta = datetime.datetime.combine(datetime.date.today(), proximo_horario) - datetime.datetime.combine(datetime.date.today(), agora)
            minutos, segundos = divmod(delta.seconds, 60)
            self.update_contagem(f"Tempo até a próxima batida: {minutos} minutos e {segundos} segundos")

            if agora.hour == proximo_horario.hour and agora.minute == proximo_horario.minute:
                self.bater_ponto(proximo_horario)
                self.schedule.remove(proximo_horario)

        self.root.after(1000, self.verificar_batida)
    
    def bater_ponto(self, horario):
        try:
            with sync_playwright() as p:
                navegador = p.chromium.launch(headless=True)
                pagina = navegador.new_page()
                pagina.goto("https://login.la.logicalis.com/adfs/ls/IdpInitiatedSignOn.aspx?loginToRp=https://sso.ehc.adp.com/samlsp/expert/")
                pagina.fill('//*[@id="userNameInput"]', self.email)
                pagina.fill('//*[@id="passwordInput"]', self.senha)
                pagina.locator('//*[@id="submitButton"]').click()
                pagina.locator('//*[@id="clock"]/button').click()
                self.update_status("Sucesso", f"Horário batido com sucesso em: {horario}")
        except Exception as e:
            self.update_status("Erro", f"Falha ao bater o ponto: {e}")
    
    def save_credentials(self):
        credentials = {
            "email": self.email,
            "senha": self.senha
        }
        with open(self.credentials_file, 'w') as file:
            json.dump(credentials, file)
    
    def load_credentials(self):
        if os.path.exists(self.credentials_file):
            with open(self.credentials_file, 'r') as file:
                credentials = json.load(file)
                self.entry_email.insert(0, credentials.get("email", ""))
                self.entry_senha.insert(0, credentials.get("senha", ""))
    
    def delete_credentials(self):
        if os.path.exists(self.credentials_file):
            os.remove(self.credentials_file)

    def update_contagem(self, message):
        self.textbox_contagem.config(state=tk.NORMAL)
        self.textbox_contagem.delete(1.0, tk.END)
        self.textbox_contagem.insert(tk.END, message)
        self.textbox_contagem.config(state=tk.DISABLED)

    def update_status(self, status, message):
        self.textbox_status.config(state=tk.NORMAL)
        self.textbox_status.delete(1.0, tk.END)
        self.textbox_status.insert(tk.END, f"{status}: {message}")
        self.textbox_status.config(state=tk.DISABLED)

    def clear_status(self):
        self.textbox_status.config(state=tk.NORMAL)
        self.textbox_status.delete(1.0, tk.END)
        self.textbox_status.config(state=tk.DISABLED)

    def on_mouse_enter(self, event):
        if self.automation_running:
            self.style.configure('Stop.TButton', background='#E50914', foreground='white')
            self.button_iniciar.config(style='Stop.TButton')

    def on_mouse_leave(self, event):
        if self.automation_running:
            self.style.configure('Running.TButton', background='#3b5998', foreground='white')
            self.button_iniciar.config(style='Running.TButton')

if __name__ == "__main__":
    root = tk.Tk()
    app = BatidaDePontoApp(root)
    root.mainloop()
