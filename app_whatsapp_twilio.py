import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import threading
import time
import re
from datetime import datetime
from twilio.rest import Client

# Configurações Twilio - PREENCHA COM SUAS CREDENCIAIS
TWILIO_ACCOUNT_SID = "SEU_ACCOUNT_SID"
TWILIO_AUTH_TOKEN = "SEU_AUTH_TOKEN"
TWILIO_WHATSAPP_NUMBER = "whatsapp:+1415XXXXXXX"

class WhatsAppSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Envio WhatsApp - Twilio")
        self.root.geometry("800x600")

        self.contatos = None
        self.excel_path = None

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Selecione a planilha (.xlsx):", font=("Arial", 12)).pack(pady=10)
        tk.Button(self.root, text="Carregar Planilha", command=self.load_excel).pack(pady=5)

        self.text_area = scrolledtext.ScrolledText(self.root, width=90, height=20)
        self.text_area.pack(pady=10)

        tk.Button(self.root, text="Enviar Mensagens", command=self.enviar_mensagens_thread).pack(pady=10)
        tk.Button(self.root, text="Salvar Planilha Atualizada", command=self.salvar_planilha).pack(pady=5)

        self.progress_label = tk.Label(self.root, text="", fg="green")
        self.progress_label.pack(pady=5)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            self.contatos = pd.read_excel(file_path)
            if "Status" not in self.contatos.columns:
                self.contatos["Status"] = ""
            self.excel_path = file_path

            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, "Contatos carregados:\n")
            self.text_area.insert(tk.END, self.contatos.to_string(index=False))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha: {e}")

    def enviar_mensagens_thread(self):
        t = threading.Thread(target=self.enviar_mensagens)
        t.start()

    def formatar_telefone(self, telefone):
        telefone = re.sub(r'\D', '', str(telefone))
        if not telefone.startswith("55"):
            telefone = "55" + telefone
        return f"whatsapp:+{telefone}"

    def montar_mensagem(self, nome, link):
        if pd.isna(link) or not str(link).strip():
            return f"Olá {nome}, seja muito bem-vindo(a)! 😄 Estamos felizes em ter você conosco."
        return f"Olá {nome}, seja muito bem-vindo(a)! 😄 Confira esta oferta especial:\n\n{link}"

    def salvar_log(self, msg):
        log_file = "logs/log_envios.txt"
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")

    def enviar_mensagens(self):
        if self.contatos is None:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return

        try:
            client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao conectar à API Twilio:\n{e}")
            return

        total = len(self.contatos)
        enviados = 0

        for index, linha in self.contatos.iterrows():
            nome = str(linha.get("nome", "")).strip()
            telefone = str(linha.get("telefone", "")).strip()
            status = str(linha.get("Status", ""))

            if status == "Enviado":
                enviados += 1
                continue

            telefone_formatado = self.formatar_telefone(telefone)

            if len(telefone_formatado) < 15:
                self.contatos.at[index, "Status"] = "Erro: número inválido"
                self.text_area.insert(tk.END, f"\n⚠️ Número inválido para {nome}: {telefone}")
                self.salvar_log(f"Número inválido para {nome}: {telefone}")
                self.text_area.see(tk.END)
                continue

            mensagem = self.montar_mensagem(nome, linha.get("mensagem", ""))

            try:
                message = client.messages.create(
                    from_=TWILIO_WHATSAPP_NUMBER,
                    body=mensagem,
                    to=telefone_formatado
                )
                self.contatos.at[index, "Status"] = "Enviado"
                self.text_area.insert(tk.END, f"\n✅ Enviado para {nome} ({telefone})")
                self.salvar_log(f"Enviado para {nome} ({telefone})")
                enviados += 1
            except Exception as e:
                self.contatos.at[index, "Status"] = f"Erro: {e}"
                self.text_area.insert(tk.END, f"\n❌ Erro ao enviar para {nome} ({telefone}): {e}")
                self.salvar_log(f"Erro ao enviar para {nome} ({telefone}): {e}")

            self.text_area.see(tk.END)
            self.progress_label.config(text=f"Enviados: {enviados}/{total}")
            time.sleep(5)  # Intervalo entre envios

        messagebox.showinfo("Concluído", "Todos os envios foram concluídos!")

    def salvar_planilha(self):
        if self.contatos is not None:
            nova_planilha = os.path.splitext(self.excel_path)[0] + "_com_status.xlsx"
            self.contatos.to_excel(nova_planilha, index=False)
            messagebox.showinfo("Salvo", f"Planilha salva como:\n{nova_planilha}")
        else:
            messagebox.showwarning("Aviso", "Nenhuma planilha carregada.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSenderApp(root)
    root.mainloop()
