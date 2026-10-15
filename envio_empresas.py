import csv
import re
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

ROOT = Path(__file__).resolve().parent
FIELDS = ["empresa", "email", "arquivos", "assunto", "mensagem", "liberacao"]
LOG = ROOT / "historico_envios.csv"

def rows_from_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        if not all(k in (reader.fieldnames or []) for k in FIELDS):
            raise ValueError("CSV precisa ter as colunas: " + ";".join(FIELDS))
        return [{k: (r.get(k) or "").strip() for k in FIELDS} for r in reader]

def recipients(s):
    return [x.strip() for x in s.split(";") if x.strip()]

def attachments(s, folder):
    result = []
    for name in s.split("|"):
        if name.strip():
            file = Path(name.strip().strip('"'))
            result.append(file if file.is_absolute() else folder / file)
    return result

def already_sent(r):
    if not LOG.exists():
        return False
    with LOG.open(newline="", encoding="utf-8-sig") as f:
        for old in csv.DictReader(f, delimiter=";"):
            if all(old.get(k) == r[k] for k in ("empresa", "email", "arquivos")) and old.get("status") == "ENVIADO":
                return True
    return False

def record(r, status, detail):
    new = not LOG.exists()
    with LOG.open("a", newline="", encoding="utf-8-sig") as f:
        fields = ["data_hora", "empresa", "email", "arquivos", "status", "detalhe"]
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        if new:
            w.writeheader()
        w.writerow({"data_hora": datetime.now().isoformat(timespec="seconds"), "empresa": r["empresa"],
                    "email": r["email"], "arquivos": r["arquivos"], "status": status, "detalhe": detail})

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Envio individualizado - Outlook classico")
        self.geometry("1080x740")
        self.csv_path = tk.StringVar(value=str(ROOT / "destinatarios.csv"))
        self.folder = tk.StringVar(value=str(ROOT / "documentos"))
        self.account = tk.StringVar()
        self.mode = tk.StringVar(value="RASCUNHOS (teste)")
        self.force = tk.BooleanVar(value=False)
        self.items = []
        self.busy = False
        self.build()

    def path_input(self, parent, title, variable, directory=False):
        line = ttk.Frame(parent)
        line.pack(fill="x", pady=4)
        ttk.Label(line, text=title, width=23).pack(side="left")
        ttk.Entry(line, textvariable=variable).pack(side="left", fill="x", expand=True)
        choose = (lambda: filedialog.askdirectory()) if directory else (lambda: filedialog.askopenfilename(filetypes=[("CSV", "*.csv")]))
        def browse():
            selected = choose()
            if selected:
                variable.set(selected)
        ttk.Button(line, text="Procurar", command=browse).pack(side="left", padx=6)

    def build(self):
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Destinatarios e anexos", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        self.path_input(frame, "Planilha CSV", self.csv_path)
        self.path_input(frame, "Pasta de documentos", self.folder, True)
        settings = ttk.LabelFrame(frame, text="Outlook classico", padding=10)
        settings.pack(fill="x", pady=10)
        line = ttk.Frame(settings)
        line.pack(fill="x")
        ttk.Label(line, text="Remetente (email exato)").pack(side="left")
        ttk.Entry(line, textvariable=self.account, width=42).pack(side="left", padx=7)
        ttk.Button(line, text="Listar contas", command=self.list_accounts).pack(side="left")
        line = ttk.Frame(frame)
        line.pack(fill="x", pady=8)
        ttk.Button(line, text="CONFERIR ANEXOS", command=self.preview).pack(side="left")
        ttk.Combobox(line, textvariable=self.mode, state="readonly", width=22,
                     values=["RASCUNHOS (teste)", "ENVIAR DE VERDADE"]).pack(side="left", padx=7)
        self.run_button = ttk.Button(line, text="PROCESSAR EMPRESAS", command=self.process_confirm)
        self.run_button.pack(side="left")
        ttk.Checkbutton(line, text="Autorizar reenvios", variable=self.force).pack(side="left", padx=8)
        cols = ("empresa", "email", "arquivos", "status")
        self.table = ttk.Treeview(frame, columns=cols, show="headings", height=12)
        for col, width in zip(cols, (220, 260, 430, 180)):
            self.table.heading(col, text=col.title())
            self.table.column(col, width=width)
        self.table.pack(fill="both", expand=True)
        self.logs = scrolledtext.ScrolledText(frame, height=8, state="disabled")
        self.logs.pack(fill="x", pady=8)
        self.note("Modo inicial: rascunhos. Confira tudo antes do envio.")

    def note(self, msg):
        self.logs.configure(state="normal")
        self.logs.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.logs.see("end")
        self.logs.configure(state="disabled")

    def list_accounts(self):
        try:
            import win32com.client
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            names = [str(a.SmtpAddress) for a in outlook.Session.Accounts]
            messagebox.showinfo("Contas Outlook", "\n".join(names) or "Nenhuma conta encontrada")
        except Exception as exc:
            messagebox.showerror("Outlook", str(exc) + "\nAbra o Outlook classico.")

    def preview(self):
        self.items = []
        self.table.delete(*self.table.get_children())
        try:
            rows = rows_from_csv(Path(self.csv_path.get()))
            folder = Path(self.folder.get())
            seen = set()
            for r in rows:
                files = attachments(r["arquivos"], folder)
                emails = recipients(r["email"])
                issues = []
                if r["liberacao"].upper() != "SIM":
                    issues.append("Nao liberado")
                if not r["empresa"]:
                    issues.append("Sem empresa")
                if not emails or any(not re.fullmatch(r"[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+", e) for e in emails):
                    issues.append("Email invalido")
                if len({e.lower() for e in emails}) != len(emails):
                    issues.append("Email duplicado")
                if not files or any(not p.is_file() for p in files):
                    issues.append("Anexo ausente")
                signature = (r["empresa"], r["email"], r["arquivos"])
                if signature in seen:
                    issues.append("Linha repetida")
                seen.add(signature)
                if not self.force.get() and already_sent(r):
                    issues.append("Ja enviado")
                status = ", ".join(issues) if issues else "PRONTO"
                self.items.append((r, files, status))
                self.table.insert("", "end", values=(r["empresa"], r["email"], " | ".join(p.name for p in files), status))
            count = sum(s == "PRONTO" for _, _, s in self.items)
            self.note(f"Conferidas {len(rows)} linhas; {count} prontas.")
            return count
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            return 0

    def process_confirm(self):
        if self.busy:
            return
        count = self.preview()
        address = self.account.get().strip()
        if not count:
            return
        if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", address):
            messagebox.showerror("Remetente", "Informe a conta exata do Outlook.")
            return
        send = self.mode.get() == "ENVIAR DE VERDADE"
        action = "ENVIAR DE VERDADE" if send else "CRIAR RASCUNHOS para"
        if not messagebox.askyesno("Confirmar", f"{action} {count} empresa(s)?\nRemetente: {address}\nUma mensagem por linha."):
            return
        tasks = [(dict(r), list(files)) for r, files, s in self.items if s == "PRONTO"]
        self.busy = True
        self.run_button.configure(state="disabled")
        threading.Thread(target=self.run_messages, args=(tasks, address, send), daemon=True).start()

    def run_messages(self, tasks, address, send):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            accounts = [a for a in outlook.Session.Accounts if str(a.SmtpAddress).lower() == address.lower()]
            if len(accounts) != 1:
                raise RuntimeError("Conta remetente nao localizada ou ambigua: " + address)
            account = accounts[0]
            for row, files in tasks:
                try:
                    if not self.force.get() and already_sent(row):
                        continue
                    if any(not p.is_file() for p in files):
                        raise FileNotFoundError("Anexo nao encontrado")
                    mail = outlook.CreateItem(0)
                    # Outlook classico: definicao COM que preserva SendUsingAccount.
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    mail.Save()
                    chosen = mail.SendUsingAccount
                    if chosen is None or str(chosen.SmtpAddress).lower() != address.lower():
                        raise RuntimeError("Outlook nao confirmou o remetente; envio bloqueado")
                    mail.To = row["email"]
                    mail.Subject = (row["assunto"] or "Documentos para {empresa}").replace("{empresa}", row["empresa"])
                    mail.Body = (row["mensagem"] or "Prezados,\n\nSegue documento para {empresa}.\n\nAtenciosamente,").replace("{empresa}", row["empresa"]).replace("\\n", "\n")
                    for file in files:
                        mail.Attachments.Add(str(file.resolve()))
                    if send:
                        mail.Send()
                        status = "ENVIADO"
                        detail = "Aceito para envio pelo Outlook; nao comprova entrega"
                    else:
                        drafts = account.DeliveryStore.GetDefaultFolder(16)
                        moved = mail.Move(drafts)
                        if moved is None or moved.SendUsingAccount is None or str(moved.SendUsingAccount.SmtpAddress).lower() != address.lower():
                            raise RuntimeError("Rascunho criado mas remetente nao confirmado; conferir manualmente")
                        status = "RASCUNHO"
                        detail = "Salvo no Outlook; nao enviado"
                    record(row, status, detail)
                    self.after(0, lambda n=row["empresa"], s=status: self.note(f"{s}: {n}"))
                except Exception as exc:
                    record(row, "ERRO", str(exc))
                    self.after(0, lambda n=row["empresa"], e=str(exc): self.note(f"ERRO {n}: {e}"))
        except Exception as exc:
            self.after(0, lambda e=str(exc): self.note("Erro geral: " + e))
        finally:
            pythoncom.CoUninitialize()
            self.after(0, self.finish)

    def finish(self):
        self.busy = False
        self.run_button.configure(state="normal")
        self.note("Concluido. Confira o Outlook e o historico.")
        self.preview()

if __name__ == "__main__":
    App().mainloop()
