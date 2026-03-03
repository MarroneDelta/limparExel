# import pandas as pd
# import re
# import tkinter as tk
# from tkinter import filedialog, messagebox, ttk
# import threading

# def executar():
#     caminho_arquivo = entry_arquivo.get()
#     if not caminho_arquivo:
#         messagebox.showerror("Erro", "Selecione um arquivo Excel!")
#         return

#     entrada_colaboradores = text_colaboradores.get("1.0", tk.END).strip()
#     entrada_municipios = text_municipios.get("1.0", tk.END).strip()

#     if not entrada_colaboradores and not entrada_municipios:
#         messagebox.showerror("Erro", "Informe ao menos um colaborador ou município!")
#         return

#     caminho_salvar = filedialog.asksaveasfilename(
#         title="Salvar arquivo como",
#         defaultextension=".xlsx",
#         filetypes=[("Arquivos Excel", "*.xlsx")]
#     )

#     if not caminho_salvar:
#         return

#     # Desabilita botão e mostra barra
#     btn_executar.config(state="disabled")
#     label_status.config(text="Carregando arquivo...")
#     progress.pack(pady=(0, 10))
#     progress.start(10)
#     root.update()

#     def processar():
#         try:
#             df = pd.read_excel(caminho_arquivo, dtype={'CNPJ': str})
#             antes = len(df)

#             root.after(0, lambda: label_status.config(text="Filtrando dados..."))

#             if entrada_colaboradores:
#                 colaboradores_remover = [c.strip().upper() for c in re.split(r'[,\n]+', entrada_colaboradores) if c.strip()]
#                 df = df[~df['COLABORADOR'].str.strip().str.upper().isin(colaboradores_remover)]

#             if entrada_municipios:
#                 municipios_remover = [m.strip().upper() for m in re.split(r'[,\n]+', entrada_municipios) if m.strip()]
#                 df = df[~df['MUNICIPIO'].str.strip().str.upper().isin(municipios_remover)]

#             depois = len(df)

#             root.after(0, lambda: label_status.config(text="Salvando arquivo..."))

#             df.to_excel(caminho_salvar, index=False)

#             root.after(0, lambda: finalizar(antes, depois, caminho_salvar))

#         except Exception as e:
#             root.after(0, lambda: erro(str(e)))

#     threading.Thread(target=processar, daemon=True).start()

# def finalizar(antes, depois, caminho_salvar):
#     progress.stop()
#     progress.pack_forget()
#     label_status.config(text="")
#     btn_executar.config(state="normal")

#     # Zera os campos
#     entry_arquivo.delete(0, tk.END)
#     text_colaboradores.delete("1.0", tk.END)
#     text_municipios.delete("1.0", tk.END)

#     messagebox.showinfo("Concluído", f"Removidas {antes - depois} linhas.\nRestaram {depois}.\n\nArquivo salvo em:\n{caminho_salvar}")

# def erro(msg):
#     progress.stop()
#     progress.pack_forget()
#     label_status.config(text="")
#     btn_executar.config(state="normal")
#     messagebox.showerror("Erro", f"Ocorreu um erro:\n{msg}")

# def selecionar_arquivo():
#     caminho = filedialog.askopenfilename(
#         title="Selecione o arquivo Excel",
#         filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
#     )
#     if caminho:
#         entry_arquivo.delete(0, tk.END)
#         entry_arquivo.insert(0, caminho)

# # Interface
# root = tk.Tk()
# root.title("Filtrar Excel")
# root.geometry("500x520")
# root.resizable(False, False)

# # Arquivo
# tk.Label(root, text="Arquivo Excel:", font=("Arial", 10)).pack(anchor="w", padx=20, pady=(20, 0))
# frame_arquivo = tk.Frame(root)
# frame_arquivo.pack(fill="x", padx=20)
# entry_arquivo = tk.Entry(frame_arquivo, font=("Arial", 10))
# entry_arquivo.pack(side="left", fill="x", expand=True)
# tk.Button(frame_arquivo, text="Procurar", command=selecionar_arquivo).pack(side="right", padx=(5, 0))

# # Colaboradores
# tk.Label(root, text="Colaboradores para remover (um por linha ou separados por vírgula):", font=("Arial", 10)).pack(anchor="w", padx=20, pady=(15, 0))
# text_colaboradores = tk.Text(root, height=5, font=("Arial", 10))
# text_colaboradores.pack(fill="x", padx=20)

# # Municípios
# tk.Label(root, text="Municípios para remover (um por linha ou separados por vírgula):", font=("Arial", 10)).pack(anchor="w", padx=20, pady=(15, 0))
# text_municipios = tk.Text(root, height=5, font=("Arial", 10))
# text_municipios.pack(fill="x", padx=20)

# # Status
# label_status = tk.Label(root, text="", font=("Arial", 9), fg="gray")
# label_status.pack(pady=(10, 0))

# # Barra de progresso (oculta inicialmente)
# progress = ttk.Progressbar(root, mode="indeterminate", length=460)

# # Botão
# btn_executar = tk.Button(root, text="Filtrar e Salvar", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", command=executar)
# btn_executar.pack(pady=10)

# root.mainloop()

import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# ── Cores e fontes ──────────────────────────────────────────────
BG        = "#0F1117"
SURFACE   = "#1A1D27"
SURFACE2  = "#22263A"
ACCENT    = "#00E5FF"
ACCENT2   = "#7C3AED"
TEXT      = "#E8EAF0"
SUBTEXT   = "#6B7280"
SUCCESS   = "#10B981"
DANGER    = "#EF4444"
FONT      = ("Consolas", 10)
FONT_B    = ("Consolas", 10, "bold")
FONT_LG   = ("Consolas", 13, "bold")
FONT_SM   = ("Consolas", 9)

colunas_disponiveis = []
filtros = []

def set_widgets_state(state):
    """Habilita ou desabilita todos os controles."""
    btn_adicionar.config(state=state)
    btn_procurar.config(state=state)
    for f in filtros:
        f['combo'].config(state="readonly" if state == "normal" else "disabled")
        f['text'].config(state=state)
        f['btn_rem'].config(state=state)

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if not caminho:
        return

    entry_arquivo.config(state="normal")
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.insert(0, caminho)
    entry_arquivo.config(state="readonly")

    try:
        df = pd.read_excel(caminho, nrows=0)
        colunas_disponiveis.clear()
        colunas_disponiveis.extend(df.columns.tolist())
        for f in filtros:
            f['combo']['values'] = colunas_disponiveis
        label_status.config(text=f"✓  {len(colunas_disponiveis)} colunas carregadas", fg=SUCCESS)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler arquivo:\n{e}")

def adicionar_filtro():
    if not colunas_disponiveis:
        messagebox.showwarning("Aviso", "Selecione um arquivo Excel primeiro!")
        return

    frame_filtro = tk.Frame(frame_filtros, bg=SURFACE2, padx=10, pady=8)
    frame_filtro.pack(fill="x", padx=0, pady=4)

    # Linha topo: label coluna + combo + botão remover
    topo = tk.Frame(frame_filtro, bg=SURFACE2)
    topo.pack(fill="x")

    tk.Label(topo, text="COLUNA", font=("Consolas", 8, "bold"),
             bg=SURFACE2, fg=ACCENT).pack(side="left")

    combo = ttk.Combobox(topo, values=colunas_disponiveis, font=FONT,
                         width=32, state="readonly")
    combo.pack(side="left", padx=(8, 0))
    combo.current(0)

    filtro_dict = {'combo': combo, 'text': None, 'frame': frame_filtro, 'btn_rem': None}

    def remover():
        frame_filtro.destroy()
        filtros.remove(filtro_dict)

    btn_rem = tk.Button(topo, text="✕", font=FONT_SM, bg=SURFACE2, fg=DANGER,
                        bd=0, cursor="hand2", activebackground=SURFACE2,
                        activeforeground="#FF6B6B", command=remover)
    btn_rem.pack(side="right")
    filtro_dict['btn_rem'] = btn_rem

    # Linha valores
    tk.Label(frame_filtro, text="VALORES  (vírgula, espaço ou uma por linha)",
             font=("Consolas", 8), bg=SURFACE2, fg=SUBTEXT).pack(anchor="w", pady=(6, 2))

    text = tk.Text(frame_filtro, height=3, font=FONT,
                   bg=SURFACE, fg=TEXT, insertbackground=ACCENT,
                   relief="flat", bd=0, padx=6, pady=4,
                   selectbackground=ACCENT2)
    text.pack(fill="x")
    filtro_dict['text'] = text
    filtros.append(filtro_dict)

def executar():
    caminho_arquivo = entry_arquivo.get()
    if not caminho_arquivo:
        messagebox.showerror("Erro", "Selecione um arquivo Excel!")
        return
    if not filtros:
        messagebox.showerror("Erro", "Adicione ao menos um filtro!")
        return

    filtros_validos = []
    for f in filtros:
        coluna = f['combo'].get()
        valores = f['text'].get("1.0", tk.END).strip()
        if coluna and valores:
            lista = [v.strip().upper() for v in re.split(r'[,\n\s]+', valores) if v.strip()]
            if lista:
                filtros_validos.append({'coluna': coluna, 'valores': lista})

    if not filtros_validos:
        messagebox.showerror("Erro", "Preencha os valores de ao menos um filtro!")
        return

    caminho_salvar = filedialog.asksaveasfilename(
        title="Salvar arquivo como",
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if not caminho_salvar:
        return

    # Desabilita tudo
    btn_executar.config(state="disabled")
    set_widgets_state("disabled")
    progress.pack(fill="x", padx=20, pady=(0, 5))
    progress.start(12)

    def processar():
        try:
            root.after(0, lambda: label_status.config(text="⏳  Carregando arquivo...", fg=ACCENT))
            df = pd.read_excel(caminho_arquivo, dtype=str)
            antes = len(df)

            root.after(0, lambda: label_status.config(text="⏳  Filtrando dados...", fg=ACCENT))
            for f in filtros_validos:
                col = f['coluna']
                if col in df.columns:
                    df = df[~df[col].str.strip().str.upper().isin(f['valores'])]

            depois = len(df)
            root.after(0, lambda: label_status.config(text="⏳  Salvando arquivo...", fg=ACCENT))
            df.to_excel(caminho_salvar, index=False)
            root.after(0, lambda: finalizar(antes, depois, caminho_salvar))
        except Exception as e:
            root.after(0, lambda: erro(str(e)))

    threading.Thread(target=processar, daemon=True).start()

def finalizar(antes, depois, caminho_salvar):
    progress.stop()
    progress.pack_forget()
    btn_executar.config(state="normal")
    set_widgets_state("normal")
    label_status.config(text=f"✓  Concluído! {antes-depois} linhas removidas · {depois} restantes", fg=SUCCESS)

    entry_arquivo.config(state="normal")
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.config(state="readonly")
    for f in filtros[:]:
        f['frame'].destroy()
    filtros.clear()
    colunas_disponiveis.clear()

    messagebox.showinfo("Concluído",
        f"Removidas:  {antes - depois} linhas\nRestantes:  {depois}\n\nSalvo em:\n{caminho_salvar}")

def erro(msg):
    progress.stop()
    progress.pack_forget()
    btn_executar.config(state="normal")
    set_widgets_state("normal")
    label_status.config(text="✕  Erro ao processar", fg=DANGER)
    messagebox.showerror("Erro", f"Ocorreu um erro:\n{msg}")

# ── Janela principal ─────────────────────────────────────────────
root = tk.Tk()
root.title("Excel Filter")
root.geometry("600x680")
root.resizable(False, False)
root.configure(bg=BG)

# Estilo ttk
style = ttk.Style()
style.theme_use("clam")
style.configure("TCombobox", fieldbackground=SURFACE, background=SURFACE,
                foreground=TEXT, selectbackground=ACCENT2, font=FONT)
style.configure("TProgressbar", troughcolor=SURFACE2, background=ACCENT,
                bordercolor=BG, lightcolor=ACCENT, darkcolor=ACCENT)

# ── Header ───────────────────────────────────────────────────────
header = tk.Frame(root, bg=SURFACE, pady=16)
header.pack(fill="x")
tk.Label(header, text="EXCEL  FILTER", font=("Consolas", 16, "bold"),
         bg=SURFACE, fg=ACCENT).pack()
tk.Label(header, text="filtre e exporte dados com precisão",
         font=("Consolas", 9), bg=SURFACE, fg=SUBTEXT).pack()

# ── Arquivo ──────────────────────────────────────────────────────
section_file = tk.Frame(root, bg=BG)
section_file.pack(fill="x", padx=20, pady=(18, 0))

tk.Label(section_file, text="ARQUIVO", font=("Consolas", 8, "bold"),
         bg=BG, fg=ACCENT).pack(anchor="w")

row_file = tk.Frame(section_file, bg=BG)
row_file.pack(fill="x", pady=(4, 0))

entry_arquivo = tk.Entry(row_file, font=FONT, bg=SURFACE, fg=TEXT,
                         insertbackground=ACCENT, relief="flat",
                         state="readonly", readonlybackground=SURFACE)
entry_arquivo.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

btn_procurar = tk.Button(row_file, text="PROCURAR", font=FONT_B,
                         bg=ACCENT2, fg=TEXT, relief="flat", bd=0,
                         padx=14, pady=6, cursor="hand2",
                         activebackground="#6D28D9", activeforeground=TEXT,
                         command=selecionar_arquivo)
btn_procurar.pack(side="right")

label_status = tk.Label(root, text="", font=FONT_SM, bg=BG, fg=SUBTEXT)
label_status.pack(anchor="w", padx=20, pady=(4, 0))

# ── Divisor ──────────────────────────────────────────────────────
tk.Frame(root, bg=SURFACE2, height=1).pack(fill="x", padx=20, pady=12)

# ── Filtros ──────────────────────────────────────────────────────
row_filtros_header = tk.Frame(root, bg=BG)
row_filtros_header.pack(fill="x", padx=20)

tk.Label(row_filtros_header, text="FILTROS", font=("Consolas", 8, "bold"),
         bg=BG, fg=ACCENT).pack(side="left")

btn_adicionar = tk.Button(row_filtros_header, text="+ ADICIONAR FILTRO",
                          font=("Consolas", 8, "bold"),
                          bg=BG, fg=ACCENT, relief="flat", bd=0,
                          cursor="hand2", activebackground=BG,
                          activeforeground=TEXT, command=adicionar_filtro)
btn_adicionar.pack(side="right")

# Scroll area
frame_outer = tk.Frame(root, bg=BG)
frame_outer.pack(fill="both", expand=True, padx=20, pady=(6, 0))

canvas = tk.Canvas(frame_outer, bg=BG, highlightthickness=0)
scrollbar = ttk.Scrollbar(frame_outer, orient="vertical", command=canvas.yview)
frame_filtros = tk.Frame(canvas, bg=BG)

frame_filtros.bind("<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=frame_filtros, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# ── Barra de progresso ───────────────────────────────────────────
progress = ttk.Progressbar(root, mode="indeterminate", length=560)

# ── Botão executar ───────────────────────────────────────────────
btn_executar = tk.Button(root, text="FILTRAR  E  SALVAR",
                         font=("Consolas", 12, "bold"),
                         bg=ACCENT, fg=BG, relief="flat", bd=0,
                         padx=20, pady=12, cursor="hand2",
                         activebackground="#00B8CC", activeforeground=BG,
                         command=executar)
btn_executar.pack(fill="x", padx=20, pady=16)

root.mainloop()