from __future__ import annotations
import random, statistics, time, tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

ALGORITHMS = ("Bubble Sort", "Selection Sort", "Insertion Sort", "Merge Sort", "Quick Sort")
SCENARIOS = ("Inversamente ordenado", "Aleatório")
COMPLEXITY = {
    "Bubble Sort": ("Ω(n) otimizado", "Θ(n²)", "O(n²)", "O(1)"),
    "Selection Sort": ("Ω(n²)", "Θ(n²)", "O(n²)", "O(1)"),
    "Insertion Sort": ("Ω(n)", "Θ(n²)", "O(n²)", "O(1)"),
    "Merge Sort": ("Θ(n log n)", "Θ(n log n)", "Θ(n log n)", "O(n)"),
    "Quick Sort": ("Θ(n log n) esperado", "Θ(n log n) esperado", "O(n²)", "O(log n) esperado"),
}

def bubble_sort(a):
    for end in range(len(a)-1, 0, -1):
        changed = False
        for j in range(end):
            if a[j] > a[j+1]:
                a[j], a[j+1] = a[j+1], a[j]
                changed = True
        if not changed: break

def selection_sort(a):
    for i in range(len(a)-1):
        m = i
        for j in range(i+1, len(a)):
            if a[j] < a[m]: m = j
        if m != i: a[i], a[m] = a[m], a[i]

def insertion_sort(a):
    for i in range(1, len(a)):
        value, j = a[i], i-1
        while j >= 0 and a[j] > value:
            a[j+1] = a[j]
            j -= 1
        a[j+1] = value

def merge_sort(a):
    n = len(a)
    if n < 2: return
    aux = [0.0] * n
    def divide(lo, hi):
        if hi-lo <= 1: return
        mid = (lo+hi)//2
        divide(lo, mid); divide(mid, hi)
        i, j, k = lo, mid, lo
        while i < mid and j < hi:
            if a[i] <= a[j]: aux[k] = a[i]; i += 1
            else: aux[k] = a[j]; j += 1
            k += 1
        while i < mid: aux[k] = a[i]; i += 1; k += 1
        while j < hi: aux[k] = a[j]; j += 1; k += 1
        a[lo:hi] = aux[lo:hi]
    divide(0, n)

def quick_sort(a):
    if len(a) < 2: return
    rng, stack = random.Random(2026), [(0, len(a)-1)]
    while stack:
        lo, hi = stack.pop()
        while lo < hi:
            pivot = a[rng.randint(lo, hi)]
            lt, i, gt = lo, lo, hi
            while i <= gt:
                if a[i] < pivot:
                    a[lt], a[i] = a[i], a[lt]; lt += 1; i += 1
                elif a[i] > pivot:
                    a[i], a[gt] = a[gt], a[i]; gt -= 1
                else: i += 1
            left, right = (lo, lt-1), (gt+1, hi)
            if left[1]-left[0] < right[1]-right[0]:
                if right[0] < right[1]: stack.append(right)
                lo, hi = left
            else:
                if left[0] < left[1]: stack.append(left)
                lo, hi = right

SORTERS = {"Bubble Sort": bubble_sort, "Selection Sort": selection_sort,
           "Insertion Sort": insertion_sort, "Merge Sort": merge_sort, "Quick Sort": quick_sort}

def norm(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c.lower() for c in s if c.isalnum() and not unicodedata.combining(c))

def load_data(filename):
    p = Path(filename)
    if p.suffix.lower() in (".xlsx", ".xls"): df = pd.read_excel(p)
    elif p.suffix.lower() == ".csv":
        try: df = pd.read_csv(p, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError: df = pd.read_csv(p, sep=None, engine="python", encoding="latin1")
    else: raise ValueError("Selecione CSV ou XLSX.")
    cols = [c for c in df.columns if norm(c) == "areamunkm"]
    if not cols: cols = [c for c in df.columns if "areamunkm" in norm(c)]
    if not cols: raise ValueError("Coluna areaMunKm não encontrada.")
    values = pd.to_numeric(df[cols[0]].astype(str).str.strip().str.replace(",", ".", regex=False),
                           errors="coerce").dropna()
    values = values[values >= 0].astype(float).tolist()
    if not values: raise ValueError("areaMunKm não contém valores numéricos válidos.")
    return values, str(cols[0]), len(df)

def benchmark(name, source, repetitions):
    expected, times = sorted(source), []
    for _ in range(repetitions):
        work = source.copy()                  # cópia fora do cronômetro
        start = time.perf_counter_ns()
        SORTERS[name](work)
        elapsed = time.perf_counter_ns() - start
        times.append(elapsed / 1e9)
        if work != expected: raise RuntimeError(f"Ordenação incorreta: {name}")
    return {"media": statistics.mean(times), "mediana": statistics.median(times),
            "min": min(times), "max": max(times), "times": times}

class App:
    def __init__(self, root):
        self.root = root
        root.title("APS UNIP | Benchmark de Algoritmos — Pro")
        root.geometry("1240x810"); root.minsize(1030, 680)
        self.datasets, self.summary, self.raw, self.fig = {}, [], [], None
        self.build()

    def build(self):
        style = ttk.Style()
        try: style.theme_use("clam")
        except tk.TclError: pass
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        top = ttk.Frame(self.root, padding=14); top.pack(fill="x")
        ttk.Label(top, text="APS | ANÁLISE DE PERFORMANCE", style="Title.TLabel").pack(anchor="w")
        ttk.Label(top, text="TerraBrasilis • Bioma Amazônia • Campo areaMunKm").pack(anchor="w")
        box = ttk.LabelFrame(self.root, text="1. Importação das bases", padding=9); box.pack(fill="x", padx=14, pady=5)
        self.labels = {}
        for key, caption in (("Com filtro", "Base COM filtro"), ("Sem filtro", "Base SEM filtro")):
            row = ttk.Frame(box); row.pack(fill="x", pady=3)
            ttk.Label(row, text=caption, width=17).pack(side="left")
            self.labels[key] = ttk.Label(row, text="Não carregada", foreground="#777")
            self.labels[key].pack(side="left", fill="x", expand=True, padx=8)
            ttk.Button(row, text="Selecionar CSV/XLSX", command=lambda k=key:self.select(k)).pack(side="right")
        cfg = ttk.LabelFrame(self.root, text="2. Configuração", padding=9); cfg.pack(fill="x", padx=14, pady=5)
        ttk.Label(cfg, text="Limite de registros (0=todos):").pack(side="left")
        self.limit = tk.StringVar(value="1500")
        ttk.Entry(cfg, textvariable=self.limit, width=8).pack(side="left", padx=(5,16))
        ttk.Label(cfg, text="Repetições:").pack(side="left")
        self.reps = tk.StringVar(value="5")
        ttk.Entry(cfg, textvariable=self.reps, width=5).pack(side="left", padx=5)
        ttk.Button(cfg, text="EXECUTAR COMPARAÇÃO", command=self.run).pack(side="right")
        ttk.Label(self.root, text="Medição com perf_counter_ns; cópia e validação fora do cronômetro; "
                 "tabela resume mediana, média, mínimo e máximo.", foreground="#555").pack(anchor="w", padx=18, pady=4)
        result = ttk.LabelFrame(self.root, text="3. Resultados", padding=7); result.pack(fill="both", expand=True, padx=14, pady=5)
        cols = ("base","scenario","algo","n","median","mean","min","max","valid")
        self.tree = ttk.Treeview(result, columns=cols, show="headings", height=10)
        heads = ("Planilha","Cenário","Algoritmo","Registros","Mediana (s)","Média (s)","Mínimo (s)","Máximo (s)","Validação")
        widths = (105,160,125,80,100,100,100,100,80)
        for c,h,w in zip(cols,heads,widths):
            self.tree.heading(c,text=h); self.tree.column(c,width=w,anchor="center")
        self.tree.pack(side="left",fill="both",expand=True)
        sb = ttk.Scrollbar(result,orient="vertical",command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set); sb.pack(side="right",fill="y")
        chart = ttk.LabelFrame(self.root,text="4. Gráficos comparativos",padding=7)
        chart.pack(fill="both",expand=True,padx=14,pady=5); self.chart_box=chart
        ttk.Label(chart,text="Os gráficos serão exibidos após a execução.").pack(anchor="w")
        bottom=ttk.Frame(self.root,padding=8); bottom.pack(fill="x")
        self.status=tk.StringVar(value="Aguardando importação.")
        ttk.Label(bottom,textvariable=self.status).pack(side="left",fill="x",expand=True)
        ttk.Button(bottom,text="Complexidade",command=self.complexity).pack(side="right",padx=2)
        ttk.Button(bottom,text="Exportar Excel",command=self.export_xlsx).pack(side="right",padx=2)
        ttk.Button(bottom,text="Exportar CSV",command=self.export_csv).pack(side="right",padx=2)
        ttk.Button(bottom,text="Salvar PNG",command=self.save_png).pack(side="right",padx=2)

    def select(self,key):
        f=filedialog.askopenfilename(filetypes=[("Planilhas","*.csv *.xlsx *.xls")])
        if not f:return
        try:
            vals,col,total=load_data(f); self.datasets[key]={"values":vals,"path":f}
            self.labels[key].config(text=f"{Path(f).name} — {len(vals):,} valores válidos".replace(",","."),foreground="green")
            self.status.set(f"Base {key} carregada; coluna {col}.")
        except Exception as e: messagebox.showerror("Importação",str(e))

    def run(self):
        if set(self.datasets)!={"Com filtro","Sem filtro"}:
            messagebox.showwarning("Bases","Carregue as duas bases."); return
        try:
            limit,reps=int(self.limit.get()),int(self.reps.get())
            if limit<0 or not 1<=reps<=30: raise ValueError
        except ValueError:
            messagebox.showerror("Parâmetros","Limite >= 0 e repetições entre 1 e 30."); return
        prepared={}
        for k,d in self.datasets.items():
            v=d["values"]
            if limit and len(v)>limit:v=random.Random(2026).sample(v,limit)
            prepared[k]=v
        n=max(map(len,prepared.values()))
        if n>5000 and not messagebox.askyesno("Aviso",f"{n:,} registros: algoritmos quadráticos podem demorar. Continuar?".replace(",",".")):return
        self.root.config(cursor="watch"); self.status.set("Executando..."); self.root.update_idletasks()
        self.summary=[]; self.raw=[]
        for item in self.tree.get_children():self.tree.delete(item)
        try:
            total=len(prepared)*len(SCENARIOS)*len(ALGORITHMS); done=0
            for base,original in prepared.items():
                for scenario in SCENARIOS:
                    data=sorted(original,reverse=True) if scenario==SCENARIOS[0] else original.copy()
                    if scenario==SCENARIOS[1]:random.Random(2026).shuffle(data)
                    for algo in ALGORITHMS:
                        r=benchmark(algo,data,reps)
                        row={"planilha":base,"cenario":scenario,"algoritmo":algo,"registros":len(data),
                             "mediana_s":r["mediana"],"media_s":r["media"],"minimo_s":r["min"],
                             "maximo_s":r["max"],"repeticoes":reps,"validacao":"Correto"}
                        self.summary.append(row)
                        for ix,t in enumerate(r["times"],1):
                            self.raw.append({"planilha":base,"cenario":scenario,"algoritmo":algo,
                                             "repeticao":ix,"tempo_s":t,"registros":len(data)})
                        self.tree.insert("","end",values=(base,scenario,algo,len(data),
                            f'{r["mediana"]:.6f}',f'{r["media"]:.6f}',f'{r["min"]:.6f}',f'{r["max"]:.6f}',"Correto"))
                        done+=1; self.status.set(f"Teste {done}/{total}: {base} — {scenario} — {algo}")
                        self.root.update_idletasks()
            self.draw(); self.status.set(f"Concluído: {done} combinações, {len(self.raw)} medições.")
            messagebox.showinfo("Concluído",f"{done} combinações testadas; {len(self.raw)} medições.")
        except Exception as e:
            messagebox.showerror("Benchmark",str(e)); self.status.set("Falha no teste.")
        finally:self.root.config(cursor="")

    def draw(self):
        for w in self.chart_box.winfo_children():w.destroy()
        fig,axes=plt.subplots(1,2,figsize=(12,3.7))
        for ax,base in zip(axes,("Com filtro","Sem filtro")):
            rows=[r for r in self.summary if r["planilha"]==base]; x=list(range(5)); width=.36
            inv=[next(r["mediana_s"] for r in rows if r["algoritmo"]==a and r["cenario"]==SCENARIOS[0]) for a in ALGORITHMS]
            rnd=[next(r["mediana_s"] for r in rows if r["algoritmo"]==a and r["cenario"]==SCENARIOS[1]) for a in ALGORITHMS]
            ax.bar([i-width/2 for i in x],inv,width,label="Inverso")
            ax.bar([i+width/2 for i in x],rnd,width,label="Aleatório")
            ax.set_title("Base "+base); ax.set_xticks(x); ax.set_xticklabels(["Bubble","Selection","Insertion","Merge","Quick"],rotation=15)
            ax.set_ylabel("Mediana (s)"); ax.grid(axis="y",alpha=.25); ax.legend(fontsize=8)
        fig.tight_layout(); self.fig=fig
        canvas=FigureCanvasTkAgg(fig,master=self.chart_box); canvas.draw()
        canvas.get_tk_widget().pack(fill="both",expand=True)

    def export_xlsx(self):
        if not self.summary:messagebox.showwarning("Sem dados","Execute primeiro.");return
        f=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],initialfile="resultado_APS.xlsx")
        if not f:return
        comp=[{"algoritmo":k,"melhor_caso":v[0],"caso_medio":v[1],"pior_caso":v[2],"espaco":v[3]} for k,v in COMPLEXITY.items()]
        with pd.ExcelWriter(f,engine="openpyxl") as w:
            pd.DataFrame(self.summary).to_excel(w,sheet_name="Resumo",index=False)
            pd.DataFrame(self.raw).to_excel(w,sheet_name="Medicoes",index=False)
            pd.DataFrame(comp).to_excel(w,sheet_name="Complexidade",index=False)
        messagebox.showinfo("Excel",f"Arquivo salvo em:\\n{f}")

    def export_csv(self):
        if not self.summary:messagebox.showwarning("Sem dados","Execute primeiro.");return
        f=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv")],initialfile="resultado_APS.csv")
        if f:
            pd.DataFrame(self.summary).to_csv(f,index=False,sep=";",decimal=",",encoding="utf-8-sig")
            messagebox.showinfo("CSV",f"Arquivo salvo em:\\n{f}")

    def save_png(self):
        if self.fig is None:messagebox.showwarning("Sem gráfico","Execute primeiro.");return
        f=filedialog.asksaveasfilename(defaultextension=".png",filetypes=[("PNG","*.png")],initialfile="graficos_APS.png")
        if f:self.fig.savefig(f,dpi=200,bbox_inches="tight");messagebox.showinfo("PNG",f"Gráfico salvo em:\\n{f}")

    def complexity(self):
        lines=[f"{k}\\nMelhor: {v[0]} | Médio: {v[1]} | Pior: {v[2]} | Espaço: {v[3]}" for k,v in COMPLEXITY.items()]
        messagebox.showinfo("Complexidade","\\n\\n".join(lines)+"\\n\\nO: limite superior; Ω: inferior; Θ: justo.")

if __name__=="__main__":
    root=tk.Tk(); App(root); root.mainloop()
