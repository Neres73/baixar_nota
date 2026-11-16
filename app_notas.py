"""
Download de Notas — interface gráfica.

Usa o main.py como "motor": esta tela só escolhe a planilha e a pasta, chama
os mesmos fluxos e mostra o andamento em linguagem simples. As perguntas que
o main.py fazia com input() viram janelas de confirmação.
"""
import builtins
import collections
import ctypes
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

PASTA_APP = Path(__file__).resolve().parent
ARQUIVO_CONFIG = PASTA_APP.parent / "config.json"   # fora de /app: sobrevive às atualizações
ICONE = PASTA_APP / "icone.ico"
sys.path.insert(0, str(PASTA_APP))

# Antes de carregar o motor: se houver versão nova na pasta da rede, atualiza e reabre.
try:
    import atualizador   # noqa: E402
    atualizador.verificar(__file__)
    VERSAO = atualizador.versao_instalada(__file__)
except Exception:
    VERSAO = ""

import main as nucleo   # noqa: E402  (o motor da automação)

# Endereço da página de busca de ocorrências do portal (já vem preenchido para a
# equipe). A busca em si — ocorrência, datas — o motor preenche sozinho.
URL_PORTAL_PADRAO = os.getenv("NIX_URL", "")


def limpar_url(url):
    """Tira a parte da busca (depois do '?') de um endereço colado do navegador."""
    return url.strip().split("?")[0].split("#")[0]


# =====================================================================
#                         CORES E FONTES
# =====================================================================
FUNDO = "#F4F6FA"
CARTAO = "#FFFFFF"
BORDA = "#E3E8EF"
TEXTO = "#1F2937"
SUAVE = "#6B7280"
CINZA_CLARO = "#EEF1F5"
HOVER_NEUTRO = "#E9EDF3"
AZUL = "#2F6FEB"
AZUL_ESCURO = "#2259C7"
AZUL_CLARO = "#EAF1FE"
AZUL_HOVER = "#D9E6FD"
AZUL_DESATIVADO = "#9DB7EE"
VERDE = "#1E9E5A"
VERDE_CLARO = "#E7F6EE"
AMARELO = "#B7791F"
AMARELO_CLARO = "#FDF4E3"
VERMELHO = "#D64545"
VERMELHO_CLARO = "#FDECEC"
FAMILIA = "Segoe UI"


def fonte(tamanho=13, peso="normal"):
    return ctk.CTkFont(family=FAMILIA, size=tamanho, weight=peso)


def botao_primario(pai, texto, comando, **kw):
    kw.setdefault("height", 38)
    return ctk.CTkButton(pai, text=texto, command=comando, corner_radius=8, fg_color=AZUL,
                         hover_color=AZUL_ESCURO, font=fonte(13, "bold"), **kw)


def botao_secundario(pai, texto, comando, **kw):
    kw.setdefault("height", 38)
    return ctk.CTkButton(pai, text=texto, command=comando, corner_radius=8, fg_color=AZUL_CLARO,
                         hover_color=AZUL_HOVER, text_color=AZUL, font=fonte(13, "bold"), **kw)


def botao_texto(pai, texto, comando, cor=SUAVE, **kw):
    kw.setdefault("height", 30)
    return ctk.CTkButton(pai, text=texto, command=comando, fg_color="transparent",
                         hover_color=HOVER_NEUTRO, text_color=cor, font=fonte(12), **kw)


# =====================================================================
#      PONTE ENTRE O MOTOR (thread de trabalho) E A TELA (thread principal)
# =====================================================================
fila_ui = queue.Queue()
ultimas_linhas = collections.deque(maxlen=15)


class EscritorFila:
    """Substitui o print(): tudo que o motor imprime chega na tela."""
    def write(self, texto):
        if texto:
            fila_ui.put(("log", texto))

    def flush(self):
        pass


def input_pela_tela(prompt=""):
    """Substitui o input(): pede a resposta numa janela e espera o usuário."""
    if threading.current_thread() is threading.main_thread():
        return ""
    pedido = {"prompt": str(prompt), "evento": threading.Event(), "resposta": None}
    fila_ui.put(("input", pedido))
    while not pedido["evento"].wait(0.2):   # espera em fatias para o botão Parar funcionar
        pass
    if pedido["resposta"] is None:
        raise KeyboardInterrupt
    return pedido["resposta"]


sys.stdout = EscritorFila()
sys.stderr = EscritorFila()
builtins.input = input_pela_tela


def interromper_thread(thread):
    """Equivale ao Ctrl+C do terminal, só que na thread que está rodando o fluxo."""
    if thread and thread.is_alive():
        ctypes.pythonapi.PyThreadState_SetAsyncExc(
            ctypes.c_ulong(thread.ident), ctypes.py_object(KeyboardInterrupt))


def avisar_etapa(nome, funcao):
    """Embrulha uma etapa do motor para a tela saber quando ela começa."""
    def embrulho(*args, **kwargs):
        fila_ui.put(("etapa", nome))
        return funcao(*args, **kwargs)
    return embrulho


nucleo.baixar_notas_sap = avisar_etapa("sap", nucleo.baixar_notas_sap)
nucleo.processar_portal = avisar_etapa("portal", nucleo.processar_portal)
nucleo.verificar = avisar_etapa("conferencia", nucleo.verificar)


# =====================================================================
#                         CONFIGURAÇÃO SALVA
# =====================================================================
def carregar_config():
    try:
        return json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_config(dados):
    try:
        ARQUIVO_CONFIG.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def aplicar_icone(janela):
    """Põe o ícone do programa na janela (barra de título e barra de tarefas)."""
    if sys.platform == "win32" and ICONE.exists():
        try:
            janela.iconbitmap(str(ICONE))
        except Exception:
            pass


def abrir_no_windows(caminho):
    if hasattr(os, "startfile"):
        os.startfile(caminho)


def formatar_tempo(segundos):
    segundos = int(segundos)
    h, resto = divmod(segundos, 3600)
    m, s = divmod(resto, 60)
    return f"{h}h{m:02d}" if h else f"{m:02d}:{s:02d}"


# =====================================================================
#                         COMPONENTES DA TELA
# =====================================================================
class Cartao(ctk.CTkFrame):
    def __init__(self, pai, **kw):
        super().__init__(pai, fg_color=CARTAO, corner_radius=14, border_width=1,
                         border_color=BORDA, **kw)


class PassoEscolha(Cartao):
    """Cartão de um passo: número, título, o que foi escolhido e botão de escolher."""
    def __init__(self, pai, numero, titulo, vazio, ao_clicar):
        super().__init__(pai)
        self.numero = str(numero)
        self.vazio = vazio
        self.grid_columnconfigure(1, weight=1)

        self.bolinha = ctk.CTkLabel(self, text=self.numero, width=34, height=34, corner_radius=17,
                                    fg_color=AZUL_CLARO, text_color=AZUL, font=fonte(15, "bold"))
        self.bolinha.grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=14)

        ctk.CTkLabel(self, text=titulo, font=fonte(14, "bold"), text_color=TEXTO, anchor="w").grid(
            row=0, column=1, sticky="sw", pady=(12, 0))
        self.lbl_valor = ctk.CTkLabel(self, text=vazio, font=fonte(12), text_color=SUAVE, anchor="w")
        self.lbl_valor.grid(row=1, column=1, sticky="nw", pady=(0, 12))

        self.botao = botao_secundario(self, "Escolher", ao_clicar, width=110, height=34)
        self.botao.grid(row=0, column=2, rowspan=2, padx=16)

    def mostrar(self, caminho):
        if caminho:
            caminho = caminho.rstrip("\\/")
            nome = os.path.basename(caminho) or caminho
            pasta = os.path.dirname(caminho)
            if len(pasta) > 48:
                pasta = "…" + pasta[-47:]
            self.lbl_valor.configure(text=f"{nome}   ·   {pasta}" if pasta else nome, text_color=TEXTO)
            self.bolinha.configure(text="✓", fg_color=VERDE_CLARO, text_color=VERDE)
            self.botao.configure(text="Trocar")
        else:
            self.lbl_valor.configure(text=self.vazio, text_color=SUAVE)
            self.bolinha.configure(text=self.numero, fg_color=AZUL_CLARO, text_color=AZUL)
            self.botao.configure(text="Escolher")

    def travar(self, travado):
        self.botao.configure(state="disabled" if travado else "normal")


class Contador(ctk.CTkFrame):
    """Quadradinho com um número grande e uma legenda. O número 'rola' até o valor novo."""
    def __init__(self, pai, legenda, cor, fundo):
        super().__init__(pai, fg_color=fundo, corner_radius=10)
        self.valor_mostrado = 0
        self.valor_alvo = 0
        self.lbl_numero = ctk.CTkLabel(self, text="0", font=fonte(24, "bold"), text_color=cor)
        self.lbl_numero.pack(pady=(10, 0))
        ctk.CTkLabel(self, text=legenda, font=fonte(12), text_color=cor).pack(pady=(0, 10))

    def definir(self, valor, animar=True):
        self.valor_alvo = max(0, valor)
        if not animar:
            self.valor_mostrado = self.valor_alvo
            self.lbl_numero.configure(text=str(self.valor_alvo))

    def passo_animacao(self):
        if self.valor_mostrado != self.valor_alvo:
            diferenca = self.valor_alvo - self.valor_mostrado
            salto = max(1, abs(diferenca) // 6)
            self.valor_mostrado += salto if diferenca > 0 else -salto
            self.lbl_numero.configure(text=str(self.valor_mostrado))


# =====================================================================
#                              A TELA
# =====================================================================
RE_TERMINO = re.compile(r"término ~(\d{2}:\d{2})")
RE_CONTADOR = re.compile(r"\[(\d+)/(\d+)\]")
RE_PENDENTES_PORTAL = re.compile(r"(\d+) nota\(s\) para baixar no portal")
RE_VERIF = {
    "ok": re.compile(r"OK \(arquivo próprio na pasta\)\s*:\s*(\d+)"),
    "faltando": re.compile(r"FALTANDO\s*:\s*(\d+)"),
    "corrompido": re.compile(r"PDF corrompido\s*:\s*(\d+)"),
    "total": re.compile(r"Linhas na planilha\s*:\s*(\d+)"),
}

ETAPAS = [("sap", "SAP"), ("portal", "Portal"), ("conferencia", "Conferência")]
TEXTOS_ETAPA = {
    "sap": ("Baixando no SAP", "Não mexa no SAP enquanto isso."),
    "portal": ("Baixando no portal", "O navegador abre sozinho. Pode deixar ele trabalhar."),
    "conferencia": ("Conferindo a pasta", "Vendo se cada nota da planilha tem o seu PDF."),
}


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=FUNDO)
        ctk.set_appearance_mode("light")
        self.title("Download de Notas" + (f"  ·  versão {VERSAO}" if VERSAO else ""))
        aplicar_icone(self)
        self.minsize(700, 640)
        self._centralizar(780, 860)

        self.thread = None
        cfg = carregar_config()
        self.planilha = cfg.get("planilha", "")
        self.pasta = cfg.get("pasta", "")
        self.url_portal = limpar_url(cfg.get("url_nix", "")) or URL_PORTAL_PADRAO

        self.contagem = {"ok": 0, "portal": 0, "erro": 0}
        self.verificacao = {}
        self.etapa_atual = None
        self.etapas_vistas = []
        self.inicio_execucao = None
        self.alvo_barra = 0.0
        self.confirmar_auto = False

        # área rolável: a janela funciona mesmo em telas pequenas de notebook
        self.corpo = ctk.CTkScrollableFrame(self, fg_color="transparent",
                                            scrollbar_button_color="#DCE1E8",
                                            scrollbar_button_hover_color="#B8C0CC")
        self.corpo.pack(fill="both", expand=True, padx=14, pady=10)
        self.corpo.grid_columnconfigure(0, weight=1)
        self.miolo = ctk.CTkFrame(self.corpo, fg_color="transparent")
        self.miolo.pack(fill="both", expand=True, padx=14, pady=8)

        self._montar_cabecalho()
        self._montar_passos()
        self._montar_botao_principal()
        self._montar_outras_opcoes()
        self._montar_andamento()
        self._montar_detalhes()

        self.passo_planilha.mostrar(self.planilha)
        self.passo_pasta.mostrar(self.pasta)
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

        self.after(100, self._processar_fila)
        self.after(16, self._animar)
        self.after(1000, self._atualizar_relogio)

    def _centralizar(self, largura, altura):
        self.update_idletasks()
        altura = min(altura, self.winfo_screenheight() - 80)
        x = max(0, (self.winfo_screenwidth() - largura) // 2)
        y = max(0, (self.winfo_screenheight() - altura) // 2 - 20)
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    # ---------- montagem ----------
    def _montar_cabecalho(self):
        topo = ctk.CTkFrame(self.miolo, fg_color="transparent")
        topo.pack(fill="x", pady=(0, 18))
        ctk.CTkLabel(topo, text="NF", font=fonte(18, "bold"), width=52, height=52, corner_radius=14,
                     fg_color=AZUL, text_color="#FFFFFF").pack(side="left", padx=(0, 14))
        textos = ctk.CTkFrame(topo, fg_color="transparent")
        textos.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(textos, text="Download de Notas", font=fonte(22, "bold"),
                     text_color=TEXTO, anchor="w").pack(anchor="w")
        ctk.CTkLabel(textos, text="Baixe as notas fiscais do mês em poucos cliques.",
                     font=fonte(13), text_color=SUAVE, anchor="w").pack(anchor="w")
        botao_texto(topo, "⚙  Configurações", self._abrir_configuracoes, width=130).pack(side="right")

    def _montar_passos(self):
        self.passo_planilha = PassoEscolha(self.miolo, 1, "Escolha a planilha do mês",
                                           "Nenhuma planilha escolhida", self._escolher_planilha)
        self.passo_planilha.pack(fill="x", pady=(0, 10))
        self.passo_pasta = PassoEscolha(self.miolo, 2, "Escolha onde salvar as notas",
                                        "Nenhuma pasta escolhida", self._escolher_pasta)
        self.passo_pasta.pack(fill="x", pady=(0, 18))

    def _montar_botao_principal(self):
        self.btn_principal = ctk.CTkButton(
            self.miolo, text="Baixar todas as notas", height=56, corner_radius=12,
            fg_color=AZUL, hover_color=AZUL_ESCURO, text_color_disabled="#FFFFFF",
            font=fonte(17, "bold"),
            command=lambda: self._rodar("Baixar todas as notas", nucleo.fluxo_completo))
        self.btn_principal.pack(fill="x")
        ctk.CTkLabel(self.miolo, text="Primeiro o SAP, depois o portal para o que faltar, e no fim uma conferência.",
                     font=fonte(12), text_color=SUAVE).pack(pady=(6, 0))

    def _montar_outras_opcoes(self):
        self.btn_mais = botao_texto(self.miolo, "Outras opções  ▾", self._alternar_opcoes,
                                    cor=AZUL, width=150)
        self.btn_mais.pack(pady=(4, 0))

        self.painel_opcoes = ctk.CTkFrame(self.miolo, fg_color="transparent")
        opcoes = [
            ("Só SAP", "Baixa as notas de serviço que estão no SAP", nucleo.fluxo_sap),
            ("Só portal", "Baixa no portal o que não saiu pelo SAP", nucleo.fluxo_nix),
            ("Conferir pasta", "Vê se falta alguma nota na pasta", nucleo.fluxo_verificar),
            ("Baixar o que faltou", "Confere e baixa só as que faltam", nucleo.fluxo_reprocessar),
        ]
        self.botoes_execucao = [self.btn_principal]
        for c, (titulo, dica, funcao) in enumerate(opcoes):
            self.painel_opcoes.grid_columnconfigure(c, weight=1, uniform="opcoes")
            caixa = Cartao(self.painel_opcoes)
            caixa.grid(row=0, column=c, sticky="nsew", padx=(0 if c == 0 else 5, 0 if c == 3 else 5))
            b = ctk.CTkButton(caixa, text=titulo, height=30, corner_radius=8, fg_color="transparent",
                              hover_color=AZUL_CLARO, text_color=AZUL, font=fonte(13, "bold"),
                              command=lambda t=titulo, f=funcao: self._rodar(t, f))
            b.pack(fill="x", padx=8, pady=(8, 0))
            ctk.CTkLabel(caixa, text=dica, font=fonte(11), text_color=SUAVE,
                         wraplength=140).pack(padx=8, pady=(0, 10))
            self.botoes_execucao.append(b)

    def _montar_andamento(self):
        self.cartao_andamento = Cartao(self.miolo)
        c = self.cartao_andamento

        # trilha das etapas + relógio
        topo = ctk.CTkFrame(c, fg_color="transparent")
        topo.pack(fill="x", padx=18, pady=(16, 6))
        self.chips = {}
        for i, (chave, nome) in enumerate(ETAPAS):
            chip = ctk.CTkLabel(topo, text=f" {i + 1}  {nome} ", font=fonte(12, "bold"),
                                corner_radius=12, height=26, fg_color=CINZA_CLARO, text_color=SUAVE)
            chip.pack(side="left", padx=(0, 6))
            self.chips[chave] = chip
            if i < len(ETAPAS) - 1:
                ctk.CTkLabel(topo, text="›", text_color="#B8C0CC", font=fonte(15)).pack(side="left", padx=(0, 6))
        self.lbl_relogio = ctk.CTkLabel(topo, text="", font=fonte(12), text_color=SUAVE)
        self.lbl_relogio.pack(side="right")

        self.lbl_titulo = ctk.CTkLabel(c, text="", font=fonte(20, "bold"), text_color=TEXTO, anchor="w")
        self.lbl_titulo.pack(fill="x", padx=18, pady=(10, 0))
        self.lbl_sub = ctk.CTkLabel(c, text="", font=fonte(13), text_color=SUAVE, anchor="w",
                                    justify="left", wraplength=640)
        self.lbl_sub.pack(fill="x", padx=18)

        self.barra = ctk.CTkProgressBar(c, height=12, corner_radius=6, progress_color=AZUL,
                                        fg_color="#E6EAF0")
        self.barra.set(0)
        self.barra.pack(fill="x", padx=18, pady=(12, 14))

        tiles = ctk.CTkFrame(c, fg_color="transparent")
        tiles.pack(fill="x", padx=18)
        for i in range(3):
            tiles.grid_columnconfigure(i, weight=1, uniform="tiles")
        self.ct_ok = Contador(tiles, "baixadas", VERDE, VERDE_CLARO)
        self.ct_portal = Contador(tiles, "aguardando o portal", AMARELO, AMARELO_CLARO)
        self.ct_erro = Contador(tiles, "com problema", VERMELHO, VERMELHO_CLARO)
        self.contadores = (self.ct_ok, self.ct_portal, self.ct_erro)
        for i, t in enumerate(self.contadores):
            t.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 5, 0 if i == 2 else 5))

        rodape = ctk.CTkFrame(c, fg_color="transparent")
        rodape.pack(fill="x", padx=18, pady=16)
        self.btn_parar = ctk.CTkButton(rodape, text="Parar", width=90, height=36, corner_radius=8,
                                       fg_color="transparent", border_width=1, border_color=VERMELHO,
                                       text_color=VERMELHO, hover_color=VERMELHO_CLARO,
                                       font=fonte(13, "bold"), command=self._parar)
        self.btn_faltou = botao_primario(rodape, "Baixar o que faltou",
                                         lambda: self._rodar("Baixar o que faltou", nucleo.fluxo_reprocessar,
                                                             confirmar_auto=True), width=170, height=36)
        self.btn_relatorio = botao_secundario(rodape, "Abrir relatório", self._abrir_relatorio,
                                              width=130, height=36)
        self.btn_relatorio.pack(side="right")
        self.btn_pasta = botao_secundario(rodape, "Abrir pasta", self._abrir_pasta, width=110, height=36)
        self.btn_pasta.pack(side="right", padx=8)

    def _montar_detalhes(self):
        self.btn_detalhes = botao_texto(self.miolo, "Ver detalhes técnicos  ▾", self._alternar_detalhes)
        self.caixa_log = ctk.CTkTextbox(self.miolo, height=220, corner_radius=10, border_width=1,
                                        border_color=BORDA, fg_color=CARTAO, text_color="#374151",
                                        font=ctk.CTkFont(family="Consolas", size=11), wrap="word")
        self.caixa_log.tag_config("ok", foreground=VERDE)
        self.caixa_log.tag_config("erro", foreground=VERMELHO)
        self.caixa_log.tag_config("portal", foreground=AMARELO)
        self.caixa_log.tag_config("titulo", foreground=AZUL)
        self.caixa_log.configure(state="disabled")

    # ---------- mostrar / esconder ----------
    def _alternar_opcoes(self):
        if self.painel_opcoes.winfo_ismapped():
            self.painel_opcoes.pack_forget()
            self.btn_mais.configure(text="Outras opções  ▾")
        else:
            self.painel_opcoes.pack(fill="x", pady=(6, 0), after=self.btn_mais)
            self.btn_mais.configure(text="Outras opções  ▴")

    def _alternar_detalhes(self):
        if self.caixa_log.winfo_ismapped():
            self.caixa_log.pack_forget()
            self.btn_detalhes.configure(text="Ver detalhes técnicos  ▾")
        else:
            self.caixa_log.pack(fill="both", expand=True, pady=(4, 0))
            self.btn_detalhes.configure(text="Esconder detalhes técnicos  ▴")
            self.after(50, lambda: self.corpo._parent_canvas.yview_moveto(1.0))

    def _mostrar_andamento(self):
        if not self.cartao_andamento.winfo_ismapped():
            self.cartao_andamento.pack(fill="x", pady=(18, 0))
            self.btn_detalhes.pack(pady=(8, 0))
        if self.painel_opcoes.winfo_ismapped():
            self._alternar_opcoes()

    # ---------- escolhas ----------
    def _escolher_planilha(self):
        inicial = os.path.dirname(self.planilha) if self.planilha else None
        caminho = filedialog.askopenfilename(title="Planilha de ocorrências", initialdir=inicial,
                                             filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if caminho:
            self.planilha = os.path.normpath(caminho)
            self.passo_planilha.mostrar(self.planilha)
            self._salvar()

    def _escolher_pasta(self):
        caminho = filedialog.askdirectory(title="Pasta onde as notas serão salvas",
                                          initialdir=self.pasta or None)
        if caminho:
            self.pasta = os.path.normpath(caminho)
            self.passo_pasta.mostrar(self.pasta)
            self._salvar()

    def _abrir_configuracoes(self):
        janela = DialogoConfiguracoes(self, self.url_portal)
        if janela.resultado is not None:
            self.url_portal = janela.resultado
            self._salvar()

    def _salvar(self):
        salvar_config({"planilha": self.planilha, "pasta": self.pasta, "url_nix": self.url_portal})

    def _abrir_pasta(self):
        if self.pasta and os.path.isdir(self.pasta):
            abrir_no_windows(self.pasta)

    def _abrir_relatorio(self):
        candidatos = [os.path.join(self.pasta, n) for n in ("_verificacao_nfse.xlsx", "_relatorio_nfse.xlsx")]
        existentes = [c for c in candidatos if os.path.exists(c)]
        if existentes:
            abrir_no_windows(max(existentes, key=os.path.getmtime))
        else:
            DialogoAviso(self, "Relatório", "O relatório aparece aqui quando a automação terminar.")

    # ---------- execução ----------
    def _aplicar_config_no_motor(self):
        nucleo.ARQUIVO_LISTA = self.planilha
        nucleo.PASTA_DESTINO = self.pasta
        nucleo.RELATORIO_EXCEL = os.path.join(self.pasta, "_relatorio_nfse.xlsx")
        nucleo.ARQUIVO_VERIFICACAO = os.path.join(self.pasta, "_verificacao_nfse.xlsx")
        nucleo.PASTA_DIAGNOSTICO = os.path.join(self.pasta, "_diagnostico_nix")
        nucleo.URL_NIX = self.url_portal

    def _validar(self, funcao):
        if funcao is not nucleo.fluxo_nix and not os.path.isfile(self.planilha):
            DialogoAviso(self, "Falta a planilha", "Primeiro escolha a planilha do mês (passo 1).")
            return False
        if not self.pasta:
            DialogoAviso(self, "Falta a pasta", "Primeiro escolha onde salvar as notas (passo 2).")
            return False
        os.makedirs(self.pasta, exist_ok=True)
        return True

    def _rodar(self, nome, funcao, confirmar_auto=False):
        if self.thread and self.thread.is_alive():
            return
        if not self._validar(funcao):
            return
        self._aplicar_config_no_motor()
        self.confirmar_auto = confirmar_auto
        self._preparar_andamento()

        def trabalho():
            try:
                try:
                    import pythoncom          # o SAP (COM) precisa disso fora da thread principal
                    pythoncom.CoInitialize()
                except ImportError:
                    pass
                print(f">>> {nome}\n")
                funcao()
                fila_ui.put(("fim", "ok"))
            except KeyboardInterrupt:
                print("\n⏹  Interrompido.")
                fila_ui.put(("fim", "parado"))
            except Exception:
                print("\n✘ Erro inesperado:")
                traceback.print_exc()
                fila_ui.put(("fim", "erro"))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    def _preparar_andamento(self):
        self.contagem = {"ok": 0, "portal": 0, "erro": 0}
        self.verificacao = {}
        for ct in self.contadores:
            ct.definir(0, animar=False)
        self.etapa_atual = None
        self.etapas_vistas = []
        self._pintar_chips()
        self.barra.configure(progress_color=AZUL)
        self._barra_indeterminada(True)
        self.lbl_titulo.configure(text="Preparando...")
        self.lbl_sub.configure(text="Lendo a planilha.")
        self.caixa_log.configure(state="normal")
        self.caixa_log.delete("1.0", "end")
        self.caixa_log.configure(state="disabled")
        self.btn_faltou.pack_forget()
        self.btn_parar.configure(state="normal")
        self.btn_parar.pack(side="left")
        self._travar_tela(True)
        self.inicio_execucao = time.time()
        self.lbl_relogio.configure(text="⏱ 00:00")
        self._mostrar_andamento()

    def _travar_tela(self, travar):
        estado = "disabled" if travar else "normal"
        for b in self.botoes_execucao:
            b.configure(state=estado)
        self.passo_planilha.travar(travar)
        self.passo_pasta.travar(travar)
        if travar:
            self.btn_principal.configure(text="Trabalhando...", fg_color=AZUL_DESATIVADO)
        else:
            self.btn_principal.configure(text="Baixar todas as notas", fg_color=AZUL)

    def _parar(self):
        if self.thread and self.thread.is_alive() and DialogoConfirmar(
                self, "Parar agora?", "As notas que já foram baixadas continuam salvas na pasta.",
                sim="Parar", nao="Continuar").resultado:
            interromper_thread(self.thread)

    def _ao_fechar(self):
        if self.thread and self.thread.is_alive():
            if not DialogoConfirmar(self, "A automação ainda está rodando",
                                    "Se fechar agora, ela para. As notas já baixadas continuam salvas.",
                                    sim="Fechar mesmo assim", nao="Voltar").resultado:
                return
            interromper_thread(self.thread)
        self.destroy()

    def _finalizar(self, como):
        self._barra_indeterminada(False)
        self.btn_parar.pack_forget()
        self._travar_tela(False)
        duracao = formatar_tempo(time.time() - self.inicio_execucao) if self.inicio_execucao else ""
        self.inicio_execucao = None
        self.lbl_relogio.configure(text=f"⏱ {duracao}" if duracao else "")

        if como == "ok":
            for chave in self.etapas_vistas:
                self.chips[chave].configure(fg_color=VERDE_CLARO, text_color=VERDE)
            self._resumo_final()
        elif como == "parado":
            self.lbl_titulo.configure(text="Parado")
            self.lbl_sub.configure(text="As notas que já foram baixadas continuam na pasta. "
                                        "Use \"Baixar o que faltou\" para continuar depois.")
        else:
            self.barra.configure(progress_color=VERMELHO)
            self.lbl_titulo.configure(text="Algo deu errado")
            self.lbl_sub.configure(text="Abra \"Ver detalhes técnicos\" e tire um print para quem cuida da automação.")

        # chama a atenção quando termina, mesmo se a janela estiver atrás de outras
        self.bell()
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)
        self.after(400, lambda: self.attributes("-topmost", False))

    def _resumo_final(self):
        v = self.verificacao
        if {"faltando", "corrompido"} <= v.keys():
            faltam = v["faltando"] + v["corrompido"]
            total = v.get("total", 0)
            # a conferência é a palavra final: ajusta os contadores para bater com a pasta
            self.ct_ok.definir(v.get("ok", self.contagem["ok"]))
            self.ct_portal.definir(0)
            self.ct_erro.definir(faltam)
            if faltam == 0:
                self._concluido("Tudo pronto!  ✓", f"As {total} notas da planilha estão na pasta.")
            else:
                self._incompleto(faltam, (total - faltam) / total if total else 0)
            return

        ok, erro = self.contagem["ok"], self.contagem["erro"]
        if erro:
            self._incompleto(erro, ok / (ok + erro))
        else:
            self._concluido("Pronto!  ✓", f"{ok} nota(s) baixada(s).")

    def _concluido(self, titulo, sub):
        self.barra.configure(progress_color=VERDE)
        self.alvo_barra = 1.0
        self.lbl_titulo.configure(text=titulo)
        self.lbl_sub.configure(text=sub)

    def _incompleto(self, faltam, fracao_pronta):
        self.barra.configure(progress_color=AMARELO)
        self.alvo_barra = fracao_pronta
        self.lbl_titulo.configure(text=f"Quase lá — falta(m) {faltam} nota(s)")
        self.lbl_sub.configure(text="Elas não saíram nem pelo SAP nem pelo portal. Tente de novo, "
                                    "ou veja o motivo de cada uma no relatório.")
        self.btn_faltou.pack(side="left")

    # ---------- animação e relógio ----------
    def _barra_indeterminada(self, ligar):
        if ligar:
            self.barra.configure(mode="indeterminate")
            self.barra.start()
        else:
            self.barra.stop()
            self.barra.configure(mode="determinate")
            self.barra.set(self.alvo_barra)

    def _animar(self):
        if self.barra.cget("mode") == "determinate":
            atual = self.barra.get()
            if abs(self.alvo_barra - atual) > 0.002:
                self.barra.set(atual + (self.alvo_barra - atual) * 0.18)
        for ct in self.contadores:
            ct.passo_animacao()
        self.after(16, self._animar)

    def _atualizar_relogio(self):
        if self.inicio_execucao:
            self.lbl_relogio.configure(text=f"⏱ {formatar_tempo(time.time() - self.inicio_execucao)}")
        self.after(1000, self._atualizar_relogio)

    def _pintar_chips(self):
        for chave, chip in self.chips.items():
            if chave == self.etapa_atual:
                chip.configure(fg_color=AZUL_CLARO, text_color=AZUL)
            elif chave in self.etapas_vistas:
                chip.configure(fg_color=VERDE_CLARO, text_color=VERDE)
            else:
                chip.configure(fg_color=CINZA_CLARO, text_color=SUAVE)

    # ---------- mensagens do motor ----------
    def _processar_fila(self):
        try:
            for _ in range(500):          # não trava a tela se chegar muita mensagem de uma vez
                tipo, dado = fila_ui.get_nowait()
                if tipo == "log":
                    self._receber_log(dado)
                elif tipo == "etapa":
                    self._mudar_etapa(dado)
                elif tipo == "input":
                    self._responder_pergunta(dado)
                elif tipo == "fim":
                    self._finalizar(dado)
        except queue.Empty:
            pass
        self.after(60, self._processar_fila)

    def _mudar_etapa(self, chave):
        self.etapa_atual = chave
        if chave not in self.etapas_vistas:
            self.etapas_vistas.append(chave)
        self._pintar_chips()
        titulo, sub = TEXTOS_ETAPA[chave]
        self.lbl_titulo.configure(text=titulo)
        self.lbl_sub.configure(text=sub)
        self.alvo_barra = 0.0
        if chave == "conferencia":
            self._barra_indeterminada(True)
        else:
            self._barra_indeterminada(False)
            self.barra.set(0)

    def _escrever_log(self, texto):
        self.caixa_log.configure(state="normal")
        for linha in texto.splitlines(keepends=True):
            if "✔" in linha or "[OK]" in linha:
                tag = "ok"
            elif "✘" in linha or "[FALHOU]" in linha or "ERRO" in linha:
                tag = "erro"
            elif "→ PORTAL" in linha:
                tag = "portal"
            elif linha.startswith(">>>"):
                tag = "titulo"
            else:
                tag = None
            self.caixa_log.insert("end", linha, tag)
        self.caixa_log.see("end")
        self.caixa_log.configure(state="disabled")

    def _receber_log(self, texto):
        self._escrever_log(texto)
        c = self.contagem

        for linha in texto.splitlines():
            if linha.strip():
                ultimas_linhas.append(linha)
            if self.etapa_atual == "sap":
                if "→ ✔" in linha:
                    c["ok"] += 1
                elif "→ PORTAL" in linha:
                    c["portal"] += 1
            elif self.etapa_atual == "portal":
                if "[OK]" in linha:
                    c["ok"] += 1
                    c["portal"] -= 1
                elif "[FALHOU]" in linha:
                    c["erro"] += 1
                    c["portal"] -= 1
            m = RE_PENDENTES_PORTAL.search(linha)
            if m:
                c["portal"] = int(m[1])
            for chave, regex in RE_VERIF.items():
                m = regex.search(linha)
                if m:
                    self.verificacao[chave] = int(m[1])

        self.ct_ok.definir(c["ok"])
        self.ct_portal.definir(c["portal"])
        self.ct_erro.definir(c["erro"])

        if self.etapa_atual in ("sap", "portal"):
            contadores = RE_CONTADOR.findall(texto)
            if contadores:
                feitas, total = map(int, contadores[-1])
                if total:
                    self.alvo_barra = feitas / total
                    onde = "no SAP" if self.etapa_atual == "sap" else "no portal"
                    self.lbl_titulo.configure(text=f"{feitas} de {total} notas  ·  {onde}")
        termino = RE_TERMINO.search(texto)
        if termino and self.etapa_atual == "sap":
            self.lbl_sub.configure(text=f"Termina por volta das {termino[1]}. Não mexa no SAP enquanto isso.")

    def _responder_pergunta(self, pedido):
        """Transforma cada input() do motor numa janela simples."""
        prompt = pedido["prompt"].strip()
        texto = prompt.lower()
        if "(s/n)" in texto:
            numero = re.search(r"(\d+)", prompt)
            if self.confirmar_auto:
                pedido["resposta"] = "s"     # a pessoa já clicou em "Baixar o que faltou"
            else:
                quantas = f"{numero[1]} nota(s)" if numero else "Algumas notas"
                ok = DialogoConfirmar(self, f"{quantas} faltando", "Quer baixar agora só as que faltam?",
                                      sim="Baixar agora", nao="Agora não").resultado
                pedido["resposta"] = "s" if ok else "n"
        elif "enter para começar" in texto:
            ok = DialogoConfirmar(self, "Antes de começar", "O SAP está aberto e logado?",
                                  sim="Sim, começar", nao="Cancelar").resultado
            pedido["resposta"] = "" if ok else None
        elif "pressione enter" in texto:
            ok = DialogoConfirmar(self, "Login no portal",
                                  "Faça o login no portal, na janela do navegador, e depois clique em Continuar.",
                                  sim="Continuar", nao="Cancelar").resultado
            pedido["resposta"] = "" if ok else None
        elif "nota ocorrência" in texto or "nota ocorrencia" in texto:
            pedido["resposta"] = DialogoColunas(self, list(ultimas_linhas)).resultado
        else:
            pedido["resposta"] = DialogoTexto(self, prompt).resultado
        pedido["evento"].set()


# =====================================================================
#                              JANELAS
# =====================================================================
class DialogoBase(ctk.CTkToplevel):
    def __init__(self, pai, titulo, largura, altura):
        super().__init__(pai, fg_color=FUNDO)
        self.title(titulo)
        self.resizable(False, False)
        self.resultado = None
        self.transient(pai)
        # o customtkinter troca o ícone das janelinhas logo depois de abrir; reaplica o nosso
        self.after(250, lambda: aplicar_icone(self))
        pai.update_idletasks()
        x = pai.winfo_rootx() + (pai.winfo_width() - largura) // 2
        y = pai.winfo_rooty() + max(40, (pai.winfo_height() - altura) // 3)
        self.geometry(f"{largura}x{altura}+{max(0, x)}+{max(0, y)}")
        self.bind("<Escape>", lambda _: self.destroy())

    def _esperar(self):
        self.lift()
        self.after(50, self.grab_set)
        self.wait_window()


class DialogoConfirmar(DialogoBase):
    """Pergunta de sim/não no mesmo visual da tela."""
    def __init__(self, pai, titulo, texto, sim="OK", nao="Cancelar", largura=440, altura=200):
        super().__init__(pai, titulo, largura, altura)
        ctk.CTkLabel(self, text=titulo, font=fonte(17, "bold"), text_color=TEXTO,
                     anchor="w").pack(fill="x", padx=24, pady=(22, 6))
        ctk.CTkLabel(self, text=texto, font=fonte(13), text_color=SUAVE, justify="left",
                     anchor="w", wraplength=largura - 48).pack(fill="x", padx=24)
        botoes = ctk.CTkFrame(self, fg_color="transparent")
        botoes.pack(side="bottom", fill="x", padx=24, pady=20)
        btn_sim = botao_primario(botoes, sim, self._sim, width=max(110, 9 * len(sim)))
        btn_sim.pack(side="right")
        if nao:
            botao_texto(botoes, nao, self.destroy, height=38, width=100).pack(side="right", padx=8)
        self.bind("<Return>", lambda _: self._sim())
        self.after(100, btn_sim.focus_set)
        self._esperar()

    def _sim(self):
        self.resultado = True
        self.destroy()


class DialogoAviso(DialogoConfirmar):
    def __init__(self, pai, titulo, texto):
        super().__init__(pai, titulo, texto, sim="Entendi", nao=None, altura=190)


class DialogoColunas(DialogoBase):
    """Confirma as colunas que o motor encontrou na planilha."""
    RE_COL = re.compile(r"(Nota|Ocorrência)\s+→ coluna (\w+)\s+\('(.*?)'\)\s+ex\.: (.*)")
    RE_INICIO = re.compile(r"a partir da linha (\d+)")

    def __init__(self, pai, contexto):
        super().__init__(pai, "Confirmar colunas", 540, 330)
        achados, inicio = {}, "2"
        for linha in contexto:
            m = self.RE_COL.search(linha)
            if m:
                achados[m[1]] = (m[2], m[3], m[4].strip())
            m = self.RE_INICIO.search(linha)
            if m:
                inicio = m[1]

        ctk.CTkLabel(self, text="Encontrei estas colunas na planilha", font=fonte(17, "bold"),
                     text_color=TEXTO).pack(anchor="w", padx=24, pady=(22, 2))
        ctk.CTkLabel(self, text="Confira se estão certas antes de começar.", font=fonte(12),
                     text_color=SUAVE).pack(anchor="w", padx=24, pady=(0, 12))

        colunas = {}
        for rotulo in ("Nota", "Ocorrência"):
            col, titulo, exemplo = achados.get(rotulo, ("?", "", ""))
            card = Cartao(self)
            card.pack(fill="x", padx=24, pady=4)
            ctk.CTkLabel(card, text=rotulo, font=fonte(13, "bold"), text_color=TEXTO, width=100,
                         anchor="w").pack(side="left", padx=(14, 4), pady=12)
            ctk.CTkLabel(card, text=f"coluna {col}  ·  “{titulo}”  ·  ex.: {exemplo}", font=fonte(12),
                         text_color=SUAVE, anchor="w").pack(side="left", fill="x", expand=True)
            colunas[rotulo] = col

        self.painel_corrigir = ctk.CTkFrame(self, fg_color="transparent")
        self.ent_nota = self._campo("Coluna da nota", colunas["Nota"])
        self.ent_ocor = self._campo("Coluna da ocorrência", colunas["Ocorrência"])
        self.ent_inicio = self._campo("Linha inicial", inicio)

        botoes = ctk.CTkFrame(self, fg_color="transparent")
        botoes.pack(side="bottom", fill="x", padx=24, pady=20)
        botao_primario(botoes, "Está certo", self._confirmar, width=130).pack(side="right")
        self.btn_corrigir = botao_secundario(botoes, "Corrigir", self._corrigir, width=100)
        self.btn_corrigir.pack(side="right", padx=8)
        botao_texto(botoes, "Cancelar", self.destroy, height=38, width=90).pack(side="left")
        self.bind("<Return>", lambda _: self._confirmar())
        self._esperar()

    def _campo(self, rotulo, valor):
        bloco = ctk.CTkFrame(self.painel_corrigir, fg_color="transparent")
        bloco.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(bloco, text=rotulo, font=fonte(11), text_color=SUAVE).pack(anchor="w")
        e = ctk.CTkEntry(bloco, width=130, height=32, corner_radius=8)
        e.insert(0, valor if valor != "?" else "")
        e.pack()
        return e

    def _corrigir(self):
        self.geometry(f"540x400")
        self.painel_corrigir.pack(fill="x", padx=24, pady=(12, 0))
        self.btn_corrigir.configure(state="disabled")
        self.ent_nota.focus_set()

    def _confirmar(self):
        if self.painel_corrigir.winfo_ismapped():
            self.resultado = f"{self.ent_nota.get()} {self.ent_ocor.get()} {self.ent_inicio.get()}".strip().upper()
        else:
            self.resultado = ""
        self.destroy()


class DialogoTexto(DialogoBase):
    """Para qualquer outra pergunta com resposta digitada."""
    def __init__(self, pai, prompt):
        super().__init__(pai, "Responder", 480, 220)
        ctk.CTkLabel(self, text=prompt, wraplength=430, justify="left", font=fonte(13),
                     text_color=TEXTO).pack(anchor="w", padx=24, pady=(22, 8))
        self.entrada = ctk.CTkEntry(self, height=34, corner_radius=8)
        self.entrada.pack(fill="x", padx=24)
        self.entrada.bind("<Return>", lambda _: self._ok())
        botao_primario(self, "OK", self._ok, width=110).pack(side="bottom", anchor="e", padx=24, pady=18)
        self.after(150, self.entrada.focus_set)
        self._esperar()

    def _ok(self):
        self.resultado = self.entrada.get()
        self.destroy()


class DialogoConfiguracoes(DialogoBase):
    def __init__(self, pai, url):
        super().__init__(pai, "Configurações", 560, 260)
        ctk.CTkLabel(self, text="Configurações", font=fonte(17, "bold"),
                     text_color=TEXTO).pack(anchor="w", padx=24, pady=(22, 10))
        ctk.CTkLabel(self, text="Endereço do portal de notas",
                     font=fonte(12, "bold"), text_color=TEXTO).pack(anchor="w", padx=24)
        ctk.CTkLabel(self, text="Já vem preenchido. Só mude se o endereço do portal mudar.",
                     font=fonte(12), text_color=SUAVE).pack(anchor="w", padx=24)
        self.entrada = ctk.CTkEntry(self, height=34, corner_radius=8)
        self.entrada.insert(0, url)
        self.entrada.pack(fill="x", padx=24, pady=(6, 0))
        self.entrada.bind("<Return>", lambda _: self._salvar())
        botoes = ctk.CTkFrame(self, fg_color="transparent")
        botoes.pack(side="bottom", fill="x", padx=24, pady=18)
        botao_primario(botoes, "Salvar", self._salvar, width=110).pack(side="right")
        botao_texto(botoes, "Cancelar", self.destroy, height=38, width=90).pack(side="right", padx=8)
        botao_texto(botoes, "Voltar ao padrão", self._padrao, cor=AZUL, height=38, width=130).pack(side="left")
        self._esperar()

    def _padrao(self):
        self.entrada.delete(0, "end")
        self.entrada.insert(0, URL_PORTAL_PADRAO)

    def _salvar(self):
        self.resultado = limpar_url(self.entrada.get()) or URL_PORTAL_PADRAO
        self.destroy()


if __name__ == "__main__":
    if sys.platform == "win32":
        try:   # ícone próprio na barra de tarefas (em vez do ícone do Python)
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Automacoes.DownloadNotas")
        except Exception:
            pass
    App().mainloop()
