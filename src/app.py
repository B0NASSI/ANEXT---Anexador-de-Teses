# -*- coding: utf-8 -*-
"""
ANEXT - Anexador de Teses
Ponto de entrada do aplicativo. Execute com: python app.py
"""

import io
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import pythoncom
from PIL import Image, ImageTk
import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, BOTTOM, CENTER, DISABLED, E, EW, LEFT, NORMAL, RIGHT, W, X, Y
from ttkbootstrap.widgets.scrolled import ScrolledText

import documento
import juntar_pdfs
import tabela
import tema
import visual
from gerar import gerar_pdf_capas, sanitizar_nome_arquivo
from pdf import ConversorPDF, ConversorPDFIndisponivel
from separar_capas import calcular_destinos, pastas_segurados_em_ordem, separar_capas

LARGURA_JANELA = 1000
ALTURA_BANNER = 116
ALTURA_JANELA = 979  # +meio cm sobre os 960 originais, pra sobrar espaço no Resultado durante o processamento
NOME_APP = "ANEXT - Anexador de Teses"
DESCRICAO_APP = "Gera, divide e junta os documentos de cada segurado da tese em um PDF"

TEXTO_MANUAL = """\
ATALHO: FAÇA TUDO DE UMA VEZ  (aba "All-in-one")

  Faz de uma vez as etapas 0, 2 e 3 abaixo: gera as capas, divide
  pelas subpastas dos segurados e junta tudo num PDF final da tese.

  Antes de usar, organize as pastas dos segurados (etapa 1, abaixo).
  Se elas ainda não existirem quando você gerar, as capas saem soltas
  na pasta da tese e a etapa de juntar é pulada — dá pra organizar
  depois e rodar de novo só pra gerar o PDF final.

  Tudo fica salvo direto na pasta da tese, sem subpasta própria: as
  capas geradas ("Capas Geradas.pdf"), as capas já divididas (uma em
  cada subpasta) e o PDF final da tese inteira.


0. GERE AS CAPAS  (aba "Gerar capas")

  Informe o número do tópico, o título da tese e a tabela de
  segurados (importada de um .docx ou colada do Word). O programa
  gera um PDF com a capa geral, a tabela completa e uma capa por
  segurado — é o arquivo que entra na aba "Dividir capas".

  Requer o Microsoft Word instalado no computador.


1. ORGANIZE AS PASTAS

  Crie uma pasta para cada segurado:

        1. NOME DO SEGURADO
        2. NOME DO SEGURADO
        3. NOME DO SEGURADO

  Dentro de cada pasta, numere os arquivos na ordem em que eles
  devem aparecer no PDF final:

        0. Capa  (já vem da aba "Dividir capas")
        1. Tela FAP
        2. CAT
        3. Laudo INSS
        4. INFBEN
        5. Extrato Previdenciário
        6. Petição Inicial
        7. Laudo da Perícia Judicial
        8. Sentença

  Quer encaixar um documento novo entre dois que já existem? Use
  um número com casa decimal, sem precisar renumerar o resto:

        2. Laudo INSS
        2.1 Documento complementar
        2.2 Novo laudo
        3. INFBEN

  IMPORTANTE: só entram no PDF final os arquivos que começam com
  número (0, 1, 2, 2.1...). Os outros são ignorados.


2. DIVIDIR CAPAS

  Selecione o PDF de capas e a pasta de destino.

  • Se as pastas dos segurados já existirem, cada capa cai sozinha,
    direto na pasta certa.
  • Se ainda não existirem, todas as capas ficam juntas na pasta de
    destino, já numeradas pra manter a ordem:

        0. Capa NOME DO SEGURADO
        0.1 Capa NOME DO SEGURADO
        0.2 Capa NOME DO SEGURADO
        ...


3. JUNTAR PDFS

  Escolha um dos modos:

  • Tese completa: junta todas as pastas numeradas em um só PDF.
  • Pasta única: junta só os arquivos da pasta escolhida.

  Antes de gerar, o programa mostra a ordem que vai usar, pra você
  confirmar.


DICAS

  • Use 2.1, 2.2, 2.3... pra encaixar documentos sem renumerar tudo.
  • Confira a lista antes de gerar. Depois de gerar, use o botão
    "👁 Abrir PDF gerado" pra revisar o resultado antes de protocolar."""


def _caminho_recurso(nome: str) -> str:
    # em modo congelado (.exe), os recursos ficam soltos em _MEIPASS; em modo
    # de desenvolvimento, app.py mora em src/ mas os recursos (modelo/,
    # assets/) ficam na raiz do projeto, um nível acima
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return str(base / nome)


def _pasta_executavel() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


MODELO_CAPAS = Path(_caminho_recurso("modelo")) / "modelo.docx"
SAIDA_PADRAO_CAPAS = _pasta_executavel() / "output"
BASE_TABELAS = Path(_caminho_recurso("modelo")) / "BASE TABELAS.docx"
CONFIG_CAPAS = _pasta_executavel() / "anext_config.json"

# larguras padrão da tabela de capas (percentuais por coluna visual).
# 7 colunas (vigência única): calibrado em render real para CNPJ, NIT e
# benefício caberem em uma linha (Item | Vigências | CNPJ | Empregados |
# NIT | Tipo | Benefício). 8 colunas de grade (vigências mescladas em
# duas sub-colunas): padrão definido pelo usuário.
_PESOS_PADRAO_7 = [6.8, 13, 23, 18.5, 16.5, 7, 15.2]
_PESOS_PADRAO_8 = [5, 5.5, 5.5, 17, 22, 12, 6, 12]
LARGURAS_PADRAO_CAPAS = {
    7: [p / sum(_PESOS_PADRAO_7) for p in _PESOS_PADRAO_7],
    8: [p / sum(_PESOS_PADRAO_8) for p in _PESOS_PADRAO_8],
}


def _carregar_config() -> dict:
    try:
        with open(CONFIG_CAPAS, encoding="utf-8") as arquivo:
            return json.load(arquivo)
    except (OSError, ValueError):
        return {}


def _salvar_config(config: dict) -> None:
    try:
        with open(CONFIG_CAPAS, "w", encoding="utf-8") as arquivo:
            json.dump(config, arquivo, ensure_ascii=False, indent=2)
    except OSError:
        pass

TEXTO_AJUDA_COLAR = (
    "Copie a tabela no Word e cole aqui (Ctrl+V). A formatação original "
    "(mesclagens, cores, bordas e fontes) é capturada automaticamente."
)


def _normalizar_espacos(texto: str) -> str:
    """Colapsa tabs/quebras de linha/espaços repetidos num único espaço."""
    return re.sub(r"\s+", " ", texto).strip()


def _normalizar_travessao(texto: str) -> str:
    """Troca "-" por "–" (travessão) só quando ele separa palavras."""
    return re.sub(r"(?<=\s)-(?=\s)", "–", texto)


def _posicionar_um_pouco_acima_do_centro(root, largura: int, altura: int) -> None:
    root.update_idletasks()
    x = (root.winfo_screenwidth() - largura) // 2
    y = max((root.winfo_screenheight() - altura) // 2 - 50, 0)
    root.geometry(f"{largura}x{altura}+{x}+{y}")


def _posicionar_sobre_janela(referencia, janela, largura: int, altura: int) -> None:
    """Centraliza `janela` sobre `referencia` (a janela principal) — usa a
    posição ATUAL dela na tela, então a caixa acompanha o monitor em que o
    ANEXT estiver, em vez de sempre abrir no monitor primário.

    Se `referencia` estiver minimizada, o Windows reporta a posição dela
    como algo em torno de -32000,-32000 (valor sentinela de janela
    iconificada) — sem esse restore, a caixa nasceria fora da tela, visível
    pro Windows (bloqueando o clique) mas invisível pra pessoa."""
    referencia.update_idletasks()
    if referencia.state() == "iconic":
        referencia.deiconify()
        referencia.update_idletasks()
    x = referencia.winfo_rootx() + (referencia.winfo_width() - largura) // 2
    y = referencia.winfo_rooty() + max((referencia.winfo_height() - altura) // 2 - 50, 0)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


def _plural(n: int, singular: str, plural: str) -> str:
    return f"{n} {singular if n == 1 else plural}"


def _mensagem_erro_amigavel(exc: Exception) -> str:
    texto = str(exc).lower()
    if "permission denied" in texto or "permissão negada" in texto:
        return (
            "Não foi possível salvar o arquivo porque ele (ou um dos PDFs de origem) "
            "está aberto em outro programa (Acrobat, Edge, etc.).\n\n"
            "Feche o arquivo e tente novamente."
        )
    if "fzerror" in texto:
        return (
            f"{exc}\n\n"
            "Isso costuma ser um bloqueio temporário de um arquivo por outro programa "
            "(antivírus, OneDrive sincronizando, etc.). Tente novamente — se persistir, "
            "verifique se a pasta da tese está totalmente sincronizada/baixada."
        )
    return str(exc)


def _carregar_imagem_altura(caminho: str, altura: int) -> ImageTk.PhotoImage:
    imagem = Image.open(caminho).convert("RGBA")
    proporcao = altura / imagem.height
    imagem = imagem.resize((int(imagem.width * proporcao), altura), Image.LANCZOS)
    return ImageTk.PhotoImage(imagem)


class AplicativoDivisorPDF:
    def __init__(self, root: ttk.Window):
        self.root = root
        root.title(NOME_APP)
        root.geometry(f"{LARGURA_JANELA}x{ALTURA_JANELA}")
        root.resizable(True, True)
        root.minsize(720, 560)
        _posicionar_um_pouco_acima_do_centro(root, LARGURA_JANELA, ALTURA_JANELA)
        tema.aplicar(root)
        try:
            root.iconbitmap(_caminho_recurso("assets/pdf.ico"))
        except tk.TclError:
            pass

        self.var_entrada = tk.StringVar()
        self.var_saida = tk.StringVar()
        self.var_origem_juntar = tk.StringVar()
        self.var_saida_juntar = tk.StringVar()
        self.var_modo_juntar = tk.StringVar(value="tese")
        self._ultimo_pdf_juntar: Path | None = None

        self.var_topico_allin = tk.StringVar()
        self.var_titulo_allin = tk.StringVar()
        self.var_origem_tabela_allin = tk.StringVar(value="docx")
        self.var_docx_tabela_allin = tk.StringVar()
        self.var_pasta_tese_allin = tk.StringVar()
        self._pasta_tese_allin_manual = False
        self.var_gerar_docx_allin = tk.BooleanVar(value=False)
        self._ultimo_pdf_allin: Path | None = None

        self.var_topico_capas = tk.StringVar()
        self.var_titulo_capas = tk.StringVar()
        self.var_origem_tabela_capas = tk.StringVar(value="docx")
        self.var_docx_tabela_capas = tk.StringVar()
        # começa vazio (não pré-preenchido com a pasta "output"): agora
        # que o campo é "Pasta da tese", faz mais sentido só preencher
        # quando o usuário importar o .docx ou escolher manualmente
        self.var_saida_capas = tk.StringVar()
        self._saida_capas_manual = False
        self.var_gerar_docx_capas = tk.BooleanVar(value=False)
        self._ultimo_pdf_capas: Path | None = None
        self._tabela_word_capturada = None
        self._captura_em_andamento = False
        self._ignorar_modificacao_colar = False
        self._config = _carregar_config()

        self._montar_banner(root)
        self._montar_barra_superior(root)
        self._montar_rodape(root)
        self._instalar_menu_contexto(root)

        notebook = ttk.Notebook(root)
        notebook.pack(fill=BOTH, expand=True, padx=14, pady=(14, 0))

        aba_allin = ttk.Frame(notebook, padding=(20, 18))
        aba_capas = ttk.Frame(notebook, padding=(20, 18))
        aba_dividir = ttk.Frame(notebook, padding=(20, 18))
        aba_juntar = ttk.Frame(notebook, padding=(20, 18))
        notebook.add(aba_allin, text="  All-in-one  ")
        notebook.add(aba_capas, text="  Gerar capas  ")
        notebook.add(aba_dividir, text="  Dividir capas  ")
        notebook.add(aba_juntar, text="  Juntar PDFs  ")

        aba_allin.columnconfigure(0, weight=1)
        self._montar_card_allin(aba_allin)
        self._montar_acoes_allin(aba_allin)
        self._montar_progresso_allin(aba_allin)
        self._montar_resultado_allin(aba_allin)

        aba_capas.columnconfigure(0, weight=1)
        self._montar_card_capas(aba_capas)
        self._montar_acoes_capas(aba_capas)
        self._montar_progresso_capas(aba_capas)
        self._montar_resultado_capas(aba_capas)

        aba_dividir.columnconfigure(0, weight=1)
        self._montar_card_arquivos(aba_dividir)
        self._montar_acoes(aba_dividir)
        self._montar_resultado(aba_dividir)

        aba_juntar.columnconfigure(0, weight=1)
        self._montar_card_juntar(aba_juntar)
        self._montar_acoes_juntar(aba_juntar)
        self._montar_progresso_juntar(aba_juntar)
        self._montar_resultado_juntar(aba_juntar)

    def _instalar_menu_contexto(self, root) -> None:
        """Menu de botão direito (Recortar/Copiar/Colar/Selecionar tudo) em
        todos os campos de texto do aplicativo. Instalado por classe de
        widget, vale também para os campos de janelas e diálogos abertos
        depois."""
        menu = tk.Menu(root, tearoff=0)

        def mostrar(event):
            w = event.widget
            try:
                if str(w.cget("state")) == "disabled":
                    return
                w.focus_set()
            except Exception:
                return

            try:
                if isinstance(w, tk.Text):
                    tem_selecao = bool(w.tag_ranges("sel"))
                else:
                    tem_selecao = w.selection_present()
            except Exception:
                tem_selecao = False
            editavel = str(w.cget("state")) == "normal"

            menu.delete(0, "end")
            menu.add_command(
                label="Recortar", accelerator="Ctrl+X",
                state="normal" if tem_selecao and editavel else "disabled",
                command=lambda: w.event_generate("<<Cut>>"),
            )
            menu.add_command(
                label="Copiar", accelerator="Ctrl+C",
                state="normal" if tem_selecao else "disabled",
                command=lambda: w.event_generate("<<Copy>>"),
            )
            menu.add_command(
                label="Colar", accelerator="Ctrl+V",
                state="normal" if editavel else "disabled",
                command=lambda: w.event_generate("<<Paste>>"),
            )
            menu.add_separator()
            menu.add_command(
                label="Selecionar tudo", accelerator="Ctrl+A",
                command=lambda: w.event_generate("<<SelectAll>>"),
            )
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
            return "break"

        for classe in ("TEntry", "Entry", "Text", "TSpinbox", "TCombobox"):
            root.bind_class(classe, "<Button-3>", mostrar)

    def _montar_banner(self, root):
        self._label_banner = tk.Label(root, borderwidth=0, anchor="w")
        self._label_banner.pack(fill=X)
        self._largura_banner_atual = 0
        self._redesenho_banner_pendente = None
        self._atualizar_banner(LARGURA_JANELA)
        # com a janela redimensionável, o banner é redesenhado na nova
        # largura (com atraso curto, para não redesenhar a cada pixel
        # durante o arraste da borda)
        root.bind("<Configure>", self._ao_redimensionar_janela, add="+")

    def _atualizar_banner(self, largura: int) -> None:
        largura = max(int(largura), 400)
        if largura == self._largura_banner_atual:
            return
        self._largura_banner_atual = largura
        imagem = visual.gerar_banner(
            largura=largura,
            altura=ALTURA_BANNER,
            cor_inicio=tema.COR_PRIMARIA,
            cor_fim="#1A1C3D",
            cor_destaque=tema.COR_SECUNDARIA,
            icone_path=_caminho_recurso("assets/folder interno.ico"),
            titulo=NOME_APP,
            subtitulo=DESCRICAO_APP,
        )
        self._imagem_banner = ImageTk.PhotoImage(imagem)
        self._label_banner.configure(image=self._imagem_banner)

    def _ao_redimensionar_janela(self, event) -> None:
        if event.widget is not self.root or event.width == self._largura_banner_atual:
            return
        if self._redesenho_banner_pendente is not None:
            self.root.after_cancel(self._redesenho_banner_pendente)
        self._redesenho_banner_pendente = self.root.after(
            120, lambda: self._atualizar_banner(self.root.winfo_width())
        )

    def _montar_barra_superior(self, root):
        barra = ttk.Frame(root, padding=(14, 6, 14, 0))
        barra.pack(fill=X)
        ttk.Button(
            barra, text="❓ Manual rápido", command=self._abrir_manual, bootstyle="primary-link",
        ).pack(side=RIGHT)

    def _montar_card_arquivos(self, pai):
        cartao = ttk.Labelframe(pai, text=" Arquivos ", padding=18, bootstyle="secondary")
        cartao.pack(fill=X, pady=(0, 18))
        cartao.columnconfigure(0, weight=1)

        ttk.Label(cartao, text="Arquivo PDF", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=W)
        linha1 = ttk.Frame(cartao)
        linha1.grid(row=1, column=0, sticky=EW, pady=(6, 16))
        linha1.columnconfigure(0, weight=1)
        ttk.Entry(linha1, textvariable=self.var_entrada).grid(row=0, column=0, sticky=EW, padx=(0, 8))
        ttk.Button(linha1, text="Procurar...", command=self._escolher_entrada, bootstyle="primary-outline").grid(row=0, column=1)

        ttk.Label(cartao, text="Pasta de salvamento", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky=W)
        linha2 = ttk.Frame(cartao)
        linha2.grid(row=3, column=0, sticky=EW, pady=(6, 0))
        linha2.columnconfigure(0, weight=1)
        ttk.Entry(linha2, textvariable=self.var_saida).grid(row=0, column=0, sticky=EW, padx=(0, 8))
        ttk.Button(linha2, text="Procurar...", command=self._escolher_saida, bootstyle="primary-outline").grid(row=0, column=1)

    def _montar_acoes(self, pai):
        acoes = ttk.Frame(pai)
        acoes.pack(fill=X, pady=(0, 18))
        self.botao_abrir_pasta = ttk.Button(
            acoes, text="📂  Abrir pasta de salvamento", command=self._abrir_pasta_saida, bootstyle="primary-outline",
            state=DISABLED,
        )
        self.botao_abrir_pasta.pack(side=LEFT)
        ttk.Button(acoes, text="🧹  Limpar", command=self._limpar, bootstyle="danger-outline").pack(side=LEFT, padx=(10, 0))
        self.botao_dividir = ttk.Button(
            acoes, text="✂  Dividir PDF", command=self._dividir, bootstyle="secondary", width=18
        )
        self.botao_dividir.pack(side=RIGHT)

    def _montar_resultado(self, pai):
        moldura = ttk.Labelframe(pai, text=" Resultado ", padding=14)
        moldura.pack(fill=BOTH, expand=True)

        self.log = ScrolledText(moldura, autohide=True, bootstyle="secondary")
        self.log.pack(fill=BOTH, expand=True)
        self.log.text.configure(font=("Segoe UI", 10), padx=8, pady=8, relief="flat")
        self.log.text.tag_configure("titulo", font=("Segoe UI", 10, "bold"))
        self.log.text.tag_configure("item", foreground=tema.COR_SUCESSO)
        self.log.text.tag_configure("erro", foreground=tema.COR_ERRO, font=("Segoe UI", 10, "bold"))
        self._mostrar_mensagem_inicial()

    def _mostrar_mensagem_inicial(self):
        self.log.text.insert("end", "Selecione o arquivo PDF e clique em \"Dividir PDF\".")
        self.log.text.configure(state=DISABLED)

    def _limpar(self):
        self.var_entrada.set("")
        self.var_saida.set("")
        self.botao_abrir_pasta.config(state=DISABLED)
        self.log.text.configure(state=NORMAL)
        self.log.text.delete("1.0", "end")
        self._mostrar_mensagem_inicial()

    def _montar_rodape(self, root):
        rodape = ttk.Frame(root, padding=(14, 11, 10, 8))
        rodape.pack(fill=X, side=BOTTOM)

        self._imagem_logo = _carregar_imagem_altura(_caminho_recurso("assets/Logo RS completa colorida.png"), 24)
        tk.Label(rodape, image=self._imagem_logo, borderwidth=0, background=tema.COR_FUNDO).pack(side=LEFT)

        ttk.Label(rodape, text="versão 3.3", bootstyle="secondary", font=("Segoe UI", 8)).pack(side=RIGHT)

    def _abrir_manual(self):
        janela = ttk.Toplevel(self.root)
        janela.title("Manual rápido")
        janela.resizable(False, False)
        _posicionar_sobre_janela(self.root, janela, 600, 560)
        try:
            janela.iconbitmap(_caminho_recurso("assets/pdf.ico"))
        except tk.TclError:
            pass

        ttk.Label(janela, text="📖 Manual rápido", font=("Segoe UI", 14, "bold")).pack(anchor=W, padx=20, pady=(18, 2))
        ttk.Label(
            janela, text="Como organizar as pastas antes de dividir e juntar os PDFs", bootstyle="secondary",
        ).pack(anchor=W, padx=20, pady=(0, 12))

        corpo = ScrolledText(janela, autohide=True, bootstyle="secondary")
        corpo.pack(fill=BOTH, expand=True, padx=20, pady=(0, 14))
        corpo.text.insert("end", TEXTO_MANUAL)
        corpo.text.configure(font=("Consolas", 9), padx=10, pady=10, state=DISABLED)

        ttk.Button(janela, text="Fechar", command=janela.destroy, bootstyle="secondary").pack(pady=(0, 18))

    def _escolher_entrada(self):
        caminho = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("PDF", "*.pdf")], parent=self.root)
        if not caminho:
            return
        self.var_entrada.set(caminho)
        if not self.var_saida.get():
            self.var_saida.set(str(Path(caminho).parent))

    def _escolher_saida(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta de salvamento", parent=self.root)
        if caminho:
            self.var_saida.set(caminho)

    def _escrever_log(self, titulo: str, itens: list[str], erro: bool = False) -> None:
        self.log.text.configure(state=NORMAL)
        self.log.text.delete("1.0", "end")
        self.log.text.insert("end", titulo + "\n", "erro" if erro else "titulo")
        for item in itens:
            self.log.text.insert("end", f"  {item}\n", "erro" if erro else "item")
        self.log.text.configure(state=DISABLED)

    def _confirmar_sobrescricao_dividir(self, entrada: Path, saida: Path) -> bool:
        try:
            destinos = calcular_destinos(entrada, saida)
        except Exception:
            return True  # deixa o erro de verdade aparecer na hora de processar

        existentes = [destino for destino in destinos if destino.exists()]
        if not existentes:
            return True

        lista = "\n".join(f"  • {destino.name}" for destino in existentes[:10])
        if len(existentes) > 10:
            lista += f"\n  ... e mais {len(existentes) - 10}"
        verbo = "já existe e será SUBSTITUÍDO" if len(existentes) == 1 else "já existem e serão SUBSTITUÍDOS"

        return messagebox.askyesno(
            "Arquivos serão substituídos",
            f"{_plural(len(existentes), 'arquivo', 'arquivos')} {verbo}:\n\n{lista}\n\nContinuar?",
            icon="warning",
            parent=self.root,
        )

    def _dividir(self):
        entrada = self.var_entrada.get().strip()
        saida = self.var_saida.get().strip()

        if not entrada:
            messagebox.showwarning("Campo obrigatório", "Selecione o arquivo PDF antes de continuar.", parent=self.root)
            return
        if not Path(entrada).is_file():
            messagebox.showerror("Arquivo não encontrado", f"Não foi possível encontrar o arquivo:\n{entrada}", parent=self.root)
            return
        if not saida:
            messagebox.showwarning("Campo obrigatório", "Selecione a pasta de salvamento antes de continuar.", parent=self.root)
            return

        if not self._confirmar_sobrescricao_dividir(Path(entrada), Path(saida)):
            return

        self.botao_dividir.config(state=DISABLED)
        self._escrever_log("Dividindo PDF, aguarde...", [])
        threading.Thread(
            target=self._dividir_worker, args=(Path(entrada), Path(saida)), daemon=True,
        ).start()

    def _dividir_worker(self, entrada: Path, saida: Path):
        try:
            arquivos = separar_capas(entrada, saida)
        except Exception as exc:
            self.root.after(0, self._dividir_falhou, _mensagem_erro_amigavel(exc))
            return
        self.root.after(0, self._dividir_concluiu, arquivos, saida)

    def _dividir_falhou(self, mensagem: str):
        self.botao_dividir.config(state=NORMAL)
        self._escrever_log("Erro ao dividir o PDF:", [mensagem], erro=True)
        messagebox.showerror("Erro ao dividir o PDF", mensagem, parent=self.root)

    def _dividir_concluiu(self, arquivos: list[Path], saida: Path):
        self.botao_dividir.config(state=NORMAL)
        titulo = f"{len(arquivos)} arquivo(s) gerado(s) em:\n  {saida}"
        itens = [f"✓ {arquivo.relative_to(saida)}" for arquivo in arquivos]
        self._escrever_log(titulo, itens)
        self.botao_abrir_pasta.config(state=NORMAL)

    def _abrir_pasta_saida(self):
        saida = self.var_saida.get().strip()
        if saida and Path(saida).is_dir():
            os.startfile(saida)

    def _montar_card_juntar(self, pai):
        cartao = ttk.Labelframe(pai, text=" Pastas ", padding=18, bootstyle="secondary")
        cartao.pack(fill=X, pady=(0, 18))
        cartao.columnconfigure(0, weight=1)

        modos = ttk.Frame(cartao)
        modos.grid(row=0, column=0, columnspan=2, sticky=W, pady=(0, 14))
        ttk.Radiobutton(
            modos, text="Tese completa (todos os segurados)", variable=self.var_modo_juntar, value="tese",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(0, 16))
        ttk.Radiobutton(
            modos, text="Pasta única", variable=self.var_modo_juntar, value="unica", bootstyle="secondary"
        ).pack(side=LEFT)

        ttk.Label(cartao, text="Pasta de origem", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=W)
        linha1 = ttk.Frame(cartao)
        linha1.grid(row=2, column=0, sticky=EW, pady=(6, 16))
        linha1.columnconfigure(0, weight=1)
        ttk.Entry(linha1, textvariable=self.var_origem_juntar).grid(row=0, column=0, sticky=EW, padx=(0, 8))
        ttk.Button(linha1, text="Procurar...", command=self._escolher_origem_juntar, bootstyle="primary-outline").grid(row=0, column=1)

        ttk.Label(cartao, text="Pasta de salvamento", font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky=W)
        linha2 = ttk.Frame(cartao)
        linha2.grid(row=4, column=0, sticky=EW, pady=(6, 0))
        linha2.columnconfigure(0, weight=1)
        ttk.Entry(linha2, textvariable=self.var_saida_juntar).grid(row=0, column=0, sticky=EW, padx=(0, 8))
        ttk.Button(linha2, text="Procurar...", command=self._escolher_saida_juntar, bootstyle="primary-outline").grid(row=0, column=1)

    def _montar_acoes_juntar(self, pai):
        acoes = ttk.Frame(pai)
        acoes.pack(fill=X, pady=(0, 18))
        self.botao_abrir_pasta_juntar = ttk.Button(
            acoes, text="📂  Abrir pasta de salvamento", command=self._abrir_pasta_juntar, bootstyle="primary-outline",
            state=DISABLED,
        )
        self.botao_abrir_pasta_juntar.pack(side=LEFT)
        self.botao_abrir_pdf_juntar = ttk.Button(
            acoes, text="👁  Abrir PDF gerado", command=self._abrir_pdf_juntar, bootstyle="primary-outline",
            state=DISABLED,
        )
        self.botao_abrir_pdf_juntar.pack(side=LEFT, padx=(10, 0))
        ttk.Button(acoes, text="🧹  Limpar", command=self._limpar_juntar, bootstyle="danger-outline").pack(side=LEFT, padx=(10, 0))
        self.botao_juntar = ttk.Button(
            acoes, text="🔗  Juntar PDFs", command=self._juntar, bootstyle="secondary", width=18
        )
        self.botao_juntar.pack(side=RIGHT)

    def _montar_progresso_juntar(self, pai):
        self.frame_progresso_juntar = ttk.Frame(pai)
        self.barra_progresso_juntar = ttk.Progressbar(
            self.frame_progresso_juntar, mode="determinate", bootstyle="secondary",
        )
        self.barra_progresso_juntar.pack(fill=X, pady=(0, 4))
        self.label_progresso_juntar = ttk.Label(
            self.frame_progresso_juntar, text="", bootstyle="secondary", font=("Segoe UI", 9),
        )
        self.label_progresso_juntar.pack(anchor=W)

    def _montar_resultado_juntar(self, pai):
        moldura = ttk.Labelframe(pai, text=" Resultado ", padding=14)
        moldura.pack(fill=BOTH, expand=True)
        self.moldura_resultado_juntar = moldura

        self.log_juntar = ScrolledText(moldura, autohide=True, bootstyle="secondary")
        self.log_juntar.pack(fill=BOTH, expand=True)
        self.log_juntar.text.configure(font=("Segoe UI", 10), padx=8, pady=8, relief="flat")
        self.log_juntar.text.tag_configure("titulo", font=("Segoe UI", 10, "bold"))
        self.log_juntar.text.tag_configure("item", foreground=tema.COR_SUCESSO)
        self.log_juntar.text.tag_configure("erro", foreground=tema.COR_ERRO, font=("Segoe UI", 10, "bold"))
        self._mostrar_mensagem_inicial_juntar()

    def _mostrar_mensagem_inicial_juntar(self):
        self.log_juntar.text.insert("end", "Escolha o modo, a pasta de origem e clique em \"Juntar PDFs\".")
        self.log_juntar.text.configure(state=DISABLED)

    def _escolher_origem_juntar(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta de origem", parent=self.root)
        if not caminho:
            return
        self.var_origem_juntar.set(caminho)
        if not self.var_saida_juntar.get():
            self.var_saida_juntar.set(str(Path(caminho).parent / "PDFs juntados"))

    def _escolher_saida_juntar(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta de salvamento", parent=self.root)
        if caminho:
            self.var_saida_juntar.set(caminho)

    def _escrever_log_juntar(self, titulo: str, itens: list[str], erro: bool = False) -> None:
        self.log_juntar.text.configure(state=NORMAL)
        self.log_juntar.text.delete("1.0", "end")
        self.log_juntar.text.insert("end", titulo + "\n", "erro" if erro else "titulo")
        for item in itens:
            self.log_juntar.text.insert("end", f"  {item}\n", "erro" if erro else "item")
        self.log_juntar.text.configure(state=DISABLED)

    def _limpar_juntar(self):
        self.var_origem_juntar.set("")
        self.var_saida_juntar.set("")
        self.botao_abrir_pasta_juntar.config(state=DISABLED)
        self.botao_abrir_pdf_juntar.config(state=DISABLED)
        self._ultimo_pdf_juntar = None
        self.frame_progresso_juntar.pack_forget()
        self.barra_progresso_juntar.configure(value=0)
        self.label_progresso_juntar.configure(text="")
        self.log_juntar.text.configure(state=NORMAL)
        self.log_juntar.text.delete("1.0", "end")
        self._mostrar_mensagem_inicial_juntar()

    def _confirmar_ordem(self, origem: Path, saida: Path, modo: str, nome_pasta_destino: str = "pasta de salvamento") -> bool:
        if modo == "tese":
            itens = juntar_pdfs.subpastas_numeradas(origem)
            titulo = "As pastas serão juntadas nesta ordem"
        else:
            itens = juntar_pdfs.pdfs_em_ordem(origem)
            titulo = "Os arquivos serão juntados nesta ordem"

        if not itens:
            return True  # deixa a função de junção levantar o erro apropriado

        linhas = []
        erros_totais = []
        for posicao, item in enumerate(itens, start=1):
            arquivos_do_item = juntar_pdfs.pdfs_em_ordem(item) if modo == "tese" else [item]
            paginas, erros = juntar_pdfs.validar_e_contar_paginas(arquivos_do_item)
            erros_totais.extend(erros)
            linhas.append((f"{posicao}. {item.name}", _plural(len(arquivos_do_item), "arquivo", "arquivos"), _plural(paginas, "página", "páginas")))

        if erros_totais:
            messagebox.showerror(
                "Arquivo com problema",
                "Não foi possível abrir os seguintes arquivos (corrompidos ou inválidos):\n\n"
                + "\n".join(erros_totais)
                + "\n\nCorrija ou remova esses arquivos antes de continuar.",
                parent=self.root,
            )
            return False

        aviso = ""
        caminho_saida = saida / juntar_pdfs.nome_saida(origem)
        if caminho_saida.exists():
            aviso = f'Já existe um arquivo "{caminho_saida.name}" na {nome_pasta_destino} — ele será SUBSTITUÍDO.'

        return self._mostrar_confirmacao(titulo, linhas, aviso)

    def _mostrar_confirmacao(self, titulo: str, linhas: list[tuple[str, str, str]], aviso: str) -> bool:
        janela = ttk.Toplevel(self.root)
        janela.title("Confirmar antes de juntar")
        janela.resizable(False, False)
        janela.transient(self.root)
        try:
            janela.iconbitmap(_caminho_recurso("assets/pdf.ico"))
        except tk.TclError:
            pass

        resultado = {"confirmado": False}

        ttk.Label(janela, text=titulo, font=("Segoe UI", 13, "bold")).pack(anchor=W, padx=24, pady=(20, 12))

        moldura = ttk.Frame(janela)
        moldura.pack(padx=24, pady=(0, 38))

        linhas_visiveis = max(1, min(len(linhas), 12))
        tabela = ttk.Treeview(moldura, columns=("nome", "arquivos", "paginas"), show="headings", height=linhas_visiveis + 1)
        tabela.heading("nome", text="Nome")
        tabela.heading("arquivos", text="Arquivos")
        tabela.heading("paginas", text="Páginas")
        tabela.column("nome", width=320, stretch=False, anchor=W)
        tabela.column("arquivos", width=100, stretch=False, anchor=CENTER)
        tabela.column("paginas", width=100, stretch=False, anchor=CENTER)
        tabela.pack(side=LEFT)
        for linha in linhas:
            tabela.insert("", "end", values=linha)
        tabela.selection_remove(*tabela.selection())
        tabela.focus("")

        if len(linhas) > linhas_visiveis:
            scroll = ttk.Scrollbar(moldura, orient="vertical", command=tabela.yview)
            tabela.configure(yscrollcommand=scroll.set)
            scroll.pack(side=LEFT, fill=Y)

        if aviso:
            ttk.Label(
                janela, text=f"⚠ {aviso}", bootstyle="danger", font=("Segoe UI", 10, "bold"), wraplength=520,
            ).pack(anchor=W, padx=24, pady=(0, 12))

        rodape = ttk.Frame(janela, padding=24)
        rodape.pack(fill=X)

        def confirmar():
            resultado["confirmado"] = True
            janela.destroy()

        botao_confirmar = ttk.Button(rodape, text="Confirmar", command=confirmar, bootstyle="secondary", width=14)
        botao_confirmar.pack(side=RIGHT)
        ttk.Button(rodape, text="Cancelar", command=janela.destroy, bootstyle="primary-outline", width=14).pack(side=RIGHT, padx=(0, 10))
        botao_confirmar.focus_set()

        janela.update_idletasks()
        largura, altura = janela.winfo_reqwidth(), janela.winfo_reqheight()
        janela.geometry(f"{largura}x{altura}")
        _posicionar_sobre_janela(self.root, janela, largura, altura)
        janela.deiconify()
        janela.lift()
        janela.grab_set()
        janela.wait_window()
        return resultado["confirmado"]

    def _juntar(self):
        origem = self.var_origem_juntar.get().strip()
        saida = self.var_saida_juntar.get().strip()

        if not origem:
            messagebox.showwarning("Campo obrigatório", "Selecione a pasta de origem antes de continuar.", parent=self.root)
            return
        if not Path(origem).is_dir():
            messagebox.showerror("Pasta não encontrada", f"Não foi possível encontrar a pasta:\n{origem}", parent=self.root)
            return
        if not saida:
            messagebox.showwarning("Campo obrigatório", "Selecione a pasta de salvamento antes de continuar.", parent=self.root)
            return

        modo = self.var_modo_juntar.get()
        if not self._confirmar_ordem(Path(origem), Path(saida), modo):
            return

        self.botao_juntar.config(state=DISABLED)
        self._escrever_log_juntar("Juntando PDFs, aguarde...", [])
        self.barra_progresso_juntar.configure(value=0, maximum=100)
        self.label_progresso_juntar.configure(text="Preparando...")
        self.frame_progresso_juntar.pack(fill=X, pady=(0, 18), before=self.moldura_resultado_juntar)
        threading.Thread(
            target=self._juntar_worker, args=(Path(origem), Path(saida), modo), daemon=True,
        ).start()

    def _juntar_worker(self, origem: Path, saida: Path, modo: str):
        def progresso(atual, total):
            self.root.after(0, self._atualizar_progresso_juntar, atual, total)

        try:
            if modo == "tese":
                arquivos = [juntar_pdfs.juntar_tese(origem, saida, progresso_callback=progresso)]
            else:
                arquivos = [juntar_pdfs.juntar_pasta(origem, saida, progresso_callback=progresso)]
        except Exception as exc:
            self.root.after(0, self._juntar_falhou, _mensagem_erro_amigavel(exc))
            return
        self.root.after(0, self._juntar_concluiu, arquivos, saida)

    def _atualizar_progresso_juntar(self, atual: int, total: int):
        self.barra_progresso_juntar.configure(value=atual, maximum=total)
        self.label_progresso_juntar.configure(text=f"Juntando arquivo {atual} de {total}...")

    def _juntar_falhou(self, mensagem: str):
        self.botao_juntar.config(state=NORMAL)
        self.frame_progresso_juntar.pack_forget()
        self._escrever_log_juntar("Erro ao juntar os PDFs:", [mensagem], erro=True)
        messagebox.showerror("Erro ao juntar os PDFs", mensagem, parent=self.root)

    def _juntar_concluiu(self, arquivos: list[Path], saida: Path):
        self.botao_juntar.config(state=NORMAL)
        self.frame_progresso_juntar.pack_forget()
        titulo = f"{len(arquivos)} arquivo(s) gerado(s) em:\n  {saida}"
        itens = [f"✓ {arquivo.relative_to(saida)}" for arquivo in arquivos]
        self._escrever_log_juntar(titulo, itens)
        self.botao_abrir_pasta_juntar.config(state=NORMAL)
        self._ultimo_pdf_juntar = arquivos[-1]
        self.botao_abrir_pdf_juntar.config(state=NORMAL)

    def _abrir_pasta_juntar(self):
        saida = self.var_saida_juntar.get().strip()
        if saida and Path(saida).is_dir():
            os.startfile(saida)

    def _abrir_pdf_juntar(self):
        if self._ultimo_pdf_juntar and self._ultimo_pdf_juntar.is_file():
            os.startfile(self._ultimo_pdf_juntar)

    # =========================================================================
    # ABA ALL-IN-ONE
    #
    # Roda em sequência automática o mesmo processamento das 3 abas
    # seguintes: gera as capas (gerar_pdf_capas), divide o PDF de capas
    # pelas subpastas dos segurados (separar_capas) e junta a tese completa
    # num único PDF (juntar_pdfs.juntar_tese) — reaproveita exatamente as
    # mesmas funções, sem duplicar a lógica de processamento.
    # =========================================================================

    def _montar_card_allin(self, pai):
        # padding/espaçamentos mais enxutos do que nas outras abas: esta é
        # a aba com mais cards empilhados (Tese + Tabela + Pastas), e
        # precisa sobrar espaço pro Resultado na altura padrão da janela
        cartao = ttk.Labelframe(pai, text=" Tese ", padding=12, bootstyle="secondary")
        cartao.pack(fill=X, pady=(0, 10))
        cartao.columnconfigure(1, weight=1)

        ttk.Label(cartao, text="Número do tópico", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=W, padx=(0, 10), pady=(0, 6))
        ttk.Entry(cartao, textvariable=self.var_topico_allin, width=8).grid(row=0, column=1, sticky=W, pady=(0, 6))
        ttk.Label(cartao, text="Título da tese", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=W, padx=(0, 10))
        ttk.Entry(cartao, textvariable=self.var_titulo_allin).grid(row=1, column=1, sticky=EW)

        tabela_card = ttk.Labelframe(pai, text=" Tabela de segurados ", padding=12, bootstyle="secondary")
        tabela_card.pack(fill=X, pady=(0, 10))
        tabela_card.columnconfigure(0, weight=1)

        linha_radios = ttk.Frame(tabela_card)
        linha_radios.grid(row=1, column=0, sticky=W, pady=(0, 6))
        ttk.Radiobutton(
            linha_radios, text="Importar de um .docx", variable=self.var_origem_tabela_allin, value="docx",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(0, 16))
        ttk.Radiobutton(
            linha_radios, text="Colar texto  (em breve)", variable=self.var_origem_tabela_allin, value="colar",
            bootstyle="secondary", state=DISABLED,
        ).pack(side=LEFT)

        frame_docx = ttk.Frame(tabela_card)
        frame_docx.grid(row=2, column=0, sticky=EW)
        frame_docx.columnconfigure(1, weight=1)
        ttk.Button(
            frame_docx, text="📂  Procurar...", command=self._escolher_tabela_docx_allin,
            bootstyle="primary", width=14,
        ).grid(row=0, column=0, padx=(0, 10))
        ttk.Entry(frame_docx, textvariable=self.var_docx_tabela_allin).grid(row=0, column=1, sticky=EW)
        ttk.Label(
            frame_docx,
            text="⚠️ Se a tabela ocupar mais de uma página, desative a opção \"Repetir "
                 "linhas de cabeçalho\" — caso contrário, o programa não conseguirá ler "
                 "os dados corretamente.",
            bootstyle="warning", font=("Segoe UI", 8, "bold"), wraplength=760, justify=LEFT,
        ).grid(row=1, column=0, columnspan=2, sticky=W, pady=(6, 0))

        linha_base = ttk.Frame(tabela_card)
        linha_base.grid(row=0, column=0, sticky=W, pady=(0, 8))
        ttk.Button(
            linha_base, text="📥  Baixar documento base das tabelas",
            command=self._baixar_base_tabelas, bootstyle="primary-outline",
        ).pack(side=LEFT)
        ttk.Label(
            linha_base, text="Documento em branco com as margens corretas — monte a sua tabela nele antes de importar.",
            bootstyle="secondary", font=("Segoe UI", 8),
        ).pack(side=LEFT, padx=(10, 0))

        pastas_card = ttk.Labelframe(pai, text=" Pastas ", padding=12, bootstyle="secondary")
        pastas_card.pack(fill=X)
        pastas_card.columnconfigure(0, weight=1)

        ttk.Label(pastas_card, text="Pasta da tese", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=W)
        ttk.Label(
            pastas_card, text="Pasta mãe com as subpastas dos segurados (ex.: \"1. NOME\", \"2. NOME\"...).",
            bootstyle="secondary", font=("Segoe UI", 8),
        ).grid(row=1, column=0, sticky=W, pady=(0, 4))
        linha_tese = ttk.Frame(pastas_card)
        linha_tese.grid(row=2, column=0, sticky=EW)
        linha_tese.columnconfigure(0, weight=1)
        ttk.Entry(linha_tese, textvariable=self.var_pasta_tese_allin).grid(row=0, column=0, sticky=EW, padx=(0, 8))
        ttk.Button(linha_tese, text="Procurar...", command=self._escolher_pasta_tese_allin, bootstyle="primary-outline").grid(row=0, column=1)

        ttk.Checkbutton(
            pastas_card, text="Gerar também um arquivo Word (.docx) das capas", variable=self.var_gerar_docx_allin,
            bootstyle="secondary",
        ).grid(row=3, column=0, sticky=W, pady=(10, 0))

    def _montar_acoes_allin(self, pai):
        acoes = ttk.Frame(pai)
        acoes.pack(fill=X, pady=(10, 10))
        self.botao_abrir_pasta_allin = ttk.Button(
            acoes, text="📂  Abrir pasta da tese", command=self._abrir_pasta_allin,
            bootstyle="primary-outline", state=DISABLED,
        )
        self.botao_abrir_pasta_allin.pack(side=LEFT)
        self.botao_abrir_pdf_allin = ttk.Button(
            acoes, text="👁  Abrir PDF gerado", command=self._abrir_pdf_allin,
            bootstyle="primary-outline", state=DISABLED,
        )
        self.botao_abrir_pdf_allin.pack(side=LEFT, padx=(10, 0))
        ttk.Button(acoes, text="🧹  Limpar", command=self._limpar_allin, bootstyle="danger-outline").pack(side=LEFT, padx=(10, 0))
        self.botao_gerar_allin = ttk.Button(
            acoes, text="🚀  Gerar anexo completo", command=self._gerar_allin, bootstyle="secondary",
        )
        self.botao_gerar_allin.pack(side=RIGHT)

    def _montar_progresso_allin(self, pai):
        self._cancelar_allin = threading.Event()
        self.frame_progresso_allin = ttk.Frame(pai)
        self.barra_progresso_allin = ttk.Progressbar(
            self.frame_progresso_allin, mode="determinate", bootstyle="secondary",
        )
        self.barra_progresso_allin.pack(fill=X, pady=(0, 4))
        linha_status = ttk.Frame(self.frame_progresso_allin)
        linha_status.pack(fill=X)
        self.label_progresso_allin = ttk.Label(
            linha_status, text="", bootstyle="secondary", font=("Segoe UI", 9),
        )
        self.label_progresso_allin.pack(side=LEFT, anchor=W)
        # só aparece (é empacotado) enquanto o processamento está rodando —
        # ver _gerar_allin / _restaurar_botao_allin
        self.botao_suspender_allin = ttk.Button(
            linha_status, text="⛔  Suspender", command=self._suspender_allin, bootstyle="danger",
        )

    def _montar_resultado_allin(self, pai):
        moldura = ttk.Labelframe(pai, text=" Resultado ", padding=14)
        moldura.pack(fill=BOTH, expand=True)
        self.moldura_resultado_allin = moldura

        self.log_allin = ScrolledText(moldura, autohide=True, bootstyle="secondary", height=6)
        self.log_allin.pack(fill=BOTH, expand=True)
        self.log_allin.text.configure(font=("Segoe UI", 10), padx=8, pady=8, relief="flat")
        self.log_allin.text.tag_configure("titulo", font=("Segoe UI", 10, "bold"))
        self.log_allin.text.tag_configure("item", foreground=tema.COR_SUCESSO)
        self.log_allin.text.tag_configure("erro", foreground=tema.COR_ERRO, font=("Segoe UI", 10, "bold"))
        self.log_allin.text.tag_configure("aviso", foreground="#B05000")
        self._mostrar_inicial_allin()

    def _mostrar_inicial_allin(self):
        self.log_allin.text.insert(
            "end", "Preencha a tese, a tabela de segurados e as pastas, e clique em \"Gerar anexo completo\".",
        )
        self.log_allin.text.configure(state=DISABLED)

    def _log_allin(self, texto: str, tag: str | None = None):
        self.log_allin.text.configure(state=NORMAL)
        self.log_allin.text.insert("end", texto + "\n", tag or ())
        self.log_allin.text.see("end")
        self.log_allin.text.configure(state=DISABLED)

    def _limpar_log_allin(self):
        self.log_allin.text.configure(state=NORMAL)
        self.log_allin.text.delete("1.0", "end")
        self.log_allin.text.configure(state=DISABLED)

    def _escolher_tabela_docx_allin(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar .docx com a tabela de segurados", filetypes=[("Documento Word", "*.docx")],
            parent=self.root,
        )
        if caminho:
            self.var_docx_tabela_allin.set(caminho)
            # a pasta do .docx importado vira a pasta da tese, a não ser
            # que o usuário já tenha escolhido uma manualmente
            if not self._pasta_tese_allin_manual:
                self.var_pasta_tese_allin.set(str(Path(caminho).parent))

    def _escolher_pasta_tese_allin(self):
        caminho = filedialog.askdirectory(title="Selecionar a pasta da tese (com as subpastas dos segurados)", parent=self.root)
        if caminho:
            self.var_pasta_tese_allin.set(caminho)
            self._pasta_tese_allin_manual = True

    def _abrir_pasta_allin(self):
        pasta_tese = self.var_pasta_tese_allin.get().strip()
        if pasta_tese and Path(pasta_tese).is_dir():
            os.startfile(pasta_tese)

    def _abrir_pdf_allin(self):
        if self._ultimo_pdf_allin and self._ultimo_pdf_allin.is_file():
            os.startfile(str(self._ultimo_pdf_allin))

    def _limpar_allin(self):
        self.var_topico_allin.set("")
        self.var_titulo_allin.set("")
        self.var_docx_tabela_allin.set("")
        self.var_pasta_tese_allin.set("")
        self._pasta_tese_allin_manual = False
        self.var_gerar_docx_allin.set(False)
        self.botao_abrir_pasta_allin.config(state=DISABLED)
        self.botao_abrir_pdf_allin.config(state=DISABLED)
        self._ultimo_pdf_allin = None
        self.frame_progresso_allin.pack_forget()
        self._limpar_log_allin()
        self._mostrar_inicial_allin()

    def _obter_tabela_allin(self):
        caminho_texto = self.var_docx_tabela_allin.get().strip()
        if not caminho_texto:
            raise ValueError("Selecione o .docx com a tabela de segurados.")
        caminho = Path(caminho_texto)
        if not caminho.is_file():
            raise ValueError(f"Arquivo da tabela não encontrado:\n{caminho}")
        return tabela.extrair_de_docx(caminho)

    def _gerar_allin(self):
        if not MODELO_CAPAS.is_file():
            messagebox.showerror(
                "Modelo não encontrado",
                f"O modelo padrão não foi encontrado em:\n{MODELO_CAPAS}",
                parent=self.root,
            )
            return

        topico = self.var_topico_allin.get().strip()
        if not topico:
            messagebox.showwarning("Campo obrigatório", "Informe o número do tópico.", parent=self.root)
            return

        titulo = self.var_titulo_allin.get().strip()
        if not titulo:
            messagebox.showwarning("Campo obrigatório", "Informe o título da tese.", parent=self.root)
            return
        titulo = _normalizar_travessao(_normalizar_espacos(titulo))

        pasta_tese_texto = self.var_pasta_tese_allin.get().strip()
        if not pasta_tese_texto:
            messagebox.showwarning("Campo obrigatório", "Selecione a pasta da tese.", parent=self.root)
            return
        pasta_tese = Path(pasta_tese_texto)
        if not pasta_tese.is_dir():
            messagebox.showerror("Pasta não encontrada", f"Não foi possível encontrar a pasta:\n{pasta_tese}", parent=self.root)
            return

        try:
            tabela_segurados, tabela_referencia = self._obter_tabela_allin()
        except ValueError as exc:
            messagebox.showwarning("Tabela de segurados", str(exc), parent=self.root)
            return

        n_segurados = len(tabela_segurados.grupos)
        subpastas = pastas_segurados_em_ordem(pasta_tese)
        aviso_subpastas = None
        if len(subpastas) != n_segurados:
            # não bloqueia: separar_capas cai no modo "soltas na pasta da
            # tese" quando as subpastas ainda não existem — só avisa que,
            # nesse caso, a etapa final de juntar os PDFs será pulada
            aviso_subpastas = (
                f"A tabela tem {_plural(n_segurados, 'segurado', 'segurados')}, mas a pasta da tese tem "
                f"{_plural(len(subpastas), 'subpasta numerada', 'subpastas numeradas')} (ex.: \"1. NOME\") — "
                "as capas vão sair soltas na pasta da tese, e a etapa de juntar os PDFs será pulada."
            )

        # avisa antes de sobrescrever um PDF/Word de capas ou o PDF final
        # já gerados (mesmo nome, direto na pasta da tese)
        nome_capas = "Capas Geradas"
        nome_final = sanitizar_nome_arquivo(f"Tópico {topico} - {pasta_tese.name}")
        existentes = [
            caminho.name
            for caminho in (
                pasta_tese / f"{nome_capas}.pdf", pasta_tese / f"{nome_capas}.docx",
                pasta_tese / f"{nome_final}.pdf",
            )
            if caminho.exists() and (caminho.suffix == ".pdf" or self.var_gerar_docx_allin.get())
        ]
        if existentes:
            lista = "\n".join(f"  • {nome}" for nome in existentes)
            verbo = "já existe e será SUBSTITUÍDO" if len(existentes) == 1 else "já existem e serão SUBSTITUÍDOS"
            if not messagebox.askyesno(
                "Arquivos serão substituídos",
                f"{_plural(len(existentes), 'arquivo', 'arquivos')} {verbo} na pasta da tese:\n\n{lista}\n\n"
                "Continuar?",
                icon="warning",
                parent=self.root,
            ):
                return

        # avisa se alguma subpasta de segurado já tem uma capa de uma
        # rodada anterior (arquivo começando com "0.") — sem isso, rodar
        # de novo com um tópico/título diferente deixaria a capa antiga
        # E a nova juntas na mesma subpasta, duplicando no PDF final
        capas_antigas = []
        for subpasta in subpastas:
            achadas = [
                arquivo for arquivo in subpasta.iterdir()
                if arquivo.is_file() and arquivo.suffix.lower() == ".pdf" and re.match(r"0\.", arquivo.name)
            ]
            if achadas:
                capas_antigas.append((subpasta, achadas))

        if capas_antigas:
            lista = "\n".join(
                f"  • {subpasta.name}: {', '.join(a.name for a in arquivos)}"
                for subpasta, arquivos in capas_antigas[:10]
            )
            if len(capas_antigas) > 10:
                lista += f"\n  ... e mais {len(capas_antigas) - 10} subpasta(s)"
            if not messagebox.askyesno(
                "Capas antigas encontradas",
                f"{_plural(len(capas_antigas), 'subpasta já tem', 'subpastas já têm')} uma capa de uma "
                f"rodada anterior (arquivo começando com \"0.\"):\n\n{lista}\n\n"
                "Para não duplicar, essas capas antigas serão apagadas antes de gerar as novas.\n\n"
                "Continuar?",
                icon="warning",
                parent=self.root,
            ):
                return
            for _subpasta, arquivos in capas_antigas:
                for arquivo in arquivos:
                    try:
                        arquivo.unlink()
                    except OSError:
                        pass

        self._limpar_log_allin()
        self._log_allin(
            f"Tabela lida: {n_segurados} segurado(s), {len(tabela_segurados.linhas)} linha(s) de benefício.",
            "titulo",
        )
        if aviso_subpastas:
            self._log_allin(f"⚠ {aviso_subpastas}", "aviso")

        self.botao_gerar_allin.config(state=DISABLED)
        self.botao_abrir_pasta_allin.config(state=DISABLED)
        self.botao_abrir_pdf_allin.config(state=DISABLED)
        self.barra_progresso_allin.configure(value=0, maximum=100)
        self.label_progresso_allin.configure(text="Etapa 1/3: gerando as capas...")
        self.frame_progresso_allin.pack(fill=X, pady=(0, 14), before=self.moldura_resultado_allin)
        self._cancelar_allin.clear()
        self.botao_suspender_allin.configure(state=NORMAL, text="⛔  Suspender")
        self.botao_suspender_allin.pack(side=RIGHT)

        threading.Thread(
            target=self._allin_worker,
            args=(titulo, topico, tabela_segurados, tabela_referencia, pasta_tese, self.var_gerar_docx_allin.get()),
            daemon=True,
        ).start()

    def _suspender_allin(self):
        self._cancelar_allin.set()
        self.botao_suspender_allin.configure(state=DISABLED, text="⛔  Suspendendo...")

    def _allin_progresso_capas(self, etapa: int, total: int, rotulo: str):
        self.root.after(0, self._allin_atualizar_progresso, "Etapa 1/3: gerando as capas", etapa, total)
        self.root.after(0, self._log_allin, f"[{etapa}/{total}] {rotulo}", "item")

    def _allin_progresso_juntar(self, atual: int, total: int):
        self.root.after(0, self._allin_atualizar_progresso, "Etapa 3/3: juntando os PDFs", atual, total)

    def _allin_atualizar_progresso(self, rotulo: str, atual: int, total: int):
        self.barra_progresso_allin.configure(value=atual, maximum=max(total, 1))
        self.label_progresso_allin.configure(text=f"{rotulo} ({atual}/{total})...")

    def _allin_confirmar_juntar(self, pasta_tese: Path) -> bool:
        """Mostra, bem no momento de juntar (com as capas já divididas nas
        subpastas — a mesma conferência da aba "Juntar PDFs", com a
        contagem de arquivos/páginas já refletindo as capas). Chamado do
        worker, mas a caixa de diálogo do Tkinter só pode ser criada na
        thread principal — um Event sincroniza as duas threads."""
        resultado = {}
        evento = threading.Event()

        def perguntar():
            resultado["ok"] = self._confirmar_ordem(
                pasta_tese, pasta_tese, "tese", nome_pasta_destino="pasta da tese",
            )
            evento.set()

        self.root.after(0, self._allin_atualizar_progresso, "Etapa 3/3: aguardando confirmação", 0, 1)
        self.root.after(0, perguntar)
        evento.wait()
        return resultado["ok"]

    def _allin_worker(self, titulo, topico, tabela_segurados, tabela_referencia, pasta_tese: Path, gerar_docx: bool):
        pythoncom.CoInitialize()
        try:
            try:
                conversor = ConversorPDF()
            except ConversorPDFIndisponivel as exc:
                self.root.after(0, self._allin_falhou, str(exc))
                return

            try:
                # gera direto na pasta da tese (mesmo lugar do PDF final,
                # sem subpasta própria), já como "Capas Geradas" — mais
                # simples do que uma subpasta extra só pra isso
                caminho_pdf_capas, caminho_docx_capas = gerar_pdf_capas(
                    MODELO_CAPAS, titulo, tabela_segurados, tabela_referencia, pasta_tese,
                    conversor, topico=topico, gerar_docx=gerar_docx,
                    callback_progresso=self._allin_progresso_capas,
                    cancelar=self._cancelar_allin,
                )
                self.root.after(0, self._log_allin, "✓ Capas geradas", "item")

                arquivos_divididos = separar_capas(caminho_pdf_capas, pasta_tese)
                self.root.after(0, self._allin_atualizar_progresso, "Etapa 2/3: dividindo as capas", 1, 1)
                self.root.after(
                    0, self._log_allin,
                    f"✓ Capas divididas ({_plural(len(arquivos_divididos), 'arquivo', 'arquivos')})", "item",
                )

                # só junta se o número de subpastas numeradas bater com o
                # número de segurados da tabela — se não bater, separar_capas
                # deixou as capas soltas na pasta da tese (não distribuídas
                # nas subpastas), e tentar juntar só daria um erro confuso
                # ("nenhuma subpasta continha PDF"). A confirmação (lista de
                # pastas + páginas + aviso de sobrescrita) só faz sentido bem
                # AQUI, no momento de juntar — antes disso as capas ainda nem
                # existiam
                caminho_final = None
                motivo_sem_juntar = None
                subpastas_atuais = pastas_segurados_em_ordem(pasta_tese)
                if subpastas_atuais and len(subpastas_atuais) == len(tabela_segurados.grupos):
                    if self._allin_confirmar_juntar(pasta_tese):
                        caminho_final = juntar_pdfs.juntar_tese(
                            pasta_tese, pasta_tese, progresso_callback=self._allin_progresso_juntar,
                            cancelar=self._cancelar_allin,
                        )
                        nome_final = sanitizar_nome_arquivo(f"Tópico {topico} - {pasta_tese.name}")
                        caminho_final = caminho_final.replace(pasta_tese / f"{nome_final}.pdf")
                    else:
                        motivo_sem_juntar = "cancelado"
                else:
                    motivo_sem_juntar = "sem_subpastas"
            except InterruptedError:
                self.root.after(0, self._allin_cancelado)
                return
            except Exception as exc:
                self.root.after(0, self._allin_falhou, _mensagem_erro_amigavel(exc))
                return
            finally:
                conversor.fechar()

            self.root.after(0, self._allin_concluiu, caminho_final, motivo_sem_juntar)
        finally:
            pythoncom.CoUninitialize()

    def _restaurar_botao_allin(self):
        self.botao_gerar_allin.config(state=NORMAL)
        self.frame_progresso_allin.pack_forget()
        self.botao_suspender_allin.pack_forget()

    def _allin_cancelado(self):
        self._restaurar_botao_allin()
        self.botao_abrir_pasta_allin.config(state=NORMAL)
        self._log_allin(
            "\n⛔ Operação suspensa pelo usuário. O que já tinha sido gerado/dividido até esse ponto "
            "continua na pasta da tese.", "aviso",
        )

    def _allin_falhou(self, mensagem: str):
        self._restaurar_botao_allin()
        self._log_allin(f"ERRO: {mensagem}", "erro")
        messagebox.showerror("Erro ao gerar o anexo completo", mensagem, parent=self.root)

    def _allin_concluiu(self, caminho_final: Path | None, motivo_sem_juntar: str | None = None):
        self._restaurar_botao_allin()
        self.botao_abrir_pasta_allin.config(state=NORMAL)
        if caminho_final is None:
            if motivo_sem_juntar == "cancelado":
                self._log_allin(
                    "\nJunção cancelada — as capas já geradas e divididas foram mantidas na pasta da "
                    "tese. Junte manualmente na aba \"Juntar PDFs\" quando quiser.", "titulo",
                )
            else:
                self._log_allin(
                    "\nCrie as subpastas numeradas dos segurados e rode novamente (ou junte manualmente "
                    "na aba \"Juntar PDFs\") para gerar o PDF final da tese.", "titulo",
                )
            return
        self._ultimo_pdf_allin = caminho_final
        self._log_allin("✓ PDFs juntados", "item")
        self._log_allin(f"\nPDF final gerado em:\n{caminho_final}", "titulo")
        self.botao_abrir_pdf_allin.config(state=NORMAL)

    # =========================================================================
    # ABA GERAR CAPAS
    # =========================================================================

    def _montar_card_capas(self, pai):
        cartao = ttk.Labelframe(pai, text=" Tese ", padding=18, bootstyle="secondary")
        cartao.pack(fill=X, pady=(0, 14))
        cartao.columnconfigure(1, weight=1)

        ttk.Label(cartao, text="Número do tópico", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=W, padx=(0, 10), pady=(0, 10))
        ttk.Entry(cartao, textvariable=self.var_topico_capas, width=8).grid(row=0, column=1, sticky=W, pady=(0, 10))
        ttk.Label(cartao, text="Título da tese", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=W, padx=(0, 10))
        ttk.Entry(cartao, textvariable=self.var_titulo_capas).grid(row=1, column=1, sticky=EW)

        tabela_card = ttk.Labelframe(pai, text=" Tabela de segurados ", padding=18, bootstyle="secondary")
        tabela_card.pack(fill=X, pady=(0, 14))
        tabela_card.columnconfigure(0, weight=1)

        linha_radios = ttk.Frame(tabela_card)
        linha_radios.grid(row=1, column=0, sticky=W, pady=(0, 10))
        ttk.Radiobutton(
            linha_radios, text="Importar de um .docx", variable=self.var_origem_tabela_capas, value="docx",
            command=self._alternar_origem_tabela_capas, bootstyle="secondary",
        ).pack(side=LEFT, padx=(0, 16))
        ttk.Radiobutton(
            linha_radios, text="Colar texto  (em breve)", variable=self.var_origem_tabela_capas, value="colar",
            command=self._alternar_origem_tabela_capas, bootstyle="secondary", state=DISABLED,
        ).pack(side=LEFT)

        self.frame_tabela_docx_capas = ttk.Frame(tabela_card)
        self.frame_tabela_docx_capas.grid(row=2, column=0, sticky=EW)
        self.frame_tabela_docx_capas.columnconfigure(1, weight=1)
        # botão à ESQUERDA (perto do rótulo "Importar de um .docx" acima) e
        # preenchido/maior — evita confusão com o "Baixar documento base",
        # que fica mais acima e continua com contorno fino
        ttk.Button(
            self.frame_tabela_docx_capas, text="📂  Procurar...", command=self._escolher_tabela_docx_capas,
            bootstyle="primary", width=14,
        ).grid(row=0, column=0, padx=(0, 10))
        ttk.Entry(self.frame_tabela_docx_capas, textvariable=self.var_docx_tabela_capas).grid(row=0, column=1, sticky=EW)
        ttk.Label(
            self.frame_tabela_docx_capas,
            text="⚠️ Se a tabela ocupar mais de uma página, desative a opção \"Repetir "
                 "linhas de cabeçalho\" — caso contrário, o programa não conseguirá ler "
                 "os dados corretamente.",
            bootstyle="warning", font=("Segoe UI", 8, "bold"), wraplength=760, justify=LEFT,
        ).grid(row=1, column=0, columnspan=2, sticky=W, pady=(8, 0))

        self.frame_tabela_colar_capas = ttk.Frame(tabela_card)
        self.frame_tabela_colar_capas.grid(row=2, column=0, sticky=EW)
        self.frame_tabela_colar_capas.columnconfigure(0, weight=1)
        ttk.Label(
            self.frame_tabela_colar_capas, text=TEXTO_AJUDA_COLAR, foreground="#666666",
            font=("Segoe UI", 8), wraplength=860,
        ).pack(anchor=W, pady=(0, 6))
        self.texto_colar_capas = ScrolledText(self.frame_tabela_colar_capas, autohide=True, bootstyle="secondary", height=5)
        self.texto_colar_capas.pack(fill=BOTH, expand=True)
        self.texto_colar_capas.text.configure(font=("Consolas", 9), padx=6, pady=6, relief="flat")
        for sequencia in ("<Control-v>", "<Control-V>", "<Shift-Insert>"):
            self.texto_colar_capas.text.bind(sequencia, self._colou_na_caixa_capas, add="+")
        self.texto_colar_capas.text.bind("<<Modified>>", self._texto_colar_modificado, add="+")

        linha_rodape_colar = ttk.Frame(self.frame_tabela_colar_capas)
        linha_rodape_colar.pack(fill=X, pady=(6, 0))
        ttk.Button(
            linha_rodape_colar, text="⚙  Padronizar tabela...", command=self._abrir_padronizar_tabela,
            bootstyle="primary-outline",
        ).pack(side=LEFT)
        self.label_larguras_capas = ttk.Label(
            linha_rodape_colar, text="", bootstyle="secondary", font=("Segoe UI", 8),
        )
        self.label_larguras_capas.pack(side=LEFT, padx=(10, 0))
        self.label_status_colar_capas = ttk.Label(
            self.frame_tabela_colar_capas, text="", font=("Segoe UI", 8, "bold"),
        )
        self.label_status_colar_capas.pack(anchor=W, pady=(6, 0))

        linha_base = ttk.Frame(tabela_card)
        linha_base.grid(row=0, column=0, sticky=W, pady=(0, 12))
        ttk.Button(
            linha_base, text="📥  Baixar documento base das tabelas",
            command=self._baixar_base_tabelas, bootstyle="primary-outline",
        ).pack(side=LEFT)
        ttk.Label(
            linha_base, text="Documento em branco com as margens corretas — monte a sua tabela nele antes de importar.",
            bootstyle="secondary", font=("Segoe UI", 8),
        ).pack(side=LEFT, padx=(10, 0))

        saida_card = ttk.Labelframe(pai, text=" Pasta da tese ", padding=18, bootstyle="secondary")
        saida_card.pack(fill=X)
        saida_card.columnconfigure(0, weight=1)
        linha_saida = ttk.Frame(saida_card)
        linha_saida.grid(row=0, column=0, sticky=EW)
        linha_saida.columnconfigure(0, weight=1)
        ttk.Entry(linha_saida, textvariable=self.var_saida_capas).grid(row=0, column=0, sticky=EW, padx=(0, 8))
        ttk.Button(linha_saida, text="Procurar...", command=self._escolher_saida_capas, bootstyle="primary-outline").grid(row=0, column=1)
        ttk.Checkbutton(
            saida_card, text="Gerar também um arquivo Word (.docx)", variable=self.var_gerar_docx_capas,
            bootstyle="secondary",
        ).grid(row=1, column=0, sticky=W, pady=(10, 0))

        self._alternar_origem_tabela_capas()

    def _montar_acoes_capas(self, pai):
        acoes = ttk.Frame(pai)
        acoes.pack(fill=X, pady=(14, 14))
        self.botao_abrir_pasta_capas = ttk.Button(
            acoes, text="📂  Abrir pasta da tese", command=self._abrir_pasta_capas,
            bootstyle="primary-outline", state=DISABLED,
        )
        self.botao_abrir_pasta_capas.pack(side=LEFT)
        self.botao_abrir_pdf_capas = ttk.Button(
            acoes, text="👁  Abrir PDF gerado", command=self._abrir_pdf_capas,
            bootstyle="primary-outline", state=DISABLED,
        )
        self.botao_abrir_pdf_capas.pack(side=LEFT, padx=(10, 0))
        ttk.Button(acoes, text="🧹  Limpar", command=self._limpar_capas, bootstyle="danger-outline").pack(side=LEFT, padx=(10, 0))
        self.botao_gerar_capas = ttk.Button(
            acoes, text="📑  Gerar capas", command=self._gerar_capas, bootstyle="secondary", width=20,
        )
        self.botao_gerar_capas.pack(side=RIGHT)

    def _montar_progresso_capas(self, pai):
        self.frame_progresso_capas = ttk.Frame(pai)
        self.barra_progresso_capas = ttk.Progressbar(
            self.frame_progresso_capas, mode="determinate", bootstyle="secondary",
        )
        self.barra_progresso_capas.pack(fill=X, pady=(0, 4))
        self.label_progresso_capas = ttk.Label(
            self.frame_progresso_capas, text="", bootstyle="secondary", font=("Segoe UI", 9),
        )
        self.label_progresso_capas.pack(anchor=W)

    def _montar_resultado_capas(self, pai):
        moldura = ttk.Labelframe(pai, text=" Resultado ", padding=14)
        moldura.pack(fill=BOTH, expand=True)
        self.moldura_resultado_capas = moldura

        self.log_capas = ScrolledText(moldura, autohide=True, bootstyle="secondary")
        self.log_capas.pack(fill=BOTH, expand=True)
        self.log_capas.text.configure(font=("Segoe UI", 10), padx=8, pady=8, relief="flat")
        self.log_capas.text.tag_configure("titulo", font=("Segoe UI", 10, "bold"))
        self.log_capas.text.tag_configure("item", foreground=tema.COR_SUCESSO)
        self.log_capas.text.tag_configure("erro", foreground=tema.COR_ERRO, font=("Segoe UI", 10, "bold"))
        self._mostrar_inicial_capas()

    def _mostrar_inicial_capas(self):
        self.log_capas.text.insert("end", "Preencha a tese, a tabela de segurados e clique em \"Gerar capas\".")
        self.log_capas.text.configure(state=DISABLED)

    def _log_capas(self, texto: str, tag: str | None = None):
        self.log_capas.text.configure(state=NORMAL)
        self.log_capas.text.insert("end", texto + "\n", tag or ())
        self.log_capas.text.see("end")
        self.log_capas.text.configure(state=DISABLED)

    def _limpar_log_capas(self):
        self.log_capas.text.configure(state=NORMAL)
        self.log_capas.text.delete("1.0", "end")
        self.log_capas.text.configure(state=DISABLED)

    def _alternar_origem_tabela_capas(self):
        if self.var_origem_tabela_capas.get() == "docx":
            self.frame_tabela_colar_capas.grid_remove()
            self.frame_tabela_docx_capas.grid()
        else:
            self.frame_tabela_docx_capas.grid_remove()
            self.frame_tabela_colar_capas.grid()

    # ── Captura da tabela colada com formatação (via Word invisível) ─────

    def _colou_na_caixa_capas(self, _evento=None):
        # deixa o Ctrl+V padrão inserir o texto simples e, logo depois,
        # tenta capturar do clipboard a tabela com formatação
        self.root.after(150, self._iniciar_captura_formatacao)

    def _texto_colar_modificado(self, _evento=None):
        widget = self.texto_colar_capas.text
        if not widget.edit_modified():
            return
        widget.edit_modified(False)
        if self._ignorar_modificacao_colar:
            return
        # edição manual invalida a formatação capturada — os dados usados
        # na geração passam a ser o texto simples da caixa
        if self._tabela_word_capturada is not None:
            self._tabela_word_capturada = None
            self.label_status_colar_capas.configure(
                text="Texto editado — será aplicado o estilo padrão da capa.", bootstyle="secondary",
            )

    def _iniciar_captura_formatacao(self):
        if self._captura_em_andamento:
            return
        self._captura_em_andamento = True
        self.label_status_colar_capas.configure(
            text="Lendo a formatação da tabela do Word...", bootstyle="secondary",
        )
        threading.Thread(target=self._capturar_formatacao_worker, daemon=True).start()

    def _capturar_formatacao_worker(self):
        pythoncom.CoInitialize()
        tmp = None
        try:
            import win32com.client

            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Add()
                doc.Range().Paste()
                try:
                    # colar num documento em branco costuma espremer colunas
                    # (ex.: CNPJ quebrado em duas linhas). Auto-ajusta ao
                    # conteúdo e depois distribui na largura da página.
                    tabela_word = doc.Tables(1)
                    tabela_word.AllowAutoFit = True
                    tabela_word.AutoFitBehavior(1)  # wdAutoFitContent
                    tabela_word.AutoFitBehavior(2)  # wdAutoFitWindow
                except Exception:
                    pass
                tmp = Path(tempfile.mktemp(suffix=".docx"))
                doc.SaveAs(str(tmp), FileFormat=16)  # wdFormatXMLDocument
                doc.Close(False)
            finally:
                word.Quit()

            resultado = tabela.extrair_de_docx(tmp)
            self.root.after(0, self._captura_formatacao_ok, resultado)
        except Exception:
            self.root.after(0, self._captura_formatacao_falhou)
        finally:
            self._captura_em_andamento = False
            if tmp is not None:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            pythoncom.CoUninitialize()

    def _captura_formatacao_ok(self, resultado):
        tabela_segurados, _referencia = resultado
        self._tabela_word_capturada = resultado

        # reescreve a caixa com a versão canônica da tabela capturada,
        # para o texto exibido e os dados usados na geração serem os mesmos.
        # células mescladas horizontalmente (span 0 = continuação) aparecem
        # uma única vez, como no Word
        def _linha_para_texto(linha):
            spans = getattr(linha, "spans", None)
            celulas = (
                (c for c, s in zip(linha, spans) if s != 0) if spans else linha
            )
            return "\t".join(c.replace("\n", " ") for c in celulas)

        linhas_texto = [_linha_para_texto(tabela_segurados.cabecalhos)]
        linhas_texto += [_linha_para_texto(linha) for linha in tabela_segurados.linhas]
        self._ignorar_modificacao_colar = True
        try:
            widget = self.texto_colar_capas.text
            widget.delete("1.0", "end")
            widget.insert("1.0", "\n".join(linhas_texto))
            widget.edit_modified(False)
        finally:
            self._ignorar_modificacao_colar = False

        n = len(tabela_segurados.grupos)
        self.label_status_colar_capas.configure(
            text=f"✔ Tabela do Word capturada ({n} segurado(s)) — mesclagens, cores, bordas e fontes serão replicadas.",
            bootstyle="success",
        )
        self._atualizar_label_larguras(len(tabela_segurados.cabecalhos))

    def _captura_formatacao_falhou(self):
        # sem Word, sem tabela no clipboard ou clipboard só com texto:
        # segue valendo o texto simples colado na caixa
        self._tabela_word_capturada = None
        if self.texto_colar_capas.text.get("1.0", "end").strip():
            self.label_status_colar_capas.configure(
                text="Texto sem formatação — será aplicado o estilo padrão da capa.", bootstyle="secondary",
            )
        else:
            self.label_status_colar_capas.configure(text="")

    # ── Padronização de larguras da tabela (modo colar) ──────────────────

    def _larguras_para(self, n_colunas: int) -> list[float] | None:
        salvas = self._config.get("larguras_tabela_capas", {}).get(str(n_colunas))
        if salvas and len(salvas) == n_colunas:
            return salvas
        return LARGURAS_PADRAO_CAPAS.get(n_colunas)

    def _salvar_larguras(self, n_colunas: int, proporcoes: list[float] | None) -> None:
        larguras = self._config.setdefault("larguras_tabela_capas", {})
        if proporcoes is None:
            larguras.pop(str(n_colunas), None)
        else:
            larguras[str(n_colunas)] = proporcoes
        _salvar_config(self._config)
        self._atualizar_label_larguras(n_colunas)

    def _atualizar_label_larguras(self, n_colunas: int | None = None) -> None:
        if n_colunas is not None and self._larguras_para(n_colunas):
            self.label_larguras_capas.configure(
                text=f"✔ Larguras personalizadas ativas ({n_colunas} colunas)")
        else:
            self.label_larguras_capas.configure(text="")

    def _tabela_atual_colar(self):
        if self._tabela_word_capturada is not None:
            return self._tabela_word_capturada[0]
        texto = self.texto_colar_capas.text.get("1.0", "end")
        if not texto.strip():
            raise ValueError("Cole a tabela de segurados antes de padronizar.")
        return tabela.extrair_de_texto(texto)

    def _abrir_padronizar_tabela(self):
        from tkinter import font as tkfont

        try:
            tabela_segurados = self._tabela_atual_colar()
        except ValueError as exc:
            messagebox.showwarning("Padronizar tabela", str(exc), parent=self.root)
            return

        referencia = self._tabela_word_capturada[1] if self._tabela_word_capturada else None
        n_colunas = len(tabela_segurados.cabecalhos)
        cabecalhos = [c.replace("\n", " ") for c in tabela_segurados.cabecalhos]
        fonte = tkfont.nametofont("TkDefaultFont")

        # colunas VISUAIS: uma célula mesclada no cabeçalho vira um único
        # controle (grupo de colunas da grade)
        spans_cab = getattr(tabela_segurados.cabecalhos, "spans", None) or [1] * n_colunas
        grupos_visuais: list[list[int]] = []
        i = 0
        while i < n_colunas:
            span = max(1, spans_cab[i])
            grupos_visuais.append(list(range(i, min(i + span, n_colunas))))
            i += span

        def _medida_grid(indice: int) -> float:
            maior = fonte.measure(cabecalhos[indice]) / max(1, spans_cab[indice] or 1)
            for linha in tabela_segurados.linhas[:60]:
                if indice < len(linha):
                    maior = max(maior, fonte.measure(linha[indice].replace("\n", " ")))
            return maior + 28

        def proporcoes_por_conteudo() -> list[float]:
            larguras = [_medida_grid(i) for i in range(n_colunas)]
            soma = sum(larguras)
            return [largura / soma for largura in larguras]

        estado = {
            "props": self._larguras_para(n_colunas) or proporcoes_por_conteudo(),
            "renderizando": False,
        }

        janela = ttk.Toplevel(self.root)
        janela.title("Padronizar tabela")
        janela.transient(self.root)
        janela.resizable(False, False)
        try:
            janela.iconbitmap(_caminho_recurso("assets/pdf.ico"))
        except tk.TclError:
            pass

        ttk.Label(
            janela, text="Defina a largura de cada coluna (%) e confira na prévia real, gerada com o Word.",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor=W, padx=18, pady=(16, 2))
        ttk.Label(
            janela, text="As proporções valem para todas as capas geradas ao colar a tabela e ficam "
                         "salvas para as próximas vezes.",
            bootstyle="secondary", font=("Segoe UI", 9),
        ).pack(anchor=W, padx=18, pady=(0, 10))

        # ── controles: um por coluna visual ──
        ajustes = ttk.Labelframe(janela, text=" Largura de cada coluna (%) ", padding=(12, 8))
        ajustes.pack(fill=X, padx=18)
        vars_percent = [tk.DoubleVar(value=0) for _ in grupos_visuais]
        sincronizando = {"ativo": False}

        def _percentuais_visuais() -> list[float]:
            return [sum(estado["props"][i] for i in grupo) for grupo in grupos_visuais]

        def _atualizar_spins():
            sincronizando["ativo"] = True
            try:
                for var, peso in zip(vars_percent, _percentuais_visuais()):
                    var.set(round(peso * 100, 1))
            finally:
                sincronizando["ativo"] = False

        def spin_alterado(*_args):
            if sincronizando["ativo"]:
                return
            try:
                pesos = [max(0.5, float(var.get())) for var in vars_percent]
            except (tk.TclError, ValueError):
                return
            soma = sum(pesos)
            novos = list(estado["props"])
            for grupo, peso in zip(grupos_visuais, pesos):
                atual_grupo = sum(estado["props"][i] for i in grupo) or 1
                for i in grupo:
                    # preserva a divisão interna do grupo (sub-colunas da mesclagem)
                    novos[i] = (peso / soma) * (estado["props"][i] / atual_grupo)
            estado["props"] = novos

        for coluna_visual, grupo in enumerate(grupos_visuais):
            bloco = ttk.Frame(ajustes)
            bloco.grid(row=0, column=coluna_visual, padx=6, pady=2)
            titulo = cabecalhos[grupo[0]]
            titulo_curto = titulo if len(titulo) <= 18 else titulo[:17] + "…"
            ttk.Label(bloco, text=titulo_curto, font=("Segoe UI", 8)).pack(anchor=W)
            spin = ttk.Spinbox(
                bloco, from_=1, to=90, increment=1, width=6,
                textvariable=vars_percent[coluna_visual], command=spin_alterado,
            )
            spin.pack()
            spin.bind("<KeyRelease>", spin_alterado)

        # ── ajustes rápidos + atualizar prévia ──
        rapidos = ttk.Frame(janela, padding=(18, 10, 18, 0))
        rapidos.pack(fill=X)

        def ajustar_conteudo():
            estado["props"] = proporcoes_por_conteudo()
            _atualizar_spins()
            renderizar()

        def distribuir_igual():
            estado["props"] = [1 / n_colunas] * n_colunas
            _atualizar_spins()
            renderizar()

        ttk.Button(rapidos, text="📐  Ajustar ao conteúdo", bootstyle="primary-outline",
                   command=ajustar_conteudo).pack(side=LEFT)
        ttk.Button(rapidos, text="▤  Distribuir igualmente", bootstyle="primary-outline",
                   command=distribuir_igual).pack(side=LEFT, padx=(8, 0))
        botao_previa = ttk.Button(rapidos, text="🔄  Atualizar prévia", bootstyle="secondary",
                                  command=lambda: renderizar())
        botao_previa.pack(side=RIGHT)

        # ── prévia real (renderizada com o Word) ──
        moldura_previa = ttk.Labelframe(janela, text=" Prévia (idêntica ao PDF final) ", padding=8)
        moldura_previa.pack(fill=BOTH, expand=True, padx=18, pady=(10, 0))
        label_previa = tk.Label(moldura_previa, text="Gerando prévia com o Word...",
                                background="white", height=18, anchor="center")
        label_previa.pack(fill=BOTH, expand=True)

        def renderizar():
            if estado["renderizando"]:
                return
            estado["renderizando"] = True
            botao_previa.configure(state=DISABLED)
            label_previa.configure(text="Gerando prévia com o Word...", image="")
            props = list(estado["props"])
            threading.Thread(target=_render_worker, args=(props,), daemon=True).start()

        def _render_worker(props):
            pythoncom.CoInitialize()
            pasta_tmp = Path(tempfile.mkdtemp(prefix="previa_capas_"))
            try:
                import fitz

                doc = documento.montar_tabela_completa(
                    str(MODELO_CAPAS), self.var_titulo_capas.get().strip() or "Prévia da tabela",
                    tabela_segurados.cabecalhos, tabela_segurados.linhas[:12],
                    tabela_segurados.indice_nome, referencia, props,
                )
                docx_tmp = pasta_tmp / "previa.docx"
                pdf_tmp = pasta_tmp / "previa.pdf"
                doc.save(str(docx_tmp))
                conversor = ConversorPDF()
                try:
                    conversor.converter(docx_tmp, pdf_tmp)
                finally:
                    conversor.fechar()

                with fitz.open(pdf_tmp) as pdf:
                    pagina = pdf[0]
                    blocos = [b for b in pagina.get_text("blocks") if b[4].strip()]
                    clip = None
                    if blocos:
                        y0 = max(0, min(b[1] for b in blocos) - 16)
                        y1 = min(pagina.rect.height, max(b[3] for b in blocos) + 16)
                        clip = fitz.Rect(0, y0, pagina.rect.width, y1)
                    pix = pagina.get_pixmap(dpi=120, clip=clip)
                    png = pix.tobytes("png")
                self.root.after(0, _mostrar_previa, png)
            except ConversorPDFIndisponivel as exc:
                self.root.after(0, _previa_falhou, str(exc))
            except Exception as exc:
                self.root.after(0, _previa_falhou, _mensagem_erro_amigavel(exc))
            finally:
                shutil.rmtree(pasta_tmp, ignore_errors=True)
                estado["renderizando"] = False
                pythoncom.CoUninitialize()

        def _mostrar_previa(png_bytes):
            if not janela.winfo_exists():
                return
            imagem = Image.open(io.BytesIO(png_bytes))
            LARGURA_MAX, ALTURA_MAX = 940, 420
            escala = min(LARGURA_MAX / imagem.width, ALTURA_MAX / imagem.height, 1.0)
            imagem = imagem.resize((int(imagem.width * escala), int(imagem.height * escala)), Image.LANCZOS)
            janela._imagem_previa = ImageTk.PhotoImage(imagem)
            label_previa.configure(image=janela._imagem_previa, text="", height=imagem.height)
            botao_previa.configure(state=NORMAL)
            janela.update_idletasks()
            janela.geometry(f"{janela.winfo_reqwidth()}x{janela.winfo_reqheight()}")

        def _previa_falhou(mensagem):
            if not janela.winfo_exists():
                return
            label_previa.configure(text=f"Não foi possível gerar a prévia:\n{mensagem}", image="")
            botao_previa.configure(state=NORMAL)

        # ── rodapé ──
        rodape = ttk.Frame(janela, padding=18)
        rodape.pack(fill=X)

        def aplicar():
            self._salvar_larguras(n_colunas, list(estado["props"]))
            janela.destroy()

        def restaurar():
            self._salvar_larguras(n_colunas, None)
            janela.destroy()

        ttk.Button(rodape, text="🧹  Restaurar padrão", command=restaurar, bootstyle="danger-outline").pack(side=LEFT)
        ttk.Button(rodape, text="Aplicar", command=aplicar, bootstyle="secondary", width=14).pack(side=RIGHT)
        ttk.Button(rodape, text="Cancelar", command=janela.destroy, bootstyle="primary-outline", width=14).pack(side=RIGHT, padx=(0, 10))

        _atualizar_spins()
        janela.update_idletasks()
        largura, altura = janela.winfo_reqwidth(), janela.winfo_reqheight()
        _posicionar_sobre_janela(self.root, janela, largura, altura)
        janela.grab_set()
        renderizar()

    def _baixar_base_tabelas(self):
        if not BASE_TABELAS.is_file():
            messagebox.showerror(
                "Arquivo não encontrado",
                f"O documento base não foi encontrado em:\n{BASE_TABELAS}",
                parent=self.root,
            )
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar documento base das tabelas",
            initialfile="BASE TABELAS.docx",
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")],
            parent=self.root,
        )
        if not destino:
            return
        try:
            shutil.copyfile(BASE_TABELAS, destino)
        except OSError as exc:
            messagebox.showerror("Erro ao salvar", _mensagem_erro_amigavel(exc), parent=self.root)
            return
        if messagebox.askyesno(
            "Documento salvo",
            f"Documento base salvo em:\n{destino}\n\nAbrir agora no Word?",
            parent=self.root,
        ):
            os.startfile(destino)

    def _escolher_tabela_docx_capas(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar .docx com a tabela de segurados", filetypes=[("Documento Word", "*.docx")],
            parent=self.root,
        )
        if caminho:
            self.var_docx_tabela_capas.set(caminho)
            # a pasta do .docx importado vira a pasta da tese,
            # a não ser que o usuário já tenha escolhido uma manualmente
            if not self._saida_capas_manual:
                self.var_saida_capas.set(str(Path(caminho).parent))

    def _escolher_saida_capas(self):
        caminho = filedialog.askdirectory(title="Selecionar pasta da tese", parent=self.root)
        if caminho:
            self.var_saida_capas.set(caminho)
            self._saida_capas_manual = True

    def _abrir_pasta_capas(self):
        saida = self.var_saida_capas.get().strip()
        if saida and Path(saida).is_dir():
            os.startfile(saida)

    def _abrir_pdf_capas(self):
        if self._ultimo_pdf_capas and self._ultimo_pdf_capas.is_file():
            os.startfile(str(self._ultimo_pdf_capas))

    def _limpar_capas(self):
        self.var_topico_capas.set("")
        self.var_titulo_capas.set("")
        self.var_docx_tabela_capas.set("")
        self._ignorar_modificacao_colar = True
        try:
            self.texto_colar_capas.text.delete("1.0", "end")
            self.texto_colar_capas.text.edit_modified(False)
        finally:
            self._ignorar_modificacao_colar = False
        self._tabela_word_capturada = None
        self.label_status_colar_capas.configure(text="")
        self.label_larguras_capas.configure(text="")
        self.var_origem_tabela_capas.set("docx")
        self._alternar_origem_tabela_capas()
        self.var_saida_capas.set("")
        self._saida_capas_manual = False
        self.var_gerar_docx_capas.set(False)
        self.botao_abrir_pasta_capas.config(state=DISABLED)
        self.botao_abrir_pdf_capas.config(state=DISABLED)
        self._ultimo_pdf_capas = None
        self.frame_progresso_capas.pack_forget()
        self._limpar_log_capas()
        self._mostrar_inicial_capas()

    def _obter_tabela_capas(self):
        if self.var_origem_tabela_capas.get() == "docx":
            caminho_texto = self.var_docx_tabela_capas.get().strip()
            if not caminho_texto:
                raise ValueError("Selecione o .docx com a tabela de segurados.")
            caminho = Path(caminho_texto)
            if not caminho.is_file():
                raise ValueError(f"Arquivo da tabela não encontrado:\n{caminho}")
            return tabela.extrair_de_docx(caminho)
        if self._tabela_word_capturada is not None:
            return self._tabela_word_capturada
        texto = self.texto_colar_capas.text.get("1.0", "end")
        if not texto.strip():
            raise ValueError("Cole a tabela de segurados no campo de texto.")
        return tabela.extrair_de_texto(texto), None

    def _gerar_capas(self):
        if not MODELO_CAPAS.is_file():
            messagebox.showerror(
                "Modelo não encontrado",
                f"O modelo padrão não foi encontrado em:\n{MODELO_CAPAS}",
                parent=self.root,
            )
            return

        topico = self.var_topico_capas.get().strip()
        if not topico:
            messagebox.showwarning("Campo obrigatório", "Informe o número do tópico.", parent=self.root)
            return

        titulo = self.var_titulo_capas.get().strip()
        if not titulo:
            messagebox.showwarning("Campo obrigatório", "Informe o título da tese.", parent=self.root)
            return
        titulo = _normalizar_travessao(_normalizar_espacos(titulo))

        saida_texto = self.var_saida_capas.get().strip()
        pasta_saida = Path(saida_texto) if saida_texto else SAIDA_PADRAO_CAPAS

        try:
            tabela_segurados, tabela_referencia = self._obter_tabela_capas()
        except ValueError as exc:
            messagebox.showwarning("Tabela de segurados", str(exc), parent=self.root)
            return

        # avisa antes de sobrescrever um PDF/Word de capas já gerado
        # (sempre salvo como "Capas Geradas", igual à aba All-in-one)
        nome_base = "Capas Geradas"
        existentes = [
            caminho.name
            for caminho in (pasta_saida / f"{nome_base}.pdf", pasta_saida / f"{nome_base}.docx")
            if caminho.exists() and (caminho.suffix == ".pdf" or self.var_gerar_docx_capas.get())
        ]
        if existentes:
            lista = "\n".join(f"  • {nome}" for nome in existentes)
            verbo = "já existe e será SUBSTITUÍDO" if len(existentes) == 1 else "já existem e serão SUBSTITUÍDOS"
            if not messagebox.askyesno(
                "Arquivos serão substituídos",
                f"{_plural(len(existentes), 'arquivo', 'arquivos')} {verbo} na pasta da tese:\n\n{lista}\n\nContinuar?",
                icon="warning",
                parent=self.root,
            ):
                return

        n_segurados = len(tabela_segurados.grupos)
        n_beneficios = len(tabela_segurados.linhas)

        self._limpar_log_capas()
        self._log_capas(f"Tabela lida: {n_segurados} segurado(s), {n_beneficios} linha(s) de benefício.", "titulo")
        if tabela_referencia is not None:
            self._log_capas("Estilo da tabela original será replicado (mesclagens, cores, bordas e fontes).")
        else:
            self._log_capas("Texto colado não tem formatação — será aplicado o estilo padrão da capa.")

        larguras_colunas = None
        if self.var_origem_tabela_capas.get() == "colar":
            larguras_colunas = self._larguras_para(len(tabela_segurados.cabecalhos))
            if larguras_colunas:
                self._log_capas("Larguras de coluna personalizadas serão aplicadas.")

        self.botao_gerar_capas.config(state=DISABLED)
        self.botao_abrir_pasta_capas.config(state=DISABLED)
        self.botao_abrir_pdf_capas.config(state=DISABLED)
        self.barra_progresso_capas.configure(value=0, maximum=2 + n_segurados)
        self.label_progresso_capas.configure(text="Iniciando geração...")
        self.frame_progresso_capas.pack(fill=X, pady=(0, 14), before=self.moldura_resultado_capas)

        threading.Thread(
            target=self._gerar_capas_worker,
            args=(titulo, topico, tabela_segurados, tabela_referencia, pasta_saida,
                  self.var_gerar_docx_capas.get(), larguras_colunas),
            daemon=True,
        ).start()

    def _gerar_capas_worker(self, titulo, topico, tabela_segurados, tabela_referencia, pasta_saida, gerar_docx,
                            larguras_colunas=None):
        pythoncom.CoInitialize()
        try:
            try:
                conversor = ConversorPDF()
            except ConversorPDFIndisponivel as exc:
                self.root.after(0, self._capas_falhou, str(exc))
                return

            try:
                # sempre salvo como "Capas Geradas" (mesmo padrão da aba
                # All-in-one), no lugar do título da tese
                caminho_pdf, caminho_docx = gerar_pdf_capas(
                    MODELO_CAPAS, titulo, tabela_segurados, tabela_referencia, pasta_saida,
                    conversor, topico=topico, gerar_docx=gerar_docx,
                    callback_progresso=lambda e, t, r: self.root.after(0, self._capas_progresso, e, t, r),
                    larguras_colunas=larguras_colunas,
                )
            except Exception as exc:
                self.root.after(0, self._capas_falhou, _mensagem_erro_amigavel(exc))
                return
            finally:
                conversor.fechar()

            self.root.after(0, self._capas_concluiu, caminho_pdf, caminho_docx)
        finally:
            pythoncom.CoUninitialize()

    def _capas_progresso(self, etapa: int, total: int, rotulo: str):
        self.barra_progresso_capas.configure(value=etapa, maximum=total)
        self.label_progresso_capas.configure(text=f"Gerando {etapa} de {total}: {rotulo}")
        self._log_capas(f"[{etapa}/{total}] {rotulo}", "item")

    def _restaurar_botao_capas(self):
        self.botao_gerar_capas.config(state=NORMAL)
        self.frame_progresso_capas.pack_forget()

    def _capas_falhou(self, mensagem: str):
        self._restaurar_botao_capas()
        self._log_capas(f"ERRO: {mensagem}", "erro")
        messagebox.showerror("Erro ao gerar capas", mensagem, parent=self.root)

    def _capas_concluiu(self, caminho_pdf: Path, caminho_docx: Path | None):
        self._restaurar_botao_capas()
        self._ultimo_pdf_capas = caminho_pdf
        self.botao_abrir_pasta_capas.config(state=NORMAL)
        self.botao_abrir_pdf_capas.config(state=NORMAL)
        self._log_capas(f"\nPDF gerado em:\n{caminho_pdf}", "titulo")
        if caminho_docx is not None:
            self._log_capas(f"Word gerado em:\n{caminho_docx}", "titulo")

def main():
    root = ttk.Window(themename="litera", iconphoto=None)
    try:
        AplicativoDivisorPDF(root)
    except Exception as exc:
        messagebox.showerror("Erro ao iniciar", f"Não foi possível iniciar o aplicativo:\n{exc}", parent=root)
        raise
    root.mainloop()


if __name__ == "__main__":
    main()
