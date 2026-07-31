# -*- coding: utf-8 -*-
"""
Launcher do ANEXT - Anexador de Teses, com auto-atualização via GitHub
Releases.

Fluxo:
  1. Lê a versão instalada (versao.txt)
  2. Consulta a API do GitHub Releases (releases/latest)
  3. Sem update (ou sem internet/GitHub fora) -> abre o ANEXT.exe e fecha, sem erro
  4. Com update -> mostra janela (versão atual, nova, progresso, botão Pular)
     - baixa o asset para arquivo temporário
     - valida o tamanho baixado
     - faz backup do ANEXT.exe atual
     - substitui o ANEXT.exe e atualiza versao.txt
     - abre o ANEXT.exe atualizado e fecha o launcher
  5. Qualquer falha após o backup -> restaura o backup e abre a versão anterior
     Falha antes do backup (download/validação) -> nada foi tocado, só abre a versão existente
     Backup inexistente numa restauração -> avisa o usuário e não mexe em nada
"""
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

import requests
import ttkbootstrap as ttk
from PIL import ImageTk
from ttkbootstrap.constants import BOTH, X
from tkinter import messagebox

import tema
import visual

# ---------------------------------------------------------------------------
# Configuração - repositório do ANEXT no GitHub
# ---------------------------------------------------------------------------
GITHUB_OWNER = "B0NASSI"
GITHUB_REPO = "ANEXT---Anexador-de-Teses"
ASSET_NAME = "ANEXT.exe"        # nome do arquivo anexado na release
REQUEST_TIMEOUT = 10             # segundos para consultas de rede
DOWNLOAD_CHUNK_SIZE = 65536

LARGURA_JANELA = 460
ALTURA_BANNER = 78


def _caminho_recurso(nome: str) -> str:
    # em modo congelado (.exe), os recursos ficam soltos em _MEIPASS; em modo
    # de desenvolvimento, launcher.py mora em src/ mas os recursos (assets/)
    # ficam na raiz do projeto, um nível acima
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return str(base / nome)


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE_DIR = get_base_dir()
APP_EXE = BASE_DIR / "ANEXT.exe"
BACKUP_EXE = BASE_DIR / "ANEXT_backup.exe"
VERSION_FILE = BASE_DIR / "versao.txt"


# ---------------------------------------------------------------------------
# Versão local / comparação
# ---------------------------------------------------------------------------
def read_local_version() -> str:
    try:
        return VERSION_FILE.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return "0.0.0"


def parse_version(version: str):
    version = version.strip().lstrip("vV")
    parts = version.split(".")
    nums = []
    for part in parts:
        digits = "".join(ch for ch in part if ch.isdigit())
        nums.append(int(digits) if digits else 0)
    return tuple(nums) if nums else (0,)


def is_newer(remote: str, local: str) -> bool:
    r, l = parse_version(remote), parse_version(local)
    length = max(len(r), len(l))
    r = r + (0,) * (length - len(r))
    l = l + (0,) * (length - len(l))
    return r > l


# ---------------------------------------------------------------------------
# GitHub API
# ---------------------------------------------------------------------------
def get_latest_release():
    """Retorna dict com {tag_name, asset_url, asset_size} ou None se indisponível."""
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    headers = {"Accept": "application/vnd.github+json", "User-Agent": "anext-auto-updater"}
    try:
        response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        data = response.json()
    except (requests.RequestException, ValueError):
        return None

    tag_name = data.get("tag_name")
    if not tag_name:
        return None

    asset_url = None
    asset_size = None
    for asset in data.get("assets", []):
        if asset.get("name") == ASSET_NAME:
            asset_url = asset.get("browser_download_url")
            asset_size = asset.get("size")
            break

    if not asset_url:
        return None

    return {"tag_name": tag_name, "asset_url": asset_url, "asset_size": asset_size}


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------
def download_asset(url: str, dest: Path, expected_size, progress_callback):
    """Baixa a URL para dest em streaming, chamando progress_callback(percent|None)."""
    headers = {"User-Agent": "anext-auto-updater"}
    with requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT, stream=True) as response:
        response.raise_for_status()
        total = int(response.headers.get("Content-Length") or 0) or expected_size
        downloaded = 0
        with open(dest, "wb") as f:
            for chunk in response.iter_content(chunk_size=DOWNLOAD_CHUNK_SIZE):
                if not chunk:
                    continue
                f.write(chunk)
                downloaded += len(chunk)
                if total:
                    progress_callback(downloaded / total * 100, downloaded, total)
                else:
                    progress_callback(None, downloaded, None)

    actual_size = dest.stat().st_size
    if actual_size == 0:
        raise IOError("Arquivo baixado está vazio.")
    if expected_size and actual_size != expected_size:
        raise IOError(f"Tamanho do download não confere (esperado {expected_size}, obtido {actual_size}).")


# ---------------------------------------------------------------------------
# Aplicar atualização
# ---------------------------------------------------------------------------
def apply_update(tmp_path: Path, new_version: str):
    """Faz backup, substitui o ANEXT.exe e atualiza versao.txt.

    Se falhar após o backup ser criado, restaura o backup automaticamente.
    """
    backup_created = False
    try:
        if APP_EXE.exists():
            shutil.copy2(APP_EXE, BACKUP_EXE)
            backup_created = True

        os.replace(tmp_path, APP_EXE)
        VERSION_FILE.write_text(new_version, encoding="utf-8")
    except Exception:
        if backup_created and BACKUP_EXE.exists():
            shutil.copy2(BACKUP_EXE, APP_EXE)
        elif not BACKUP_EXE.exists():
            messagebox.showwarning(
                "Atualização",
                "Falha ao aplicar a atualização e nenhum backup foi encontrado.\n"
                "Nada foi substituído.",
            )
        raise


def log_error(context: str, exc: Exception):
    import datetime
    import traceback

    log_path = BASE_DIR / "launcher_log.txt"
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now().isoformat()}] {context}\n")
            f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except OSError:
        pass


def launch_app():
    if APP_EXE.exists():
        subprocess.Popen([str(APP_EXE)], cwd=str(BASE_DIR))
    else:
        messagebox.showerror("Erro", "ANEXT.exe não encontrado.")


# ---------------------------------------------------------------------------
# Interface gráfica (só aparece se houver atualização)
# ---------------------------------------------------------------------------
class UpdaterUI:
    def __init__(self, local_version, remote_version, release):
        self.release = release
        self.skipped = False
        self._indeterminate_running = False

        self.root = ttk.Window(themename="litera", iconphoto=None)
        tema.aplicar(self.root)
        self.root.title("ANEXT - Atualização disponível")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.on_skip)
        try:
            self.root.iconbitmap(_caminho_recurso("assets/pdf.ico"))
        except Exception:
            pass

        self._imagem_banner = ImageTk.PhotoImage(
            visual.gerar_banner(
                largura=LARGURA_JANELA,
                altura=ALTURA_BANNER,
                cor_inicio=tema.COR_PRIMARIA,
                cor_fim="#1A1C3D",
                cor_destaque=tema.COR_SECUNDARIA,
                icone_path=_caminho_recurso("assets/pdf.ico"),
                titulo="Nova versão disponível",
                subtitulo="ANEXT - Anexador de Teses",
            )
        )
        ttk.Label(self.root, image=self._imagem_banner, borderwidth=0).pack(fill=X)

        corpo = ttk.Frame(self.root, padding=(24, 18, 24, 20))
        corpo.pack(fill=BOTH, expand=True)

        ttk.Label(corpo, text=f"Versão atual: {local_version}", font=("Segoe UI", 10)).pack(anchor="w")
        ttk.Label(corpo, text=f"Nova versão: {remote_version}", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(2, 14)
        )

        self.progress = ttk.Progressbar(corpo, orient="horizontal", mode="determinate", bootstyle="secondary")
        self.progress.pack(fill=X, pady=(0, 8))

        self.status_label = ttk.Label(corpo, text="Baixando atualização...", bootstyle="secondary")
        self.status_label.pack(anchor="w")

        self.skip_button = ttk.Button(corpo, text="Pular", command=self.on_skip, bootstyle="primary-outline")
        self.skip_button.pack(pady=(14, 0))

    def run(self):
        thread = threading.Thread(target=self._download_and_apply, daemon=True)
        thread.start()
        self.root.mainloop()

    def on_skip(self):
        self.skipped = True
        self.root.destroy()
        launch_app()

    def set_progress(self, percent, downloaded=None, total=None):
        if percent is None:
            self.progress.configure(mode="indeterminate")
            if not self._indeterminate_running:
                self.progress.start(10)
                self._indeterminate_running = True
        else:
            self.progress.configure(mode="determinate")
            self.progress.stop()
            self._indeterminate_running = False
            self.progress["value"] = percent

        if downloaded is not None:
            mb_downloaded = downloaded / (1024 * 1024)
            if total:
                mb_total = total / (1024 * 1024)
                self.set_status(f"Baixando atualização... {percent:.0f}% ({mb_downloaded:.1f} MB / {mb_total:.1f} MB)")
            else:
                self.set_status(f"Baixando atualização... {mb_downloaded:.1f} MB")
        self.root.update_idletasks()

    def set_status(self, text):
        self.status_label.configure(text=text)
        self.root.update_idletasks()

    def finish(self, success: bool, message: str = ""):
        if self.skipped:
            return
        self.root.destroy()
        if not success and message:
            messagebox.showerror("Atualização", message)
        launch_app()

    def _download_and_apply(self):
        tmp_path = None
        try:
            fd, tmp_name = tempfile.mkstemp(dir=str(BASE_DIR), suffix=".tmp")
            os.close(fd)
            tmp_path = Path(tmp_name)

            download_asset(
                self.release["asset_url"],
                tmp_path,
                self.release.get("asset_size"),
                lambda p, d, t: self.root.after(0, self.set_progress, p, d, t),
            )
            self.root.after(0, self.set_status, "Aplicando atualização...")

            apply_update(tmp_path, self.release["tag_name"])

            self.root.after(0, self.finish, True, "")
        except Exception as exc:
            log_error("Falha ao baixar/aplicar atualização", exc)
            if tmp_path and tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass
            self.root.after(
                0,
                self.finish,
                False,
                f"Não foi possível concluir a atualização. A versão anterior será aberta.\n\nDetalhes: {exc}",
            )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    local_version = read_local_version()
    release = get_latest_release()

    if release is None:
        # Sem internet, GitHub fora do ar, ou sem asset compatível: segue sem erro.
        launch_app()
        return

    if not is_newer(release["tag_name"], local_version):
        launch_app()
        return

    ui = UpdaterUI(local_version, release["tag_name"], release)
    ui.run()


if __name__ == "__main__":
    main()
