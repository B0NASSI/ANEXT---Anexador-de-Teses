# -*- coding: utf-8 -*-
"""Automatiza compilar + empacotar + checksum de uma release do ANEXT, com
uma conferência que faltava: garante que a nota de atualização da versão
atual está de fato dentro do zip gerado antes de deixar prosseguir.

Isso existe porque já aconteceu de reaproveitar um build antigo (feito
antes de escrever a nota da versão) pra empacotar uma release — o app lê
"NOTAS DE ATUALIZAÇÃO/" de dentro do próprio .exe, então isso silenciosamente
publica uma release sem a nota, sem erro nenhum na hora.

NÃO compila nem publica nada sozinho além disso — não roda `gh release
create`. Só deixa o .zip e o .sha256 prontos em installer/release/, e
imprime no final o comando pra publicar (pra rodar manualmente, ou pedir
pra alguém revisar/publicar).

Uso (de dentro de ANEXT/, com o venv ativo):
    .venv\\Scripts\\python.exe installer\\build_release.py
"""
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SPEC = RAIZ / "ANEXT - Anexador de Teses.spec"
DIST_APP = RAIZ / "dist" / "ANEXT"
RELEASE_DIR = RAIZ / "installer" / "release"
ZIP_PATH = RELEASE_DIR / "ANEXT-app.zip"


def ler_versao() -> str:
    return (RAIZ / "versao.txt").read_text(encoding="utf-8-sig").strip()


def checar_nota_existe(versao: str) -> None:
    caminho = RAIZ / "NOTAS DE ATUALIZAÇÃO" / f"{versao}.txt"
    if not caminho.is_file():
        sys.exit(
            f'ERRO: não existe "NOTAS DE ATUALIZAÇÃO/{versao}.txt" para a versão '
            f"atual (versao.txt = {versao}). Crie a nota ANTES de compilar — ela "
            "precisa existir antes do build pra ficar embutida no .exe."
        )


def compilar() -> None:
    print("Compilando (PyInstaller)...")
    shutil.rmtree(RAIZ / "build", ignore_errors=True)
    shutil.rmtree(DIST_APP, ignore_errors=True)
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", str(SPEC), "--noconfirm"],
        cwd=str(RAIZ), check=True,
    )


def empacotar() -> None:
    print("Empacotando...")
    if not (DIST_APP / "ANEXT.exe").is_file():
        sys.exit(f"ERRO: {DIST_APP / 'ANEXT.exe'} não existe — o build falhou ou não rodou.")
    RELEASE_DIR.mkdir(parents=True, exist_ok=True)
    ZIP_PATH.unlink(missing_ok=True)
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for arquivo in DIST_APP.rglob("*"):
            if arquivo.is_file():
                zf.write(arquivo, arquivo.relative_to(DIST_APP))


def conferir_zip(versao: str) -> None:
    with zipfile.ZipFile(ZIP_PATH) as zf:
        nomes = zf.namelist()
        tem_exe = "ANEXT.exe" in nomes
        tem_internal = any(n.startswith("_internal/") for n in nomes)
        alvo_nota = f"NOTAS DE ATUALIZAÇÃO/{versao}.txt"
        tem_nota = any(n.endswith(alvo_nota) for n in nomes)

    if not (tem_exe and tem_internal):
        sys.exit("ERRO: o zip não tem a estrutura esperada (ANEXT.exe + _internal/). Não publique.")
    if not tem_nota:
        sys.exit(
            f"ERRO: o zip NÃO contém a nota da versão {versao}. O build deve ter sido "
            "feito antes de criar a nota, ou algo mudou no .spec. Não publique este zip."
        )
    print(f"OK: estrutura do zip e nota da versão {versao} confirmadas.")


def gerar_checksum() -> None:
    print("Gerando checksum...")
    resultado = subprocess.run(
        [sys.executable, str(RAIZ / "installer" / "gerar_checksum.py"), str(ZIP_PATH)],
        cwd=str(RAIZ), check=True, capture_output=True, text=True, encoding="utf-8",
    )
    print(resultado.stdout.strip())


def main() -> None:
    versao = ler_versao()
    print(f"Versão em versao.txt: {versao}")
    checar_nota_existe(versao)
    compilar()
    empacotar()
    conferir_zip(versao)
    gerar_checksum()
    print()
    print("Tudo pronto em installer/release/. Revise e, se estiver tudo certo, publique com:")
    print(
        f'  gh release create {versao} "installer/release/ANEXT-app.zip" '
        f'"installer/release/ANEXT-app.zip.sha256" --repo B0NASSI/ANEXT---Anexador-de-Teses '
        f'--title "v{versao}" --notes-file "NOTAS DE ATUALIZAÇÃO/{versao}.txt"'
    )


if __name__ == "__main__":
    main()
