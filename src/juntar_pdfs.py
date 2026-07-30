"""
Junta os PDFs de uma pasta em um único arquivo, na ordem numérica do nome
(ex: "0.", "1.", "3.", "3.1", "3.2", "4.", "5.", "5.1"...). Subpastas,
arquivos que não sejam PDF e arquivos sem numeração no início do nome
(ex: "CAT 2022 - Mohawk.pdf") são ignorados.

Modo tese completa: dado uma pasta "mãe" com subpastas numeradas por
segurado (ex: "1. NOME", "2. NOME"...), percorre as subpastas em ordem
numérica e, dentro de cada uma, os arquivos em ordem numérica, gerando
um único PDF com tudo (a tese inteira), não um por segurado.
"""

import re
import time
from pathlib import Path

import fitz  # PyMuPDF

PADRAO_NUMERADO = re.compile(r"\s*\d+\s*\.")


def _com_repeticao(acao, tentativas: int = 4, espera: float = 0.5):
    # PDFs recem-criados podem ficar momentaneamente bloqueados por
    # antivirus/indexacao/sincronizacao (OneDrive) - tentar de novo depois
    # de uma pequena espera resolve a maioria desses casos transitorios.
    for tentativa in range(tentativas):
        try:
            return acao()
        except Exception:
            if tentativa == tentativas - 1:
                raise
            time.sleep(espera)


def _comeca_com_numero(nome: str) -> bool:
    return bool(PADRAO_NUMERADO.match(nome))


def _chave_ordenacao(caminho: Path) -> tuple[float, str]:
    match = re.match(r"\s*(\d+(?:\.\d+)?)", caminho.stem)
    numero = float(match.group(1)) if match else float("inf")
    return (numero, caminho.stem.lower())


def pdfs_em_ordem(pasta: Path) -> list[Path]:
    arquivos = [
        item
        for item in pasta.iterdir()
        if item.is_file() and item.suffix.lower() == ".pdf" and _comeca_com_numero(item.stem)
    ]
    arquivos.sort(key=_chave_ordenacao)
    return arquivos


def subpastas_numeradas(pasta_mae: Path) -> list[Path]:
    candidatas = [item for item in pasta_mae.iterdir() if item.is_dir() and _comeca_com_numero(item.name)]
    candidatas.sort(key=_chave_ordenacao)
    return candidatas


def nome_saida(origem: Path) -> str:
    return f"{origem.name}.pdf"


def validar_e_contar_paginas(arquivos: list[Path]) -> tuple[int, list[str]]:
    """Abre cada PDF só para validar que é legível e contar páginas (não insere nada).
    Retorna (total_de_paginas, lista_de_erros) — erros vazios significa que todos abriram bem."""
    total = 0
    erros = []
    for arquivo in arquivos:
        try:
            with _com_repeticao(lambda: fitz.open(arquivo)) as doc:
                total += doc.page_count
        except Exception as exc:
            erros.append(f'"{arquivo.name}": {exc}')
    return total, erros


def juntar_pasta(pasta: Path, pasta_saida: Path, progresso_callback=None) -> Path:
    arquivos = pdfs_em_ordem(pasta)
    if not arquivos:
        raise ValueError(f'Nenhum PDF encontrado em "{pasta.name}".')

    pasta_saida.mkdir(parents=True, exist_ok=True)
    combinado = fitz.open()
    try:
        total = len(arquivos)
        for indice, arquivo in enumerate(arquivos, start=1):
            with _com_repeticao(lambda: fitz.open(arquivo)) as doc:
                combinado.insert_pdf(doc)
            if progresso_callback:
                progresso_callback(indice, total)
        caminho_saida = pasta_saida / nome_saida(pasta)
        _com_repeticao(lambda: combinado.save(caminho_saida))
    finally:
        combinado.close()
    return caminho_saida


def juntar_tese(pasta_mae: Path, pasta_saida: Path, progresso_callback=None, cancelar=None) -> Path:
    subpastas = subpastas_numeradas(pasta_mae)
    if not subpastas:
        raise ValueError('Nenhuma subpasta numerada encontrada (ex: "1. NOME").')

    arquivos_em_ordem = []
    for subpasta in subpastas:
        arquivos_em_ordem.extend(pdfs_em_ordem(subpasta))

    if not arquivos_em_ordem:
        raise ValueError("Nenhuma das subpastas continha arquivos PDF.")

    pasta_saida.mkdir(parents=True, exist_ok=True)
    combinado = fitz.open()
    try:
        total = len(arquivos_em_ordem)
        for indice, arquivo in enumerate(arquivos_em_ordem, start=1):
            if cancelar is not None and cancelar.is_set():
                raise InterruptedError("Cancelado pelo usuário.")
            with _com_repeticao(lambda: fitz.open(arquivo)) as doc:
                combinado.insert_pdf(doc)
            if progresso_callback:
                progresso_callback(indice, total)
        caminho_saida = pasta_saida / nome_saida(pasta_mae)
        _com_repeticao(lambda: combinado.save(caminho_saida))
    finally:
        combinado.close()
    return caminho_saida
