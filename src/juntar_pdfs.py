"""
Junta os PDFs de uma pasta em um único arquivo, na ordem numérica do nome
(ex: "0.", "1.", "3.", "3.1", "3.2", "4.", "5.", "5.1"...). Subpastas,
arquivos que não sejam PDF e arquivos sem numeração no início do nome
(ex: "CAT 2022 - Mohawk.pdf") são ignorados.

Modo tese completa: dado uma pasta "mãe" com subpastas numeradas por
segurado (ex: "1. NOME", "2. NOME"...), percorre as subpastas em ordem
numérica e, dentro de cada uma, os arquivos em ordem numérica, gerando
um único PDF com tudo (a tese inteira), não um por segurado.

Toda junção registra no log (anext.log), documento por documento e no
momento exato em que ele é inserido, a ordem real usada: posição, pasta,
arquivo, páginas que ele ocupou no PDF final, tamanho e o SHA256 do arquivo
de origem. É o registro que comprova a ordem em que os documentos foram
juntados (a ordem vem só da numeração do nome — ver _chave_ordenacao).
"""

import hashlib
import logging
import re
import time
from pathlib import Path

import fitz  # PyMuPDF

from diagnostico import metadados_anext

PADRAO_NUMERADO = re.compile(r"\s*\d+\s*\.")

logger = logging.getLogger(__name__)


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


def ultimo_documento(pasta_mae: Path) -> Path | None:
    """Devolve o último PDF na mesma ordem que `juntar_tese` usaria pra
    juntar a tese inteira, ou None se não houver nenhum. Usado para checar
    se um documento (ex.: anexo fixo, ver anexo_fixo.py) já foi incluído
    manualmente antes de anexar de novo."""
    for subpasta in reversed(subpastas_numeradas(pasta_mae)):
        arquivos = pdfs_em_ordem(subpasta)
        if arquivos:
            return arquivos[-1]
    return None


def _sha256(caminho: Path) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def _inserir_registrando(combinado, arquivo: Path, posicao: int, total: int, base: Path) -> None:
    """Insere `arquivo` no fim de `combinado` e registra no log a posição e
    as páginas exatas que ele passou a ocupar no PDF final."""
    pagina_inicial = combinado.page_count + 1
    with _com_repeticao(lambda: fitz.open(arquivo)) as doc:
        combinado.insert_pdf(doc)
    try:
        nome = arquivo.relative_to(base).as_posix()
    except ValueError:  # anexo fixo, que não mora dentro da pasta juntada
        nome = f"(anexo fixo) {arquivo.name}"
    logger.info(
        "  [%d/%d] %s → páginas %d-%d (%d pág.) | %d bytes | sha256 %s",
        posicao, total, nome, pagina_inicial, combinado.page_count,
        combinado.page_count - pagina_inicial + 1, arquivo.stat().st_size, _sha256(arquivo),
    )


def _registrar_inicio(origem: Path, arquivos: list[Path], extra: Path | None = None) -> None:
    logger.info(
        "Ordem de junção (%d documento(s)%s, origem=%s) — critério: numeração no início do nome "
        "das subpastas e dos arquivos:",
        len(arquivos), " + anexo fixo" if extra is not None else "", origem,
    )


def _registrar_fim(caminho_saida: Path, combinado) -> None:
    logger.info(
        "PDF juntado salvo: %s (%d páginas, sha256 %s)", caminho_saida, combinado.page_count,
        _sha256(caminho_saida),
    )


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
        _registrar_inicio(pasta, arquivos)
        for indice, arquivo in enumerate(arquivos, start=1):
            _inserir_registrando(combinado, arquivo, indice, total, pasta)
            if progresso_callback:
                progresso_callback(indice, total)
        caminho_saida = pasta_saida / nome_saida(pasta)
        combinado.set_metadata(metadados_anext())
        _com_repeticao(lambda: combinado.save(caminho_saida))
        _registrar_fim(caminho_saida, combinado)
    finally:
        combinado.close()
    return caminho_saida


def juntar_tese(
    pasta_mae: Path, pasta_saida: Path, progresso_callback=None, cancelar=None,
    anexo_extra: Path | None = None,
) -> Path:
    """anexo_extra: PDF fixo a acrescentar como últimas páginas, depois de
    todos os documentos das subpastas (ver `anexo_fixo.py` — usado para
    teses que sempre levam um mesmo documento de referência ao final)."""
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
        total_registro = total + (1 if anexo_extra is not None else 0)
        _registrar_inicio(pasta_mae, arquivos_em_ordem, anexo_extra)
        for indice, arquivo in enumerate(arquivos_em_ordem, start=1):
            if cancelar is not None and cancelar.is_set():
                logger.info("Junção cancelada pelo usuário antes do documento %d/%d.", indice, total)
                raise InterruptedError("Cancelado pelo usuário.")
            _inserir_registrando(combinado, arquivo, indice, total_registro, pasta_mae)
            if progresso_callback:
                progresso_callback(indice, total)
        if anexo_extra is not None:
            _inserir_registrando(combinado, anexo_extra, total_registro, total_registro, pasta_mae)
        caminho_saida = pasta_saida / nome_saida(pasta_mae)
        combinado.set_metadata(metadados_anext())
        _com_repeticao(lambda: combinado.save(caminho_saida))
        _registrar_fim(caminho_saida, combinado)
    finally:
        combinado.close()
    return caminho_saida
