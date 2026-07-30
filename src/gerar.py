# -*- coding: utf-8 -*-
"""
Orquestra a geração do PDF final: para cada página (capa geral, tabela
completa, uma por segurado) monta uma cópia do modelo já com o conteúdo
substituído, converte para PDF via Word (uma única instância do Word fica
aberta durante todo o lote) e junta todas as páginas com fitz. Os .docx/.pdf
intermediários ficam numa pasta temporária, apagada ao final.
"""

import re
import shutil
import tempfile
import time
from pathlib import Path
from typing import Callable

import fitz

from documento import (
    carregar_catalogo_tabelas,
    mesclar_documentos,
    montar_capa,
    montar_pagina_segurado,
    montar_tabela_completa,
)
from pdf import ConversorPDF
from tabela import TabelaSegurados

CallbackProgresso = Callable[[int, int, str], None]


def _com_repeticao(acao: Callable[[], None], tentativas: int = 4, espera: float = 0.5) -> None:
    # PDFs recém-criados podem ficar momentaneamente bloqueados por
    # antivírus/indexação/sincronização (OneDrive) - uma nova tentativa
    # após uma pequena espera resolve a maioria desses casos transitórios.
    for tentativa in range(tentativas):
        try:
            acao()
            return
        except Exception:
            if tentativa == tentativas - 1:
                raise
            time.sleep(espera)


def sanitizar_nome_arquivo(texto: str) -> str:
    # tab/quebra de linha/outros caracteres de controle não estão na lista
    # de caracteres reservados do Windows, mas quebram operações de
    # arquivo (rename/remove) de forma imprevisível — normaliza pra espaço
    # antes de remover só os caracteres realmente proibidos.
    sem_controle = re.sub(r'[\x00-\x1f\x7f]+', ' ', texto)
    sem_invalidos = re.sub(r'[\\/:*?"<>|]', "", sem_controle)
    colapsado = re.sub(r'\s+', ' ', sem_invalidos).strip()
    limpo = colapsado.rstrip(".")
    return limpo or "Capa FAP"


def gerar_pdf_capas(
    caminho_modelo: Path,
    titulo: str,
    tabela: TabelaSegurados,
    tabela_referencia,
    pasta_saida: Path,
    conversor: ConversorPDF,
    topico: str | None = None,
    gerar_docx: bool = False,
    callback_progresso: CallbackProgresso | None = None,
    larguras_colunas: list[float] | None = None,
    cancelar=None,
) -> tuple[Path, Path | None]:
    if not tabela.grupos:
        raise ValueError("Nenhum segurado foi identificado na tabela informada.")

    carregar_catalogo_tabelas(Path(caminho_modelo).parent / "TABELAS.docx")

    # Com um único segurado, a "tabela completa" já mostra o nome dele na
    # única linha - uma página individual repetiria a mesma informação, então
    # ela é pulada (a tabela completa cumpre esse papel sozinha).
    gerar_paginas_individuais = len(tabela.grupos) > 1

    pasta_temp = Path(tempfile.mkdtemp(prefix="capas_fap_"))
    total_etapas = 2 + (len(tabela.grupos) if gerar_paginas_individuais else 0)
    etapa = 0
    pdf_final = fitz.open()
    documentos = []
    try:
        def _converter_e_anexar(doc, prefixo: str) -> None:
            if gerar_docx:
                documentos.append(doc)
            docx_tmp = pasta_temp / f"{prefixo}.docx"
            pdf_tmp = pasta_temp / f"{prefixo}.pdf"
            doc.save(str(docx_tmp))
            conversor.converter(docx_tmp, pdf_tmp)

            def _anexar():
                with fitz.open(pdf_tmp) as paginas:
                    pdf_final.insert_pdf(paginas)

            _com_repeticao(_anexar)

        doc = montar_capa(caminho_modelo, titulo, topico)
        _converter_e_anexar(doc, "00_capa")
        etapa += 1
        if callback_progresso:
            callback_progresso(etapa, total_etapas, "Capa geral")

        doc = montar_tabela_completa(
            caminho_modelo, titulo, tabela.cabecalhos, tabela.linhas, tabela.indice_nome, tabela_referencia,
            larguras_colunas,
        )
        _converter_e_anexar(doc, "01_tabela_completa")
        etapa += 1
        if callback_progresso:
            callback_progresso(etapa, total_etapas, "Tabela completa")

        if gerar_paginas_individuais:
            for i, (nome, linhas_segurado) in enumerate(tabela.grupos, start=1):
                if cancelar is not None and cancelar.is_set():
                    raise InterruptedError("Cancelado pelo usuário.")
                doc = montar_pagina_segurado(
                    caminho_modelo, nome, linhas_segurado, tabela.cabecalhos, tabela.indice_nome, tabela_referencia,
                    larguras_colunas,
                )
                _converter_e_anexar(doc, f"{i + 1:03d}_{sanitizar_nome_arquivo(nome)[:40]}")
                etapa += 1
                if callback_progresso:
                    callback_progresso(etapa, total_etapas, nome)

        # nome fixo (curto) em vez do título da tese: os dois lugares que
        # chamam essa função sempre salvam como "Capas Geradas" de qualquer
        # forma - usar o título aqui só arriscava estourar o limite de 260
        # caracteres do Windows em pastas de tese já bem aninhadas.
        pasta_saida.mkdir(parents=True, exist_ok=True)
        caminho_pdf = pasta_saida / "Capas Geradas.pdf"
        _com_repeticao(lambda: pdf_final.save(str(caminho_pdf)))

        caminho_docx = None
        if gerar_docx:
            caminho_docx = pasta_saida / "Capas Geradas.docx"
            mesclar_documentos(documentos).save(str(caminho_docx))
    finally:
        pdf_final.close()
        shutil.rmtree(pasta_temp, ignore_errors=True)

    return caminho_pdf, caminho_docx
