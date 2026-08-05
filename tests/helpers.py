# -*- coding: utf-8 -*-
"""Fixtures sintéticas para os testes — sem depender de ANEXT/Exemplos/
(gitignorado por ter nome/CNPJ/NIT reais de clientes)."""
import fitz


def criar_pdf_paginas(caminho, paginas: list[list[str]]) -> None:
    """Cria um PDF com uma página por item de `paginas`; cada item é a lista
    de linhas de texto dessa página, empilhadas de cima para baixo — o
    suficiente para bater com a leitura de `page.get_text()` linha a linha
    usada por separar_capas.py e juntar_pdfs.py."""
    doc = fitz.open()
    try:
        for linhas in paginas:
            pagina = doc.new_page()
            y = 72
            for linha in linhas:
                pagina.insert_text((72, y), linha, fontsize=12)
                y += 20
        doc.save(str(caminho))
    finally:
        doc.close()
