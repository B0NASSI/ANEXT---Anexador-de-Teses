# -*- coding: utf-8 -*-
import fitz

from helpers import criar_pdf_paginas
from juntar_pdfs import juntar_pasta, juntar_tese


def _textos_das_paginas(caminho):
    with fitz.open(caminho) as doc:
        return [pagina.get_text().strip() for pagina in doc]


def test_juntar_pasta_respeita_ordem_numerica_e_ignora_nao_numerados(tmp_path):
    pasta = tmp_path / "tese"
    pasta.mkdir()
    criar_pdf_paginas(pasta / "0. capa.pdf", [["MARCA_0"]])
    criar_pdf_paginas(pasta / "1. tabela.pdf", [["MARCA_1A"], ["MARCA_1B"]])
    criar_pdf_paginas(pasta / "2. segurado.pdf", [["MARCA_2"]])
    criar_pdf_paginas(pasta / "CAT 2022 - Mohawk.pdf", [["NAO_ENTRA"]])  # sem número no início: ignorado
    (pasta / "subpasta_ignorada").mkdir()

    saida = juntar_pasta(pasta, tmp_path / "saida")

    assert _textos_das_paginas(saida) == ["MARCA_0", "MARCA_1A", "MARCA_1B", "MARCA_2"]


def test_juntar_tese_percorre_subpastas_numeradas_em_ordem(tmp_path):
    pasta_mae = tmp_path / "tese completa"
    pasta_mae.mkdir()
    seg1 = pasta_mae / "1. FULANO DE TAL"
    seg2 = pasta_mae / "2. CICLANO DA SILVA"
    seg1.mkdir()
    seg2.mkdir()
    criar_pdf_paginas(seg1 / "0. capa.pdf", [["SEG1_CAPA"]])
    criar_pdf_paginas(seg2 / "0. capa.pdf", [["SEG2_CAPA"]])
    criar_pdf_paginas(seg2 / "1. anexo.pdf", [["SEG2_ANEXO"]])

    saida = juntar_tese(pasta_mae, tmp_path / "saida")

    assert _textos_das_paginas(saida) == ["SEG1_CAPA", "SEG2_CAPA", "SEG2_ANEXO"]


def test_juntar_tese_acrescenta_anexo_extra_como_ultimo_documento(tmp_path):
    pasta_mae = tmp_path / "tese completa"
    pasta_mae.mkdir()
    seg1 = pasta_mae / "1. FULANO DE TAL"
    seg1.mkdir()
    criar_pdf_paginas(seg1 / "0. capa.pdf", [["SEG1_CAPA"]])
    anexo_extra = tmp_path / "ACORDAO STJ.pdf"
    criar_pdf_paginas(anexo_extra, [["PAGINA_DO_ANEXO"]])

    saida = juntar_tese(pasta_mae, tmp_path / "saida", anexo_extra=anexo_extra)

    assert _textos_das_paginas(saida) == ["SEG1_CAPA", "PAGINA_DO_ANEXO"]
