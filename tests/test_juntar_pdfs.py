# -*- coding: utf-8 -*-
import fitz

from helpers import criar_pdf_paginas
from juntar_pdfs import juntar_pasta, juntar_tese, ultimo_documento


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


def test_ultimo_documento_acha_o_ultimo_arquivo_na_ordem_de_juncao(tmp_path):
    pasta_mae = tmp_path / "tese completa"
    pasta_mae.mkdir()
    seg1 = pasta_mae / "1. FULANO DE TAL"
    seg2 = pasta_mae / "2. CICLANO DA SILVA"
    seg1.mkdir()
    seg2.mkdir()
    criar_pdf_paginas(seg1 / "0. capa.pdf", [["SEG1_CAPA"]])
    criar_pdf_paginas(seg2 / "0. capa.pdf", [["SEG2_CAPA"]])
    criar_pdf_paginas(seg2 / "1. anexo.pdf", [["SEG2_ANEXO"]])

    assert ultimo_documento(pasta_mae).name == "1. anexo.pdf"


def test_ultimo_documento_devolve_none_sem_nenhum_pdf(tmp_path):
    pasta_mae = tmp_path / "tese vazia"
    pasta_mae.mkdir()
    (pasta_mae / "1. FULANO DE TAL").mkdir()

    assert ultimo_documento(pasta_mae) is None


def test_log_registra_ordem_real_e_paginas_de_cada_documento(tmp_path, caplog):
    import hashlib
    import logging

    pasta_mae = tmp_path / "tese completa"
    seg1 = pasta_mae / "1. FULANO DE TAL"
    seg10 = pasta_mae / "10. BELTRANO"  # 10 depois de 2, não depois de 1 (ordem numérica)
    seg2 = pasta_mae / "2. CICLANO DA SILVA"
    for pasta in (seg1, seg10, seg2):
        pasta.mkdir(parents=True)
    criar_pdf_paginas(seg1 / "0. capa.pdf", [["A"]])
    criar_pdf_paginas(seg1 / "1. cnis.pdf", [["B"], ["C"]])
    criar_pdf_paginas(seg2 / "0. capa.pdf", [["D"]])
    criar_pdf_paginas(seg10 / "0. capa.pdf", [["E"]])
    anexo = tmp_path / "ACORDAO STJ.pdf"
    criar_pdf_paginas(anexo, [["F"]])

    with caplog.at_level(logging.INFO, logger="juntar_pdfs"):
        saida = juntar_tese(pasta_mae, tmp_path / "saida", anexo_extra=anexo)

    linhas = [r.getMessage() for r in caplog.records if r.getMessage().lstrip().startswith("[")]
    assert [linha.split(" → ")[0].strip() for linha in linhas] == [
        "[1/5] 1. FULANO DE TAL/0. capa.pdf",
        "[2/5] 1. FULANO DE TAL/1. cnis.pdf",
        "[3/5] 2. CICLANO DA SILVA/0. capa.pdf",
        "[4/5] 10. BELTRANO/0. capa.pdf",
        "[5/5] (anexo fixo) ACORDAO STJ.pdf",
    ]
    assert "páginas 2-3 (2 pág.)" in linhas[1] and "páginas 6-6" in linhas[4]
    sha_cnis = hashlib.sha256((seg1 / "1. cnis.pdf").read_bytes()).hexdigest()
    assert sha_cnis in linhas[1]
    # o que o log diz bate com o PDF gerado de verdade
    assert _textos_das_paginas(saida) == ["A", "B", "C", "D", "E", "F"]
    assert any("PDF juntado salvo" in r.getMessage() and "6 páginas" in r.getMessage() for r in caplog.records)
