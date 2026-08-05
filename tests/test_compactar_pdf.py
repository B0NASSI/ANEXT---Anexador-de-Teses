# -*- coding: utf-8 -*-
import fitz

from compactar_pdf import compress_lossless, compress_to_target
from helpers import criar_pdf_paginas


def test_compress_lossless_gera_pdf_valido_com_mesmas_paginas(tmp_path):
    origem = tmp_path / "origem.pdf"
    criar_pdf_paginas(origem, [["Página 1"], ["Página 2"]])
    destino = tmp_path / "destino.pdf"

    tamanho = compress_lossless(str(origem), str(destino))

    assert tamanho == destino.stat().st_size
    with fitz.open(destino) as doc:
        assert doc.page_count == 2


def test_compress_to_target_aceita_no_primeiro_nivel_quando_alvo_e_generoso(tmp_path):
    origem = tmp_path / "origem.pdf"
    criar_pdf_paginas(origem, [["Página única"]])
    destino = tmp_path / "destino.pdf"

    tamanho, nivel, coube = compress_to_target(str(origem), str(destino), max_bytes=10_000_000)

    assert coube is True
    assert nivel == "qualidade alta — 200 dpi"  # 1º nível tentado, já dentro do alvo


def test_compress_to_target_avisa_quando_nenhum_nivel_atinge_o_alvo(tmp_path):
    origem = tmp_path / "origem.pdf"
    criar_pdf_paginas(origem, [["Página única"]])
    destino = tmp_path / "destino.pdf"

    tamanho, nivel, coube = compress_to_target(str(origem), str(destino), max_bytes=1)

    assert coube is False
    assert nivel == "máxima compressão — 54 dpi"  # tentou todos os níveis até o último
